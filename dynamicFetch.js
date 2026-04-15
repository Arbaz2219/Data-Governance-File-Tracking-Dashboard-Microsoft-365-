require('dotenv').config();
const axios = require('axios');
const fs = require('fs');
const path = require('path');

const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;

const AUTH_URL = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
const API_BASE_URL = `https://manage.office.com/api/v1.0/${TENANT_ID}/activity/feed/subscriptions/content`;
const OUTPUT_FILE = process.env.AUDIT_LOG_PATH || path.join(__dirname, 'm365_audit_logs.json');

// Run every 15 seconds for near-instant detection
const POLL_INTERVAL_MS = 15 * 1000; 
// Fetch window: last 6 hours to handle any M365 delays securely, while keeping payload small
const FETCH_WINDOW_HOURS = 6; 

// EMAIL ALERT SYSTEM 
const ALERT_EMAIL = 'help-desk@ldplogistics.com';

async function sendEmailAlert(event) {
    try {
        console.log(`[ALERT] Triggering Deletion Email for ${event.ObjectId}`);
        
        // 1. Get a specific token for Microsoft Graph
        const params = new URLSearchParams();
        params.append('client_id', CLIENT_ID);
        params.append('client_secret', CLIENT_SECRET);
        params.append('scope', 'https://graph.microsoft.com/.default');
        params.append('grant_type', 'client_credentials');

        const authResponse = await axios.post(AUTH_URL, params, {
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
        });
        const graphToken = authResponse.data.access_token;
        
        // Extract filename
        let filename = event.ObjectId.split('/').pop();
        try { filename = decodeURIComponent(filename); } catch(e){}

        // 2. Send the Email via Graph API
        const mailPayload = {
            message: {
                subject: `🚨 SECURITY ALERT: File Deleted by ${event.UserId?.split('@')[0]}`,
                body: {
                    contentType: "HTML",
                    content: `
                        <h2 style="color:red;">File Deletion Detected</h2>
                        <table style="border: 1px solid black; padding: 10px;">
                            <tr><td><b>Time:</b></td><td>${new Date(event.CreationTime).toLocaleString('en-US', { timeZone: 'America/New_York'})} EST</td></tr>
                            <tr><td><b>User:</b></td><td>${event.UserId}</td></tr>
                            <tr><td><b>File Path:</b></td><td>${event.ObjectId}</td></tr>
                            <tr><td><b>Extracted File:</b></td><td>${filename}</td></tr>
                        </table>
                    `
                },
                toRecipients: [{ emailAddress: { address: ALERT_EMAIL } }]
            },
            saveToSentItems: "false"
        };

        // Note: we use users/help-desk@ldplogistics.com to send *from* the help-desk mailbox itself,
        // or any valid member mailbox. Make sure Mail.Send is granted!
        await axios.post(`https://graph.microsoft.com/v1.0/users/${ALERT_EMAIL}/sendMail`, mailPayload, {
            headers: { 'Authorization': `Bearer ${graphToken}`, 'Content-Type': 'application/json' }
        });
        
        console.log(`[ALERT] Email successfully sent to ${ALERT_EMAIL}!`);
    } catch (e) {
        console.error("[ALERT ERROR] Failed to send email alert. Did you add 'Mail.Send' Application Permission in Entra ID?", e.response?.data?.error || e.message);
    }
}

async function fetchAuditData() {
    try {
        console.log(`[${new Date().toISOString()}] Waking up to check M365 Logs...`);
        
        let existingEventsMap = new Map();
        if (fs.existsSync(OUTPUT_FILE)) {
            const rawData = fs.readFileSync(OUTPUT_FILE);
            const existingEvents = JSON.parse(rawData);
            // Deduplicate by the unique ID Microsoft assigns to each Audit event
            existingEvents.forEach(e => existingEventsMap.set(e.Id, e));
        }
        
        // 1. Authenticate
        const params = new URLSearchParams();
        params.append('client_id', CLIENT_ID);
        params.append('client_secret', CLIENT_SECRET);
        params.append('scope', 'https://manage.office.com/.default');
        params.append('grant_type', 'client_credentials');

        const authResponse = await axios.post(AUTH_URL, params, {
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
        });
        const accessToken = authResponse.data.access_token;

        // 2. Query Time Window
        const endTime = new Date();
        const startTime = new Date(endTime.getTime() - (FETCH_WINDOW_HOURS * 60 * 60 * 1000));
        
        const feedUrl = `${API_BASE_URL}?contentType=Audit.SharePoint&startTime=${startTime.toISOString()}&endTime=${endTime.toISOString()}`;

        const contentResponse = await axios.get(feedUrl, {
            headers: { 'Authorization': `Bearer ${accessToken}` }
        });

        const contentUris = contentResponse.data;
        if (!contentUris || contentUris.length === 0) {
            console.log("No new blob URIs discovered.");
            return;
        }
        
        // 3. Fetch blobs
        let fetchedEvents = [];
        for (const blob of contentUris) {
            const blobResponse = await axios.get(blob.contentUri, {
                headers: { 'Authorization': `Bearer ${accessToken}` }
            });
            fetchedEvents = fetchedEvents.concat(blobResponse.data);
        }

        // 4. Merge and Save
        let newEventsAdded = 0;
        let deletedEventsToAlert = [];
        
        fetchedEvents.forEach(e => {
            if (!existingEventsMap.has(e.Id)) {
                existingEventsMap.set(e.Id, e);
                newEventsAdded++;
                
                // If it's a deletion event!
                if (e.Operation === 'FileDeleted') {
                    deletedEventsToAlert.push(e);
                }
            }
        });

        // Convert Map back to array and ensure it is saved
        const updatedEventsList = Array.from(existingEventsMap.values());
        fs.writeFileSync(OUTPUT_FILE, JSON.stringify(updatedEventsList, null, 2));
        
        console.log(`[${new Date().toISOString()}] Synced! Added ${newEventsAdded} new events to database. Total events: ${updatedEventsList.length}`);
        
        // 5. Fire asynchronous alerts for deeply nested new events securely
        for (const deletedEvent of deletedEventsToAlert) {
            await sendEmailAlert(deletedEvent);
        }
        
    } catch (error) {
        console.error("Error during poll:");
        if (error.response) {
            console.error(error.response.data);
        } else {
            console.error(error.message);
        }
    }
}

// Perform instant first run, then start ticker
fetchAuditData().then(() => {
    console.log(`Monitoring Started! Checking for fresh data every ${POLL_INTERVAL_MS / 1000 / 60} minutes.`);
    setInterval(fetchAuditData, POLL_INTERVAL_MS);
});

// START EXPRESS API SERVER FOR DASHBOARD
const express = require('express');
const cors = require('cors');

const app = express();
app.use(cors());

// Expose the JSON file to frontend (Optimized Slicing for Speed)
app.get('/api/audit-logs', (req, res) => {
    if (fs.existsSync(OUTPUT_FILE)) {
        const rawData = fs.readFileSync(OUTPUT_FILE);
        const data = JSON.parse(rawData);
        // Only return the latest 1000 records to ensure the dashboard feels instant
        const slicedData = data.slice(-1000); 
        res.json(slicedData);
    } else {
        res.json([]);
    }
});

// Fetch Real-time OneDrive Storage and AD Profile details
app.get('/api/user/:email/profile', async (req, res) => {
    try {
        const userEmail = req.params.email;
        
        // 1. Get Graph Token
        const params = new URLSearchParams();
        params.append('client_id', CLIENT_ID);
        params.append('client_secret', CLIENT_SECRET);
        params.append('scope', 'https://graph.microsoft.com/.default');
        params.append('grant_type', 'client_credentials');

        const authResponse = await axios.post(AUTH_URL, params, {
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
        });
        const graphToken = authResponse.data.access_token;
        
        // 2. Fetch AD User Profile (Job Title, Dept, etc.)
        const profileUrl = `https://graph.microsoft.com/v1.0/users/${userEmail}?$select=displayName,jobTitle,department`;
        const profileResponse = await axios.get(profileUrl, {
            headers: { 'Authorization': `Bearer ${graphToken}` } 
        });

        // 3. Fetch OneDrive Storage
        const driveUrl = `https://graph.microsoft.com/v1.0/users/${userEmail}/drive`;
        const driveResponse = await axios.get(driveUrl, {
            headers: { 'Authorization': `Bearer ${graphToken}` } 
        });

        // 4. Calculate local stats from m365_audit_logs.json
        let stats = { totalEvents: 0, uniqueFileCount: 0 };
        if (fs.existsSync(OUTPUT_FILE)) {
            const rawData = fs.readFileSync(OUTPUT_FILE);
            const allLogs = JSON.parse(rawData);
            const userLogs = allLogs.filter(log => log.UserId?.toLowerCase() === userEmail.toLowerCase());
            
            stats.totalEvents = userLogs.length;
            const uniqueFiles = new Set(userLogs.map(log => log.ObjectId).filter(Boolean));
            stats.uniqueFileCount = uniqueFiles.size;
        }
        
        res.json({
            profile: profileResponse.data,
            storage: driveResponse.data.quota,
            stats: stats
        });
        
    } catch (e) {
        console.error("Profile fetch failed for " + req.params.email, e.response?.data || e.message);
        res.status(500).json({ error: "Failed to fetch profile details" });
    }
});

// NEW: List User's Managed Devices (Intune)
app.get('/api/user/:email/devices', async (req, res) => {
    try {
        const userEmail = req.params.email;
        
        // 1. Get Graph Token
        const params = new URLSearchParams();
        params.append('client_id', CLIENT_ID);
        params.append('client_secret', CLIENT_SECRET);
        params.append('scope', 'https://graph.microsoft.com/.default');
        params.append('grant_type', 'client_credentials');

        const authResponse = await axios.post(AUTH_URL, params, {
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
        });
        const graphToken = authResponse.data.access_token;

        // 2. Resolve Email to User Object ID
        const userRes = await axios.get(`https://graph.microsoft.com/v1.0/users/${userEmail}?$select=id`, {
            headers: { 'Authorization': `Bearer ${graphToken}` } 
        });
        const userId = userRes.data.id;

        // 3. Fetch Managed Devices for this User
        const devicesUrl = `https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?$filter=userId eq '${userId}'&$select=id,deviceName,model,operatingSystem,complianceState,managementAgent`;
        const devicesRes = await axios.get(devicesUrl, {
            headers: { 'Authorization': `Bearer ${graphToken}` } 
        });

        res.json(devicesRes.data.value || []);
    } catch (e) {
        console.error("Device fetch failed:", e.response?.data || e.message);
        res.status(500).json({ error: "Intune Sync Pending or No Managed Devices Found" });
    }
});

// NEW: List ALL Corporate Managed Devices (Global View)
app.get('/api/devices/all', async (req, res) => {
    try {
        // 1. Get Graph Token
        const params = new URLSearchParams();
        params.append('client_id', CLIENT_ID);
        params.append('client_secret', CLIENT_SECRET);
        params.append('scope', 'https://graph.microsoft.com/.default');
        params.append('grant_type', 'client_credentials');

        const authResponse = await axios.post(AUTH_URL, params, {
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
        });
        const graphToken = authResponse.data.access_token;

        // 2. Fetch All Managed Devices
        const devicesUrl = `https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?$select=id,deviceName,userDisplayName,userPrincipalName,model,operatingSystem,complianceState,managementAgent,lastSyncDateTime`;
        const devicesRes = await axios.get(devicesUrl, {
            headers: { 'Authorization': `Bearer ${graphToken}` } 
        });

        res.json(devicesRes.data.value || []);
    } catch (e) {
        console.error("Global device fetch failed:", e.response?.data || e.message);
        res.status(500).json({ error: "Failed to fetch corporate fleet data" });
    }
});

// NEW: Fetch Browser / Web Search History (Microsoft Defender)
app.get('/api/security/web-activity', async (req, res) => {
    try {
        // 1. Get Graph Token
        const params = new URLSearchParams();
        params.append('client_id', CLIENT_ID);
        params.append('client_secret', CLIENT_SECRET);
        params.append('scope', 'https://graph.microsoft.com/.default');
        params.append('grant_type', 'client_credentials');

        const authResponse = await axios.post(AUTH_URL, params, {
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
        });
        const graphToken = authResponse.data.access_token;

        // 2. Query Defender Advanced Hunting
        // We look for HttpConnection events which typically include browser search patterns
        const huntingQuery = { 
            Query: `DeviceNetworkEvents 
                    | where ActionType == 'HttpConnection' 
                    | where InitiatingProcessAccountName !in~ ('help-desk', 'kundan')
                    | project Timestamp, DeviceName, LocalIP, RemoteUrl, RemoteIP, InitiatingProcessAccountName 
                    | order by Timestamp desc 
                    | take 50` 
        };

        const defenderRes = await axios.post('https://graph.microsoft.com/v1.0/security/runHuntingQuery', huntingQuery, {
            headers: { 
                'Authorization': `Bearer ${graphToken}`,
                'Content-Type': 'application/json'
            } 
        });

        res.json(defenderRes.data.results || []);
    } catch (e) {
        console.error("Defender fetch failed:", e.response?.data || e.message);
        res.status(500).json({ 
            error: "Defender Sync Pending", 
            details: e.response?.data?.error?.message || "Verify Advanced Hunting permissions"
        });
    }
});

// NEW: Trigger Remote Reboot
app.post('/api/device/:deviceId/reboot', async (req, res) => {
    try {
        const deviceId = req.params.deviceId;
        
        // 1. Get Graph Token
        const params = new URLSearchParams();
        params.append('client_id', CLIENT_ID);
        params.append('client_secret', CLIENT_SECRET);
        params.append('scope', 'https://graph.microsoft.com/.default');
        params.append('grant_type', 'client_credentials');

        const authResponse = await axios.post(AUTH_URL, params, {
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
        });
        const graphToken = authResponse.data.access_token;

        // 2. Trigger Reboot Command
        const rebootUrl = `https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/${deviceId}/rebootNow`;
        await axios.post(rebootUrl, {}, {
            headers: { 'Authorization': `Bearer ${graphToken}` } 
        });

        res.json({ success: true, message: "Reboot command dispatched successfully" });
    } catch (e) {
        console.error("Reboot command failed:", e.response?.data || e.message);
        res.status(500).json({ error: "Failed to dispatch reboot command. Check Intune permissions." });
    }
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
    console.log(`Dashboard Backend API running on http://localhost:${PORT}/api/audit-logs`);
});
