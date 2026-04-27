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

// PATHS FOR LOG STORAGE
const WEB_LOG_FILE = path.join(__dirname, 'web_activity_logs.json');
const AUTHORIZED_FILE = path.join(__dirname, 'authorized_users.json');

// Initialize Authorized Users if missing
function initAuthorizedUsers() {
    if (!fs.existsSync(AUTHORIZED_FILE)) {
        const initialAdmins = ['kundan@ldplogistics.com', 'help-desk@ldplogistics.com'];
        fs.writeFileSync(AUTHORIZED_FILE, JSON.stringify(initialAdmins, null, 2));
        console.log("[ADMIN] Authorized users list initialized.");
    }
}
initAuthorizedUsers();

// SENSITIVE CONTENT KEYWORDS
const SENSITIVE_KEYWORDS = ['salary', 'invoice', 'password', 'confidential', 'secret', 'contract', 'finance', 'payment', 'tax', 'bonus'];
function checkSensitivity(path) {
    if (!path) return false;
    return SENSITIVE_KEYWORDS.some(kw => path.toLowerCase().includes(kw));
}

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

// SELF-DESTRUCT SYSTEM: Purge logs older than 24 hours
function purgeOldLogs() {
    try {
        const cutoff = new Date(Date.now() - (24 * 60 * 60 * 1000)).getTime();
        
        // 1. Purge Audit Logs
        if (fs.existsSync(OUTPUT_FILE)) {
            const auditLogs = JSON.parse(fs.readFileSync(OUTPUT_FILE));
            const filteredAudit = auditLogs.filter(e => new Date(e.CreationTime).getTime() > cutoff);
            if (auditLogs.length !== filteredAudit.length) {
                fs.writeFileSync(OUTPUT_FILE, JSON.stringify(filteredAudit, null, 2));
                console.log(`[SELF-DESTRUCT] Purged ${auditLogs.length - filteredAudit.length} old Audit records.`);
            }
        }
        
        // 2. Purge Web History Agent Logs
        if (fs.existsSync(WEB_LOG_FILE)) {
            const webLogs = JSON.parse(fs.readFileSync(WEB_LOG_FILE));
            const filteredWeb = webLogs.filter(e => new Date(e.Timestamp).getTime() > cutoff);
            if (webLogs.length !== filteredWeb.length) {
                fs.writeFileSync(WEB_LOG_FILE, JSON.stringify(filteredWeb, null, 2));
                console.log(`[SELF-DESTRUCT] Purged ${webLogs.length - filteredWeb.length} old Web History records.`);
            }
        }
    } catch (e) {
        console.error("[SELF-DESTRUCT ERROR]", e.message);
    }
}

async function fetchAuditData() {
    try {
        console.log(`[${new Date().toISOString()}] Waking up to check M365 Logs...`);
        
        // Run the 24-hour purge before processing new data
        purgeOldLogs();

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
        
        const feedUrlSP = `${API_BASE_URL}?contentType=Audit.SharePoint&startTime=${startTime.toISOString()}&endTime=${endTime.toISOString()}`;
        const feedUrlAAD = `${API_BASE_URL}?contentType=Audit.AzureActiveDirectory&startTime=${startTime.toISOString()}&endTime=${endTime.toISOString()}`;

        let contentUris = [];
        try {
            const [spRes, aadRes] = await Promise.all([
                axios.get(feedUrlSP, { headers: { 'Authorization': `Bearer ${accessToken}` } }),
                axios.get(feedUrlAAD, { headers: { 'Authorization': `Bearer ${accessToken}` } })
            ]);
            if (spRes.data) contentUris = contentUris.concat(spRes.data);
            if (aadRes.data) contentUris = contentUris.concat(aadRes.data);
        } catch (err) {
            console.error("Error fetching blob URIs:", err.message);
        }

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
                // ADD SECURITY INTELLIGENCE
                e.isSensitive = checkSensitivity(e.ObjectId);
                
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
app.use(express.json()); // Enable JSON parsing for incoming desktop agent logs

// SECURITY MIDDLEWARE: Block direct endpoint access
const DASHBOARD_SECRET = process.env.DASHBOARD_SECRET_KEY || "LDP_SECURE_9821_!@#$";
app.use('/api', (req, res, next) => {
    // Allow the desktop agent to report activity without the dashboard key
    if (req.path === '/security/report-activity') return next();
    
    const key = req.headers['x-dashboard-key'];
    if (key === DASHBOARD_SECRET) {
        next();
    } else {
        console.warn(`Blocked unauthorized access attempt to: ${req.path} from IP: ${req.ip}`);
        res.status(403).json({ error: "Access Denied: Secure Terminal Only" });
    }
});

// ADMIN API: User Authorization Management
app.get('/api/admin/authorized-users', (req, res) => {
    try {
        const users = JSON.parse(fs.readFileSync(AUTHORIZED_FILE));
        res.json(users);
    } catch (e) {
        res.status(500).json({ error: "Failed to read authorized list" });
    }
});

app.post('/api/admin/authorized-users', (req, res) => {
    try {
        const { email } = req.body;
        if (!email) return res.status(400).json({ error: "Email required" });
        
        let users = JSON.parse(fs.readFileSync(AUTHORIZED_FILE));
        if (!users.includes(email.toLowerCase())) {
            users.push(email.toLowerCase());
            fs.writeFileSync(AUTHORIZED_FILE, JSON.stringify(users, null, 2));
        }
        res.json({ success: true, users });
    } catch (e) {
        res.status(500).json({ error: "Failed to update authorized list" });
    }
});

app.delete('/api/admin/authorized-users', (req, res) => {
    try {
        const { email } = req.body;
        let users = JSON.parse(fs.readFileSync(AUTHORIZED_FILE));
        
        // Prevent deleting the super admin
        if (email.toLowerCase() === 'kundan@ldplogistics.com') {
            return res.status(403).json({ error: "Cannot revoke Super Admin access" });
        }

        const filtered = users.filter(u => u !== email.toLowerCase());
        fs.writeFileSync(AUTHORIZED_FILE, JSON.stringify(filtered, null, 2));
        res.json({ success: true, users: filtered });
    } catch (e) {
        res.status(500).json({ error: "Failed to update authorized list" });
    }
});

// Serve Static Frontend Files from 'client/dist'
app.use(express.static(path.join(__dirname, 'client/dist'), { setHeaders: (res, path) => {
    res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
    res.setHeader('Pragma', 'no-cache');
    res.setHeader('Expires', '0');
    res.setHeader('Surrogate-Control', 'no-store');
}}));

// SHADOW IT CATEGORIZATION ENGINE
const SHADOW_APPS = [
    { name: 'Personal Cloud Storage', domains: ['dropbox.com', 'mega.nz', 'mediafire.com', 'wetransfer.com', 'pcloud.com'] },
    { name: 'Unauthorized AI / LLM', domains: ['chatgpt.com', 'openai.com', 'claude.ai', 'gemini.google.com', 'perplexity.ai'] },
    { name: 'Social / Video Sprawl', domains: ['tiktok.com', 'netflix.com', 'instagram.com', 'facebook.com', 'reddit.com'] },
    { name: 'Proxy / VPN circumvention', domains: ['nordvpn.com', 'expressvpn.com', 'protonvpn.com', 'hide.me'] }
];

function categorizeUrl(url) {
    if (!url) return 'Unknown';
    const match = SHADOW_APPS.find(app => app.domains.some(d => url.toLowerCase().includes(d)));
    return match ? match.name : 'Authorized / Standard';
}

// RISK SCORING ENGINE
function calculateUserRisk() {
    try {
        const auditLogs = fs.existsSync(OUTPUT_FILE) ? JSON.parse(fs.readFileSync(OUTPUT_FILE)) : [];
        const webLogs = fs.existsSync(WEB_LOG_FILE) ? JSON.parse(fs.readFileSync(WEB_LOG_FILE)) : [];
        
        const riskProfiles = {};

        // 1. Audit Log Risks (Deletions are high risk)
        auditLogs.forEach(entry => {
            const user = entry.UserId;
            if (!riskProfiles[user]) riskProfiles[user] = { score: 0, flags: [], events: 0, files: new Set() };
            
            riskProfiles[user].events++;
            if (entry.ObjectId) riskProfiles[user].files.add(entry.ObjectId);
            if (entry.Operation === 'FileDeleted') {
                riskProfiles[user].score += 15;
                riskProfiles[user].flags.push('File Deletion Detected');
            }
            if (entry.Operation === 'FileShared') {
                riskProfiles[user].score += 10;
                riskProfiles[user].flags.push('External Sharing Activity');
            }
        });

        // 2. Web History Risks (Shadow IT usage)
        webLogs.forEach(entry => {
            const user = entry.InitiatingProcessAccountName;
            if (!user || user === 'Unknown User') return;
            if (!riskProfiles[user]) riskProfiles[user] = { score: 0, flags: [], events: 0, files: new Set() };
            
            const category = categorizeUrl(entry.RemoteUrl);
            if (category !== 'Authorized / Standard') {
                riskProfiles[user].score += 5;
                if (!riskProfiles[user].flags.includes('Shadow IT Usage')) {
                    riskProfiles[user].flags.push(`Shadow IT (${category})`);
                }
            }
        });

        // Convert to sorted array
        return Object.entries(riskProfiles).map(([user, data]) => ({
            user,
            score: Math.min(data.score, 100), // Cap at 100
            level: data.score > 70 ? 'Critical' : data.score > 30 ? 'Moderate' : 'Low',
            flags: [...new Set(data.flags)].slice(0, 3), // Unique top 3 flags
            activityCount: data.events,
            fileCount: data.files.size
        })).sort((a,b) => b.score - a.score);

    } catch (e) {
        console.error("Risk calc failed:", e.message);
        return [];
    }
}

// NEW: Data Receiver for License-Free Desktop Agent
app.post('/api/security/report-activity', (req, res) => {
    try {
        const { deviceName, accountName, remoteUrl, searchTerm, timestamp } = req.body;
        
        let logs = [];
        if (fs.existsSync(WEB_LOG_FILE)) {
            logs = JSON.parse(fs.readFileSync(WEB_LOG_FILE));
        }
        
        // Add new log entry with categorization
        logs.push({
            Timestamp: timestamp || new Date().toISOString(),
            DeviceName: deviceName || 'Unknown Device',
            InitiatingProcessAccountName: accountName || 'Unknown User',
            RemoteUrl: remoteUrl || '',
            SearchTerm: searchTerm || '',
            Category: categorizeUrl(remoteUrl),
            Source: 'Desktop Agent'
        });
        
        // Keep only last 5000 logs to prevent bloat
        if (logs.length > 5000) logs = logs.slice(-5000);
        
        fs.writeFileSync(WEB_LOG_FILE, JSON.stringify(logs, null, 2));
        res.json({ success: true });
    } catch (e) {
        console.error("Failed to save local log:", e.message);
        res.status(500).json({ error: "Storage Error" });
    }
});

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

// NEW: List all Unique Users from Audit Logs
app.get('/api/users/list', (req, res) => {
    try {
        if (fs.existsSync(OUTPUT_FILE)) {
            const rawData = fs.readFileSync(OUTPUT_FILE);
            const allLogs = JSON.parse(rawData);
            // Get unique user IDs, filter out system accounts
            const users = [...new Set(allLogs.map(log => log.UserId))]
                .filter(Boolean)
                .filter(email => !email.includes('app@sharepoint') && !email.includes('urn:spo'));
            res.json(users);
        } else {
            res.json([]);
        }
    } catch (e) {
        res.status(500).json({ error: "Failed to list users" });
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
        // Query Defender Advanced Hunting for Browser Activity & Searches
        const huntingQuery = { 
            Query: `DeviceNetworkEvents 
                    | where InitiatingProcessFileName in~ ("msedge.exe", "chrome.exe", "firefox.exe")
                    | where isnotempty(RemoteUrl)
                    | project Timestamp, DeviceName, InitiatingProcessAccountName, RemoteUrl, ActionType = "Visit"
                    | union (
                        UrlClickEvents
                        | project Timestamp, DeviceName, InitiatingProcessAccountName = AccountUpn, RemoteUrl = Url, ActionType = "Click"
                    )
                    | extend SearchTerm = extract("q=([^&]+)", 1, RemoteUrl)
                    | extend SearchTerm = ifelse(isnotempty(SearchTerm), replace_string(SearchTerm, "+", " "), "")
                    | extend SearchTerm = ifelse(isnotempty(SearchTerm), url_decode(SearchTerm), "")
                    | order by Timestamp desc 
                    | take 100` 
        };

        const defenderRes = await axios.post('https://graph.microsoft.com/v1.0/security/runHuntingQuery', huntingQuery, {
            headers: { 
                'Authorization': `Bearer ${graphToken}`,
                'Content-Type': 'application/json'
            } 
        });

        let results = defenderRes.data.results || [];
        
        // Merge with Local Agent Logs if they exist
        if (fs.existsSync(WEB_LOG_FILE)) {
            const localLogs = JSON.parse(fs.readFileSync(WEB_LOG_FILE));
            results = [...localLogs, ...results].sort((a,b) => new Date(b.Timestamp) - new Date(a.Timestamp)).slice(0, 500);
        }

        res.json(results);
    } catch (e) {
        const errorMsg = e.response?.data?.error?.message || e.message;
        console.error("Defender fetch failed, falling back to local logs:", errorMsg);
        
        // Fallback: If Defender fails (due to license), still show Local Agent logs
        if (fs.existsSync(WEB_LOG_FILE)) {
            const localLogs = JSON.parse(fs.readFileSync(WEB_LOG_FILE));
            return res.json(localLogs.sort((a,b) => new Date(b.Timestamp) - new Date(a.Timestamp)));
        }

        // If no local logs and no defender, then show error
        if (errorMsg.includes("Failed to resolve table") || errorMsg.includes("DeviceNetworkEvents")) {
            return res.status(200).json({ 
                error: "License Restriction", 
                details: "Microsoft Defender Advanced Hunting is restricted by your license. Please deploy the 'Local Desktop Agent' to track browser history for free." 
            });
        }
        
        res.status(500).json({ error: "Sync Error", details: errorMsg });
    }
});

// NEW: Shadow IT Detection Feed
app.get('/api/security/shadow-it', (req, res) => {
    try {
        if (!fs.existsSync(WEB_LOG_FILE)) return res.json([]);
        const logs = JSON.parse(fs.readFileSync(WEB_LOG_FILE));
        
        // Filter for Shadow IT categories
        const shadowLogs = logs.filter(l => l.Category !== 'Authorized / Standard' && l.Category !== 'Unknown');
        res.json(shadowLogs.reverse().slice(0, 100));
    } catch (e) {
        res.status(500).json({ error: "Feed Logic Error" });
    }
});

// NEW: User Behavior Risk Scoring
app.get('/api/security/risk-stats', (req, res) => {
    try {
        const riskProfiles = calculateUserRisk();
        res.json(riskProfiles);
    } catch (e) {
        res.status(500).json({ error: "Risk Calc Error" });
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

// NEW: List Live Teams Calls (Communications Intelligence)
app.get('/api/security/active-calls', async (req, res) => {
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

        // 2. Query Live Calls
        // Note: Calls.Read.All Application permission is required for tenant-wide visibility
        const callsRes = await axios.get('https://graph.microsoft.com/v1.0/communications/calls', {
            headers: { 'Authorization': `Bearer ${graphToken}` } 
        });

        // Map calls to a clean format for the dashboard
        const activeCalls = (callsRes.data.value || []).map(call => ({
            id: call.id,
            state: call.state,
            subject: call.subject || "Confidential Intelligence Call",
            participants: call.participants?.map(p => p.info?.identity?.user?.displayName || "External Participant") || [],
            startTime: call.createdDateTime,
            direction: call.direction
        }));

        res.json(activeCalls);
    } catch (e) {
        const errorMsg = e.response?.data?.error?.message || e.message;
        console.error("Teams Call Intel failed:", errorMsg);
        
        // Handle common permission errors gracefully
        if (errorMsg.includes("Permissions")) {
             return res.status(200).json({ 
                error: "Permission Denied", 
                details: "Monitoring live calls requires 'Calls.Read.All' Application Permissions in Entra ID." 
            });
        }

        res.status(500).json({ error: "Comms Sync Error", details: errorMsg });
    }
});

// SPA Fallback: Serve index.html for any unknown routes
app.use((req, res) => {
    res.sendFile(path.join(__dirname, 'client/dist', 'index.html'));
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
    console.log(`Unified Dashboard running on http://localhost:${PORT}`);
});
