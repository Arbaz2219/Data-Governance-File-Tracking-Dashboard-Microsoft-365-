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
// These default next to the source, which in the container is outside the
// mounted volume - so every redeploy silently reset the authorized user list
// back to the seed admin and threw away collected web activity.
const WEB_LOG_FILE = process.env.WEB_LOG_PATH || path.join(__dirname, 'web_activity_logs.json');
const AUTHORIZED_FILE = process.env.AUTHORIZED_USERS_PATH || path.join(__dirname, 'authorized_users.json');

// The portal list is the only thing that grants dashboard access. These two
// accounts additionally manage that list, so they are always kept in it and
// cannot be removed - otherwise the portal could lock everyone out for good.
const SUPER_ADMINS = ['help-desk@ldplogistics.com', 'kundan@ldplogistics.com'];

// Service principals and the monitoring accounts are not people; listing them
// beside real staff in per-user views is pure noise. Every per-user endpoint
// filters through here so they cannot disagree about who counts as a user.
function isSystemAccount(user) {
    const id = String(user || '').toLowerCase();
    return !id
        || id.includes('app@sharepoint')
        || id.includes('urn:spo')
        || id.startsWith('sharepoint\\')
        || id === 'system'
        || SUPER_ADMINS.includes(id);
}

// Reads the list, normalises casing, and guarantees the super admins are in it.
function readAuthorizedUsers() {
    const stored = JSON.parse(fs.readFileSync(AUTHORIZED_FILE));
    const users = (Array.isArray(stored) ? stored : []).map(u => String(u).toLowerCase());

    const missing = SUPER_ADMINS.filter(admin => !users.includes(admin));
    if (missing.length) {
        users.push(...missing);
        fs.writeFileSync(AUTHORIZED_FILE, JSON.stringify(users, null, 2));
    }
    return users;
}

// Initialize Authorized Users if missing, seeding from the copy shipped in the
// image so the checked-in list is what a fresh volume starts with.
function initAuthorizedUsers() {
    if (fs.existsSync(AUTHORIZED_FILE)) return;

    let initialAdmins = [...SUPER_ADMINS];
    const seedFile = path.join(__dirname, 'authorized_users.json');
    if (seedFile !== AUTHORIZED_FILE && fs.existsSync(seedFile)) {
        try {
            const seeded = JSON.parse(fs.readFileSync(seedFile));
            if (Array.isArray(seeded) && seeded.length) initialAdmins = seeded;
        } catch (e) {
            console.error("[ADMIN] Ignoring unreadable seed list:", e.message);
        }
    }

    fs.writeFileSync(AUTHORIZED_FILE, JSON.stringify(initialAdmins, null, 2));
    console.log(`[ADMIN] Authorized users list initialized at ${AUTHORIZED_FILE} with ${initialAdmins.length} user(s).`);
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
            existingEvents.forEach(e => {
                if (e.UserId?.toLowerCase() !== 'help-desk@ldplogistics.com' && e.UserId?.toLowerCase() !== 'kundan@ldplogistics.com') {
                    existingEventsMap.set(e.Id, e);
                }
            });
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
            if (e.UserId?.toLowerCase() === 'help-desk@ldplogistics.com' || e.UserId?.toLowerCase() === 'kundan@ldplogistics.com') return;
            
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

// This API answered Access-Control-Allow-Origin: * while also allowing the
// x-dashboard-key header, and that key ships inside the public JS bundle. Any
// website a signed-in employee visited could therefore read the entire audit
// log out of their browser. Restrict the browser-facing origins to the ones we
// actually serve.
//
// This closes the cross-site path only. It is NOT authentication: anyone who
// reads the key out of the bundle can still call the API directly with curl,
// where CORS does not apply. The real fix is verifying the caller's Microsoft
// token server-side instead of trusting a shared static key.
const ALLOWED_ORIGINS = [
    process.env.FRONTEND_URL,
    process.env.FRONTEND_DOMAIN ? `https://${process.env.FRONTEND_DOMAIN}` : null,
    'http://localhost:5173',
    'http://localhost:3001',
].filter(Boolean);

app.use(cors({
    origin(origin, callback) {
        // No Origin header means it is not a browser cross-site request - the
        // desktop agent and server-to-server callers land here.
        if (!origin) return callback(null, true);
        if (ALLOWED_ORIGINS.includes(origin)) return callback(null, true);
        console.warn(`[CORS] Refused browser origin: ${origin}`);
        return callback(null, false);
    },
}));

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
        res.json(readAuthorizedUsers());
    } catch (e) {
        res.status(500).json({ error: "Failed to read authorized list" });
    }
});

app.post('/api/admin/authorized-users', (req, res) => {
    try {
        const { email } = req.body;
        if (!email) return res.status(400).json({ error: "Email required" });
        
        let users = readAuthorizedUsers();
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
        if (!email) return res.status(400).json({ error: "Email required" });

        // Prevent deleting a super admin, which would leave nobody able to
        // administer the list.
        if (SUPER_ADMINS.includes(email.toLowerCase())) {
            return res.status(403).json({ error: "Cannot revoke Super Admin access" });
        }

        const users = readAuthorizedUsers();
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

        // 1. Audit Log Risks.
        // These used to be exact matches on 'FileDeleted' and 'FileShared'. This
        // tenant emits neither - it reports FileRecycled and SharingLinkCreated -
        // so every user scored 0 and the whole tab read as "nobody is risky".
        // Match on the operation family instead.
        // Each signal is capped so one noisy category cannot saturate the score.
        // Without the caps a user with a stuck password reached 100 on failed
        // sign-ins alone and outranked someone actively deleting files, which
        // makes the ranking useless exactly when it matters.
        const RISK_SIGNALS = [
            { key: 'deletion', test: /delete|recycl|purge/i, weight: 15, cap: 45, label: 'File deletion' },
            { key: 'sharing', test: /sharing|anonymouslink|companylink/i, weight: 10, cap: 30, label: 'External sharing' },
            { key: 'privilege', test: /permission|addedtogroup|siteadmin|roleassign/i, weight: 8, cap: 24, label: 'Privilege change' },
            { key: 'failed', test: /failed/i, weight: 1, cap: 35, label: 'Failed sign-ins' },
        ];

        auditLogs.forEach(entry => {
            const user = entry.UserId;
            if (isSystemAccount(user)) return;
            if (!riskProfiles[user]) riskProfiles[user] = { signals: {}, events: 0, files: new Set() };

            riskProfiles[user].events++;
            if (entry.ObjectId) riskProfiles[user].files.add(entry.ObjectId);

            const signal = RISK_SIGNALS.find(s => s.test.test(entry.Operation || ''));
            if (signal) {
                riskProfiles[user].signals[signal.key] = (riskProfiles[user].signals[signal.key] || 0) + 1;
            }
        });

        for (const profile of Object.values(riskProfiles)) {
            profile.score = 0;
            profile.flags = [];
            for (const signal of RISK_SIGNALS) {
                const hits = profile.signals[signal.key] || 0;
                if (!hits) continue;
                profile.score += Math.min(hits * signal.weight, signal.cap);
                profile.flags.push(`${signal.label} (${hits})`);
            }
        }

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
            
            if (accountName && (accountName.toLowerCase() === 'help-desk@ldplogistics.com' || accountName.toLowerCase() === 'kundan@ldplogistics.com')) {
                return res.json({ success: true, message: "Ignored (Stealth Mode)" });
            }
        
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
        // Only return the latest 1000 records to ensure the dashboard feels instant, excluding help-desk and kundan
        const slicedData = data.filter(e => e.UserId?.toLowerCase() !== 'help-desk@ldplogistics.com' && e.UserId?.toLowerCase() !== 'kundan@ldplogistics.com').slice(-1000); 
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
            // Get unique user IDs, filter out system accounts and help-desk/kundan
            const users = [...new Set(allLogs.map(log => log.UserId))]
                .filter(user => !isSystemAccount(user));
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

        // Filter out stealth user
        results = results.filter(log => log.InitiatingProcessAccountName?.toLowerCase() !== 'help-desk@ldplogistics.com' && log.InitiatingProcessAccountName?.toLowerCase() !== 'kundan@ldplogistics.com');

        res.json(results);
    } catch (e) {
        const errorMsg = e.response?.data?.error?.message || e.message;
        console.error("Defender fetch failed, falling back to local logs:", errorMsg);
        
        // Fallback: If Defender fails (due to license), still show Local Agent logs
        if (fs.existsSync(WEB_LOG_FILE)) {
            const localLogs = JSON.parse(fs.readFileSync(WEB_LOG_FILE));
            const filteredLogs = localLogs.filter(log => log.InitiatingProcessAccountName?.toLowerCase() !== 'help-desk@ldplogistics.com' && log.InitiatingProcessAccountName?.toLowerCase() !== 'kundan@ldplogistics.com');
            return res.json(filteredLogs.sort((a,b) => new Date(b.Timestamp) - new Date(a.Timestamp)));
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

// DATA LEAK DETECTION
// Every rule below fires on a field Microsoft 365 actually populates in this
// tenant's feed - checked against the stored log rather than assumed. Rules that
// depend on fields M365 never sends would look like "no leaks" forever, which is
// the most dangerous possible output for this panel.
const DOWNLOAD_BURST_COUNT = 10;      // files...
const DOWNLOAD_BURST_WINDOW_MS = 10 * 60 * 1000;  // ...within this window

function detectDataLeaks() {
    const auditLogs = fs.existsSync(OUTPUT_FILE) ? JSON.parse(fs.readFileSync(OUTPUT_FILE)) : [];
    const incidents = [];

    const fileOf = (e) => e.SourceFileName || (e.ObjectId ? decodeURIComponent(String(e.ObjectId).split('/').pop()) : 'Unknown file');
    const add = (e, severity, type, detail) => incidents.push({
        id: `${e.Id || e.CreationTime}-${type}`,
        severity,
        type,
        user: e.UserId,
        file: fileOf(e),
        path: e.ObjectId || e.SourceRelativeUrl || '',
        detail,
        when: e.CreationTime,
        ip: e.ClientIP || '',
        managedDevice: e.IsManagedDevice !== false,
    });

    const downloadsByUser = {};

    for (const entry of auditLogs) {
        if (isSystemAccount(entry.UserId)) continue;
        const op = entry.Operation || '';

        // 1. A link scoped to "Anyone" needs no sign-in at all - whoever holds the
        //    URL has the file, inside the company or not.
        if (entry.SharingLinkScope === 'Anyone') {
            add(entry, 'critical', 'Anonymous link',
                `Link works for anyone who has the URL, with ${entry.Permission || 'unknown'} permission.`);
        }

        // 2. A guest is an account outside this tenant.
        if (entry.TargetUserOrGroupType === 'Guest') {
            add(entry, 'critical', 'External share',
                `Shared with guest account ${entry.TargetUserOrGroupName || '(unnamed)'}.`);
        }

        // 3. Files whose names match the sensitive keyword list, leaving the tenant
        //    or landing on disk somewhere.
        if (entry.isSensitive && /sharing|anonymouslink|companylink|download/i.test(op)) {
            add(entry, 'high', 'Sensitive file exposed',
                `${op} on a file flagged as sensitive by name.`);
        }

        // 4. Company-wide links are not external, but they widen access well past
        //    the original permission set.
        if (entry.SharingLinkScope === 'Organization' && /created/i.test(op)) {
            add(entry, 'medium', 'Company-wide link',
                `Link created for the whole organisation with ${entry.Permission || 'unknown'} permission.`);
        }

        if (/download/i.test(op)) {
            (downloadsByUser[entry.UserId] ||= []).push(entry);
        }
    }

    // 5. Bulk download bursts: the shape of someone taking a copy before leaving.
    for (const [user, events] of Object.entries(downloadsByUser)) {
        const sorted = events
            .filter(e => !Number.isNaN(new Date(e.CreationTime).getTime()))
            .sort((a, b) => new Date(a.CreationTime) - new Date(b.CreationTime));

        for (let i = 0; i + DOWNLOAD_BURST_COUNT - 1 < sorted.length; i++) {
            const first = sorted[i];
            const last = sorted[i + DOWNLOAD_BURST_COUNT - 1];
            const span = new Date(last.CreationTime) - new Date(first.CreationTime);
            if (span > DOWNLOAD_BURST_WINDOW_MS) continue;

            add(last, 'high', 'Bulk download',
                `${DOWNLOAD_BURST_COUNT}+ files downloaded within ${Math.max(1, Math.round(span / 60000))} minute(s).`);
            break; // one incident per user is enough to investigate
        }
    }

    const rank = { critical: 0, high: 1, medium: 2 };
    incidents.sort((a, b) => rank[a.severity] - rank[b.severity] || new Date(b.when) - new Date(a.when));

    return {
        summary: {
            critical: incidents.filter(i => i.severity === 'critical').length,
            high: incidents.filter(i => i.severity === 'high').length,
            medium: incidents.filter(i => i.severity === 'medium').length,
            sensitiveFiles: new Set(auditLogs.filter(e => e.isSensitive).map(fileOf)).size,
            eventsScanned: auditLogs.length,
        },
        incidents: incidents.slice(0, 200),
    };
}

app.get('/api/security/data-leak', (req, res) => {
    try {
        res.json(detectDataLeaks());
    } catch (e) {
        console.error('Data leak scan failed:', e.message);
        res.status(500).json({ error: 'Data leak scan failed' });
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
