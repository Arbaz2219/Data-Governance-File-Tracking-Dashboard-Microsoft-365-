require('dotenv').config();
const axios = require('axios');

const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
const AUTH_URL = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;

async function testDefender() {
    try {
        console.log("Testing Microsoft Defender Advanced Hunting API...");
        
        const params = new URLSearchParams();
        params.append('client_id', CLIENT_ID);
        params.append('client_secret', CLIENT_SECRET);
        params.append('scope', 'https://graph.microsoft.com/.default');
        params.append('grant_type', 'client_credentials');

        const authResponse = await axios.post(AUTH_URL, params);
        const token = authResponse.data.access_token;

        // Query for last 5 web-related events
        const query = "DeviceNetworkEvents | where ActionType == 'HttpConnection' | project Timestamp, DeviceName, RemoteUrl, InitiatingProcessAccountName | order by Timestamp desc | take 5";
        
        const res = await axios.post("https://graph.microsoft.com/v1.0/security/runHuntingQuery", { Query: query }, {
            headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' }
        });
        
        console.log("✅ Success! Defender API responded.");
        console.log("Found " + (res.data.results?.length || 0) + " web activity events.");
        if(res.data.results && res.data.results.length > 0) {
            console.log("Sample Activity: " + res.data.results[0].RemoteUrl);
        }
    } catch (e) {
        console.error("❌ Defender Diagnostic Failed!");
        console.error(e.response?.data?.error || e.message);
    }
}

testDefender();
