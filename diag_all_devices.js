require('dotenv').config();
const axios = require('axios');

const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
const AUTH_URL = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;

async function testAllDevices() {
    try {
        console.log("Testing GLOBAL Device List API...");
        
        const params = new URLSearchParams();
        params.append('client_id', CLIENT_ID);
        params.append('client_secret', CLIENT_SECRET);
        params.append('scope', 'https://graph.microsoft.com/.default');
        params.append('grant_type', 'client_credentials');

        const authResponse = await axios.post(AUTH_URL, params);
        const token = authResponse.data.access_token;

        const res = await axios.get("https://graph.microsoft.com/v1.0/deviceManagement/managedDevices", {
            headers: { 'Authorization': `Bearer ${token}` }
        });
        
        console.log("✅ Success! Found " + (res.data.value?.length || 0) + " total devices.");
        if(res.data.value && res.data.value.length > 0) {
            console.log("Device 1: " + res.data.value[0].deviceName + " Owner: " + res.data.value[0].userPrincipalName);
        }
    } catch (e) {
        console.error("❌ Global Diagnostic Failed!");
        console.error(e.response?.data?.error || e.message);
    }
}

testAllDevices();
