require('dotenv').config();
const axios = require('axios');

const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
const AUTH_URL = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;

async function testIntune() {
    try {
        console.log("Testing Intune API Connection...");
        
        // 1. Get Token
        const params = new URLSearchParams();
        params.append('client_id', CLIENT_ID);
        params.append('client_secret', CLIENT_SECRET);
        params.append('scope', 'https://graph.microsoft.com/.default');
        params.append('grant_type', 'client_credentials');

        const authResponse = await axios.post(AUTH_URL, params);
        const token = authResponse.data.access_token;
        console.log("✅ Token Acquired");

        // 2. Fetch Managed Devices (Limited to 1 for test)
        const res = await axios.get("https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?$top=1", {
            headers: { 'Authorization': `Bearer ${token}` }
        });
        
        console.log("✅ Intune API Response received!");
        console.log("Found " + (res.data.value?.length || 0) + " devices in initial test scan.");
        
        if (res.data.value && res.data.value.length > 0) {
            console.log("Sample Device: " + res.data.value[0].deviceName);
        } else {
            console.log("No devices found (this is normal if no devices are enrolled in Intune yet).");
        }

    } catch (e) {
        console.error("❌ Diagnostic Failed!");
        console.error(e.response?.data?.error || e.message);
    }
}

testIntune();
