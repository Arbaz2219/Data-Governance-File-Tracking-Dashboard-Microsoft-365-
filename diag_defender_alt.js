require('dotenv').config();
const axios = require('axios');

const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;

async function testDefenderAlternative() {
    try {
        console.log("Testing Defender API with dedicated security scope...");
        
        // Use Defender-specific scope
        const params = new URLSearchParams();
        params.append('client_id', CLIENT_ID);
        params.append('client_secret', CLIENT_SECRET);
        params.append('scope', 'https://api.security.microsoft.com/.default');
        params.append('grant_type', 'client_credentials');

        const authUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
        const authResponse = await axios.post(authUrl, params);
        const token = authResponse.data.access_token;
        console.log("✅ Token Acquired (Security Scope)");

        const query = { Query: "DeviceNetworkEvents | take 1" };
        const res = await axios.post("https://api.security.microsoft.com/api/advancedqueries/run", query, {
            headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' }
        });
        
        console.log("✅ Success! Dedicate Defender API responded.");
        console.log("Results: " + JSON.stringify(res.data.results));

    } catch (e) {
        console.error("❌ Alternative Diagnostic Failed!");
        console.error(e.response?.data?.error || e.message);
        if(e.response?.data) console.log(JSON.stringify(e.response.data, null, 2));
    }
}

testDefenderAlternative();
