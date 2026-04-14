require('dotenv').config();
const axios = require('axios');

const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;

async function checkTokenRoles() {
    try {
        console.log("Analyzing App Token Permissions...");
        
        const params = new URLSearchParams();
        params.append('client_id', CLIENT_ID);
        params.append('client_secret', CLIENT_SECRET);
        params.append('scope', 'https://graph.microsoft.com/.default');
        params.append('grant_type', 'client_credentials');

        const authUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
        const authResponse = await axios.post(authUrl, params);
        const token = authResponse.data.access_token;

        const base64Url = token.split('.')[1];
        const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
        const jsonPayload = decodeURIComponent(atob(base64).split('').map(function(c) {
            return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
        }).join(''));

        const decoded = JSON.parse(jsonPayload);
        console.log("✅ Token Decoded!");
        console.log("Active Roles (Permissions):");
        console.log(JSON.stringify(decoded.roles, null, 2));

    } catch (e) {
        console.error("❌ Token Analysis Failed!");
        console.error(e.message);
    }
}

checkTokenRoles();
