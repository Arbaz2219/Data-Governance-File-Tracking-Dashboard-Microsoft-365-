require('dotenv').config();
const axios = require('axios');

const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;

const AUTH_URL = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
const START_SUB_URL = `https://manage.office.com/api/v1.0/${TENANT_ID}/activity/feed/subscriptions/start?contentType=Audit.SharePoint`;

async function startSubscription() {
    try {
        console.log("1. Authenticating...");
        const params = new URLSearchParams();
        params.append('client_id', CLIENT_ID);
        params.append('client_secret', CLIENT_SECRET);
        params.append('scope', 'https://manage.office.com/.default');
        params.append('grant_type', 'client_credentials');

        const authResponse = await axios.post(AUTH_URL, params, {
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
        });
        const accessToken = authResponse.data.access_token;
        
        console.log("2. Starting the SharePoint Audit Subscription...");
        const startResponse = await axios.post(START_SUB_URL, {}, {
            headers: { 'Authorization': `Bearer ${accessToken}` }
        });

        console.log("Success! Subscription started:", startResponse.data);
    } catch(err) {
        if(err.response) {
            console.error(err.response.data);
        } else {
            console.error(err.message);
        }
    }
}

startSubscription();
