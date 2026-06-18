require('dotenv').config();
const axios = require('axios');

const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;

const AUTH_URL = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
const API_START_URL_AAD = `https://manage.office.com/api/v1.0/${TENANT_ID}/activity/feed/subscriptions/start?contentType=Audit.AzureActiveDirectory`;
const API_START_URL_SP = `https://manage.office.com/api/v1.0/${TENANT_ID}/activity/feed/subscriptions/start?contentType=Audit.SharePoint`;

async function startSubscriptions() {
    try {
        console.log("Authenticating...");
        const params = new URLSearchParams();
        params.append('client_id', CLIENT_ID);
        params.append('client_secret', CLIENT_SECRET);
        params.append('scope', 'https://manage.office.com/.default');
        params.append('grant_type', 'client_credentials');

        const authResponse = await axios.post(AUTH_URL, params, {
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
        });
        const accessToken = authResponse.data.access_token;
        console.log("Authenticated successfully.");

        console.log("Starting Audit.AzureActiveDirectory subscription...");
        try {
            const resAad = await axios.post(API_START_URL_AAD, null, {
                headers: { 'Authorization': `Bearer ${accessToken}` }
            });
            console.log("AAD Subscription Status:", resAad.status, resAad.data);
        } catch (e) {
            console.error("AAD Error:", e.response?.data || e.message);
        }

        console.log("Starting Audit.SharePoint subscription just in case...");
        try {
            const resSp = await axios.post(API_START_URL_SP, null, {
                headers: { 'Authorization': `Bearer ${accessToken}` }
            });
            console.log("SP Subscription Status:", resSp.status, resSp.data);
        } catch (e) {
            console.error("SP Error:", e.response?.data || e.message);
        }

    } catch (error) {
        console.error("Authentication Error:", error.message);
    }
}

startSubscriptions();
