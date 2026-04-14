require('dotenv').config();
const axios = require('axios');
const fs = require('fs');
const path = require('path');

// Extract variables from the .env file
const {
    TENANT_ID,
    CLIENT_ID,
    CLIENT_SECRET
} = process.env;

const AUTH_URL = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
const API_BASE_URL = `https://manage.office.com/api/v1.0/${TENANT_ID}/activity/feed/subscriptions/content`;

async function fetchAuditData() {
    try {
        console.log("1. Authenticating with Microsoft Entra ID...");
        const params = new URLSearchParams();
        params.append('client_id', CLIENT_ID);
        params.append('client_secret', CLIENT_SECRET);
        params.append('scope', 'https://manage.office.com/.default');
        params.append('grant_type', 'client_credentials');

        // Fetch Access Token
        const authResponse = await axios.post(AUTH_URL, params, {
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
        });
        
        const accessToken = authResponse.data.access_token;
        console.log("Authentication successful! Token acquired.");
        try {
            const base64Url = accessToken.split('.')[1];
            const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
            const jsonPayload = Buffer.from(base64, 'base64').toString('utf8');
            console.log("Token Roles:", JSON.parse(jsonPayload).roles);
        } catch(e) {}


        // Define the time window for the logs (Last 24 hours)
        const endTime = new Date();
        const startTime = new Date(endTime.getTime() - (24 * 60 * 60 * 1000));
        
        console.log("2. Fetching Content URIs from the Management Activity API...");
        const feedUrl = `${API_BASE_URL}?contentType=Audit.SharePoint&startTime=${startTime.toISOString()}&endTime=${endTime.toISOString()}`;

        // Get the list of blob URIs
        const contentResponse = await axios.get(feedUrl, {
            headers: { 'Authorization': `Bearer ${accessToken}` }
        });

        const contentUris = contentResponse.data;
        if (!contentUris || contentUris.length === 0) {
            console.log("No audit logs found for the specified time window.");
            return;
        }

        console.log(`Found ${contentUris.length} audit data blobs. Downloading events...`);
        
        // Fetch the actual event data from each blob
        let allEvents = [];
        for (const blob of contentUris) {
            const blobResponse = await axios.get(blob.contentUri, {
                headers: { 'Authorization': `Bearer ${accessToken}` }
            });
            allEvents = allEvents.concat(blobResponse.data);
        }

        console.log(`Successfully fetched ${allEvents.length} events!`);

        // Save events to a JSON file for Power BI to pick up
        const outputPath = path.join(__dirname, 'm365_audit_logs.json');
        fs.writeFileSync(outputPath, JSON.stringify(allEvents, null, 2));
        
        console.log(`Data saved to: ${outputPath} - You can now map Power BI to this file.`);
        
    } catch (error) {
        console.error("Error fetching audit logs:");
        if (error.response) {
            console.error(error.response.data);
        } else {
            console.error(error.message);
        }
    }
}

fetchAuditData();
