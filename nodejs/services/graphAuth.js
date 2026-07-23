const axios = require("axios");
require("dotenv").config();

async function getGraphToken() {
    try {
        const url = `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`;

        const params = new URLSearchParams();

        params.append("client_id", process.env.CLIENT_ID);
        params.append("client_secret", process.env.CLIENT_SECRET);
        params.append("scope", "https://graph.microsoft.com/.default");
        params.append("grant_type", "client_credentials");

        const response = await axios.post(url, params, {
            headers: {
                "Content-Type": "application/x-www-form-urlencoded"
            }
        });

        console.log("TOKEN RESPONSE:", response.data)

        return response.data.access_token;
    } catch (error) {
        console.error("Token Error:", error.response?.data || error.message);
        throw error;
    }
}

module.exports = { getGraphToken }