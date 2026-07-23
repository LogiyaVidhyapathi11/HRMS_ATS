const axios = require("axios");
require("dotenv").config()
const { getGraphToken } = require("./graphAuth");
 
/**
 * Parse a Teams meeting join URL to extract threadId, tenantId, and organizerId.
 * Example URL: https://teams.microsoft.com/l/meetup-join/19%3ameeting_xxx%40thread.v2/0?context={"Tid":"...","Oid":"..."}
 */
function parseJoinUrl(joinUrl) {
    const url = new URL(joinUrl);
    const pathParts = url.pathname.split("/");
 
    // The path is: /l/meetup-join/<encoded-threadId>/0
    // pathParts = ["", "l", "meetup-join", "19%3ameeting_...%40thread.v2", "0"]
    const threadId = decodeURIComponent(pathParts[3]);
 
    // The context query param contains Tid (tenant) and Oid (organizer)
    const contextParam = url.searchParams.get("context");
    const context = JSON.parse(contextParam);
 
    return {
        threadId,
        tenantId: context.Tid,
        organizerId: context.Oid
    };
}
 
async function joinMeeting(joinUrl) {
    console.log("Joining Teams Meeting...");
    console.log("joinUrl:", joinUrl);
 
    const token = await getGraphToken();
    const { threadId, tenantId, organizerId } = parseJoinUrl(joinUrl);
 
    console.log("Parsed from join URL:");
    console.log("  Thread ID:", threadId);
    console.log("  Tenant ID:", tenantId);
    console.log("  Organizer ID:", organizerId);
 
    const payload = {
        "@odata.type": "#microsoft.graph.call",
 
        "callbackUri": "https://interview.adamsbridgestage.com/bot-test/calls-test",
 
        "requestedModalities": ["audio"],
 
        "mediaConfig": {
            "@odata.type": "#microsoft.graph.serviceHostedMediaConfig"
        },
 
        "chatInfo": {
            "@odata.type": "#microsoft.graph.chatInfo",
            "threadId": threadId,
            "messageId": "0"
        },
 
        "source": {
            "@odata.type": "#microsoft.graph.participantInfo",
            "identity": {
                "@odata.type": "#microsoft.graph.identitySet",
                "application": {
                    "@odata.type": "#microsoft.graph.identity",
                    "displayName": "AI Interview Bot",
                    "id": process.env.CLIENT_ID
                }
            }
        },
 
        "meetingInfo": {
            "@odata.type": "#microsoft.graph.organizerMeetingInfo",
            "organizer": {
                "@odata.type": "#microsoft.graph.identitySet",
                "user": {
                    "@odata.type": "#microsoft.graph.identity",
                    "id": organizerId,
                    "tenantId": tenantId
                }
            }
        },
 
        "tenantId": tenantId
    };
 
    try {
        const response = await axios.post(
            "https://graph.microsoft.com/v1.0/communications/calls",
            payload,
            {
                headers: {
                    Authorization: `Bearer ${token}`,
                    "Content-Type": "application/json"
                }
            }
        );
 
        console.log("Graph Join Response:", response.data);
        return response.data;
    } catch (error) {
        console.log("-------FULL ERROR DEBUG-------");
        if (error.response) {
            console.log("Status:", error.response.status);
            console.log("Headers:", JSON.stringify(error.response.headers, null, 2));
            console.log("Data:", JSON.stringify(error.response.data, null, 2));
            if (error.response.data?.error) {
                console.log("Error Code:", error.response.data.error.code);
                console.log("Error Message:", error.response.data.error.message);
            }
        } else {
            console.log("Error:", error.message);
        }
        console.log("--------------------------------");
        throw error;
    }
}

module.exports = { joinMeeting, parseJoinUrl };