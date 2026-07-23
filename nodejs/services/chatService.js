/**
 * chatService.js
 * ─────────────────────────────────────────────────────────────────────────────
 * Sends messages to a Teams chat thread (meeting chat).
 * ─────────────────────────────────────────────────────────────────────────────
 */

const axios = require("axios");
const { getGraphToken } = require("./graphAuth");

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

/**
 * Sends a plain text message to the specified chat thread.
 * 
 * @param {string} threadId - The Teams thread ID (e.g. 19:meeting_xxx@thread.v2)
 * @param {string} content - The message content.
 */
async function sendChatMessage(threadId, content) {
    try {
        const token = await getGraphToken();

        const payload = {
            body: {
                content: content
            }
        };

        const response = await axios.post(
            `${GRAPH_BASE}/chats/${threadId}/messages`,
            payload,
            {
                headers: {
                    Authorization: `Bearer ${token}`,
                    "Content-Type": "application/json"
                }
            }
        );

        console.log(`[ChatService] ✅ Message sent to thread ${threadId}: "${content}"`);
        return response.data;
    } catch (error) {
        console.error(`[ChatService] ❌ Failed to send message to ${threadId}:`, error.response?.data || error.message);
        // Fallback or silent failure
    }
}

module.exports = { sendChatMessage };
