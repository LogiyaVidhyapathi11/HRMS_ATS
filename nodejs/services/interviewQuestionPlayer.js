/**
 * interviewQuestionPlayer.js
 * ─────────────────────────────────────────────────────────────────────────────
 * Plays an interview question audio file inside an active Teams call and
 * records the candidate's spoken response via the Microsoft Graph
 * Communications API — recordResponse action.
 *
 * Graph API reference:
 *   POST /communications/calls/{callId}/recordResponse
 *   https://learn.microsoft.com/en-us/graph/api/call-recordresponse
 *
 * Required permissions (Azure AD app):
 *   Calls.AccessMedia.All
 *   Calls.JoinGroupCall.All
 * ─────────────────────────────────────────────────────────────────────────────
 */
const axios = require("axios");
const { getGraphToken } = require("./graphAuth");

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
/**
 * Instructs the bot to play an audio prompt and then record the candidate's
 * spoken answer. Microsoft Graph fires a `recordingCompleted` notification
 * to the callback URI when recording finishes.
 *
 * @param {string} callId    - The active Graph call ID.
 * @param {string} audioUrl  - Public HTTPS URL of the .wav question audio.
 * @param {number} questionIndex - 0-based question index (for logging).
 * @returns {Promise<object>} - Graph API response data.
 */
async function playQuestionAndRecord(callId, audioUrl, questionIndex) {
    const token = await getGraphToken();
    const payload = {
        // The client context is echoed back in the recordingCompleted event
        // so the handler knows which call/question this belongs to.
        clientContext: `interview-q${questionIndex + 1}-${callId}`,
        // Prompt: play our TTS audio file to the candidate
        prompts: [
            {
                "@odata.type": "#microsoft.graph.mediaPrompt",
                mediaInfo: {
                    "@odata.type": "#microsoft.graph.mediaInfo",
                    uri: audioUrl,
                    resourceId: `q${questionIndex + 1}`
                }
            }
        ],

        // Max recording duration per answer (10 minutes)
        maxRecordDurationInSeconds: 600,

        // Stop recording after 15 seconds of silence or `#` key
        initialSilenceTimeoutInSeconds: 15,
        finalSilenceTimeoutInSeconds: 1,  // Reduced from 3 to 1 for faster transition,

        // Play a beep before recording starts (like a voicemail)
        playBeep: false,

        // Stop tones candidates can press to submit early (# key)
        stopTones: ["#"]
    };

    console.log(`[Player] 🎙 Playing Q${questionIndex + 1} | callId: ${callId}`);
    console.log(`[Player]    Audio URL: ${audioUrl}`);
    
    try {
        const response = await axios.post(
            `${GRAPH_BASE}/communications/calls/${callId}/recordResponse`,
            payload,
            {
                headers: {
                    Authorization: `Bearer ${token}`,
                    "Content-Type": "application/json"
                }
            }
        );
        console.log(`[Player] ✅ recordResponse initiated for Q${questionIndex + 1}`);
        return response.data;
    } catch (error) {
        console.error(`[Player] ❌ recordResponse failed for Q${questionIndex + 1}:`);
        if (error.response) {
            console.error("  Status:", error.response.status);
            console.error("  Body:", JSON.stringify(error.response.data, null, 2));
        } else {
            console.error(" ", error.message);
        }
        throw error;
    }
}

/**
 * Plays a simple audio prompt to the call without recording
 * Useful for warnings and final notes
 */

async function playPrompt(callId, audioUrl)
{
    const token = await getGraphToken();

    const payload = {
        clientContext: `prompt-${Date.now()}-${callId}`, 
        prompts: [
            {
                "@odata.type": "#microsoft.graph.mediaPrompt", 
                mediaInfo: {
                    "@odata.type": "#microsoft.graph.mediaInfo", 
                    uri: audioUrl
                }
            }
        ]
    };

    console.log(`[Player] Playing prompt | CallId: ${callId} | URL: ${audioUrl}`);

    try {
        await axios.post(
            `${GRAPH_BASE}/communications/calls/${callId}/playPrompt`, 
            payload, 
            {
                headers: {
                    Authorization: `Bearer ${token}`, 
                    "Content-Type": "application/json"
                }
            }
        );

        console.log(`[Player] Prompt initiated.`);
    } catch (error) {
        console.error(`[Player] playPrompt failed:`, error.response?.data || error.message);
    }
}

/**
 * Hangs up (terminates) the bot's active call via Graph API.
 * Optionally terminates the entire meeting room for all participants.
 *
 * @param {string} callId - The active Graph call ID to terminate.
 * @param {string} organizerId - The organizer's OID (to delete the meeting).
 * @param {string} meetingId - The online meeting ID (to delete the meeting).
 * 
 * @returns {Promise<void>}
 */
async function terminateCall(callId) {
    const token = await getGraphToken();
    console.log(`[Player] 📴 Terminating call: ${callId}`);
    try {
        await axios.delete(
            `${GRAPH_BASE}/communications/calls/${callId}`,
            {
                headers: {
                    Authorization: `Bearer ${token}`
                }
            }
        );
        console.log(`[Player] ✅ Call terminated successfully.`);

        // If meeting info is provided terminate the entire room for everyone
        // if (organizerId && meetingId) {
        //     await deleteMeeting(organizerId, meetingId);
        // }
        
    } catch (error) {
        console.error("[Player] ❌ Failed to terminate call:", error.response?.data || error.message);
    }
}

// /**
//  * Kicks a specific participant out of the active call via Graph API.
//  * @param {string} callId - The active Graph call ID.
//  * @param {string} participantId - The Graph participant ID to kick.
//  * @returns {Promise<void>}
//  */
// async function kickParticipant(callId, participantId) {
//     if (!callId || !participantId) {
//         console.warn("[Player] Cannot kick participant: callId or participantId missing.")
//         return;
//     }

//     const token = await getGraphToken();
//     console.log(`[Player] Kicking participant: ${participantId} from call: ${callId}`);

//     try {
//         await axios.delete(
//             `${GRAPH_BASE}/communications/calls/${callId}/participants/${participantId}`, 
//             {
//                 headers: { Authorization: `Bearer ${token}` }
//             }
//         );
//         console.log(`[Player] Participant kicked successfully.`)
//     } catch (error) {
//         console.error("[Player] Failed to kick participant:", error.response?.data || error.message);
//     }
// }


/**
 * Terminate the entire meeting to kick everyone out.
 * Priority 1: Delete Calendar Event (Kick out + remove from calendar).
 * Priority 2: Delete Online Meeting (Kick out)
 */

async function terminateMeeting(organizerId, eventId, meetingId) {
    console.log(`[Player] terminateMeeting called with: organizerId = ${organizerId}, eventId = ${eventId}, meetingId = ${meetingId}`);
    
    if (!organizerId) {
        console.error("[Player] Cannot terminate meeting: organizerId is not defined.")
        return;
    }

    const token = await getGraphToken();

    try {
        if (eventId) {
            console.log(`[Player] Deleting Calendar Event: ${eventId}`);
            await axios.delete(
                `${GRAPH_BASE}/users/${organizerId}/events/${eventId}`, 
                {
                    headers: {
                        Authorization: `Bearer ${token}`
                    }
                }
            ); 
            console.log(`[Player] Event deleted. Meeting ended for all`);

        } else if (meetingId) {
            console.log(`[Player] Deleting Online Meeting: ${meetingId}`);

            await axios.delete(
                `${GRAPH_BASE}/users/${organizerId}/onlineMeetings/${meetingId}`, 
                {
                    headers: {
                        Authorization: `Bearer ${token}`
                    }
                }
            ); 
            console.log(`[Player] Online Meeting deleted. Meeting ended for all.`);

        } else {
            console.warn("[Player] No eventId or meetingId available to terminate the meeting.")
        }
    } catch (error) {
        console.log("[Player] Failed to terminate meeting:", error.response?.data || error.message);
    }
}

module.exports = { playQuestionAndRecord, terminateCall, playPrompt, terminateMeeting, kickParticipant };
