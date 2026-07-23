/**
 * callHandlers.js
 * ─────────────────────────────────────────────────────────────────────────────
 * Microsoft Graph Communications API — Call Event Handler
 *
 * Microsoft Graph posts real-time call lifecycle events to the bot's
 * callback URI (configured in teamsJoinService.js as CALLBACK_URI).
 *
 * This handler processes every event type and drives the interview flow:
 *
 *   established         → Bot is in the call → start Q1 immediately
 *   participantsUpdated → Log when candidate joins / leaves
 *   recordingCompleted  → Candidate finished answering → play next question
 *   terminated          → Call ended → clean up state
 *   (all others)        → Logged and acknowledged with 200
 *
 * ─────────────────────────────────────────────────────────────────────────────
 */

const { getInterviewStateByCallId, incrementQuestion, deleteInterviewState } = require("../services/interviewState");
const { playQuestionAndRecord, terminateCall, playPrompt, terminateMeeting, kickParticipant } = require("../services/interviewQuestionPlayer");
const { sendChatMessage } = require("../services/chatService");
const { 
    fetchRecordingAfterCall, 
    downloadRecordingFile, 
    notifyBackendRecordingReady
} = require("../services/recordingService");
    
/**
 * Main entry point — called by Express for every POST to /bot-test/calls-test
 */
async function handleCall(req, res) {
    // Always respond 200 immediately — Graph requires a fast acknowledgement
    res.sendStatus(200);
    const body = req.body;
    // Graph sends an array of event notifications
    const notifications = body?.value;
    
    if (!notifications || !Array.isArray(notifications)) {
        console.log("[CallHandler] Received non-notification payload:", JSON.stringify(body, null, 2));
        return;
    }

    for (const notification of notifications) {
        await processNotification(notification);
    }
}
/**
 * Routes a single Graph notification to the correct handler.
 */
async function processNotification(notification) {
    const changeType = notification?.changeType;
    const resourceData = notification?.resourceData || {};
    const resource = notification?.resource || "";

    //Microsoft Graph operations use resourceData.id for the operation ID, not the call ID.
    //The true call ID is always embedded in the resource URI.

    const resourceUri = notification?.resource || "";
    const callIdMatch = resourceUri.match(/calls\/([^/]+)/);
    let callId = callIdMatch ? callIdMatch[1] : resourceData?.id;

    if (!callId && resource) {
        const match = resource.match(/calls\/([a-zA-Z0-9-]+)/i);
        if (match) {
            callId = match[1];
        }
    }

    const callState = resourceData?.state;
    let resourceType = resourceData?.["@odata.type"] || "";
 
    if (!resourceType) {
        if (resource.includes("/operations")) {
            resourceType = "#microsoft.graph.recordOperation";
        } else if (resource.includes("/participants")) {
            resourceType = "#microsoft.graph.participant";
        } else if (resource.includes("calls/")) {
            resourceType = "#microsoft.graph.call";
        }
    }

    console.log(`\n[CallHandler] ── Event ──────────────────────────────────`);
    console.log(`  changeType   : ${changeType}`);
    console.log(`  resourceType : ${resourceType}`);
    console.log(`  callId       : ${callId}`);
    console.log(`  callState    : ${callState || "N/A"}`);

    // ── Call state transitions ───────────────────────────────────────────────
    if (resourceType.includes("call")) {
        switch (callState) {
            case "established":
                await onCallEstablished(callId);
                break;
            case "terminated":
                await onCallTerminated(callId);
                break;
            default:
                console.log(`[CallHandler] Unhandled call state: ${callState}`);
        }
        return;
    }

    // ── Recording completed ─────────────────────────────────────────────────
    // Graph fires this when recordResponse finishes (either timeout, # key, or silence)
    if (resourceType.includes("recordOperation")) {
        const operationStatus = resourceData?.status;
        console.log(`  operationStatus: ${operationStatus}`);

        if (operationStatus === "completed") {
            await onRecordingCompleted(callId, resourceData);
        } else if (operationStatus === "failed") {
            console.error(`[callHandler] Recording operation FAILED for callId: ${callId}`);
            console.error(`[CallHandler] Failure Details:`, JSON.stringify(resourceData, null, 2));
            
            // Handle timeout by advancing to the next question so the bot doesn't hang
            const completionReason = resourceData?.completionReason;
            if (completionReason === "initialSilenceTimeout" || completionReason === "timeout") {
                console.log(`[CallHandler] Timeout reached without response. Advancing interview...`);
                
                const state = getInterviewStateByCallId(callId);
                if (state) {
                    state.consecutiveTimeouts = (state.consecutiveTimeouts || 0) + 1;
                    console.log(`[CallHandler] Consecutive Timeouts: ${state.consecutiveTimeouts} / 3`);
                    
                    if (state.consecutiveTimeouts >= 4) {
                        console.error(`[CallHandler] ❌ 3 consecutive timeouts reached. Terminating meeting.`);
                        await sendChatMessage(state.threadId, "The interview has been terminated due to 3 consecutive silence timeouts.");
                        await playFinalNoteAndEnd(callId, state);
                        return;
                    }
                }
                
                await onRecordingCompleted(callId, resourceData);
            }
        }
        return;
    }

    // ── Participants updated ─────────────────────────────────────────────────
    if (resourceType.includes("participant")) {
        const participants = [];

        // Handle all possible ways Graph sends participant lists
        if (Array.isArray(resourceData)) {
            participants.push(...resourceData);
        } else if (Array.isArray(resourceData?.value)) {
            participants.push(...resourceData.value);
        } else if (Array.isArray(resourceData?.members)) {
            participants.push(...resourceData.members);
        } else if (resourceData && typeof resourceData === "object") {
            // It is a singular participant notification
            participants.push(resourceData);
        }

        console.log(`[CallHandler] Participants updated: ${participants.length} total member(s) seen`);

        const state = getInterviewStateByCallId(callId); 
        if (!state) {
            console.warn(`[CallHandler] ⚠️ No state found for callId: ${callId} - cannot detect humans.`);
            return;
        }
        
        for (const p of participants) {
            // Identity can be in p.info.identity or just p.identity depending on Graph version
            const identity = p.info?.identity || p.identity;

            console.log(`[CallHandler] Inspecting participant: ID = ${p.id}, Info = ${JSON.stringify(p.info || p)}`);
            
            // If it's a "user" or "guest", it's a human
            if (identity?.user || identity?.guest) {
                const name = identity.user?.displayName || identity.guest?.displayName || "Human";

                const userId = identity.user?.id || identity.guest?.id;

                // Check if this is the organizer to avoid misidentifying organizer as candidate
                if (userId && state.organizerId && userId === state.organizerId) {
                    console.log(`[CallHandler] Organizer "${name}" detected (ID: ${userId}). Skipping candidate mapping.`);
                    continue;
                }

                if (!state.hasHumanJoined) {
                    console.log(`[CallHandler] 👤 Human candidate detected: "${name}". Triggering interview!`);
                    state.hasHumanJoined = true;
                }

                // Store candidate participant ID so we can kick them when the call completes
                if (p.id) {
                    state.candidateParticipantId = p.id;
                    console.log(`[CallHandler] Stored candidateParticipantId: ${p.id} for "${name}"`);
                }
            }
        }
        return;
    }
    console.log("[CallHandler] Unhandled notification:", JSON.stringify(notification, null, 2));
}
/**
 * Fires when the bot has successfully joined the Teams meeting.
 * Starts the interview by playing Question 1.
 */
async function onCallEstablished(callId) {
    console.log(`\n[CallHandler] ✅ Call ESTABLISHED — callId: ${callId}`);
    console.log(`[CallHandler] Looking up interview state for callId...`);

    const state = getInterviewStateByCallId(callId);
    if (!state) {
        console.warn(`[CallHandler] ⚠️  No interview state found for callId: ${callId}`);
        console.warn(`[CallHandler]    The bot may have joined before state was registered.`);
        return;
    }

    console.log(`[CallHandler] Found interview: candidate="${state.candidateName}", questions=${state.questions.length}`);

    // Guard: Prevent multiple interview loops
    if (state.interviewStarted) {
        console.log(`[CallHandler] Interview logic already started for callId: ${callId}. Ignoring duplicate event.`);
        return;
    }
    state.interviewStarted = true;

    // Step 1: Wait for TTS audio to be ready
    await waitForAudioReady(state, 30_000);
    if (state.audioUrls.length === 0) {
        console.error("[CallHandler] ❌ No audio URLs available — cannot start interview.");
        return;
    }

    // ── Step 2: Wait for Candidate to physically join ────────────────────────
    // This prevents the bot from speaking to an empty room.
    console.log(`[CallHandler] ⏳ Waiting for candidate "${state.candidateName}" to join the room...`);
    const joinResult = await waitForCandidateJoin(state, 600_000); // 10 minute timeout

    if (!joinResult) {
        console.error(`[CallHandler] ❌ Candidate failed to join within 10 minutes. Hanging up.`);
        await endInterview(callId, state);
        return;
    }

    // Enforce Meeting Time Controls

    // This monitors the end time and plays warnings at 5 min and 1 min marks
    startTimeControl(callId);

    // ── Step 3: Small buffer for audio setup ────────────────────────────────
    console.log(`[CallHandler] ✅ Candidate in room. Waiting 2s for audio to settle...`);
    await delay(2000);

    console.log(`[CallHandler] 🎙 Starting interview with Q1 for "${state.candidateName}"`);
    state.status = "waiting"; // Initialize status
    await askQuestion(callId, state, 0);
}
/**
 * Fires when the candidate's answer recording is complete.
 * Saves the recording reference and plays the next question.
 */
async function onRecordingCompleted(callId, resourceData) {

    const recordingLocation = resourceData?.recordedStreams?.[0]?.contentLocation;
    const state = getInterviewStateByCallId(callId);

    if (!state) {
        console.error(`[CallHandler] State not found for callId: ${callId}`);
        return;
    }

    console.log(`\n[CallHandler] 🎙 Recording COMPLETED — callId: ${callId}`);

    if (recordingLocation) {
        console.log(`[CallHandler]    Recording URL: ${recordingLocation}`);

        // Reset consecutive timeouts because the candidate spoke successfully
        state.consecutiveTimeouts = 0;

        // TODO: Download and store this recording in MongoDB for HR review
    }

    state.status = "waiting" ; // Reset status after recording finished

    // Check if we are past the end time
    if (state.endTime && new Date() >= new Date(state.endTime)) {
        if (!state.warningsSent.timesUp) {
            state.warningsSent.timesUp = true;

            console.log(`[CallHandler] End Time reached during recording. Concluding interview.`);
            
            await concludeInterviewBeyondTime(callId, state);
        } else {
            console.log(`[CallHandler] End Time reached, and conclusion already in progress. Ignoring.`);
        }
        return;
    }

    const nextIndex = incrementQuestion(callId);
    console.log(`[CallHandler] ➡️  Moving to question ${nextIndex + 1} / ${state.questions.length}`);

    if (nextIndex >= state.questions.length) {
        console.log(`[CallHandler] 🏁 All ${state.questions.length} questions completed! Ending call.`);
        await playFinalNoteAndEnd(callId, state);
        return;
    }

    // // 1 second pause between questions for natural conversation flow
    // await delay(1000);
    // Immediately ask the next question without waiting
    await askQuestion(callId, state, nextIndex);
}
/**
 * Fires when the call is terminated (by bot, candidate, or timeout).
 */
async function onCallTerminated(callId) {
    console.log(`\n[CallHandler] 📴 Call TERMINATED — callId: ${callId}`);

    const state = getInterviewStateByCallId(callId);
    if (state) {
        console.log(`[CallHandler]    Cleaning up state for candidate: ${state.candidateName}`);

        // ── Fetch and Download full-session recording in background 
        if (state.meetingId) {
            console.log(`[CallHandler] 🎬 Scheduling recording fetch for meetingId: ${state.meetingId}`);
            (async () => {
                try {
                        // Wait 2 minutes before first poll — Teams needs time to finalise the file
                        console.log("[Recording] ⏳ Waiting 2 minutes for Teams to process the recording...");

                        await delay(120_000);

                        const result = await fetchRecordingAfterCall(state.meetingId);

                        if (result && result.contentUrl) {
                            console.log(`\n [Recording] Full session recording ready!`);

                            // Professional naming convention: Recording_Name_MeetingID.mp4

                            const safeName = state.candidateName.replace(/\s+/g, "_");

                            // Sanitize Meeting ID (remove characters like :, *, @ which are illegal in filenames)

                            const safeMeetingId = state.meetingId.replace(/[:*@]/g, "_");

                            const filename = `Recording_${safeName}_${safeMeetingId}.mp4`;

                            // Step 1: Download the file locally
                            const localPath = await downloadRecordingFile(result.contentUrl, filename);

                            if (localPath) {
                                // Step 2: Notify backend to start AI Analysis
                                const candidateDetails = {
                                    name: state.candidateName, 
                                    email: state.candidateEmail, 
                                    threadId: state.threadId, 
                                    questions: state.questions
                                }; 

                                const recordingData = {
                                    recordingId: result.recordingId, 
                                    localPath: localPath,
                                    contentUrl: result.contentUrl
                                }; 

                                await notifyBackendRecordingReady(candidateDetails, recordingData);
                            }
                        } else {
                            console.error("[Recording] Recording not retreieved - it may not have been enabled for this meeting.");
                        }

                    } catch (error) {
                        console.error("[Recording] Background recording process failed:", error.message);
                    } finally {
                        // Meeting cleanup only after recording is done.
                        if (state.organizerId && (state.eventId || state.meetingId)) {
                            await terminateMeeting(state.organizerId, state.eventId, state.meetingId)
                        }
                    }
            })();
        } else {
            console.warn("[CallHandler] No meetingId in state - skipping recording fetch.")
        }

        // Find and delete by callId
        deleteInterviewState(state.threadId);
    }
}
// ── Helpers ──────────────────────────────────────────────────────────────────
/**
 * Plays the question audio for `questionIndex` and starts recording.
 */
async function askQuestion(callId, state, questionIndex) {
    const audioUrl = state.audioUrls[questionIndex];

    console.log("==================================");
    console.log("ASK QUESTION");
    console.log("Question Index:", questionIndex);
    console.log("Total Questions:", state.questions.length);
    console.log("==================================");

    if (!audioUrl) {
        console.error(`[CallHandler] ❌ No audio URL for Q${questionIndex + 1}. Skipping.`);
        // Try the next question
        const nextIndex = incrementQuestion(callId);
        if (nextIndex < state.questions.length) {
            await askQuestion(callId, state, nextIndex);
        } else {
            await playFinalNoteAndEnd(callId, state);
        }
        return;
    }
    const questionText = state.questions[questionIndex]?.text;
    console.log(`\n[CallHandler] 📢 Q${questionIndex + 1}: ${questionText}`);
    try {
        state.status = "listening"; // Set status to listening during recordResponse.
        await playQuestionAndRecord(callId, audioUrl, questionIndex);
    } catch (err) {
        state.status = "waiting"; // Reset on error
        console.error(`[CallHandler] ❌ Failed to play Q${questionIndex + 1}:`, err.message);
    }
}
/**
 * Gracefully ends the interview call and cleans up state.
 */
async function endInterview(callId, state) {
    console.log(`\n[CallHandler] 🏁 Interview complete for "${state.candidateName}". Hanging up.`);

    console.log(`[CallHandler] State diagnostics at hangup:`, {
        candidateName: state.candidateName, 
        candidateEmail: state.candidateEmail, 
        hasHumanJoined: state.hasHumanJoined, 
        interviewStarted: state.interviewStarted, 
        candidateParticipantId: state.candidateParticipantId, 
        meetingId: state.meetingId, 
        organizerId: state.organizerId
    });

    // 1. Kick candidate out of the meeting immediately
    if (state.candidateParticipantId) {
        console.log(`[CallHandler] Attempting to kick candidate: ${state.candidateParticipantId}`);
        await kickParticipant(callId, state.candidateParticipantId);
    } else {
        console.warn("[CallHandler] candidateParticipantId not found - cannot kick candidate.");
    }

    // Small delay to ensure kick request completes
    await delay(1000);

    // 2. Bot hangs up its own call leg
    await terminateCall(callId);

    // Note: terminateMeeting is now handled in the finally block of the background recording process in onCallTerminated.
}

/**
 * Plays the final note and then ends the interview.
 */
async function playFinalNoteAndEnd(callId, state) {
    if (state.finalNote) {
        console.log(`[CallHandler] Playing final thank you note...`)
        await playPrompt(callId, state.finalNote);
        await delay(10000); // Wait for audio to finish + 5s buffer before leaving
    } else {
        console.warn(`[CallHandler] finalNote URL not found in state, skipping thank you note.`)
    }
    await endInterview(callId, state);
}

/**
 * Handles the "beyond time" conclusion logic.
 */
async function concludeInterviewBeyondTime(callId, state) {
    // Send "Time's up" popup notification
    const message = "The scheduled time for this interview has reached its end. We will conclude in 15 seconds.";
    await sendChatMessage(state.threadId, message);
    
    console.log(`[CallHandler] ⏳ Waiting 15 seconds as per requirements...`);
    await delay(15000);

    await playFinalNoteAndEnd(callId, state);
}

/**
 * Starts a periodic check for interview time controls (warnings).
 */
function startTimeControl(callId) {
    const timer = setInterval(async () => {
        const state = getInterviewStateByCallId(callId);
        if (!state || !state.endTime) {
            clearInterval(timer);
            return;
        }

        const now = new Date();
        const endTime = new Date(state.endTime);
        const remainingMs = endTime - now;

        // 5 minute warning
        if (remainingMs <= 300_000 && remainingMs > 240_000 && !state.warningsSent.min5) {
            state.warningsSent.min5 = true;
            console.log(`[CallHandler] ⏱️ 5-minute warning popup for "${state.candidateName}"`);
            await sendChatMessage(state.threadId, "Reminder: There are 5 minutes left in the interview.");
        }

        // 1 minute warning
        if (remainingMs <= 60_000 && remainingMs > 0 && !state.warningsSent.min1) {
            state.warningsSent.min1 = true;
            console.log(`[CallHandler] ⏱️ 1-minute warning popup for "${state.candidateName}"`);
            await sendChatMessage(state.threadId, "Final Reminder: There is only 1 minute left in the interview.");
        }

        // End time reached
        if (remainingMs <= 0) {
            clearInterval(timer);
            console.log(`[CallHandler] ⏱️ End time reached for "${state.candidateName}". Checking status: ${state.status}`);
            
            if (!state.warningsSent.timesUp) {
                state.warningsSent.timesUp = true;
                
                if (state.status === "listening") {
                    console.log(`[CallHandler] 🗣️ Candidate is currently answering. Providing 15s grace time.`);
                    await concludeInterviewBeyondTime(callId, state);
                } else {
                    console.log(`[CallHandler] 🤫 Candidate is not answering. Ending immediately.`);
                    await playFinalNoteAndEnd(callId, state);
                }
            }
        }
    }, 10000); // Check every 10 seconds
}

/**
 * Polls every 2s until audio URLs are available (TTS generating in background).
 * Times out after `maxWaitMs` milliseconds.
 */
async function waitForAudioReady(state, maxWaitMs = 30_000) {
    const POLL_INTERVAL = 2000;
    let waited = 0;
    while (state.audioUrls.length === 0 && waited < maxWaitMs) {
        console.log(`[CallHandler] ⏳ Waiting for TTS audio... (${waited / 1000}s elapsed)`);
        await delay(POLL_INTERVAL);
        waited += POLL_INTERVAL;
    }
    if (state.audioUrls.length === 0) {
        console.warn(`[CallHandler] ⚠️  TTS audio not ready after ${maxWaitMs / 1000}s.`);
    } else {
        console.log(`[CallHandler] ✅ TTS audio ready (${state.audioUrls.length} files).`);
    }
}

/**
 * Polls every 2s until a human participant is detected (state.hasHumanJoined).
 * Times out after `maxWaitMs` milliseconds.
 */
async function waitForCandidateJoin(state, maxWaitMs = 300_000) {
    const POLL_INTERVAL = 2000;
    let waited = 0;

    while (!state.hasHumanJoined && waited < maxWaitMs) {
        await delay(POLL_INTERVAL);
        waited += POLL_INTERVAL;
        
        // Log every 10s to keep the console alive
        if (waited % 10000 === 0) {
            console.log(`[CallHandler] ... still waiting for candidate join (${waited / 1000}s)`);
        }
    }

    return state.hasHumanJoined;
}

/**
 * Simple promise-based delay helper.
 */
function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}
module.exports = { handleCall };
