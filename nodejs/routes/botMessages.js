/**
 * botMessages.js
 * ─────────────────────────────────────────────────────────────────────────────
 * Express router for the Node.js Teams Bot service.
 *
 * Routes:
 *   POST /bot-test/join-meeting-test  — Python backend triggers bot to join
 *   POST /bot-test/calls-test         — Microsoft Graph sends call events here
 * ─────────────────────────────────────────────────────────────────────────────
 */

const express = require("express");
const router = express.Router();

const { joinMeeting, parseJoinUrl } = require("../services/teamsJoinService");
const { handleCall } = require("../handlers/callHandler");
const { generateQuestionAudio } = require("../services/ttsService");
const { createInterviewState, setCallIdForThread,
        getInterviewState } = require("../services/interviewState");

// Public base URL where audio files are served (set in .env)
const AUDIO_BASE_URL = process.env.AUDIO_BASE_URL
    || "https://interview.adamsbridgestage.com/public/audio";
// ─────────────────────────────────────────────────────────────────────────────
// POST /bot-test/join-meeting-test
// Called by the Python backend (via APScheduler → bot_trigger_service.py)
// at the scheduled interview start time.
// ─────────────────────────────────────────────────────────────────────────────
router.post("/join-meeting-test", async (req, res) => {
    try {
        
        console.log("\n[Bot] ─── Join Meeting Request Received ────────────────");

        const { meetingUrl, candidateName, candidateEmail, questions, endTime, meetingId, eventId, event_id } = req.body;

        const finalEventId = eventId || event_id;
        
        if (!meetingUrl) {
            return res.status(400).json({ error: "meetingUrl is required" });
        }

        console.log("[Bot] Meeting URL :", meetingUrl);
        console.log("[Bot] Candidate   :", candidateName || "Unknown");
        console.log("[Bot] Questions   :", questions?.length || 0);
        console.log("[Bot] End Time    :", endTime || "N/A");
        console.log("[Bot] meetingId:", meetingId);
        console.log("[Bot] eventId:", eventId);

        // Parse the Teams join URL to extract the thread ID and organizer ID
        const { threadId, organizerId } = parseJoinUrl(meetingUrl);
        console.log("[Bot] organizerId:", organizerId);
        
        // ── Step 1: Initialise interview state (if not already exists)

        const existingState = getInterviewState(threadId);
        if (existingState) {
            console.warn(`[Bot] Interview state already exists for thread: ${threadId}. Skipping reset.`);
        } 
        else {
            console.log("[Bot Creating new interview state for thread:", threadId);

            createInterviewState(
                threadId, 
                candidateName || "Candidate", 
                candidateEmail || "unknown@example.com", 
                questions || [], 
                [], // audioUrls will be filled asynchronously by TTS
                endTime, 
                meetingId || null, 
                organizerId || null, 
                eventId || null
            );
       
        }

        // ── Step 2: Join the meeting immediately ────────────────────────────
        // We join BEFORE TTS finishes so we don't miss the meeting window.
        // The call handler waits for audio URLs before playing Q1.
        const joinResponse = await joinMeeting(meetingUrl);
        
        // Link the Graph-assigned Call ID to our interview state
        if (joinResponse?.id) {
            setCallIdForThread(threadId, joinResponse.id);
            console.log(`[Bot] ✅ Joined meeting — Graph callId: ${joinResponse.id}`);
        }
       
        // ── Step 3: Generate TTS audio for all questions (background) ──────
        // Runs async — audio is ready well before candidate joins (~30–60s)
        if (Array.isArray(questions) && questions.length > 0) {
            (async () => {
                
                console.log(`\n[Bot] 🔊 Starting Azure TTS for ${questions.length} questions...`);

                // Prepare dynamic time-based greeting in IST (UTC+5:30)
                const now = new Date();
                const istMs = now.getTime() + (5.5 * 60 * 60 * 1000);
                const istHour = new Date(istMs).getUTCHours();
                let greeting;
                if (istHour < 12) {
                    greeting = "Good Morning";
                } else if (istHour < 17) {
                    greeting = "Good Afternoon";
                } else {
                    greeting = "Good Evening";
                }
                const introText = `Hello, ${greeting} ${candidateName || "Candidate"}. Welcome to your interview. To start with, please introduce yourself.`;
                const introQuestion = { text: introText, type: "Introductory" };

                // Safely clone questions array to avoid reference duplication
                let questionsList = Array.isArray(questions) ? [...questions] : [];

                // Check if index 0 is already an Introductory question
                const firstType = (questionsList[0]?.type || "").toLowerCase();
                const firstText = (typeof questionsList[0] === "string" ? questionsList[0] : (questionsList[0]?.text || "")).toLowerCase();
                const isAlreadyIntro = firstType === "introductory" || firstText.includes("introduce yourself") || firstText.includes("hello");


                if (isAlreadyIntro) {
                    // Update index 0 with the personalized time-based greeting
                    if (typeof questionsList[0] === "object") {
                        questionsList[0].text = introText;
                        questionsList[0].type = "Introductory";
                    } else {
                        questionsList[0] = { text: introText, type: "Introductory" };
                    }
                } else {
                    // Insert intro ONCE at the beginning if not present
                    questionsList.unshift(introQuestion);
                }

                // Update state questions array reference cleanly
                const state = getInterviewState(threadId);
                if (state) {
                    state.questions = questionsList;
                }

                const audioUrls = [];
                const safeThread = threadId.replace(/[^a-zA-Z0-9_-]/g, "");

                for (let i = 0; i < questionsList.length; i++) {
                    const questionText = questionsList[i]?.text || questionsList[i];
                    const filename = `${safeThread}_q${i + 1}.wav`;
                    try {
                        await generateQuestionAudio(questionText, filename);
                        
                        const url = `${AUDIO_BASE_URL}/${filename}`;
                        audioUrls.push(url);
                        console.log(`[Bot] ✅ TTS done — Q${i + 1}: ${filename}`);
                    } catch (err) {
                        // Non-fatal — push null so index alignment is preserved
                        audioUrls.push(null);
                        console.error(`[Bot] ❌ TTS failed for Q${i + 1}:`, err.message);
                    }
                }

                // Step 4: Generate Final Note audio (background)
                try {
                    const finalNoteText = "Thank you for your time. The interview is now complete. Have a great day.";
                    const finalNoteFilename = `${safeThread}_final_note.wav`;

                    await generateQuestionAudio(finalNoteText, finalNoteFilename);

                    if (state) {
                        state.finalNote = `${AUDIO_BASE_URL}/${finalNoteFilename}`;
                    }
                    console.log("[Bot] Final Note audio ready");
                } catch (err) {
                    console.error("[Bot] Final Note TTS failed:", err.message);
                }
                
                // Attach all generated audio URLs to the interview state
                if (state) {
                    state.audioUrls = audioUrls;
                    console.log(`\n[Bot] ✅ All TTS complete — ${audioUrls.filter(Boolean).length}/${questions.length} audio files ready.`);
                }
            })();
        }
        
        res.json({
            message: "Bot is joining the meeting. Interview will begin shortly.",
            callId:  joinResponse?.id,
            threadId
        });
    } catch (error) {
        console.error("[Bot] ❌ Error joining meeting:", error.response?.data || error.message);
        res.status(500).json({
            message: "Error joining meeting",
            error:   error.response?.data || error.message
        });
    }
});

// ─────────────────────────────────────────────────────────────────────────────
// POST /bot-test/calls-test
// Microsoft Graph posts call lifecycle events here (established, recording
// completed, participants updated, terminated, etc.)
// ─────────────────────────────────────────────────────────────────────────────
router.post("/calls-test", handleCall);

// router.post("/calls", handleCall);

// // ─────────────────────────────────────────────────────────────────────────────
// // POST /bot/messages
// // Handle standard Bot Framework messaging events (used for "Test in Web Chat")
// // ─────────────────────────────────────────────────────────────────────────────
// router.post("/bot/messages", (req, res) => {
//     console.log("[Bot] 💬 Received message in chat:", req.body?.text);
//     // Standard bot response to keep the Web Chat happy
//     res.status(200).send();
// });

module.exports = router;
