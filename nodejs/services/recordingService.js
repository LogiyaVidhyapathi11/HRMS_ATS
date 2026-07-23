/**
 * recordingService.js
 * ─────────────────────────────────────────────────────────────────────────────
 * Fetches the full-session recording produced by Teams' native auto-recording.
 *
 * When a meeting is created with `recordAutomatically: true`, Teams records the
 * entire session as an MP4 and stores it in the organizer's OneDrive.
 *
 * After the call ends, the recording takes 1–5 minutes to process. This service
 * polls the Graph API with retries until the recording file is available.
 *
 * Graph API reference:
 *   GET /users/{userId}/onlineMeetings/{meetingId}/recordings
 *   https://learn.microsoft.com/en-us/graph/api/onlinemeeting-list-recordings
 *
 * Required permissions:
 *   OnlineMeetingRecording.Read.All
 * ─────────────────────────────────────────────────────────────────────────────
 */

const axios = require("axios");
const fs = require("fs");
const path = require("path");
const https = require("https");
const { pipeline } = require("stream/promises");
const { getGraphToken } = require("./graphAuth");

/**
 * Uses Node's native https module to resolve the Graph API redirect.
 * This is more reliable than Axios for capturing 302 Location headers.
 *
 * @param {string} url   - Graph content URL.
 * @param {string} token - Bearer token.
 * @returns {Promise<string|null>} - The SharePoint download URL, or null.
 */
function resolveGraphRedirect(url, token) {
    return new Promise((resolve, reject) => {
        const parsedUrl = new URL(url);
        const options = {
            hostname: parsedUrl.hostname,
            path: parsedUrl.pathname + parsedUrl.search,
            method: "GET",
            headers: { Authorization: `Bearer ${token}` }
        };

        const req = https.request(options, (res) => {
            console.log(`[Recording] 🔍 Graph response status: ${res.statusCode}`);
            if (res.statusCode === 302 || res.statusCode === 301) {
                const location = res.headers?.location || null;
                console.log(`[Recording] ✅ Got redirect location: ${location ? "yes" : "none"}`);
                resolve(location);
            } else {
                console.log(`[Recording] ⚠️  Unexpected status ${res.statusCode} — no redirect.`);
                resolve(null);
            }
            res.resume(); // Consume response to free memory
        });

        req.on("error", (err) => {
            console.error(`[Recording] ❌ https.request error:`, err.message);
            reject(err);
        });

        req.end();
    });
}

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const ORGANIZER_ID = process.env.ORGANIZER_ID; // The user whose onlineMeeting was created
const BACKEND_URL = process.env.BACKEND_URL || "http://localhost:8000";

/**
 * Polls Graph for the recording of a completed meeting.
 * Retries up to `maxRetries` times, waiting `retryDelayMs` between each attempt.
 *
 * @param {string} meetingId    - The Graph onlineMeeting ID.
 * @param {number} maxRetries   - Max number of polling attempts (default: 20).
 * @param {number} retryDelayMs - Milliseconds between retries (default: 30s).
 * @returns {Promise<object|null>} - Recording metadata object, or null if not found.
 */
async function fetchRecordingAfterCall(meetingId, maxRetries = 20, retryDelayMs = 30_000) {
    if (!meetingId) {
        console.warn("[Recording] ⚠️  No meetingId provided — cannot fetch recording.");
        return null;
    }

    if (!ORGANIZER_ID) {
        console.error("[Recording] ❌ ORGANIZER_ID env variable not set.");
        return null;
    }

    console.log(`\n[Recording] 🎬 Starting recording poll for meetingId: ${meetingId}`);
    console.log(`[Recording]    Will retry up to ${maxRetries} times every ${retryDelayMs / 1000}s`);

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
        try {
            const token = await getGraphToken();

            const response = await axios.get(
                `${GRAPH_BASE}/users/${ORGANIZER_ID}/onlineMeetings/${meetingId}/recordings`,
                {
                    headers: { Authorization: `Bearer ${token}` }
                }
            );

            const recordings = response.data?.value || [];

            if (recordings.length > 0) {
                const recording = recordings[0]; // Latest recording
                console.log(`\n[Recording] ✅ Recording available! (attempt ${attempt}/${maxRetries})`);
                console.log(`[Recording]    Recording ID  : ${recording.id}`);
                console.log(`[Recording]    Created At    : ${recording.createdDateTime}`);

                // The content URL gives access to the actual MP4 file
                const contentUrl = await fetchRecordingContentUrl(meetingId, recording.id);
                if (contentUrl) {
                    console.log(`[Recording]    Download URL  : ${contentUrl}`);
                }

                return {
                    recordingId: recording.id,
                    meetingId,
                    createdDateTime: recording.createdDateTime,
                    contentUrl
                };
            }

            console.log(`[Recording] ⏳ Attempt ${attempt}/${maxRetries} — recording not ready yet. Waiting ${retryDelayMs / 1000}s...`);

        } catch (error) {
            const status = error.response?.status;
            const msg = error.response?.data?.error?.message || error.message;

            // 404 means the meeting has no recordings yet — keep polling
            if (status === 404) {
                console.log(`[Recording] ⏳ Attempt ${attempt}/${maxRetries} — no recordings yet (404). Waiting...`);
            } else {
                console.error(`[Recording] ❌ Attempt ${attempt}/${maxRetries} — Graph error ${status}: ${msg}`);
            }
        }

        if (attempt < maxRetries) {
            await delay(retryDelayMs);
        }
    }

    console.error(`[Recording] ❌ Recording not found after ${maxRetries} attempts. Giving up.`);
    return null;
}

/**
 * Fetches the content (download) URL for a specific recording.
 *
 * @param {string} meetingId   - Graph onlineMeeting ID.
 * @param {string} recordingId - ID of the recording resource.
 * @returns {Promise<string|null>} - Content URL or null.
 */
async function fetchRecordingContentUrl(meetingId, recordingId) {
    // Return the Graph content URL directly.
    // downloadRecordingFile() will handle the 302 redirect to SharePoint automatically.
    const contentUrl = `${GRAPH_BASE}/users/${ORGANIZER_ID}/onlineMeetings/${meetingId}/recordings/${recordingId}/content`;
    console.log(`[Recording]    Content URL resolved (Graph endpoint): ${contentUrl}`);
    return contentUrl;
}

/**
 * Downloads the recording MP4 from Microsoft Graph to the local server.
 *
 * @param {string} contentUrl - The temporary download URL from Graph.
 * @param {string} filename   - The name to save the file as.
 * @returns {Promise<string|null>} - The local path to the saved file.
 */
async function downloadRecordingFile(contentUrl, filename) {
    try {
        const downloadDir = path.join(__dirname, "..", "public", "recordings");

        if (!fs.existsSync(downloadDir)) {
            fs.mkdirSync(downloadDir, { recursive: true });
        }

        const localPath = path.join(downloadDir, filename);
        console.log(`[Recording] 📥 Downloading recording to: ${localPath}`);

        // Graph API can behave in two ways for the /content endpoint:
        //   200 OK  → Graph streams the file directly (app-only permissions)
        //   302     → Graph redirects to a SharePoint pre-signed URL
        //
        // We use Node's native https module to check which case we are in,
        // then handle each case with the correct auth strategy.

        if (contentUrl.includes("graph.microsoft.com")) {
            const token = await getGraphToken();
            console.log(`[Recording] 🔍 Checking Graph response type (200 or 302)...`);

            const graphStatus = await getGraphResponseStatus(contentUrl, token);
            console.log(`[Recording]    Graph status: ${graphStatus.status}`);

            if (graphStatus.status === 302 && graphStatus.location) {
                // Case A: Graph redirects → download from SharePoint WITHOUT token
                console.log(`[Recording] ↪️  302 redirect — downloading from SharePoint URL...`);
                const response = await axios({ method: "get", url: graphStatus.location, responseType: "stream" });
                await pipeline(response.data, fs.createWriteStream(localPath));

            } else if (graphStatus.status === 200) {
                // Case B: Graph streams directly → download from Graph WITH token
                console.log(`[Recording] ⬇️  200 OK — downloading directly from Graph with token...`);
                const response = await axios({
                    method: "get",
                    url: contentUrl,
                    responseType: "stream",
                    headers: { Authorization: `Bearer ${token}` }
                });
                await pipeline(response.data, fs.createWriteStream(localPath));

            } else {
                throw new Error(`Unexpected Graph response: status ${graphStatus.status}`);
            }

        } else {
            // Non-Graph URL — download directly without any token
            const response = await axios({ method: "get", url: contentUrl, responseType: "stream" });
            await pipeline(response.data, fs.createWriteStream(localPath));
        }

        console.log(`[Recording] ✅ Download complete: ${filename}`);
        return localPath;

    } catch (error) {
        console.error(`[Recording] ❌ Download failed:`, error.message);
        return null;
    }
}

/**
 * Makes a HEAD-like GET request (stops at first response) to detect
 * whether Graph returns 200 (direct stream) or 302 (redirect to SharePoint).
 *
 * @param {string} url   - Graph content URL.
 * @param {string} token - Bearer token.
 * @returns {Promise<{status: number, location: string|null}>}
 */
function getGraphResponseStatus(url, token) {
    return new Promise((resolve, reject) => {
        const parsedUrl = new URL(url);
        const options = {
            hostname: parsedUrl.hostname,
            path: parsedUrl.pathname + parsedUrl.search,
            method: "GET",
            headers: { Authorization: `Bearer ${token}` }
        };

        const req = https.request(options, (res) => {
            const location = res.headers?.location || null;
            resolve({ status: res.statusCode, location });
            res.destroy(); // Stop reading body — we only needed the status/headers
        });

        req.on("error", (err) => {
            console.error(`[Recording] ❌ getGraphResponseStatus error:`, err.message);
            reject(err);
        });

        req.end();
    });
}


/**
 * Notifies the Python backend that a recording is ready for AI analysis.
 *
 * @param {object} candidateDetails - { name, email, threadId }
 * @param {object} recordingData   - { recordingId, localPath, contentUrl }
 */
async function notifyBackendRecordingReady(candidateDetails, recordingData) {
    try {
        const payload = {
            candidate_name: candidateDetails.name,
            candidate_email: candidateDetails.email,
            questions: candidateDetails.questions, 
            thread_id: candidateDetails.threadId,
            recording_id: recordingData.recordingId,
            local_path: recordingData.localPath,
            content_url: recordingData.contentUrl
        };

        const baseUrl = (BACKEND_URL).replace(/\/+$/, '');
        const targetUrl = `${baseUrl}/api/recording-ready`;

        console.log(`[Recording] 🔔 Notifying backend for AI analysis at URL: ${targetUrl}`);

        const response = await axios.post(targetUrl, payload);

        if (response.status === 200) {
            console.log(`[Recording] ✅ Backend notified successfully.`);
            console.log(`\n======================================================`);
            console.log(`📄 AI ATS SCORECARD FOR FOR ${candidateDetails.name}:`);
            console.log(`======================================================\n`);
            console.log(response.data.ats_scorecard, { depth: null });
            console.log(`\n======================================================\n`);
        } else {
            console.warn(`[Recording] ⚠️ Backend returned status ${response.status}: ${response.data}`);
        }

    } catch (error) {
        console.error(`[Recording] ❌ Failed to notify backend at ${error.config?.url}:`, error.response?.data || error.message);
    }
}

function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

module.exports = {
    fetchRecordingAfterCall,
    downloadRecordingFile,
    notifyBackendRecordingReady
};
