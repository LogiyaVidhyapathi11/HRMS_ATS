/**
 * ttsService.js
 * ─────────────────────────────────────────────────────────────────────────────
 * Thin wrapper over azureTTSService.
 *
 * Resolves the output path to `public/audio/<filename>` relative to the
 * nodejs project root, then delegates synthesis to azureTTSService.
 *
 * Usage (called from botMessages.js background task):
 *   await generateQuestionAudio("Tell me about yourself?", "thread123_q1.wav");
 * ─────────────────────────────────────────────────────────────────────────────
 */
const path = require("path");
const fs   = require("fs");
const { synthesizeToFile } = require("./azureTTSService");
// Absolute path to the public/audio directory (served as static files)
const AUDIO_DIR = path.join(__dirname, "..", "public", "audio");
/**
 * Generates a WAV audio file for a single interview question.
 *
 * @param {string} questionText - The question text to convert to speech.
 * @param {string} filename     - Output filename (e.g. "thread123_q1.wav").
 * @returns {Promise<string>}   - Absolute path of the generated audio file.
 */
async function generateQuestionAudio(questionText, filename) {
    // Ensure the audio directory exists
    if (!fs.existsSync(AUDIO_DIR)) {
        fs.mkdirSync(AUDIO_DIR, { recursive: true });
        console.log(`[TTS] Created audio directory: ${AUDIO_DIR}`);
    }
    const outputPath = path.join(AUDIO_DIR, filename);
    await synthesizeToFile(questionText, outputPath);
    console.log(`[TTS] Audio ready: ${filename}`);
    return outputPath;
}
module.exports = { generateQuestionAudio };