/*
* Azure Cognitive Services Speech SDK - Text-to-Speech synthesis.

* Converts a plain-text string into a .wav audio file using Azure's

* neural TTS voices. The synthesized file is saved to the local filesystem.

* So it can be served as a public URL for Graph API recordResponse.

* Required env vars:
- AZURE_SPEECH_KEY - Azure Cognitive Services Speech Key
- AZURE_SPEECH_REGION - Azure region (e.g. "eastus", "centralindia")
*/


const sdk = require("microsoft-cognitiveservices-speech-sdk");
const path = require("path");
require("dotenv").config()

/*
* Synthesizes `text` to a WAV audio file at `outputFilePath`.
* @param {string} text - The question text to speech
* @param {string} outputFilePath - Absolute path for the output .wav file.
* @returns {Promise<void>}
*/

const synthesizeToFile = async (text, filePath) => {
    return new Promise((resolve, reject) => {
        const speechKey = process.env.AZURE_SPEECH_KEY;
        const speechRegion = process.env.AZURE_SPEECH_REGION;

        if (!speechKey || !speechRegion) {
            return reject(new Error(
                "Azure Speech credentials missing. Set AZURE_SPEECH_KEY and AZURE_SPEECH_REGION in .env"
            ));
        }


        // Build Speech Config
        const speechConfig = sdk.SpeechConfig.fromSubscription(speechKey, speechRegion);

        // Use a warm, professional neural voice (en-US-JennyNeural)
        const voiceName = "en-US-JennyNeural";
        speechConfig.speechSynthesisVoiceName = voiceName;

        // Output as 16kHz 16-bit mono WAV - compatible with Graph recordResponse
        speechConfig.speechSynthesisOutputFormat = sdk.SpeechSynthesisOutputFormat.Riff16khz16BitMonoPcm;

        // Output to file
        const audioConfig = sdk.AudioConfig.fromAudioFileOutput(filePath);

        const synthesizer = new sdk.SpeechSynthesizer(speechConfig, audioConfig);

        console.log(`[Azure TTS] Synthesizing: "${text.substring(0, 60)}..."`)

        const escapeXml = (unsafe) => unsafe.replace(/[<>&'"]/g, (c) => {
            switch(c) {
                case '<': return '&lt;';
                case '>': return '&gt;';
                case '&': return '&amp;';
                case '\'': return '&apos;';
                case '"': return '&quot;';
            }
        });

        const safeText = escapeXml(text)

        // Wrap the text in SSML to reduce the speaking rate by 15%
        const ssml = `
        <speak version = "1.0" xmlns = "http://www.w3.org/2003/10/synthesis" xml:lang="en-US">
            <voice name="${voiceName}">
                <prosody rate="-12%">
                    ${safeText}
                </prosody>
            </voice>
        </speak>
        `

        synthesizer.speakTextAsync(
            text, 
            (result) => {
                synthesizer.close();

                if (result.reason === sdk.ResultReason.SynthesizingAudioCompleted) {
                    console.log(`[Azure TTS] Audio saved -> ${path.basename(filePath)}`);
                    resolve();
                } else {
                    const err = result.errorDetails || "Unknown synthesis error";
                    console.error(`[Azure TTS] Synthesis failed: ${err}`);
                    reject(new Error(`Azure TTS synthesis failed: ${err}`));
                }
            }, 
            (error) => {
                synthesizer.close();
                console.error(`[Azure TTS] Exception:`, error);
                reject(new Error(`Azure TTS exception: ${error}`));
            }
        );
    });
}

module.exports = { synthesizeToFile };
