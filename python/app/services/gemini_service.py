import json
import asyncio
import requests
import os
import re
# import google.generativeai as genai
from google import genai 
from google.genai import types
from app.core.config import settings
from app.schemas.candidate_schema import GeminiOutputModel, ATSScorecardModel

# Initialize the modern Google GenAI SDK Client
client = genai.Client(api_key = settings.GEMINI_API_KEY)

model_name = "gemini-2.5-flash"

# ─────────────────────────────────────────────────────────────────────────────
# System Instruction for AI Interview Generation
# ─────────────────────────────────────────────────────────────────────────────
SYSTEM_PROMPT = """
You are an expert technical interviewer conducting a live first-round screening interview directly with the candidate.

Your task has two parts:
1. Extract the candidate's full name and primary email address from the provided Resume.
2. Generate EXACTLY 3 technical interview questions addressed DIRECTLY TO THE CANDIDATE.

CRITICAL RULE: 
1. Speak exactly as if you are talking to the candidate face-to-face. Use "you" and "your".
2. Do NOT ever use the candidate's name or refer to them in the third person.
3. CONCISENESS IS MANDATORY: Each question MUST be a single, short, and professional sentence.
4. NO MULTI-PART QUESTIONS: Do not ask two or more things in one prompt. Focus on one clear topic per question.
5. FIRST-ROUND SCREENING: Keep questions high-level but professional, evaluating core fit and primary technical achievements.
6. TECHNICAL ONLY: Do NOT generate introductory or 'walk through your background' questions. Focus purely on technical evaluation.

Structure your 3 technical questions logically:
- Question 1 (Technical): Evaluate a core technical skill or experience mentioned in the resume and job description.
- Question 2 (Technical): Evaluate a specific project achievement or technical challenge.
- Question 3 (Technical): Evaluate problem-solving capabilities or architectural/production concepts.

JSON FORMAT REQUIREMENT:
{
  "candidate_name": "Full Name",
  "candidate_email": "email@example.com",
  "questions": [
    {"text": "Your first concise technical question here?", "type": "Technical"}, 
    {"text": "Your second concise technical question here?", "type": "Technical"}, 
    {"text": "Your third concise technical question here?", "type": "Technical"}
  ]
}
"""

async def generate_interview_questions(resume_text: str, jd_text: str) -> GeminiOutputModel:
    """
        Calls Gemini using the modern google-genai SDK to extract candidate info and generate 3 questions. Force strcutured JSON matching GeminiOutputModel.
    """

    prompt_content = f"Resume:\n{resume_text}\n\nJob Description:\n{jd_text}"

    try:
        loop = asyncio.get_event_loop()

        # Call the new Client API (using executor to keep it non-blocking)

        response = await loop.run_in_executor(
            None, 
            lambda: client.models.generate_content(
                model = model_name,
                contents = prompt_content, 
                config = types.GenerateContentConfig(
                    system_instruction = SYSTEM_PROMPT, 
                    response_mime_type = "application/json", 
                    response_schema = GeminiOutputModel
                )
            )
        )

        # Fallback: manually parse the raw JSON text
        data = json.loads(response.text)
        print('Gemini raw response type:', type(data).__name__)

        # Gemini sometimes wraps the object in a list - unwrap it
        if isinstance(data, list):
            if len(data) == 0:
                raise ValueError("Gemini returned an empty list")
            data = data[0]

        # ─────────────────────────────────────────────────────────────────────
        # Robust Unwrapping Logic (Handles nested/wrapped response shapes)
        # ─────────────────────────────────────────────────────────────────────
        candidate_name = data.get("candidate_name") or \
                         data.get("candidate_details", {}).get("name") or \
                         data.get("candidate", {}).get("name")

        candidate_email = data.get("candidate_email") or \
                          data.get("candidate_details", {}).get("email") or \
                          data.get("candidate", {}).get("email")

        questions = data.get("questions") or \
                    data.get("interview_questions") or \
                    data.get("job_match", {}).get("questions") or []

        # Map types inside questions safely
        formatted_questions = []
        for q in questions:
            if isinstance(q, str):
                formatted_questions.append({"text": q, "type": "Technical"})
            elif isinstance(q, dict):
                formatted_questions.append({
                    "text": q.get("text") or q.get("question") or "",
                    "type": q.get("type") or q.get("category") or "Technical"
                })

        # Prepend Introductory question as Question 0 (total = 1 Intro + 3 Technical)
        intro_q = {
            "text": "Hello, Welcome to your interview. To start with, please introduce yourself.",
            "type": "Introductory"
        }
        formatted_questions.insert(0, intro_q)

        # Reconstruct clean data structure matching GeminiOutputModel
        final_data = {
            "candidate_name": candidate_name,
            "candidate_email": candidate_email,
            "questions": formatted_questions
        }

        return GeminiOutputModel(**final_data)

    except Exception as e:
        print(f"Gemini SDK Question Generation Error: {e}")
        if 'response' in locals() and hasattr(response, 'text'):
            print(f"[DEBUG] AI RETURNED: {response.text}")
        raise e

# System Instruction for Candidate Feedback Analysis and ATS Scorecard Generation

FEEDBACK_SYSTEM_PROMPT = """
    You are a senior hiring manager, technical assessor, and talent acquisition expert.

    Your task is to analyze:
    1. The attached video recording of the candidate's interview.
    2. The specific Interview Questions asked.
    3. The server-side Proctoring Analysis metrics and incident logs.

    Please generate a structured, highly objective ATS Scorecard. Evaluate the candidate's performance in the video:

    - Rate their technical proficiency, communication skills, and organization/culture fit.
    - Match candidate responses against the requested questions to assess specific skills.
    - List observed strengths and growth areas/weaknesses.
    - Give an overall evaluation score (0-100) and recommendation decision ('Strong Hire', 'Hire', 'Proceed with Caution', 'No Hire'.).
    - Critically evaluate candidate integrity and the provided proctoring flags. Verify whether the flags (e.g. looking away, phone detected) correspond to actual integrity violations in the video or were false alrams. Document this in 'integrity_check_commentary'.

    ADDITIONAL FIELDS REQUIRED:
    - strengths_summary: Generate 4-6 very short, direct bullet-point sentences (max 8 words each) for the Recruiter Summary panel. These should highlight what impressed you about the candidate. Examples: 'Good communication skills', 'Answered most technical questions', 'No multiple person detected', 'Overall presence was stable'.
    - concerns_summary: Generate 4-6 very short, direct concern bullet-points (max 8 words each) for the Recruiter Summary panel. These should capture integrity flags AND performance gaps. Examples: 'Frequent eye diversion observed', 'Mobile phone detected twice', 'Candidate absent for 17 seconds', 'Background voice detected once'.
    - overall_summary_points: Generate EXACTLY 5 concise, direct summary bullet-point sentences (max 12 words each) summarizing the candidate's technical performance, overall performance for the AI Interview Feedback panel.
    - recommendation_reasons: Generate EXACTLY 4 concise, professional bullet-points sentences (maximum of 10 words each) explaining the technical/integrity grounds for the hiring recommendation decision ('Strong Hire', 'Hire', 'Proceed with Caution', 'No Hire'). Keep them short, concise and clean.

    Tone: Professional, constructive, and formal.
    Language: English.

"""

async def generate_interview_feedback(candidate_data: dict, proctoring_results: dict = None) -> ATSScorecardModel:

    """
        Generates a professional structure AI feedback report and ATS Scorecard by analyzing the interview MP4 video and local proctoring results.
    """

    local_path = candidate_data.get('local_path')
    candidate_name = candidate_data.get('candidate_name', 'Unknown')
    questions = candidate_data.get('questions', [])

    if not local_path or not os.path.exists(local_path):
        print(f"[Gemini] Recording file not found at: {local_path}")
        return get_fallback_scorecard(candidate_name, "Error: The recording file could not be found for analysis.")
    
    print(f"[Gemini] Uploading video to Google AI servers: {local_path}")

    try:
        # 1. Upload the video file (using executor to avoid blocking the event loop)

        loop = asyncio.get_event_loop()

        # Upload using the new SDK client files API
        video_file = await loop.run_in_executor(None, lambda: client.files.upload(file = local_path))
        print(f"[Gemini] Video uploaded ({video_file.name}). Waiting for Google to process frames.")

        # 2. Poll until the video is ACTIVE
        file_info = await loop.run_in_executor(None, lambda: client.files.get(name = video_file.name))

        while file_info.state.name == 'PROCESSING':
            await asyncio.sleep(5)
            file_info = await loop.run_in_executor(None, lambda: client.files.get(name = video_file.name))

        if file_info.state.name == "FAILED":
            raise ValueError("Gemini video processing failed in the cloud")
        
        print(f"\n[Gemini] Video is ACTIVE. Starting multimodal scorecard analysis with {model_name}.")

        proctoring_summary = ""
        if proctoring_results:
            proctoring_summary = f"""
            Proctoring Analysis Metrics:
            - Cheating Score (0-100): {proctoring_results.get('cheating_score', 0)}
            - Gaze Away Time: {proctoring_results.get('gaze_away_seconds', 0)} seconds
            - No Face Time: {proctoring_results.get('no_face_seconds', 0)} seconds
            - Multiple People In Frame Time: {proctoring_results.get('multi_face_seconds', 0)} seconds
            - Cell Phone Detected Time: {proctoring_results.get('phone_detected_seconds', 0)} seconds
            - Book/Cheating Material Detected Time: {proctoring_results.get('book_detected_seconds', 0)} seconds
            - Total Sampled Time: {proctoring_results.get('total_sampled_seconds', 0)} seconds

            Incidents Log:
            {json.dumps(proctoring_results.get('logs', []), indent=2)}
        """

        prompt = (
            f"Candidate Name: {candidate_name}\n"
            f"Questions asked during the interview:\n{json.dumps(questions, indent = 2)}\n\n"
            f"{proctoring_summary}\n\n"
            f"Generate the structured ATS Scorecard based on the candidate's response video and these proctoring metrics."
        )

        # Call generate_content on client using the file object and schema configuration
        response = await loop.run_in_executor(
            None, 
            lambda: client.models.generate_content(
                model = model_name, 
                contents = [video_file, prompt], 
                config = types.GenerateContentConfig(
                    system_instruction = FEEDBACK_SYSTEM_PROMPT,
                    response_mime_type = "application/json",
                    response_schema = ATSScorecardModel
                )
            )
        )
        
        # Parse the structured JSON response
        data = json.loads(response.text)
        if isinstance(data, list):
            if len(data) == 0:
                raise ValueError("Gemini returned empty list")
            data = data[0]

        # Convert to ATSScorecardModel Pydantic model
        scorecard = ATSScorecardModel(**data)

        # 4. Mandatory Privacy Cleanup: Delete the video from Google's servers
        print(f"[Gemini] Deleting video from Google servers to protect candidate privacy.")
        
        await loop.run_in_executor(None, lambda: client.files.delete(name = video_file.name))

        return scorecard

    except Exception as e:
        print(f"Gemini Native Analysis Error: {e}")
        # Clean up video file if it was uploaded and we encountered an error
        try:
            if 'video_file' in locals() and video_file:
                await loop.run_in_executor(None, lambda: client.files.delete(name = video_file.name))
        except Exception:
            pass
        return get_fallback_scorecard(candidate_name, f"AI was unable to generate a feedback report. Error: {e}")


def get_fallback_scorecard(candidate_name: str, error_msg: str) -> ATSScorecardModel:
    """Helper to return a placeholder scorecard in case of errors."""
    return ATSScorecardModel(
        overall_score=0,
        technical_rating=1,
        communication_rating=1,
        culture_fit_rating=1,
        skills_assessment=[
            {"skill": "General Evaluation", "rating": 1, "comment": "Evaluation failed due to service error."}
        ],
        strengths=["Service fallback enabled"],
        weaknesses=["Could not process recording"],
        recommendation="Proceed with Caution",
        detailed_feedback=f"Analysis failed. Details: {error_msg}",
        integrity_check_commentary="Integrity check could not be completed.", 
        strengths_summary = ["Analysis pending"], 
        concerns_summary = ["Recording could not be processed"], 
        overall_summary_points=[], 
        recommendation_reasons = []
    )
