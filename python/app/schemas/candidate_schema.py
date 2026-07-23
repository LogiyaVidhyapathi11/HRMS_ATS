from pydantic import BaseModel, Field, EmailStr
from typing import List, Optional, Dict
from enum import Enum
class QuestionType(str, Enum):
    introductory = "Introductory"
    basic = "Basic"
    technical = "Technical"
class QuestionModel(BaseModel):
    text: str = Field(description = "The exact interview question to ask the candidate directly. Must be an interrogative sentence ending in a question mark. Do NOT output a summary statement.")
    type: QuestionType
class GeminiOutputModel(BaseModel):
    candidate_name: str = Field(description = "The full name of the candidate extracted from the resume")
    candidate_email: Optional[EmailStr] = Field(description = "The primary email address of the candidate extracted from the resume. If not found in the resume, you must return null.")
    questions: List[QuestionModel] = Field(
        description = "A list of exactly 3 generated interview questions."
    )
class CandidatePrepareRequest(BaseModel):
    resume_text: str
    job_description_text: str
class InterviewResponseModel(BaseModel):
    candidate_name: str
    candidate_email: str
    resume_text: str
    job_description_text: str
    questions: List[QuestionModel]
    meeting_link: str
    event_id: str
    start_time: str
    end_time: str
# ─── PROCTORING SCHEMA ───
class ProctoringLogModel(BaseModel):
    timestamp: str = Field(description="MM:SS formatted timestamp when the proctoring flag occurred")
    event: str = Field(description="Description of the event (e.g. 'Unauthorised Object (Cell Phone) detected')")
    severity: str = Field(description="Severity of the event: 'low', 'medium', or 'high'")
    confidence: float = Field(description="Confidence of the detection model (0.0 to 1.0)")
    image_url: Optional[str] = Field(default=None, description="Relative URL of the captured screenshot showing the malpractice incident")
class AudioAnalysisModel(BaseModel):
    """Audio-track analysis results extracted from the interview video."""
    multiple_voices_count: int = Field(default=0, description="Number of times more than one speaker voice was detected")
    long_silence_count: int = Field(default=0, description="Number of silence gaps longer than 5 seconds detected")
    background_noise_level: str = Field(default="Low", description="Overall background noise assessment: 'Low', 'Medium', or 'High'")
class DetectionSummaryItemModel(BaseModel):
    """One row in the AI Proctoring Summary detection table."""
    status: str = Field(description="Human-readable status label, e.g. 'Stable', 'Detected', 'None', 'Moderate', 'Frequent'")
    detail: str = Field(description="Supplementary detail string, e.g. '99.2%', '2 Events', '17 sec'")
    severity: str = Field(description="Color-coding severity level: 'none', 'low', 'medium', 'high', 'normal'")
class RiskBreakdownItemModel(BaseModel):
    """Per-category integrity risk score used for the horizontal bar chart."""
    score: int = Field(description="Earned score for this category")
    max: int = Field(description="Maximum possible score for this category")
class ProctoringAnalysisModel(BaseModel):
    cheating_score: int = Field(description="Overall integrity violation/cheating score from 0 (clean) to 100 (high flag count)", ge=0, le=100)
    gaze_away_seconds: int = Field(description="Total seconds candidate spent looking away from the screen", ge=0)
    no_face_seconds: int = Field(description="Total seconds candidate face was not visible", ge=0)
    multi_face_seconds: int = Field(description="Total seconds multiple people were in frame", ge=0)
    phone_detected_seconds: int = Field(description="Total seconds candidate had a cell phone visible", ge=0)
    book_detected_seconds: int = Field(description="Total seconds candidate had a book visible", ge=0)
    total_sampled_seconds: int = Field(description="Total video duration in seconds analyzed", ge=0)
    logs: List[ProctoringLogModel] = Field(default=[], description="Chronological log of proctoring events")
    # ─── Extended fields for the Proctoring Report Dashboard ───
    attention_score: int = Field(default=0, description="Percentage of time candidate face was visible and engaged (0-100)", ge=0, le=100)
    ai_confidence: float = Field(default=0.0, description="Weighted average confidence of all AI detections across the session (0.0-100.0)")
    candidate_presence_pct: float = Field(default=0.0, description="Percentage of video frames in which the candidate face was detected", ge=0.0, le=100.0)
    head_turn_seconds: int = Field(default=0, description="Total seconds candidate's head was turned more than 45 degrees away", ge=0)
    camera_blocked_seconds: int = Field(default=0, description="Total seconds the camera appeared blocked (very dark frame)", ge=0)
    integrity_score: int = Field(default=0, description="Composite integrity score calculated as the sum of all risk_breakdown sub-scores (0-100)", ge=0, le=100)
    detection_summary: Optional[Dict[str, DetectionSummaryItemModel]] = Field(
        default=None,
        description="Structured detection rows for the AI Proctoring Summary table"
    )
    risk_breakdown: Optional[Dict[str, RiskBreakdownItemModel]] = Field(
        default=None,
        description="Per-category integrity scores for the Integrity Factors bar chart"
    )
    audio_analysis: Optional[AudioAnalysisModel] = Field(
        default=None,
        description="Audio track analysis results"
    )
# ─── ATS SCORECARD SCHEMA ───
class SkillAssessmentModel(BaseModel):
    skill: str = Field(description="Name of the skill assessed (e.g. 'Python Programming')")
    rating: int = Field(description="Rating from 1 (low) to 10 (expert)", ge=1, le=10)
    comment: str = Field(description="Brief assessment comment explaining the rating")
class ATSScorecardModel(BaseModel):
    overall_score: int = Field(description="Overall match and evaluation score from 0 to 100", ge=0, le=100)
    technical_rating: int = Field(description="Technical capabilities score from 1 to 10", ge=1, le=10)
    communication_rating: int = Field(description="Communication clarity score from 1 to 10", ge=1, le=10)
    culture_fit_rating: int = Field(description="Organizational and culture fit rating from 1 to 10", ge=1, le=10)
    skills_assessment: List[SkillAssessmentModel] = Field(description="Breakdown of specific skill match evaluations")
    strengths: List[str] = Field(description="List of primary strengths observed")
    weaknesses: List[str] = Field(description="List of growth areas or weaknesses noted")
    recommendation: str = Field(description="Hiring decision recommendation. Allowed: 'Strong Hire', 'Hire', 'Proceed with Caution', 'No Hire'")
    detailed_feedback: str = Field(description="Comprehensive technical and professional feedback summary")
    integrity_check_commentary: str = Field(description="Hiring manager's assessment of candidate integrity and any proctoring flags")
    # ─── Recruiter Summary fields (for the dashboard panel) ───
    strengths_summary: List[str] = Field(
        default=[],
        description="Concise recruiter-facing strength bullets (max 6). Short, direct statements like 'Good communication skills', 'Answered most technical questions'."
    )
    concerns_summary: List[str] = Field(
        default=[],
        description="Concise recruiter-facing concern bullets (max 6). Short statements about proctoring flags or performance gaps like 'Frequent eye diversion', 'Mobile phone detected twice'."
    )
    integrity_check_commentary: str = Field(description="Hiring manager's assessment of candidate integrity and any proctoring flags")
    overall_summary_points: List[str] = Field(
        default=[],
        description="Exactly 5 concise, direct summary bullet points of the overall feedback (max 12 words each) to show in the AI Interview Feedback dashboard card."
    )
    recommendation_reasons: List[str] = Field(
        default = [], 
        description = "Exactly 4 concise, direct bullet points explaining the recommendation decision (max 12 words each) to show in the AI Recommendation dashboard card."
    )