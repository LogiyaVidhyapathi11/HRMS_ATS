from fastapi import APIRouter, UploadFile, File, HTTPException, status
from fastapi.responses import RedirectResponse, HTMLResponse
from datetime import datetime, timedelta
from app.core.config import settings
from app.services.auth_service import get_graph_token
from app.services.interviewer_scheduler import schedule_interview
from app.services.gemini_service import generate_interview_questions, generate_interview_feedback
from app.utils.file_parser import extract_text_from_file
from app.schemas.candidate_schema import InterviewResponseModel
import requests
import base64
from pydantic import BaseModel
from app.db.database import candidates_collection
from app.services.proctoring_service import ProctoringService
import os
import json
import shutil
import re

# Initialize the proctoring service globally to shared loaded models
proctoring_service = ProctoringService()
 
router = APIRouter()

class RecordingReadyRequest(BaseModel):
    candidate_name: str
    candidate_email: str
    questions: list = []
    thread_id: str
    recording_id: str
    local_path: str
    content_url: str

@router.get("/debug-incidents")
async def debug_incidents():
    import os
    current_dir = os.path.dirname(os.path.abspath(__file__))
    backend_root = os.path.abspath(os.path.join(current_dir, "..", "..", ".."))
    res = {}
    for candidate in ["teams-bot", "nodejs", "node", "bot"]:
        path = os.path.join(backend_root, candidate, "public", "recordings", "incidents")
        exists = os.path.exists(path)
        files = os.listdir(path) if exists else []
        res[candidate] = {
            "path": path,
            "exists": exists,
            "file_count": len(files),
            "files": files[:30]  # list first 30 files
        }
    # Check what is inside candidate collection
    return res


@router.post("/recording-ready")
async def recording_ready(data: RecordingReadyRequest):
    """
        Called by the Node.js bot when a recording is downloaded and ready for analysis.
        Triggers proctoring analysis and AI feedback generation.
    """

    print(f"[Backend] Recording ready for {data.candidate_name} ({data.candidate_email})")

    try:
        # Redirect local path to teams-bot folder if it points to nodejs
        local_path = data.local_path
        if "/nodejs/" in local_path:
            target_path = local_path.replace("/nodejs/", "/teams-bot/")
            target_dir = os.path.dirname(target_path)
            # Ensure target directory exists
            os.makedirs(target_dir, exist_ok=True)
            # If the source file exists, move it to the target directory
            if os.path.exists(local_path) and not os.path.exists(target_path):
                try:
                    shutil.move(local_path, target_path)
                    print(f"[Backend] Automatically moved recording from nodejs to teams-bot: {target_path}")
                except Exception as e:
                    print(f"[Backend] Failed to move recording file: {e}")
            if os.path.exists(target_path):
                local_path = target_path

        # 1. Run local proctoring analysis
        print(f"[Backend] Running Proctoring Analysis for candidate.")

        proctoring_results = proctoring_service.analyze_video(local_path)

        # 2. Generate ATS Scorecard using Gemini
        print(f"[Backend] Generating AI ATS Scorecard feedback.")
        scorecard = await generate_interview_feedback({
            "candidate_name": data.candidate_name, 
            "candidate_email": data.candidate_email, 
            "questions": data.questions, 
            "local_path": local_path
        }, proctoring_results)

        print(f"[Backend] AI Scorecard generated successfully.")

        # 3. Update MongoDB with the proctoring results and the feedback report
        await candidates_collection.update_one(
            {"candidate_email": data.candidate_email}, 
            {
                "$set": { 
                    "status": "completed", 
                    "recording_local_path": local_path, 
                    "recording_content_url": data.content_url, 
                    "proctoring_analysis": proctoring_results, 
                    "ats_scorecard": scorecard.dict(), 
                    "ai_feedback_report": scorecard.detailed_feedback, # for backward compatibility
                    "questions": data.questions,  # Store actual questions asked (including intro greeting), 
                    "updated_at": datetime.utcnow()
                }
            }
        )

        print(f"[Backend] MongoDB updated with AI report for {data.candidate_name}")

        return {
            "status": "success", 
            "message": "Recording processed, proctoring analysis completed, and feedback scorecard generated.", 
            "ats_scorecard": scorecard.dict(),
            "ai_feedback_report": scorecard.detailed_feedback
        }
    
    except Exception as e:
        print(f"[Backend] Error processing recording: {e}")
        raise HTTPException(status_code = 500, detail = str(e))

@router.get("/join-redirect")
def join_redirect(start_time: int, url: str = None, meeting_id: str = None):
    """
        Validates the 10-minute expiry window and redirects to the Teams meeting.
        Supports either a base64 encoded 'url' or a 'meeting_id' for lookup.
    """

    decoded_url = None

    if meeting_id:
        try: 
            print(f"[Redirect] Resolving Teams URL for meeting_id {meeting_id}")
            token = get_graph_token()
            headers = {"Authorization": f"Bearer {token}"}
            organizer = settings.TEAMS_ORGANIZER_EMAIL

            # Fetch meeting details from Graph
            res = requests.get(
                f"https://graph.microsoft.com/v1.0/users/{organizer}/onlineMeetings/{meeting_id}", 
                headers = headers
            )

            if res.status_code == 200:
                decoded_url = res.json().get("joinWebUrl")
                print(f"[Redirect] Resolved to: {decoded_url}")
            else:
                print(f"[Redirect] Graph Lookup Failed ({res.status_code}): {res.text}")
        
        except Exception as e:
            print(f"[Redirect] Resolution error: {e}")

    # Fallback to base64 url if meeting_id failed or wasn't provided
    if not decoded_url and url:
        try:
            decoded_url = base64.b64decode(url).decode('utf-8')
        except Exception:
            decoded_url = url

    if not decoded_url:
        raise HTTPException(status_code = 400, detail = "Invalid meeting link components.")

    # 10-minute expiry check
    current_time = datetime.utcnow().timestamp()

    if current_time > (start_time + 600): # 600 seconds = 10 minutes
        return HTMLResponse(
            content = 
            """
                <!DOCTYPE html>
                    <html lang="en">
                        <head>
                            <meta charset="UTF-8">
                            <meta name="viewport" content="width=device-width, initial-scale=1.0">
                            <title>Interview Link Expired</title>
                            <style>
                                body {
                                    font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
                                    background-color: #f8fafc;
                                    display: flex;
                                    align-items: center;
                                    justify-content: center;
                                    height: 100vh;
                                    margin: 0;
                                    color: #1e293b;
                                }
                                .container {
                                    background: white;
                                    padding: 3rem;
                                    border-radius: 16px;
                                    box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.1), 0 8px 10px -6px rgba(0, 0, 0, 0.1);
                                    max-width: 480px;
                                    width: 90%;
                                    text-align: center;
                                }
                                .icon {
                                    background-color: #fee2e2;
                                    color: #dc2626;
                                    width: 64px;
                                    height: 64px;
                                    border-radius: 50%;
                                    display: flex;
                                    align-items: center;
                                    justify-content: center;
                                    margin: 0 auto 1.5rem;
                                    font-size: 32px;
                                }
                                h1 {
                                    font-size: 1.5rem;
                                    font-weight: 700;
                                    margin-bottom: 1rem;
                                    color: #0f172a;
                                }
                                p {
                                    line-height: 1.6;
                                    color: #64748b;
                                    margin-bottom: 2rem;
                                }
                                .btn {
                                    display: inline-block;
                                    background-color: #2563eb;
                                    color: white;
                                    text-decoration: none;
                                    padding: 0.75rem 1.5rem;
                                    border-radius: 8px;
                                    font-weight: 600;
                                    transition: background-color 0.2s;
                                }
                                .btn:hover {
                                    background-color: #1d4ed8;
                                }
                                .footer {
                                    margin-top: 2rem;
                                    font-size: 0.875rem;
                                    color: #94a3b8;
                                }
                            </style>
                        </head>
                        <body>
                            <div class="container">
                                <div class="icon">⚠️</div>
                                <h1>Interview Link Expired</h1>
                                <p>
                                    For security reasons, this interview link is only valid for 10 minutes after the scheduled start time. 
                                    It appears this window has passed.
                                </p>
                                <p>Please contact your HR representative or the recruitment team to request a new interview link.</p>
                                <div class="footer">
                                    Adams Bridge AI Interview Platform
                                </div>
                            </div>
                        </body>
                    </html>
                """,
            status_code = 403
        )
    
    return RedirectResponse(url = decoded_url)
 
@router.post("/start-interview", response_model = InterviewResponseModel)
 
async def start_interview(
    resume_file: UploadFile = File(...),
    jd_file: UploadFile = File(...)
):
   
    """
    Full AI Interview Pipeline:
    1. Parses uploaded Resume and Job Description files.
    2. Gemini extracts candidate name, email and generates 10 interview questions.
    3. Schedules a Microsoft Teams meeting for the extracted candidate email.
    4. Returns the complete response including meeting link and questions.
    """
 
    try:
        # Step 1: Extract raw text from Uploaded Files
        resume_text = await extract_text_from_file(resume_file)
        jd_text = await extract_text_from_file(jd_file)
 
        # Step 2: Gemini extracts candidate name, email, and generate questions
        gemini_output = await generate_interview_questions(
            resume_text = resume_text,
            jd_text = jd_text
        )
 
        print(f"Gemini extracted - Name: {gemini_output.candidate_name}, Email: {gemini_output.candidate_email}")
 
        # Step 3: Use the dynamically extracted email to schedule the Teams meeting
        start_time = datetime.utcnow() + timedelta(minutes = 2)
        end_time = start_time + timedelta(minutes = 60)
 
        meeting = schedule_interview(
            candidate_name = gemini_output.candidate_name,
            candidate_email = str(gemini_output.candidate_email),
            start_time = start_time,
            end_time = end_time, 
            questions = [q.dict() for q in gemini_output.questions]
        )

        # Step 4: Save initial record to MongoDB
        candidate_record = {
            "candidate_name": gemini_output.candidate_name, 
            "candidate_email": str(gemini_output.candidate_email), 
            "resume_text": resume_text, 
            "job_description_text": jd_text, 
            "questions": [q.dict() for q in gemini_output.questions], 
            "meeting_link": meeting["meeting_link"], 
            "event_id": meeting["event_id"], 
            "meeting_id": meeting.get("meeting_id"), 
            "start_time": meeting["start_time"], 
            "end_time": meeting["end_time"], 
            "status": "scheduled", 
            "created_at": datetime.utcnow(), 
            "ai_feedback_report": None
        }

        await candidates_collection.insert_one(candidate_record)

        print(f"[Backend] Initial record saved to MongoDB for {gemini_output.candidate_name}.")
 
        # Step 5: Return full response
        return InterviewResponseModel(
            candidate_name = gemini_output.candidate_name,
            candidate_email = str(gemini_output.candidate_email),
            resume_text = resume_text,
            job_description_text = jd_text,
            questions = gemini_output.questions,
            meeting_link = meeting["meeting_link"],
            event_id = meeting["event_id"],
            start_time = meeting["start_time"],
            end_time = meeting["end_time"]
        )
   
    except Exception as e:
        print(f"Error in start_interview pipeline: {e}")
        raise HTTPException(
            status_code = status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail = str(e)
        )
 
 
# ─── DASHBOARD & SCORECARD ROUTES ───

@router.get("/candidates")
async def get_candidates():
    """
    Retrieves all candidates from MongoDB.
    """
    try:
        candidates = await candidates_collection.find().to_list(length=100)
        # Convert ObjectId and datetime objects for serialization
        for c in candidates:
            c["_id"] = str(c["_id"])
            if "created_at" in c and c["created_at"]:
                c["created_at"] = c["created_at"].isoformat()
            if "updated_at" in c and c["updated_at"]:
                c["updated_at"] = c["updated_at"].isoformat()
        return candidates
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/dashboard", response_class=HTMLResponse)
async def serve_dashboard():
    """
    Serves the HR candidate management dashboard matching the mockup UI exactly.
    Light-theme, full-width layout with stat cards, sparklines, and search/filter candidate table.
    """
    html_content = """
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Candidate Evaluation Dashboard — Adams Bridge</title>
        <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
        <style>
            :root {
                --bg-main: #F8FAFC;
                --bg-card: #FFFFFF;
                --border: #E2E8F0;
                --text: #1E293B;
                --text-muted: #64748B;
                --accent: #4F46E5;
                --success: #10B981;
                --warning: #F97316;
                --danger: #EF4444;
            }
            * { box-sizing: border-box; margin: 0; padding: 0; }
            body { font-family: 'Inter', sans-serif; background: var(--bg-main); color: var(--text); min-height: 100vh; padding: 2.5rem 5rem; }

            .container { max-width: 1200px; margin: 0 auto; display: flex; flex-direction: column; gap: 24px; }

            /* ── HEADER ──────────────────────────────── */
            header {
                display: flex; justify-content: space-between; align-items: center;
                padding-bottom: 20px; border-bottom: 1px solid var(--border);
                margin-bottom: 8px;
            }
            .header-left { display: flex; align-items: center; gap: 12px; }
            .logo-icon {
                width: 42px; height: 42px; border-radius: 12px;
                background: #4F46E5;
                display: flex; align-items: center; justify-content: center;
                font-weight: 800; font-size: 1.15rem; color: #fff;
                font-family: 'Outfit', sans-serif; flex-shrink: 0;
            }
            .header-titles h1 { font-family: 'Outfit', sans-serif; font-weight: 700; font-size: 1.6rem; color: #1E293B; }
            .header-titles p { font-size: 0.8rem; color: var(--text-muted); font-weight: 500; margin-top: 2px; }

            .header-right { display: flex; align-items: center; gap: 20px; }
            .bell-btn { background: none; border: none; cursor: pointer; position: relative; font-size: 1.35rem; color: #64748B; display: flex; align-items: center; }
            .bell-dot { position: absolute; top: 1px; right: 1px; width: 6px; height: 6px; background: #EF4444; border-radius: 50%; }
            
            .user-profile { display: flex; align-items: center; gap: 10px; cursor: pointer; }
            .user-avatar {
                width: 38px; height: 38px; border-radius: 50%;
                background: #EEF2FF; border: 1px solid #C7D2FE;
                display: flex; align-items: center; justify-content: center;
                font-weight: 700; font-size: 0.9rem; color: #4F46E5;
            }
            .user-info { display: flex; flex-direction: column; }
            .user-name { font-size: 0.82rem; font-weight: 600; color: #1E293B; }
            .user-role { font-size: 0.72rem; color: var(--text-muted); }
            .chevron-icon { font-size: 0.65rem; color: var(--text-muted); margin-left: 2px; }

            /* ── STATS CARDS ──────────────────────────── */
            .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; }
            .stat-card {
                background: var(--bg-card); border: 1px solid var(--border);
                border-radius: 16px; padding: 24px;
                display: flex; flex-direction: column; gap: 16px;
                box-shadow: 0 1px 3px rgba(0,0,0,0.02);
            }
            .stat-top-row { display: flex; align-items: center; gap: 16px; }
            .stat-icon-box {
                width: 48px; height: 48px; border-radius: 12px;
                display: flex; align-items: center; justify-content: center;
            }
            .stat-icon-box.purple { background: #EEF2FF; color: #4F46E5; }
            .stat-icon-box.green { background: #E6FDF4; color: #10B981; }
            .stat-icon-box.orange { background: #FFF7ED; color: #F97316; }
            .stat-icon-box.blue { background: #EFF6FF; color: #3B82F6; }

            .stat-title-val { display: flex; flex-direction: column; gap: 2px; }
            .stat-label { font-size: 0.82rem; color: #4B5563; font-weight: 500; }
            .stat-value { font-size: 1.8rem; font-weight: 800; font-family: 'Outfit', sans-serif; color: #111827; line-height: 1.1; }
            
            .stat-bottom-row { display: flex; justify-content: space-between; align-items: flex-end; margin-top: auto; }
            .stat-subtext { font-size: 0.72rem; color: #9CA3AF; font-weight: 500; padding-bottom: 2px; }
            
            .sparkline-container { width: 100px; height: 32px; }
            .sparkline-svg { width: 100%; height: 100%; overflow: visible; }

            /* ── TABLE CARD ──────────────────────────── */
            .tcard {
                background: var(--bg-card); border: 1px solid var(--border);
                border-radius: 16px; display: flex; flex-direction: column;
                box-shadow: 0 1px 3px rgba(0,0,0,0.02);
                overflow: hidden;
            }
            .tcard-hdr {
                padding: 20px 24px; border-bottom: 1px solid var(--border);
                display: flex; align-items: center; justify-content: space-between;
                flex-wrap: wrap; gap: 12px;
            }
            .tcard-title { font-family: 'Outfit', sans-serif; font-size: 1.05rem; font-weight: 700; color: #1E293B; }
            .tcard-tools { display: flex; align-items: center; gap: 8px; }
            
            .search-wrap { position: relative; width: 240px; }
            .search-wrap input {
                width: 100%; padding: 8px 12px 8px 32px;
                border: 1px solid var(--border); border-radius: 8px;
                font-size: 0.8rem; outline: none; transition: border-color 0.15s;
            }
            .search-wrap input:focus { border-color: #4F46E5; }
            .search-icon { position: absolute; left: 10px; top: 50%; transform: translateY(-50%); color: var(--text-muted); font-size: 0.8rem; }
            
            .btn-tool {
                background: #FFFFFF; border: 1px solid var(--border); border-radius: 8px;
                padding: 8px 14px; font-size: 0.8rem; font-weight: 500; color: #4B5563;
                cursor: pointer; display: flex; align-items: center; gap: 6px;
                transition: all 0.15s;
            }
            .btn-tool:hover { background: #F8FAFC; border-color: #CBD5E1; }

            /* Table Styles */
            table { width: 100%; border-collapse: collapse; }
            thead tr { border-bottom: 1px solid var(--border); }
            th {
                padding: 14px 24px; font-size: 0.7rem; font-weight: 700;
                text-transform: uppercase; letter-spacing: 0.05em; color: var(--text-muted);
                text-align: left;
            }
            td { padding: 16px 24px; font-size: 0.85rem; color: #1E293B; border-bottom: 1px solid #F1F5F9; vertical-align: middle; }
            tr:last-child td { border-bottom: none; }
            tbody tr:hover td { background: #F8FAFC; }

            .cand-cell { display: flex; align-items: center; gap: 12px; }
            .cand-av {
                width: 36px; height: 36px; border-radius: 50%;
                display: flex; align-items: center; justify-content: center;
                font-weight: 700; font-size: 0.82rem;
                background: #EEF2FF; color: #4F46E5;
            }
            .cand-name { font-weight: 700; font-size: 0.9rem; color: #1E293B; }
            .cand-email { font-size: 0.74rem; color: var(--text-muted); margin-top: 1px; }

            /* Badges */
            .badge-status { display: inline-flex; align-items: center; padding: 4px 12px; border-radius: 20px; font-size: 0.74rem; font-weight: 600; }
            .badge-status.completed { background: #ECFDF5; color: #10B981; border: 1px solid rgba(16,185,129,0.15); }
            .badge-status.scheduled { background: #FFF7ED; color: #F97316; border: 1px solid rgba(249,115,22,0.15); }

            .badge-score { display: inline-flex; align-items: center; justify-content: center; padding: 4px 10px; min-width: 48px; border-radius: 6px; font-weight: 700; font-size: 0.78rem; }
            .score-high { background: #ECFDF5; color: #10B981; border: 1px solid rgba(16,185,129,0.15); }
            .score-medium { background: #FFFBEB; color: #D97706; border: 1px solid rgba(217,119,6,0.15); }
            .score-low { background: #FEF2F2; color: #EF4444; border: 1px solid rgba(239,68,68,0.15); }

            .text-na { color: var(--text-muted); font-style: italic; }

            .btn-action-view {
                display: inline-flex; align-items: center; gap: 6px;
                padding: 8px 16px; border-radius: 8px;
                background: #4F46E5; color: #FFFFFF;
                font-size: 0.8rem; font-weight: 600; text-decoration: none;
                transition: all 0.15s; box-shadow: 0 1px 2px rgba(79,70,229,0.15);
            }
            .btn-action-view:hover { background: #4338CA; transform: translateY(-1px); }
            .btn-action-nr {
                display: inline-flex; align-items: center;
                padding: 8px 16px; border-radius: 8px;
                background: #F8FAFC; color: #94A3B8; border: 1px solid var(--border);
                font-size: 0.8rem; font-weight: 600; cursor: not-allowed;
            }

            /* Pagination */
            .pagination-container {
                display: flex; align-items: center; justify-content: space-between;
                padding: 16px 24px; border-top: 1px solid var(--border);
                font-size: 0.8rem; color: var(--text-muted);
            }
            .pagination-pages { display: flex; gap: 4px; }
            .page-btn {
                width: 32px; height: 32px; border-radius: 6px; border: 1px solid var(--border);
                background: #FFFFFF; color: var(--text-muted); cursor: pointer;
                display: flex; align-items: center; justify-content: center;
                font-size: 0.8rem; font-weight: 500; transition: all 0.15s;
            }
            .page-btn:hover:not(:disabled) { border-color: #CBD5E1; color: #1E293B; }
            .page-btn.active { background: #EEF2FF; border-color: #C7D2FE; color: #4F46E5; font-weight: 600; }
            .page-btn:disabled { opacity: 0.4; cursor: not-allowed; }

             /* Dropdown and Date Picker CSS */
            .tool-dropdown-container {
                position: relative;
                display: inline-block;
            }
            .dropdown-menu {
                position: absolute;
                top: calc(100% + 6px);
                right: 0;
                background: #FFFFFF;
                border: 1px solid var(--border);
                border-radius: 10px;
                box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -2px rgba(0, 0, 0, 0.05);
                min-width: 170px;
                z-index: 100;
                display: none;
                flex-direction: column;
                padding: 4px;
            }
            .dropdown-menu.open {
                display: flex;
            }
            .dropdown-item {
                padding: 8px 12px;
                font-size: 0.78rem;
                font-weight: 500;
                color: #4B5563;
                border-radius: 6px;
                cursor: pointer;
                display: flex;
                align-items: center;
                gap: 8px;
                transition: all 0.1s;
                text-align: left;
            }
            .dropdown-item:hover {
                background: #F3F4F6;
                color: #1F2937;
            }
            .dropdown-item.active {
                background: #EEF2FF;
                color: #4F46E5;
                font-weight: 600;
            }
            
            .date-picker-box {
                min-width: 240px;
                padding: 12px;
            }
            .date-picker-box label {
                font-size: 0.65rem;
                font-weight: 700;
                color: #6B7280;
                text-transform: uppercase;
                margin-bottom: 4px;
                display: block;
            }
            .date-picker-box input[type="date"] {
                width: 100%;
                padding: 6px;
                border: 1px solid var(--border);
                border-radius: 6px;
                font-size: 0.75rem;
                margin-bottom: 8px;
                outline: none;
                box-sizing: border-box;
            }
            .date-picker-box input[type="date"]:focus {
                border-color: #4F46E5;
            }
            .date-presets {
                border-top: 1px solid var(--border);
                padding-top: 8px;
                margin-top: 4px;
                display: grid;
                grid-template-columns: 1fr 1fr;
                gap: 4px;
            }
            .preset-btn {
                background: #F9FAFB;
                border: 1px solid var(--border);
                border-radius: 4px;
                padding: 4px;
                font-size: 0.7rem;
                font-weight: 500;
                color: #4B5563;
                cursor: pointer;
                text-align: center;
            }
            .preset-btn:hover {
                background: #F3F4F6;
            }
            .date-actions {
                display: flex;
                justify-content: space-between;
                margin-top: 8px;
                border-top: 1px solid var(--border);
                padding-top: 8px;
            }
            .date-btn {
                padding: 4px 8px;
                font-size: 0.72rem;
                font-weight: 600;
                border-radius: 4px;
                cursor: pointer;
            }
            .date-btn.clear {
                background: #FFFFFF;
                border: 1px solid var(--border);
                color: #6B7280;
            }
            .date-btn.apply {
                background: #4F46E5;
                color: #FFFFFF;
                border: 1px solid transparent;
            }
            .date-btn.apply:hover {
                background: #4338CA;
            }

            @media (max-width: 1024px) {
                body { padding: 2rem; }
                .stats-grid { grid-template-columns: repeat(2, 1fr); }
            }
            @media (max-width: 640px) {
                body { padding: 1rem; }
                .stats-grid { grid-template-columns: 1fr; }
                .tcard-hdr { flex-direction: column; align-items: flex-start; }
            }
        </style>
    </head>
    <body>

    <div class="container">
        <!-- HEADER -->
        <header>
            <div class="header-left">
                <div class="logo-icon">AB</div>
                <div class="header-titles">
                    <h1>Candidate Evaluation Dashboard</h1>
                    <p>Adams Bridge AI HRMS Integration</p>
                </div>
            </div>
            <!-- <div class="header-right">
                <button class="bell-btn">
                    🔔<span class="bell-dot"></span>
                </button>
                <div class="user-profile">
                    <div class="user-avatar">AB</div>
                    <div class="user-info">
                        <span class="user-name">Admin User</span>
                        <span class="user-role">HR Manager</span>
                    </div>
                    <span class="chevron-icon">▼</span>
                </div>
            </div> -->
        </header>

        <!-- STATS CARDS -->
        <div class="stats-grid">
            <!-- Card 1: Total Candidates -->
            <div class="stat-card">
                <div class="stat-top-row">
                    <div class="stat-icon-box purple">
                        <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M22 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>
                    </div>
                    <div class="stat-title-val">
                        <span class="stat-label">Total Candidates</span>
                        <span class="stat-value" id="stat-total">0</span>
                    </div>
                </div>
                <div class="stat-bottom-row">
                    <span class="stat-subtext">All time</span>
                    <div class="sparkline-container">
                        <svg class="sparkline-svg" viewBox="0 0 100 30" preserveAspectRatio="none">
                            <defs>
                                <linearGradient id="purple-grad" x1="0" y1="0" x2="0" y2="1">
                                    <stop offset="0%" stop-color="#7C3AED" stop-opacity="0.2"/>
                                    <stop offset="100%" stop-color="#7C3AED" stop-opacity="0"/>
                                </linearGradient>
                            </defs>
                            <path d="M0,25 Q15,22 30,12 T60,18 T90,8 T100,5" fill="none" stroke="#7C3AED" stroke-width="2" stroke-linecap="round"/>
                            <path d="M0,25 Q15,22 30,12 T60,18 T90,8 T100,5 L100,30 L0,30 Z" fill="url(#purple-grad)"/>
                        </svg>
                    </div>
                </div>
            </div>

            <!-- Card 2: Completed Interviews -->
            <div class="stat-card">
                <div class="stat-top-row">
                    <div class="stat-icon-box green">
                        <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><path d="m9 12 2 2 4-4"/></svg>
                    </div>
                    <div class="stat-title-val">
                        <span class="stat-label">Completed Interviews</span>
                        <span class="stat-value" id="stat-completed">0</span>
                    </div>
                </div>
                <div class="stat-bottom-row">
                    <span class="stat-subtext">This month</span>
                    <div class="sparkline-container">
                        <svg class="sparkline-svg" viewBox="0 0 100 30" preserveAspectRatio="none">
                            <defs>
                                <linearGradient id="green-grad" x1="0" y1="0" x2="0" y2="1">
                                    <stop offset="0%" stop-color="#10B981" stop-opacity="0.2"/>
                                    <stop offset="100%" stop-color="#10B981" stop-opacity="0"/>
                                </linearGradient>
                            </defs>
                            <path d="M0,28 Q20,18 40,24 T80,10 T100,12" fill="none" stroke="#10B981" stroke-width="2" stroke-linecap="round"/>
                            <path d="M0,28 Q20,18 40,24 T80,10 T100,12 L100,30 L0,30 Z" fill="url(#green-grad)"/>
                        </svg>
                    </div>
                </div>
            </div>

            <!-- Card 3: Scheduled Interviews -->
            <div class="stat-card">
                <div class="stat-top-row">
                    <div class="stat-icon-box orange">
                        <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="18" height="18" x="3" y="4" rx="2" ry="2"/><line x1="16" x2="16" y1="2" y2="6"/><line x1="8" x2="8" y1="2" y2="6"/><line x1="3" x2="21" y1="10" y2="10"/></svg>
                    </div>
                    <div class="stat-title-val">
                        <span class="stat-label">Scheduled Interviews</span>
                        <span class="stat-value" id="stat-scheduled">0</span>
                    </div>
                </div>
                <div class="stat-bottom-row">
                    <span class="stat-subtext">Upcoming</span>
                    <div class="sparkline-container">
                        <svg class="sparkline-svg" viewBox="0 0 100 30" preserveAspectRatio="none">
                            <defs>
                                <linearGradient id="orange-grad" x1="0" y1="0" x2="0" y2="1">
                                    <stop offset="0%" stop-color="#F97316" stop-opacity="0.2"/>
                                    <stop offset="100%" stop-color="#F97316" stop-opacity="0"/>
                                </linearGradient>
                            </defs>
                            <path d="M0,26 Q15,22 30,26 T60,18 T90,24 T100,10" fill="none" stroke="#F97316" stroke-width="2" stroke-linecap="round"/>
                            <path d="M0,26 Q15,22 30,26 T60,18 T90,24 T100,10 L100,30 L0,30 Z" fill="url(#orange-grad)"/>
                        </svg>
                    </div>
                </div>
            </div>

            <!-- Card 4: Average ATS Match Score -->
            <div class="stat-card">
                <div class="stat-top-row">
                    <div class="stat-icon-box blue">
                        <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="2"/></svg>
                    </div>
                    <div class="stat-title-val">
                        <span class="stat-label">Average ATS Match Score</span>
                        <span class="stat-value" id="stat-avg">0%</span>
                    </div>
                </div>
                <div class="stat-bottom-row">
                    <span class="stat-subtext">Across all candidates</span>
                    <div class="sparkline-container">
                        <svg class="sparkline-svg" viewBox="0 0 100 30" preserveAspectRatio="none">
                            <defs>
                                <linearGradient id="blue-grad" x1="0" y1="0" x2="0" y2="1">
                                    <stop offset="0%" stop-color="#3B82F6" stop-opacity="0.2"/>
                                    <stop offset="100%" stop-color="#3B82F6" stop-opacity="0"/>
                                </linearGradient>
                            </defs>
                            <path d="M0,22 Q30,12 60,25 T100,8" fill="none" stroke="#3B82F6" stroke-width="2" stroke-linecap="round"/>
                            <path d="M0,22 Q30,12 60,25 T100,8 L100,30 L0,30 Z" fill="url(#blue-grad)"/>
                        </svg>
                    </div>
                </div>
            </div>
        </div>

        <!-- TABLE CARD -->
        <div class="tcard">
            <div class="tcard-hdr">
                <div class="tcard-title">Recent Candidates</div>
                <div class="tcard-tools">
                    <div class="search-wrap">
                        <span class="search-icon">🔍</span>
                        <input type="text" id="searchQ" placeholder="Search candidate..." oninput="applyFilter()">
                    </div>
                    <!-- FILTER DROPDOWN -->
                    <div class="tool-dropdown-container">
                        <button class="btn-tool" id="btnFilter" onclick="toggleDropdown(event, 'filterDropdown')">
                            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3"/></svg>
                            <span>Filter: ALL</span>
                            <span style="font-size: 0.6rem; margin-left: 2px;">▼</span>
                        </button>
                        <div class="dropdown-menu" id="filterDropdown">
                            <div class="dropdown-item active" onclick="selectFilter('all')">All Candidates</div>
                            <div class="dropdown-item" onclick="selectFilter('completed')">Completed</div>
                            <div class="dropdown-item" onclick="selectFilter('scheduled')">Scheduled</div>
                        </div>
                    </div>

                    <!-- DATE DROPDOWN -->
                    <div class="tool-dropdown-container">
                        <button class="btn-tool" id="btnDate" onclick="toggleDropdown(event, 'dateDropdown')">
                            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><rect width="18" height="18" x="3" y="4" rx="2" ry="2"/><line x1="16" x2="16" y1="2" y2="6"/><line x1="8" x2="8" y1="2" y2="6"/><line x1="3" x2="21" y1="10" y2="10"/></svg>
                            <span id="lblDateRange">All Time</span>
                            <span style="font-size: 0.6rem; margin-left: 2px;">▼</span>
                        </button>
                        <div class="dropdown-menu date-picker-box" id="dateDropdown">
                            <div style="margin-bottom: 8px">
                                <label>Start Date</label>
                                <input type="date" id="dateStart">
                            </div>
                            <div style="margin-bottom: 8px">
                                <label>End Date</label>
                                <input type="date" id="dateEnd">
                            </div>
                            <div class="date-presets">
                                <button class="preset-btn" onclick="applyPreset('today')">Today</button>
                                <button class="preset-btn" onclick="applyPreset('yesterday')">Yesterday</button>
                                <button class="preset-btn" onclick="applyPreset('7days')">Last 7 Days</button>
                                <button class="preset-btn" onclick="applyPreset('30days')">Last 30 Days</button>
                            </div>
                            <div class="date-actions">
                                <button class="date-btn clear" onclick="clearDateFilter()">Clear</button>
                                <button class="date-btn apply" onclick="applyDateFilter()">Apply</button>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            
            <table>
                <thead>
                    <tr>
                        <th>Candidate Details</th>
                        <th>Status</th>
                        <th>ATS Match Score</th>
                        <th>Cheating Risk</th>
                        <th>Interview Date</th>
                        <th>Action</th>
                    </tr>
                </thead>
                <tbody id="tb">
                    <tr><td colspan="6" style="text-align:center;padding:44px;color:#6B7280"><div style="font-size:1.5rem;margin-bottom:8px">⏳</div>Loading candidates…</td></tr>
                </tbody>
            </table>
            
            <div class="pagination-container">
                <span id="pInfo">Showing — to — of — candidates</span>
                <div class="pagination-pages" id="pBtns"></div>
            </div>
        </div>
    </div>

    <script>
        let ALL = [], curF = 'all', curP = 1;
        const PS = 5;
        const COLORS = ['#4F46E5','#06B6D4','#10B981','#F59E0B','#EF4444','#8B5CF6','#EC4899','#F97316'];
        
        function aColor(n){ let h=0; for(let i=0;i<(n||'').length;i++) h=(h*31+n.charCodeAt(i))&0xFFFFFF; return COLORS[Math.abs(h)%COLORS.length]; }
        function ini(n){ if(!n) return '?'; const p=n.trim().split(' '); return (p[0][0]+(p[1]?.[0]||'')).toUpperCase(); }
        
        function fdt(iso){ 
            if(!iso) return 'N/A'; 
            try{ 
                const d = new Date(iso);
                const yr = d.getFullYear();
                const mo = String(d.getMonth()+1).padStart(2,'0');
                const dy = String(d.getDate()).padStart(2,'0');
                let hr = d.getHours();
                const ampm = hr >= 12 ? 'PM' : 'AM';
                hr = hr % 12;
                hr = hr ? hr : 12;
                const min = String(d.getMinutes()).padStart(2,'0');
                return `${yr}-${mo}-${dy} ${String(hr).padStart(2,'0')}:${min} ${ampm}`;
            } catch{ 
                return iso; 
            } 
        }

        function renderTable(list){
            const tb = document.getElementById('tb');
            const off = (curP-1)*PS;
            const page = list.slice(off, off+PS);
            
            if(!list.length){
                tb.innerHTML = `<tr><td colspan="6" style="text-align:center;padding:44px;color:#6B7280">No candidates match search filter.</td></tr>`;
                document.getElementById('pInfo').textContent='Showing 0 of 0 candidates';
                document.getElementById('pBtns').innerHTML='';
                return;
            }
            
            tb.innerHTML='';
            page.forEach((c,i)=>{
                const col=aColor(c.candidate_name), inits=ini(c.candidate_name);
                const sb = c.status==='completed'
                    ? '<span class="badge-status completed">Completed</span>'
                    : '<span class="badge-status scheduled">Scheduled</span>';
                
                let ats='<span class="text-na">N/A</span>';
                let ch='<span class="text-na">N/A</span>';
                let act='<span class="btn-action-nr">No Report</span>';
                
                if(c.status==='completed'){
                    const sc=c.ats_scorecard?.overall_score??0;
                    ats=`<span class="badge-score ${sc>=75?'score-high':sc>=50?'score-medium':'score-low'}">${sc}%</span>`;
                    const chv=c.proctoring_analysis?.cheating_score??0;
                    ch=`<span class="badge-score ${chv>=50?'score-low':chv>=20?'score-medium':'score-high'}">${chv}%</span>`;
                    act=`<a href="/api/report/${encodeURIComponent(c.candidate_email)}" class="btn-action-view"><svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" style="margin-right:2px"><path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z"/><polyline points="14 2 14 8 20 8"/></svg>View Scorecard</a>`;
                }
                
                tb.innerHTML+=`
                <tr>
                  <td>
                    <div class="cand-cell">
                      <div class="cand-av" style="background:${col};color:#fff">${inits}</div>
                      <div>
                        <div class="cand-name">${c.candidate_name||'Unknown'}</div>
                        <div class="cand-email">${c.candidate_email||''}</div>
                      </div>
                    </div>
                  </td>
                  <td>${sb}</td>
                  <td>${ats}</td>
                  <td>${ch}</td>
                  <td><div style="display:flex;align-items:center;gap:6px;font-size:0.78rem;color:#4B5563"><svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><rect width="18" height="18" x="3" y="4" rx="2" ry="2"/><line x1="16" x2="16" y1="2" y2="6"/><line x1="8" x2="8" y1="2" y2="6"/><line x1="3" x2="21" y1="10" y2="10"/></svg>${fdt(c.start_time||c.created_at)}</div></td>
                  <td>${act}</td>
                </tr>`;
            });
            
            const tp=Math.ceil(list.length/PS);
            document.getElementById('pInfo').textContent=`Showing ${off+1} to ${Math.min(off+PS,list.length)} of ${list.length} candidates`;
            
            const pb=document.getElementById('pBtns'); pb.innerHTML='';
            
            const prev = document.createElement('button');
            prev.className = 'page-btn';
            prev.innerHTML = '‹';
            prev.disabled = curP === 1;
            prev.onclick = () => { if(curP>1){ curP--; renderTable(getFiltered()); } };
            pb.appendChild(prev);

            for(let p=1;p<=tp;p++){
                const b=document.createElement('button');
                b.className='page-btn'+(p===curP?' active':'');
                b.textContent=p;
                b.onclick=()=>{curP=p; renderTable(getFiltered());};
                pb.appendChild(b);
            }

            const next = document.createElement('button');
            next.className = 'page-btn';
            next.innerHTML = '›';
            next.disabled = curP === tp;
            next.onclick = () => { if(curP<tp){ curP++; renderTable(getFiltered()); } };
            pb.appendChild(next);
        }

        let filterStartDate = null;
        let filterEndDate = null;

        function getFiltered(){
            const q=(document.getElementById('searchQ').value||'').toLowerCase();
            return ALL.filter(c=>{
                // 1. Status Filter

                const mf=curF==='all'||c.status===curF;

                // 2. Search Text Filter
                const ms=!q||(c.candidate_name||'').toLowerCase().includes(q)||(c.candidate_email||'').toLowerCase().includes(q);

                // 3. Date Range Filter
                let md = true;
                if (filterStartDate || filterEndDate) {
                    const cDate = new Date(c.created_at || c.start_time);
                    if (filterStartDate) {
                        const start = new Date(filterStartDate);
                        start.setHours(0,0,0,0);
                        if (cDate < start) md = false;
                    }
                    if (filterEndDate) {
                        const end = new Date(filterEndDate);
                        end.setHours(23,59,59,999);
                        if (cDate > end) md = false;
                    }
                }
                
                return mf&&ms&&md;
            });
        }
        
        function applyFilter(){ curP=1; renderTable(getFiltered()); }
        
        function toggleDropdown(e, id) {
            e.stopPropagation();
            const dropdowns = document.querySelectorAll('.dropdown-menu');
            dropdowns.forEach(d => {
                if (d.id !== id) d.classList.remove('open');
            });
            document.getElementById(id).classList.toggle('open');
        }
        
        document.addEventListener('click', function(e) {
            if (!e.target.closest('.tool-dropdown-container')) {
                const dropdowns = document.querySelectorAll('.dropdown-menu');
                dropdowns.forEach(d => d.classList.remove('open'));
            }
        });

        function selectFilter(status) {
            curF = status;
            const label = status === 'all' ? 'ALL' : status.toUpperCase();
            document.getElementById('btnFilter').querySelector('span').textContent = `Filter: ${label}`;
            
            const items = document.getElementById('filterDropdown').querySelectorAll('.dropdown-item');
            items.forEach(item => {
                if (item.getAttribute('onclick').includes(status)) {
                    item.classList.add('active');
                } else {
                    item.classList.remove('active');
                }
            });
            
            document.getElementById('filterDropdown').classList.remove('open');

            applyFilter();
        }

        function formatShortDate(dateStr) {
            if (!dateStr) return '';
            const d = new Date(dateStr);
            const opt = { month: 'short', day: '2-digit' };
            return d.toLocaleDateString('en-US', opt);
        }

        function applyDateFilter() {
            const startVal = document.getElementById('dateStart').value;
            const endVal = document.getElementById('dateEnd').value;
            
            filterStartDate = startVal ? startVal : null;
            filterEndDate = endVal ? endVal : null;
            
            const lbl = document.getElementById('lblDateRange');
            if (filterStartDate && filterEndDate) {
                lbl.textContent = `${formatShortDate(filterStartDate)} - ${formatShortDate(filterEndDate)}`;
            } else if (filterStartDate) {
                lbl.textContent = `From ${formatShortDate(filterStartDate)}`;
            } else if (filterEndDate) {
                lbl.textContent = `Until ${formatShortDate(filterEndDate)}`;
            } else {
                lbl.textContent = `All Time`;
            }
            
            document.getElementById('dateDropdown').classList.remove('open');
            applyFilter();
        }

        function clearDateFilter() {
            document.getElementById('dateStart').value = '';
            document.getElementById('dateEnd').value = '';
            filterStartDate = null;
            filterEndDate = null;
            document.getElementById('lblDateRange').textContent = 'All Time';
            document.getElementById('dateDropdown').classList.remove('open');
            applyFilter();
        }

        function applyPreset(preset) {
            const today = new Date();
            let start = new Date();
            let end = new Date();
            
            if (preset === 'today') {
                start = today;
                end = today;
            } else if (preset === 'yesterday') {
                start.setDate(today.getDate() - 1);
                end.setDate(today.getDate() - 1);
            } else if (preset === '7days') {
                start.setDate(today.getDate() - 6);
            } else if (preset === '30days') {
                start.setDate(today.getDate() - 29);
            }
            
            const formatDate = (d) => {
                const yr = d.getFullYear();
                const mo = String(d.getMonth() + 1).padStart(2, '0');
                const dy = String(d.getDate()).padStart(2, '0');
                return `${yr}-${mo}-${dy}`;
            };
            
            document.getElementById('dateStart').value = formatDate(start);
            document.getElementById('dateEnd').value = formatDate(end);
            applyDateFilter();
        }

        async function load(){
            document.getElementById('tb').innerHTML=`<tr><td colspan="6" style="text-align:center;padding:44px;color:#6B7280"><div style="font-size:1.5rem;margin-bottom:8px">⏳</div>Loading candidates…</td></tr>`;
            try{
                ALL=await (await fetch('/api/candidates')).json();
                
                ALL.sort((a,b) => {
                    if(a.status === 'completed' && b.status !== 'completed') return -1;
                    if(a.status !== 'completed' && b.status === 'completed') return 1;
                    return new Date(b.created_at || b.start_time) - new Date(a.created_at || a.start_time);
                });

                const tot=ALL.length, comp=ALL.filter(c=>c.status==='completed').length, sched=ALL.filter(c=>c.status==='scheduled').length;
                let ts=0,sc=0;
                
                ALL.forEach(c=>{ 
                    if(c.status==='completed'&&c.ats_scorecard?.overall_score!=null){
                        ts+=c.ats_scorecard.overall_score;
                        sc++;
                    } 
                });
                
                document.getElementById('stat-total').textContent=tot;
                document.getElementById('stat-completed').textContent=comp;
                document.getElementById('stat-scheduled').textContent=sched;
                document.getElementById('stat-avg').textContent=(sc>0?Math.round(ts/sc):0)+'%';
                
                renderTable(getFiltered());
            }catch(e){
                document.getElementById('tb').innerHTML=`<tr><td colspan="6" style="text-align:center;padding:44px;color:#EF4444"><div style="font-size:1.5rem;margin-bottom:8px">⚠️</div>Failed to connect to backend server. Make sure MongoDB is active.</td></tr>`;
            }
        }
        
        window.onload=load;
    </script>
    </body>
    </html>
    """
    return HTMLResponse(content=html_content)

# ─────────────────────────────────────────────────────────────────────────────
# AI INTERVIEW PROCTORING REPORT  ·  GET /api/report/{email}
# ─────────────────────────────────────────────────────────────────────────────

@router.get("/report/{email:path}")
async def serve_proctoring_report(email: str):
    """
    Full-page AI Interview Proctoring Report for HR review.
    Renders the premium dark-mode dashboard with:
      - 6 KPI cards (ATS, Integrity, Risk, AI Confidence, Recommendation, Cheating)
      - AI Proctoring Summary table
      - Integrity Score gauge + Integrity Factors bar chart
      - Integrity Timeline with evidence screenshots
      - Candidate Performance (skill bars + radar chart)
      - Evidence Snapshots carousel
      - Recruiter Summary (strengths / concerns)
      - AI Recommendation card
      - Download PDF button
    """
    candidate = await candidates_collection.find_one({"candidate_email": email})
    if not candidate:
        return HTMLResponse(content="<h2 style='color:white;font-family:sans-serif;padding:2rem'>Candidate not found.</h2>", status_code=404)

    # ── Safely extract nested documents ──
    proctor   = candidate.get("proctoring_analysis", {}) or {}
    scorecard = candidate.get("ats_scorecard", {}) or {}
    logs      = proctor.get("logs", []) or []
    risk_bd   = proctor.get("risk_breakdown", {}) or {}
    det_sum   = proctor.get("detection_summary", {}) or {}
    audio     = proctor.get("audio_analysis", {}) or {}
    skills    = scorecard.get("skills_assessment", []) or []
    strengths_s = scorecard.get("strengths_summary", scorecard.get("strengths", []))
    concerns_s  = scorecard.get("concerns_summary",  scorecard.get("weaknesses", []))

    # ── Computed display values ──
    ats_score       = scorecard.get("overall_score", 0)
    integrity_score = proctor.get("integrity_score", 0)
    cheating_score  = proctor.get("cheating_score", 0)
    ai_conf         = proctor.get("ai_confidence", 96.7)
    attention       = proctor.get("attention_score", 0)
    recommendation  = scorecard.get("recommendation", "Proceed with Caution")
    integrity_commentary = scorecard.get("integrity_check_commentary", "")
    detailed_feedback    = scorecard.get("detailed_feedback", "")

    # Risk label / color derived from cheating score
    if cheating_score <= 30:
        risk_label, risk_color, risk_icon = "Low Risk", "#4ade80", "🛡️"
    elif cheating_score <= 60:
        risk_label, risk_color, risk_icon = "Medium Risk", "#fb923c", "⚠️"
    else:
        risk_label, risk_color, risk_icon = "High Risk", "#f87171", "🚨"

    # Recommendation card styling
    rec_map = {
        "Strong Hire":         ("#166534", "#4ade80", "✅", "Strong candidate — proceed to next round"),
        "Hire":                ("#14532d", "#86efac", "✅", "Recommended for the role"),
        "Proceed with Caution":("#78350f", "#fb923c", "⚠️", "Integrity concerns observed — review required"),
        "No Hire":             ("#7f1d1d", "#f87171", "🚫", "Not recommended for this role"),
    }
    rec_bg, rec_text_color, rec_icon, rec_sub = rec_map.get(recommendation, rec_map["Proceed with Caution"])

    # Interview metadata
    name     = candidate.get("candidate_name", "Candidate")
    position = candidate.get("job_description_text", "")[:40] if candidate.get("job_description_text") else "Role"
    initials = "".join(p[0].upper() for p in name.split()[:2])
    interview_date = candidate.get("start_time", "")[:10] if candidate.get("start_time") else "N/A"
    interview_time = candidate.get("start_time", "")
    duration_s = proctor.get("total_sampled_seconds", 0)
    duration_str = f"{duration_s // 60}m {duration_s % 60}s" if duration_s else "N/A"
    video_url = ""
    if candidate.get("recording_local_path"):
        vfn = os.path.basename(candidate["recording_local_path"])
        video_url = f"/static/recordings/{vfn}"
    report_id = f"ATS-{interview_date.replace('-', '')}001"

    # ── Interview Statistics ──
    total_questions   = len(candidate.get("questions", []))
    answered_q        = scorecard.get("answered_questions", total_questions)
    total_incidents   = len(logs)
    evidence_captured = len([l for l in logs if l.get("image_url")])
    avg_response_s    = proctor.get("avg_response_time", 0)
    avg_response_str  = f"{avg_response_s}s" if avg_response_s else "N/A"
    dur_min  = duration_s // 60 if duration_s else 0
    dur_sec  = duration_s % 60  if duration_s else 0
    dur_disp = f"{dur_min}m {dur_sec}s" if duration_s else "N/A"

    # ── Integrity score color ──
    if integrity_score >= 71:
        ig_label, ig_color = "Good Integrity", "#0B6623"
    elif integrity_score >= 41:
        ig_label, ig_color = "Moderate Integrity", "#D97706"
    else:
        ig_label, ig_color = "Low Integrity", "#DC2626"

    # ── Risk breakdown for bar chart ──
    rb_labels  = ["Face Presence", "Eye Behavior", "Head Pose", "Phone Detection", "Audio Analysis"]
    rb_keys    = ["face_presence", "eye_behavior", "head_pose", "phone_detection", "audio_analysis"]
    rb_scores  = [risk_bd.get(k, {}).get("score", 0) for k in rb_keys]
    rb_maxes   = [risk_bd.get(k, {}).get("max", 25 if k in ["face_presence","eye_behavior"] else (15 if k in ["head_pose","audio_analysis"] else 20)) for k in rb_keys]
    rb_total   = sum(rb_scores)
    rb_total_max = sum(rb_maxes)

    # ── Skills for radar + bars ──
    def find_rating(keywords, fallback_val):
        for s in skills:
            skill_name = s.get("skill", "").lower()
            if any(kw in skill_name for kw in keywords):
                return s.get("rating", fallback_val)
        return fallback_val

    tech_fallback = scorecard.get("technical_rating", round(ats_score / 10, 1) if ats_score else 7.5)
    tech_rating = find_rating(["technical", "knowledge", "programming", "development", "coding", "experience"], tech_fallback)

    prob_fallback = round(ats_score / 10, 1) if ats_score else 6.9
    prob_rating = find_rating(["problem", "solving", "analytical", "logic", "reasoning"], prob_fallback)

    comm_fallback = scorecard.get("communication_rating", round(ats_score / 10, 1) if ats_score else 8.0)
    comm_rating = find_rating(["communication", "verbal", "articulate", "english"], comm_fallback)

    conf_fallback = scorecard.get("culture_fit_rating", round(ats_score / 10, 1) if ats_score else 7.8)
    conf_rating = find_rating(["confidence", "behavioral", "demeanor", "presentation", "personality"], conf_fallback)

    # Enforce exactly these four standard competencies
    skill_labels  = ["Technical Knowledge", "Problem Solving", "Communication", "Confidence"]
    skill_ratings = [tech_rating, prob_rating, comm_rating, conf_rating]

    # ── Detection table rows (9 metrics) ──
    dt_rows = [
        ("face_presence",  "👤", "Face Presence"),
        ("eye_attention",  "👁️", "Eye Attention"),
        ("phone_detected", "📱", "Mobile Phone Detected"),
        ("multiple_persons","👥","Multiple Person Detected"),
        ("face_missing",   "🚫", "Face Missing"),
        ("looking_away",   "↔️", "Looking Away (Eye Gaze)"),
        ("head_turn",      "↩️", "Head Turn > 45°"),
        ("background_voice","🔊","Background Voice"),
        ("camera_blocked", "📷", "Camera Blocked"),
    ]
    sev_classes = {
        "none":   ("color:#4B5563", "#F3F4F6"),
        "normal": ("color:#15803D", "#DCFCE7"),
        "low":    ("color:#B45309", "#FEF9C3"),
        "medium": ("color:#B45309", "#FFEDD5"),
        "high":   ("color:#B91C1C", "#FEE2E2"),
    }

    def status_html(item):
        sev = item.get("severity", "none")
        col, bg = sev_classes.get(sev, sev_classes["none"])
        return f'<span style="padding:2px 8px;border-radius:4px;font-size:0.75rem;font-weight:600;{col};background:{bg}">{item.get("status","—")}</span>'
    
    
    # ── Normalize image_url: convert absolute file paths → /static/recordings/... URL ──
    def normalize_image_url(url: str) -> str:
        """Convert an absolute file path stored in MongoDB to a browser-accessible /static URL."""
        if not url:
            return url
        # Already a web URL — return as-is
        if url.startswith("http://") or url.startswith("https://") or url.startswith("/static/"):
            return url
        # Extract just the filename (handles both Linux and Windows separators)
        fname = os.path.basename(url.replace("\\", "/"))

        # All incident/flag screenshots live under incidents/
        return f"/static/recordings/incidents/{fname}"

    # Apply normalisation to every log entry
    for _l in logs:
        if _l.get("image_url"):
            _l["image_url"] = normalize_image_url(_l["image_url"])

    # ── Evidence snapshot images from logs ──
    snapshot_logs = [l for l in logs if l.get("image_url")]

    logs_json     = json.dumps(logs)
    skills_json   = json.dumps([{"skill": s.get("skill"), "rating": s.get("rating", 5), "comment": s.get("comment","")} for s in skills[:6]])
    snapshots_json = json.dumps(snapshot_logs)
    rb_scores_json = json.dumps(rb_scores)
    rb_maxes_json  = json.dumps(rb_maxes)
    rb_labels_json = json.dumps(rb_labels)
    skill_labels_json  = json.dumps(skill_labels)
    skill_ratings_json = json.dumps(skill_ratings)
    strengths_json = json.dumps(strengths_s)
    concerns_json  = json.dumps(concerns_s)

    # Build detection table HTML
    dt_html = ""
    for key, icon, label in dt_rows:
        item = det_sum.get(key, {"status": "N/A", "detail": "—", "severity": "none"})
        detail_txt = item.get("detail", "—")
        dt_html += f"""
        <tr style="border-bottom:1px solid #F3F4F6">
            <td style="padding:8px 6px;display:flex;align-items:center;gap:6px;font-size:0.78rem;color:#111827">{icon} {label}</td>
            <td style="padding:8px 6px">{status_html(item)}</td>
            <td style="padding:8px 6px;font-size:0.77rem;color:#6B7280">{detail_txt}</td>
        </tr>"""

    # Build recruiter summary HTML
    str_html = "".join(f"""
    <div class="rec-item strength">
        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" style="flex-shrink:0;margin-top:2px"><polyline points="20 6 9 17 4 12"/></svg>
        <span>{s}</span>
    </div>""" for s in strengths_s[:5])

    conc_html = "".join(f"""
    <div class="rec-item concern">
        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" style="flex-shrink:0;margin-top:2px"><path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3Z"/><line x1="12" x2="12" y1="9" y2="13"/><line x1="12" x2="12.01" y1="17" y2="17"/></svg>
        <span>{c}</span>
    </div>""" for c in concerns_s[:5])

    if not str_html:
        str_html = '<div style="font-size:0.8rem;color:#6B7280;font-style:italic">No strengths noted</div>'
    if not conc_html:
        conc_html = '<div style="font-size:0.8rem;color:#6B7280;font-style:italic">No concerns flagged</div>'

    # Extract 5 points for Overall Summary
    raw_points = scorecard.get("overall_summary_points", []) or []
    
    # Fallback to parsing detailed_feedback (which is also generated by AI) if overall_summary_points is not populated
    if not raw_points and detailed_feedback:
        # First, split by newlines and see if we have short, bullet-like lines
        candidate_lines = []
        for line in detailed_feedback.split('\n'):
            line_str = line.strip()
            if not line_str or line_str.startswith('#'):
                continue
            # If the line is an actual list item or a short sentence
            cleaned = line_str.lstrip('*-•1234567890.').strip()
            if len(cleaned) > 10 and len(cleaned) < 150:
                candidate_lines.append(cleaned)
        
        # If we have at least 3 clean bullet-like lines, use them
        if len(candidate_lines) >= 3:
            raw_points = candidate_lines
        else:
            # Otherwise, split the text into individual sentences
            
            sentences = re.split(r'(?<=[.!?])\s+', detailed_feedback.strip())
            for s in sentences:
                s_clean = s.strip().lstrip('*-•1234567890.').strip()
                if len(s_clean) > 12 and not s_clean.startswith('#'):
                    # Truncate very long sentences for clean bullet presentation
                    if len(s_clean) > 100:
                        s_clean = s_clean[:97] + "..."
                    raw_points.append(s_clean)
     
    # Limit to maximum of 5 points
    summary_points = raw_points[:5]
    
    # Render points with the 5 different SVG icons
    icons = [
        # Blue profile
        '<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#2563EB" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" style="flex-shrink:0;margin-top:2px"><path d="M19 21v-2a4 4 0 0 0-4-4H9a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>',
        # Green check circle
        '<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#10B981" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" style="flex-shrink:0;margin-top:2px"><circle cx="12" cy="12" r="10"/><path d="m9 12 2 2 4-4"/></svg>',
        # Red cross circle
        '<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#EF4444" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" style="flex-shrink:0;margin-top:2px"><circle cx="12" cy="12" r="10"/><path d="m15 9-6 6"/><path d="m9 9 6 6"/></svg>',
        # Orange checkmark
        '<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#F97316" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" style="flex-shrink:0;margin-top:2px"><polyline points="20 6 9 17 4 12"/></svg>',
        # Green shield check
        '<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#10B981" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" style="flex-shrink:0;margin-top:2px"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>'
    ]
    
    fb_points_html = ""
    if summary_points:
        for idx, pt in enumerate(summary_points):
            icon = icons[idx % len(icons)]
            fb_points_html += f"""
            <div class="feedback-item">
                {icon}
                <span>{pt}</span>
            </div>"""
    else:
        fb_points_html = """
        <div style="font-size:0.82rem;color:#6B7280;font-style:italic;padding:12px 0;text-align:center;width:100%">
            AI Technical Evaluation is pending for this candidate.
        </div>"""

     # Extract points for Recommendation Reasons (generated by AI)
    rec_reasons = scorecard.get("recommendation_reasons", []) or []
        
    rec_reasons_html = ""
    if rec_reasons:
        for pt in rec_reasons:
            rec_reasons_html += f"""
            <div style="display:flex;align-items:start;gap:8px;font-size:0.74rem;color:#1E293B;line-height:1.4">
                <span style="color:#2563EB;font-size:0.9rem;line-height:1;margin-top:0px">•</span>
                <span>{pt}</span>
            </div>"""
    else:
        rec_reasons_html = """
        <div style="font-size:0.72rem;color:#6B7280;font-style:italic">
            No recommendation reason details available.
        </div>"""

    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>AI Proctoring Report — {name}</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
  :root {{
    --bg-main:   #F5F7FA;
    --bg-card:   #FFFFFF;
    --bg-card2:  #F9FAFB;
    --border:    #E5E7EB;
    --text:      #111827;
    --text-muted:#6B7280;
    --green:     #0B6623;
    --orange:    #F59E0B;
    --red:       #DC2626;
    --blue:      #2563EB;
    --purple:    #7C3AED;
    --yellow:    #D97706;
    --white:     #F8FAFC;
  }}
  * {{ margin:0; padding:0; box-sizing:border-box; }}
  body {{ font-family:'Inter',sans-serif; background:var(--bg-main); color:var(--text); min-height:100vh; display:flex; }}
  
  /* SIDEBAR */
  .sidebar {{ width:200px; min-height:100vh; background:#1E293B; border-right:1px solid #334155; display:flex; flex-direction:column; flex-shrink:0; position:fixed; top:0; left:0; z-index:10; }}
  .sidebar-logo {{ padding:18px 16px; border-bottom:1px solid #334155; display:flex;align-items:center;gap:8px; }}
  .sidebar-logo .logo-icon {{ width:28px;height:28px;background:linear-gradient(135deg,#2563EB,#4F46E5);border-radius:6px;display:flex;align-items:center;justify-content:center;font-size:0.75rem;font-weight:700;color:white; }}
  .sidebar-logo span {{ font-size:0.85rem;font-weight:700;color:#F8FAFC; }}
  .sidebar-nav {{ padding:8px 0; flex:1; }}
  .sidebar-nav a {{ text-decoration:none;color:inherit;outline:none; }}
  .nav-item {{ display:flex;align-items:center;gap:10px;padding:9px 16px;font-size:0.8rem;color:#94A3B8;cursor:pointer;transition:all 0.2s; }}
  .nav-item:hover {{ background:rgba(99,102,241,0.12);color:#F8FAFC; }}
  .nav-item.active {{ background:rgba(37,99,235,0.18);color:#93C5FD;border-right:2px solid #2563EB; }}
  .sidebar-footer {{ padding:12px 16px;border-top:1px solid #334155;font-size:0.7rem;color:#64748B; }}
  .sidebar-footer .ai-badge {{ display:flex;align-items:center;gap:6px;margin-top:4px; }}
  
  /* MAIN */
  .main {{ margin-left:200px; flex:1; display:flex; flex-direction:column; min-height:100vh; }}
  
  /* TOP BAR */
  .topbar {{ background:#FFFFFF; border-bottom:1px solid var(--border); padding:12px 24px; display:flex; align-items:center; justify-content:space-between; position:sticky;top:0;z-index:9;box-shadow:0 1px 4px rgba(0,0,0,0.06); }}
  .topbar h1 {{ font-size:1.05rem;font-weight:700;color:#111827; }}
  .topbar-actions {{ display:flex;gap:8px; }}
  .btn {{ padding:6px 14px;border-radius:7px;font-size:0.78rem;font-weight:600;cursor:pointer;transition:all 0.2s;border:1px solid var(--border);background:#F9FAFB;color:#374151;display:flex;align-items:center;gap:6px; }}
  .btn:hover {{ background:#F3F4F6;border-color:#D1D5DB;color:#111827; }}
  .btn-primary {{ background:#2563EB;border-color:#2563EB;color:white; }}
  .btn-primary:hover {{ background:#1D4ED8;border-color:#1D4ED8; }}
  
  /* CANDIDATE HEADER */
  .candidate-header {{ background:#FFFFFF; border-bottom:1px solid var(--border); padding:14px 24px; display:flex; align-items:center; gap:16px; box-shadow:0 1px 3px rgba(0,0,0,0.05); }}
  .avatar {{ width:52px;height:52px;border-radius:50%;background:linear-gradient(135deg,#2563EB,#4F46E5);display:flex;align-items:center;justify-content:center;font-size:1.1rem;font-weight:700;color:white;flex-shrink:0; }}
  .cand-info h2 {{ font-size:1rem;font-weight:700;color:#111827; }}
  .cand-info .subtitle {{ font-size:0.73rem;color:var(--text-muted);margin-top:2px; }}
  .status-badge {{ padding:2px 10px;border-radius:20px;font-size:0.7rem;font-weight:600;background:#DCFCE7;color:#0B6623;border:1px solid #BBF7D0; }}
  .meta-chips {{ display:flex;gap:16px;flex-wrap:wrap;margin-left:auto; }}
  .meta-chip {{ display:flex;align-items:center;gap:5px;font-size:0.72rem;color:var(--text-muted); }}
  .meta-chip strong {{ color:#111827; }}
  
  /* CONTENT */
  .content {{ padding:18px 24px; flex:1; }}
  
  /* KPI CARDS */
  .kpi-row {{ display:grid;grid-template-columns:repeat(6,1fr);gap:12px;margin-bottom:16px; }}
  .kpi-card {{ background:#FFFFFF;border:1px solid var(--border);border-radius:12px;padding:14px;text-align:center;position:relative;box-shadow:0 1px 4px rgba(0,0,0,0.05);display:flex;flex-direction:column;justify-content:center;align-items:center;overflow:hidden; }}
  .kpi-card::before {{ content:'';position:absolute;top:0;left:0;right:0;height:3px;background:var(--accent,#2563EB);border-radius:12px 12px 0 0; }}
  .kpi-label {{ font-size:0.65rem;font-weight:600;text-transform:uppercase;letter-spacing:0.05em;color:#6B7280;margin-bottom:8px; }}
  .kpi-value {{ font-size:1.6rem;font-weight:800;line-height:1;color:#111827; }}
  .kpi-sub {{ font-size:0.68rem;color:#6B7280;margin-top:4px; }}
  .gauge-wrap {{ position:relative;width:80px;height:45px;margin:0 auto 4px; }}
  .gauge-wrap canvas {{ width:80px!important;height:45px!important; }}
  .gauge-center {{ position:absolute;bottom:0;left:50%;transform:translateX(-50%);font-size:0.9rem;font-weight:800; }}
  
  /* MAIN GRID */
  .grid-3 {{ display:grid;grid-template-columns:1.1fr 0.9fr 1fr;gap:14px;margin-bottom:14px; }}
  .grid-3-equal {{ display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:14px; }}
  .grid-4 {{ display:grid;grid-template-columns:1.6fr 2.2fr 1.8fr 1.2fr;gap:14px;margin-bottom:14px; }}
  @media (max-width: 1024px) {{
    .grid-4 {{ grid-template-columns: 1fr!important; }}
  }}

  .audio-strip {{ display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:14px; }}

  /* AUDIO STATS */
  .audio-stat {{ background:#FFFFFF;border:1px solid var(--border);border-radius:10px;padding:12px 16px;display:flex;align-items:center;gap:12px;box-shadow:0 1px 3px rgba(0,0,0,0.05); }}

  .audio-stat-icon {{ font-size:1.4rem; }}
  .audio-stat-label {{ font-size:0.68rem;color:#6B7280;font-weight:600;text-transform:uppercase;letter-spacing:0.04em; }}
  .audio-stat-val {{ font-size:1.1rem;font-weight:700;color:#111827; }}
  
  /* CARDS */
  .card {{ background:#FFFFFF;border:1px solid var(--border);border-radius:12px;padding:16px;box-shadow:0 1px 4px rgba(0,0,0,0.05);display:flex;flex-direction:column; }}
  .card-title {{ font-size:0.72rem;font-weight:700;text-transform:uppercase;letter-spacing:0.05em;color:#6B7280;margin-bottom:12px; }}
  
  /* DETECTION TABLE */
  .det-table {{ width:100%;border-collapse:collapse; }}
  .det-table thead td {{ font-size:0.65rem;font-weight:700;text-transform:uppercase;color:#6B7280;padding:4px 6px;border-bottom:1px solid var(--border); }}
  
  /* INTEGRITY GAUGE */
  .ig-wrap {{ position:relative;width:160px;height:90px;margin:8px auto;overflow:visible; }}
  .ig-wrap canvas {{ display:block; }}
  .ig-center {{ position:absolute;bottom:4px;left:50%;transform:translateX(-50%);text-align:center; }}
  .ig-score {{ font-size:1.8rem;font-weight:800; }}
  .ig-label {{ font-size:0.65rem;font-weight:600;margin-top:1px; }}
  .ig-legend {{ display:flex;gap:10px;justify-content:center;margin-top:8px;flex-wrap:wrap; }}
  .ig-legend-item {{ display:flex;align-items:center;gap:4px;font-size:0.65rem;color:var(--text-muted); }}
  .dot {{ width:8px;height:8px;border-radius:50%; }}
  
  /* FACTOR BARS */
  .factor-row {{ margin-bottom:9px; }}
  .factor-header {{ display:flex;justify-content:space-between;margin-bottom:3px;font-size:0.73rem;color:#374151; }}
  .factor-bar-bg {{ height:7px;background:#F3F4F6;border-radius:10px;overflow:hidden;border:1px solid #E5E7EB; }}
  .factor-bar {{ height:100%;border-radius:10px;transition:width 0.8s ease; }}
  
  /* TIMELINE */
  .timeline-list {{ max-height:340px;overflow-y:auto;padding-right:4px; }}
  .timeline-list::-webkit-scrollbar {{ width:4px; }}
  .timeline-list::-webkit-scrollbar-thumb {{ background:#D1D5DB;border-radius:4px; }}
  .tl-item {{ display:flex;align-items:flex-start;gap:8px;padding:8px 0;border-bottom:1px solid #F3F4F6; }}
  .tl-icon {{ font-size:0.85rem;flex-shrink:0;margin-top:1px; }}
  .tl-ts {{ font-size:0.68rem;color:#2563EB;font-weight:600;flex-shrink:0;min-width:40px; }}
  .tl-event {{ font-size:0.73rem;color:#111827;flex:1; }}
  .tl-conf {{ font-size:0.68rem;color:#6B7280;flex-shrink:0; }}
  .tl-img-toggle {{ font-size:0.65rem;color:#7C3AED;cursor:pointer;text-decoration:underline;display:block;margin-top:2px; }}
  .tl-img {{ display:none;margin-top:5px;width:100%;max-width:220px;border-radius:6px;border:1px solid var(--border); }}
  
  /* PERFORMANCE */
  .skill-bar-row {{ margin-bottom:9px; }}
  .skill-bar-header {{ display:flex;justify-content:space-between;font-size:0.73rem;margin-bottom:3px;color:#374151; }}
  .skill-bar-bg {{ height:6px;background:#F3F4F6;border-radius:10px;overflow:hidden;border:1px solid #E5E7EB; }}
  .skill-bar {{ height:100%;border-radius:10px;background:linear-gradient(90deg,#2563EB,#4F46E5); }}
  
 /* SNAPSHOT CAROUSEL — 3-up card layout */
  .snap-section-header {{ display:flex;justify-content:space-between;align-items:center;margin-bottom:14px; }}
  .snap-view-all {{ font-size:0.75rem;font-weight:600;color:#818cf8;cursor:pointer;background:rgba(129,140,248,0.1);padding:5px 14px;border-radius:20px;border:1px solid rgba(129,140,248,0.3);transition:all 0.2s;white-space:nowrap; }}
  .snap-view-all:hover {{ background:rgba(129,140,248,0.22);border-color:#818cf8; }}

  .carousel-outer {{ position:relative; display:flex; align-items:center; width:100%; }}
  .carousel-arrow {{ position:absolute; top:50%; transform:translateY(-50%); width:32px; height:32px; background:#FFFFFF; border:1px solid var(--border); border-radius:50%; color:#374151; font-size:1.1rem; font-weight:bold; display:flex; align-items:center; justify-content:center; cursor:pointer; transition:all 0.2s; user-select:none; z-index:10; box-shadow:0 2px 5px rgba(0,0,0,0.12); }}
  .carousel-arrow:hover {{ background:#2563EB; border-color:#2563EB; color:white; }}
  .carousel-arrow.prev {{ left:-10px; }}
  .carousel-arrow.next {{ right:-10px; }}
  .carousel-arrow.hidden {{ display:none!important; }}

  .snap-3-grid {{ display:grid;grid-template-columns:repeat(3,1fr);gap:10px;flex:1;min-width:0; }}
  .snap-card {{ background:#FFFFFF;border:2px solid var(--border);border-radius:10px;overflow:hidden;cursor:pointer;transition:all 0.25s;box-shadow:0 1px 3px rgba(0,0,0,0.06); }}
  .snap-card:hover {{ border-color:#2563EB;box-shadow:0 0 0 3px rgba(37,99,235,0.12); }}
  .snap-card.active {{ border-color:#2563EB;box-shadow:0 0 0 3px rgba(37,99,235,0.15); }}
  .snap-thumb {{ width:100%;aspect-ratio:4/3;object-fit:cover;display:block;background:#F3F4F6; }}
  .snap-card-body {{ padding:9px 11px; }}
  .snap-badge {{ display:inline-block;font-size:0.64rem;font-weight:700;padding:3px 8px;border-radius:5px;margin-bottom:4px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:100%; }}
  .snap-badge-red {{ background:#FEF2F2;color:#DC2626;border:1px solid #FECACA; }}
  .snap-badge-orange {{ background:#FFFBEB;color:#D97706;border:1px solid #FDE68A; }}
  .snap-badge-yellow {{ background:#FEFCE8;color:#A16207;border:1px solid #FEF08A; }}
  .snap-conf-text {{ font-size:0.68rem;color:#6B7280; }}

  /* CAROUSEL DOTS */
  .carousel-dots {{
    display: flex;
    justify-content: center;
    gap: 6px;
    margin-top: auto;
    padding-top: 14px;
  }}
  .carousel-dot {{
    width: 8px;
    height: 8px;
    border-radius: 50%;
    background: #E5E7EB;
    cursor: pointer;
    transition: all 0.25s ease;
  }}
  .carousel-dot.active {{
    background: #7C3AED;
    width: 20px;
    border-radius: 4px;
  }}
  
  /* VIEW-ALL MODAL */
  .va-overlay {{ display:none;position:fixed;inset:0;background:rgba(0,0,0,0.6);z-index:200;overflow-y:auto;padding:32px 20px; }}
  .va-overlay.open {{ display:block; }}
  .va-box {{ background:#FFFFFF;border:1px solid var(--border);border-radius:16px;padding:26px;width:100%;max-width:920px;margin:0 auto;position:relative;box-shadow:0 20px 60px rgba(0,0,0,0.2); }}
  .va-header {{ display:flex;justify-content:space-between;align-items:center;margin-bottom:20px; }}
  .va-title {{ font-size:1rem;font-weight:700;color:#111827; }}
  .va-close {{ font-size:1.5rem;cursor:pointer;color:#6B7280;line-height:1;transition:color 0.2s;padding:2px 6px; }}
  .va-close:hover {{ color:#111827; }}
  .va-grid {{ display:grid;grid-template-columns:repeat(4,1fr);gap:14px; }}
  .va-card {{ background:#FFFFFF;border:2px solid var(--border);border-radius:10px;overflow:hidden;cursor:pointer;transition:all 0.2s;box-shadow:0 1px 3px rgba(0,0,0,0.06); }}
  .va-card:hover {{ border-color:#2563EB;transform:translateY(-2px);box-shadow:0 6px 20px rgba(37,99,235,0.15); }}
  .va-thumb {{ width:100%;aspect-ratio:4/3;object-fit:cover;display:block;background:#F3F4F6; }}
  .va-body {{ padding:8px 10px; }}
  .va-badge {{ display:inline-block;font-size:0.62rem;font-weight:700;padding:2px 7px;border-radius:4px;margin-bottom:3px;overflow:hidden;text-overflow:ellipsis;max-width:100%; }}
  .va-conf {{ font-size:0.65rem;color:#6B7280; }}

  /* SINGLE IMAGE MODAL */
  .modal-overlay {{ display:none;position:fixed;inset:0;background:rgba(0,0,0,0.6);z-index:300;align-items:center;justify-content:center;backdrop-filter:blur(4px); }}
  .modal-overlay.open {{ display:flex; }}
  .modal-box {{ background:#FFFFFF;border:1px solid #E5E7EB;border-radius:16px;padding:24px;max-width:900px;width:92%;position:relative;box-shadow:0 25px 50px -12px rgba(0,0,0,0.15);display:flex;flex-direction:column;gap:16px; }}
  .modal-hdr {{ display:flex;justify-content:space-between;align-items:center;padding-bottom:4px; }}
  .modal-hdr-left {{ display:flex;align-items:center;gap:16px; }}
  .modal-timestamp {{ font-size:0.95rem;font-weight:700;color:#111827;font-family:monospace; }}
  .modal-event {{ font-size:0.95rem;font-weight:700;color:#EF4444; }}
  .modal-hdr-right {{ display:flex;align-items:center;gap:20px; }}
  .modal-confidence {{ font-size:0.85rem;color:#4B5563;font-weight:600; }}
  .modal-close-btn {{ font-size:1.4rem;cursor:pointer;color:#6B7280;transition:color 0.15s;line-height:1;padding:2px 6px; }}
  .modal-close-btn:hover {{ color:#111827; }}
  .modal-body-wrap {{ display:flex;align-items:center;justify-content:space-between;position:relative;height:50vh;gap:16px; }}
  .modal-img-container {{ flex:1;height:100%;display:flex;align-items:center;justify-content:center;overflow:hidden;background:#F9FAFB;border-radius:12px;border:1px solid #E5E7EB;position:relative; }}
  .modal-img {{ max-width:100%;max-height:100%;object-fit:contain; }}
  .modal-nav-btn {{ width:38px;height:38px;border-radius:50%;background:rgba(255,255,255,0.9);border:1px solid #E5E7EB;color:#374151;font-size:1.4rem;display:flex;align-items:center;justify-content:center;cursor:pointer;transition:all 0.2s;user-select:none;z-index:10;box-shadow:0 2px 5px rgba(0,0,0,0.05); }}
  .modal-nav-btn:hover {{ background:#4F46E5;border-color:#4F46E5;color:#FFFFFF;transform:scale(1.05); }}
  .modal-filmstrip-wrap {{ display:flex;flex-direction:column;align-items:center;gap:12px;width:100%;border-top:1px solid #E5E7EB;padding-top:12px; }}
  .modal-filmstrip-track {{ display:flex;gap:12px;padding:4px 16px;overflow-x:auto;width:100%;justify-content:flex-start;scrollbar-width:thin; }}
  .modal-filmstrip-track::-webkit-scrollbar {{ height:6px; }}
  .modal-filmstrip-track::-webkit-scrollbar-track {{ background:transparent; }}
  .modal-filmstrip-track::-webkit-scrollbar-thumb {{ background:#E5E7EB;border-radius:4px; }}
  .modal-filmstrip-track::-webkit-scrollbar-thumb:hover {{ background:#D1D5DB; }}
  .modal-thumb-container {{ position:relative;width:80px;height:52px;flex-shrink:0;cursor:pointer;border-radius:8px;overflow:hidden;border:2px solid transparent;opacity:0.6;transition:all 0.15s;background:#F3F4F6; }}
  .modal-thumb-container:hover {{ opacity:0.9; }}
  .modal-thumb-container.active {{ border-color:#4F46E5;opacity:1;transform:scale(1.03);box-shadow:0 0 10px rgba(79,70,229,0.25); }}
  .modal-thumb-img {{ width:100%;height:100%;object-fit:cover; }}
  .modal-thumb-time {{ position:absolute;bottom:4px;left:4px;background:rgba(0,0,0,0.7);color:#FFFFFF;font-size:0.58rem;font-family:monospace;padding:1px 4px;border-radius:3px;pointer-events:none; }}
  .modal-indicator {{ font-size:0.75rem;color:#4B5563;font-weight:600;margin-top:2px; }}
    
  /* RECRUITER SUMMARY */
  .rec-summary-container {{
      display: flex;
      flex-direction: column;
      gap: 16px;
      margin-top: 14px;
  }}
  .rec-box {{
      border-radius: 12px;
      padding: 16px 20px;
  }}
  .rec-box.strengths {{
      background: #F0FDF4;
      border: 1px solid rgba(16, 185, 129, 0.15);
  }}
  .rec-box.concerns {{
      background: #FEF2F2;
      border: 1px solid rgba(239, 68, 68, 0.15);
  }}
  .rec-hdr {{
      font-size: 0.8rem;
      font-weight: 700;
      letter-spacing: 0.05em;
      margin-bottom: 12px;
      display: flex;
      align-items: center;
      gap: 6px;
  }}
  .rec-hdr.strengths {{ color: #15803D; }}
  .rec-hdr.concerns {{ color: #B91C1C; }}

  .rec-list {{
      display: flex;
      flex-direction: column;
      gap: 8px;
  }}
  .rec-item {{
      display: flex;
      align-items: flex-start;
      gap: 8px;
      font-size: 0.78rem;
      line-height: 1.4;
  }}
  .rec-item.strength {{ color: #166534; }}
  .rec-item.concern {{ color: #991B1B; }}
  
  /* AI RECOMMENDATION CARD */
  .reco-card {{ border-radius:10px;padding:14px;margin-bottom:10px; }}
  .reco-badge {{ font-size:1rem;font-weight:800;letter-spacing:0.03em;margin-bottom:2px; }}
  .reco-sub {{ font-size:0.7rem;color:rgba(255,255,255,0.75);margin-bottom:8px; }}
  .reco-reason {{ font-size:0.72rem;color:rgba(255,255,255,0.95);line-height:1.5; }}
  .conf-bar-wrap {{ margin-top:10px; }}
  .conf-bar-label {{ display:flex;justify-content:space-between;font-size:0.68rem;color:rgba(255,255,255,0.75);margin-bottom:4px; }}
  .conf-bar-bg {{ height:5px;background:rgba(255,255,255,0.2);border-radius:10px; }}
  .conf-bar-fill {{ height:100%;border-radius:10px;background:#FFFFFF; }}

  /* AI INTERVIEW FEEDBACK CARD */
  .feedback-box {{
      background: linear-gradient(135deg, #F5F3FF 0%, #EFF6FF 100%);
      border: 1px solid #E5E7EB;
      border-radius: 12px;
      padding: 18px 20px;
      margin-top: 14px;
      display: flex;
      flex-direction: column;
      gap: 12px;
  }}
  .feedback-hdr {{
      font-size: 0.8rem;
      font-weight: 700;
      color: #2563EB;
      letter-spacing: 0.05em;
  }}
  .feedback-list {{
      display: flex;
      flex-direction: column;
      gap: 10px;
  }}
  .feedback-item {{
      display: flex;
      align-items: flex-start;
      gap: 10px;
      font-size: 0.82rem;
      color: #374151;
      line-height: 1.4;
  }}
  
  .feedback-link-row {{
      margin-top: 14px;
      border-top: 1px solid #E5E7EB;
      padding-top: 14px;
      display: flex;
      align-items: center;
  }}
  .feedback-link {{
      font-size: 0.82rem;
      font-weight: 700;
      color: #2563EB;
      cursor: pointer;
      display: flex;
      align-items: center;
      gap: 6px;
      text-decoration: none;
      transition: color 0.15s;
  }}
  .feedback-link:hover {{
      color: #1D4ED8;
  }}
  
  /* FULL FEEDBACK MODAL */
  .ff-overlay {{ display:none;position:fixed;inset:0;background:rgba(0,0,0,0.6);z-index:400;align-items:center;justify-content:center;backdrop-filter:blur(4px); }}
  .ff-overlay.open {{ display:flex; }}
  .ff-box {{ background:#FFFFFF;border:1px solid #E5E7EB;border-radius:16px;padding:26px;width:100%;max-width:640px;position:relative;box-shadow:0 25px 50px -12px rgba(0,0,0,0.15);display:flex;flex-direction:column;gap:16px; }}
  .ff-header {{ display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid #F3F4F6;padding-bottom:12px; }}
  .ff-title {{ font-size:1rem;font-weight:700;color:#111827; }}
  .ff-close {{ font-size:1.4rem;cursor:pointer;color:#6B7280;line-height:1;transition:color 0.15s;padding:2px 6px; }}
  .ff-close:hover {{ color:#111827; }}
  .ff-body {{ font-size:0.85rem;color:#374151;line-height:1.6;max-height:400px;overflow-y:auto;padding-right:8px;text-align:justify; }}

  /* PRINT STYLES */
  @media print {{
    .sidebar,.topbar-actions {{ display:none!important; }}
    .main {{ margin-left:0!important; }}
    .topbar {{ position:static!important; }}
    body {{ background:white!important;color:black!important; }}
    .card,.kpi-card {{ border:1px solid #ddd!important;background:#f9f9f9!important; }}
  }}

    /* SKILL RATING BADGES */
  .skill-rating-badge {{
    display: inline-block;
    font-size: 0.62rem;
    font-weight: 700;
    padding: 2px 8px;
    border-radius: 6px;
    white-space: nowrap;
    text-transform: uppercase;
  }}
  .skill-rating-badge.very-low {{
    background: #FEF2F2 !important;
    color: #EF4444 !important;
    border: 1px solid rgba(239, 68, 68, 0.15);
  }}
  .skill-rating-badge.medium {{
    background: #FFFBEB !important;
    color: #D97706 !important;
    border: 1px solid rgba(217, 119, 6, 0.15);
  }}
  .skill-rating-badge.high {{
    background: #ECFDF5 !important;
    color: #10B981 !important;
    border: 1px solid rgba(16, 185, 129, 0.15);
  }}

</style>
</head>
<body>

<!-- SIDEBAR -->
<div class="sidebar">
  <div class="sidebar-logo">
    <div class="logo-icon">AI</div>
    <span>AI ATS</span>
  </div>
  <nav class="sidebar-nav">
    <a href="/api/dashboard">
        <div class="nav-item">🏠 Dashboard</div>
    </a>
    <div class="nav-item active">📁 Reports</div>
  </nav>
  <div class="sidebar-footer">
    <div>Report ID</div>
    <div style="color:var(--white);font-weight:600;font-size:0.72rem">{report_id}</div>
    <div style="margin-top:6px">Generated On</div>
    <div style="color:var(--white);font-weight:600;font-size:0.72rem">{interview_date}</div>
    <div class="ai-badge" style="margin-top:8px">
      <span style="color:#a78bfa;font-size:0.65rem">✦ Gemini 2.5 Flash</span>
    </div>
  </div>
</div>

<!-- MAIN -->
<div class="main">

  <!-- TOP BAR -->
  <div class="topbar">
    <h1>🔍 AI Interview Proctoring Report</h1>
    <div class="topbar-actions">
      <button class="btn btn-primary" onclick="window.print()">⬇ Download PDF</button>
      <button class="btn" onclick="navigator.share && navigator.share({{title:'Proctoring Report',url:window.location.href}})">↗ Share Report</button>
      <button class="btn" onclick="history.back()">← Back to Candidates</button>
    </div>
  </div>

  <!-- CANDIDATE HEADER -->
  <div class="candidate-header">
    <div class="avatar">{initials}</div>
    <div class="cand-info">
      <div style="display:flex;align-items:center;gap:8px">
        <h2>{name}</h2>
        <span class="status-badge">Completed</span>
      </div>
      <div class="subtitle">{email} &nbsp;·&nbsp; {position}</div>
    </div>
    <div class="meta-chips">
      <div class="meta-chip">📅 <div><div style="font-size:0.6rem">Interview Date</div><strong>{interview_date}</strong></div></div>
      <div class="meta-chip">⏱️ <div><div style="font-size:0.6rem">Duration</div><strong>{duration_str}</strong></div></div>
      <div class="meta-chip">💻 <div><div style="font-size:0.6rem">Platform</div><strong>Microsoft Teams</strong></div></div>
      <div class="meta-chip">{"✅" if video_url else "⏳"} <div><div style="font-size:0.6rem">Recording Status</div><strong style="color:{'#0B6623' if video_url else '#D97706'}">{"Available" if video_url else "Processing"}</strong></div></div>
      <div class="meta-chip">🔖 <div><div style="font-size:0.6rem">Report ID</div><strong>{report_id}</strong></div></div>
    </div>
  </div>

  <div class="content">

    <!-- 6 KPI CARDS -->
    <div class="kpi-row">
      <div class="kpi-card" style="--accent:#0B6623">
        <div class="kpi-label">ATS Match Score</div>
        <div class="gauge-wrap"><canvas id="gATS" width="120" height="66"></canvas><div class="gauge-center" style="color:#16A34A">{ats_score}%</div></div>
        <div class="kpi-sub" style="color:#0B6623">{("Excellent" if ats_score>=80 else ("Good Match" if ats_score>=65 else "Fair Match"))}</div>
      </div>
      <div class="kpi-card" style="--accent:{ig_color}">
        <div class="kpi-label">Interview Integrity Score</div>
        <div class="gauge-wrap"><canvas id="gInteg" width="120" height="66"></canvas><div class="gauge-center" style="color:{ig_color}">{integrity_score}%</div></div>
        <div class="kpi-sub" style="color:{ig_color}">{ig_label}</div>
      </div>
      <div class="kpi-card" style="--accent:{risk_color}">
        <div class="kpi-label">Risk Level</div>
        <div style="font-size:1.6rem;margin:4px 0">{risk_icon}</div>
        <div class="kpi-value" style="color:{risk_color};font-size:1rem">{risk_label}</div>
        <div class="kpi-sub">Cheating score: {cheating_score}/100</div>
      </div>
      <div class="kpi-card" style="--accent:#7C3AED">
        <div class="kpi-label">AI Confidence</div>
        <div style="font-size:1.6rem;margin:4px 0">🤖</div>
        <div class="kpi-value" style="color:#7C3AED;font-size:1.05rem;font-weight:800">{ai_conf}%</div>
        <div class="kpi-sub" style="color:#7C3AED">{'High Confidence' if ai_conf>=75 else ('Moderate Confidence' if ai_conf>=50 else 'Low Confidence')}</div>
      </div>
      <div class="kpi-card" style="--accent:{rec_text_color}">
        <div class="kpi-label">Overall Recommendation</div>
        <div style="font-size:1.4rem;margin:4px 0">{rec_icon}</div>
        <div class="kpi-value" style="color:{rec_text_color};font-size:0.8rem;font-weight:800">{recommendation.upper()}</div>
        <div class="kpi-sub">{rec_sub}</div>
      </div>
      <div class="kpi-card" style="--accent:#DC2626">
        <div class="kpi-label">Cheating Risk Score</div>
        <div class="gauge-wrap"><canvas id="gCheat" width="120" height="66"></canvas><div class="gauge-center" style="color:#DC2626">{cheating_score}%</div></div>
        <div class="kpi-sub" style="color:#DC2626">{("High Risk" if cheating_score>60 else ("Moderate Risk" if cheating_score>30 else "Low Risk"))}</div>
      </div>
    </div>

    <!-- MIDDLE ROW: AI Summary | Integrity Gauge+Factors | Timeline -->
    <div class="grid-3">

      <!-- AI PROCTORING SUMMARY -->
      <div class="card">
        <div class="card-title">AI Proctoring Summary</div>
        <table class="det-table">
          <thead>
            <tr>
              <td>Metric</td>
              <td>Status</td>
              <td>Details</td>
            </tr>
          </thead>
          <tbody>{dt_html}</tbody>
        </table>
      </div>

      <!-- CANDIDATE PERFORMANCE -->
      <div class="card" style="padding: 16px; display:flex; flex-direction:column; justify-content:space-between;">
        <div>
          <div class = "card-title" style = "margin-bottom:25px">Candidate Performance</div>
          <div id="skillBars" style="display:flex; flex-direction:column; gap:25px;"></div>
        </div>

        <!-- Overall Performance Footer Box -->
        <div style="background:#F4F7FE; border:1px solid #E2E8F0; border-radius:12px; padding:12px 16px; display:flex; align-items:center; margin-top:16px;">
          <!-- 1. Score -->
          <div style="display:flex; align-items:baseline; gap:2px; flex-shrink:0;">
            <span style="font-size:1.6rem; font-weight:800; color:#2563EB;" id="lblOverallPerfScore">1.0</span>
            <span style="font-size:0.75rem; color:#64748B; font-weight:600;">/10</span>
          </div>

          <!-- Divider 1 -->
          <div style="width:1px; height:28px; background:#E2E8F0; margin:0 16px; flex-shrink:0;"></div>

          <!-- 2. Match tag -->
          <div style="flex-shrink:0; display:flex; flex-direction:column; justify-content:center;">
            <div style="font-size:0.65rem; color:#64748B; font-weight:700; text-transform:uppercase; letter-spacing:0.05em; margin-bottom:2px;">Overall Performance</div>
            <div style="display:inline-block; font-size:0.68rem; font-weight:700; color:#EF4444; background:#FEF2F2; padding:2px 8px; border-radius:6px; width:fit-content; white-space:nowrap;" id="lblOverallPerfLabel">Needs Significant Improvement</div>
          </div>

          <!-- Divider 2 -->
          <div style="width:1px; height:28px; background:#E2E8F0; margin:0 16px; flex-shrink:0;"></div>

          <!-- 3. Commentary -->
          <div style="flex:1; display:flex; align-items:center; gap:8px; min-width:0;">
            <div style="width:20px; height:20px; border-radius:50%; border:1.5px solid #64748B; color:#64748B; display:flex; align-items:center; justify-content:center; font-size:0.7rem; font-weight:700; flex-shrink:0;">i</div>
           <span style="font-size:0.68rem; color:#64748B; font-weight:500; line-height:1.3; overflow:hidden; text-overflow:ellipsis;" id="lblOverallPerfDesc">Performance is below expectations in all areas.</span>
          </div>
        </div>
      </div>

      <!-- INTEGRITY TIMELINE -->
      <div class="card">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px">
          <div class="card-title" style="margin-bottom:0">Integrity Timeline</div>
          <span style="font-size:0.65rem;color:var(--text-muted)">{len(logs)} events</span>
        </div>
        <div class="timeline-list" id="timelineList"></div>
      </div>
    </div>

    <!-- BOTTOM ROW: Performance | Snapshots | Recruiter | Recommendation -->
    <div class="grid-4">

     <!-- AI INTERVIEW FEEDBACK -->
      <div class="card" style="display:flex;flex-direction:column;justify-content:space-between">
        <div>
          <div class="card-title" style="display:flex;align-items:center;gap:8px;text-transform:uppercase;font-size:0.85rem;letter-spacing:0.05em">
            <div class = "card-title">AI Interview Feedback</div>
          </div>
          <div class="feedback-box">
            <div class="feedback-hdr">OVERALL SUMMARY</div>
            <div class="feedback-list">
              {fb_points_html}
            </div>
          </div>
        </div>
        <div class="feedback-link-row">
          <span class="feedback-link" onclick="openFullFeedback()">
            Read Full AI Feedback
            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="5" x2="19" y1="12" y2="12"/><polyline points="12 5 19 12 12 19"/></svg>
          </span>
        </div>
      </div>
      <!-- EVIDENCE SNAPSHOTS -->
      <div class="card">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px">
          <div class="card-title" style="margin-bottom:0">Evidence Snapshots</div>
          <span class="snap-view-all" id="viewAllBtn" onclick="openViewAll()">View All</span>
        </div>
        <div class="carousel-outer" style="margin-top:12px">
          <div class="carousel-arrow prev hidden" id="snapPrev" onclick="snapNav(-1)">&#8249;</div>
          <div class="snap-3-grid" id="snap3Grid"></div>
          <div class="carousel-arrow next hidden" id="snapNext" onclick="snapNav(1)">&#8250;</div>
        </div>
        <div id="noSnapMsg" style="font-size:0.75rem;color:var(--text-muted);text-align:center;padding:24px;display:none">No malpractice screenshots captured</div>
      </div>
    
      <!-- RECRUITER SUMMARY -->
      <div class="card">
        <div class="card-title">Recruiter Summary</div>
        <div class="rec-summary-container">
          <div class="rec-box strengths">
            <div class="rec-hdr strengths">
              <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>
              STRENGTHS
            </div>
            <div class="rec-list">
              {str_html}
            </div>
          </div>
         <div class="rec-box concerns">
            <div class="rec-hdr concerns">
              <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3Z"/><line x1="12" x2="12" y1="9" y2="13"/><line x1="12" x2="12.01" y1="17" y2="17"/></svg>
              CONCERNS
            </div>
            <div class="rec-list">
              {conc_html}
            </div>
          </div>
        </div>
      </div>

      <!-- AI RECOMMENDATION -->
      <div class="card" style="justify-content:space-between; display:flex; flex-direction:column; min-height:100%;">
        <div>
          <div class="card-title" style="margin-bottom:10px">AI Recommendation</div>
          <div class="reco-card" style="background:{rec_bg};border:1px solid {rec_text_color}33;padding:10px 12px;margin-bottom:12px;border-radius:8px">
            <div class="reco-badge" style="color:{rec_text_color};font-size:0.85rem;font-weight:800">{rec_icon} {recommendation.upper()}</div>
            <div class="reco-sub" style="font-size:0.65rem;color:rgba(255,255,255,0.8);margin-bottom:0">{rec_sub}</div>
          </div>
          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px 14px; margin-bottom:16px;">
            <div style="font-size:0.7rem; font-weight:700; color:#2563EB; letter-spacing:0.06em; text-transform:uppercase; margin-bottom:10px;">Reason</div>
            <div style="display:flex; flex-direction:column; gap:8px;">
              {rec_reasons_html}
            </div>
          </div>
        </div>
        <div style="margin-top:auto;">
          <div style="display:flex;justify-content:space-between;font-size:0.65rem;color:#4B5563;margin-bottom:6px;font-weight:600">
            <span>AI Analysis Confidence</span>
            <span style="color:#2563EB;font-weight:700">{ai_conf}%</span>
          </div>
          <div style="height:6px;background:#F3F4F6;border-radius:10px;overflow:hidden;border:1px solid #E5E7EB;margin-bottom:16px;">
            <div style="height:100%;width:{min(ai_conf,100)}%;background:linear-gradient(90deg,#2563EB,#4F46E5);border-radius:10px"></div>
          </div>

          <div style="border-top:1px solid var(--border); padding-top:12px; display:flex; justify-content:center;">
            <span class="feedback-link" onclick="openFullFeedback()" style="cursor:pointer; display:flex; align-items:center; gap:4px; transition:color 0.15s;">
              View Full Analysis Report
              <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="5" x2="19" y1="12" y2="12"/><polyline points="12 5 19 12 12 19"/></svg>
            </span>
          </div>

        </div>
      </div>

    </div><!-- /grid-4 -->
  </div><!-- /content -->

  <div style="padding:10px 24px;font-size:0.65rem;color:var(--text-muted);border-top:1px solid var(--border);text-align:center">
    Note: This report is AI-generated based on video, audio and behavioral analysis. Results are indicative and should be reviewed by the recruiter before making a final decision.
  </div>
</div><!-- /main -->

<!-- VIEW-ALL SNAPSHOTS POPUP -->
<div class="va-overlay" id="vaOverlay" onclick="if(event.target===this)closeViewAll()">
  <div class="va-box">
    <div class="va-header">
      <div class="va-title" id="vaTitle">All Evidence Snapshots</div>
      <span class="va-close" onclick="closeViewAll()">&#10005;</span>
    </div>
    <div class="va-grid" id="vaGrid"></div>
  </div>
</div>

<!-- MODAL for single full-size screenshot -->
<div class="modal-overlay" id="modalOverlay" onclick="if(event.target===this)closeModal()">
  <div class="modal-box">
    <!-- Top Header Bar -->
    <div class="modal-hdr">
      <div class="modal-hdr-left">
        <span class="modal-timestamp" id="modalTimestamp">00:00:00</span>
        <span class="modal-event" id="modalTitle">Event Title</span>
      </div>
      <div class="modal-hdr-right">
        <span class="modal-confidence" id="modalConfidence">Confidence: 100%</span>
        <span class="modal-close-btn" onclick="closeModal()">&#10005;</span>
      </div>
    </div>
    
    <!-- Main Content Area with Image and Arrows -->
    <div class="modal-body-wrap">
      <button class="modal-nav-btn prev" id="modalPrev" onclick="modalNavigate(-1)">&#8249;</button>
      <div class="modal-img-container">
        <img class="modal-img" id="modalImg" src="" alt="Evidence" onerror="if(this.src.includes('/incidents/')) {{ this.src=this.src.replace('/incidents/', '/'); }} else {{ this.style.opacity='0.3'; }}">
      </div>
      <button class="modal-nav-btn next" id="modalNext" onclick="modalNavigate(1)">&#8250;</button>
    </div>

    <!-- Thumbnails Filmstrip Carousel at the Bottom -->
    <div class="modal-filmstrip-wrap">
      <div class="modal-filmstrip-track" id="modalFilmstrip">
        <!-- Thumbnails injected via JS -->
      </div>
      <div class="modal-indicator" id="modalIndicator">1 / 1</div>
    </div>
  </div>
</div>

<script>
// ── Data from Python backend ──────────────────────────────────────────────
const logs        = {logs_json};
const snapshots   = {snapshots_json};

// ── Client-side image URL normalizer ─────────────────────────────────────
// Converts absolute file paths stored in MongoDB to browser-accessible URLs.
// Works as a safety net even if server-side normalization hasn't run yet.
function normalizeImgUrl(url) {{
  if (!url) return url;
  if (url.startsWith('/static/') || url.startsWith('http://') || url.startsWith('https://')) return url;
  // Extract just the filename from any absolute Linux or Windows path
  const fname = url.replace(/\\\\/g, '/').split('/').pop();
  return '/static/recordings/incidents/' + fname;
}}
logs.forEach(l      => {{ if (l.image_url) l.image_url = normalizeImgUrl(l.image_url); }});
snapshots.forEach(s => {{ if (s.image_url) s.image_url = normalizeImgUrl(s.image_url); }});
// ─────────────────────────────────────────────────────────────────────────

const rbScores    = {rb_scores_json};
const rbMaxes     = {rb_maxes_json};
const rbLabels    = {rb_labels_json};
const skillLabels = {skill_labels_json};
const skillRatings= {skill_ratings_json};
const strengths   = {strengths_json};
const concerns    = {concerns_json};
const atsScore    = {ats_score};
const integScore  = {integrity_score};
const cheatScore  = {cheating_score};
const aiConf      = {ai_conf};

// ── Gauge helper (semicircle doughnut) ───────────────────────────────────
function makeGauge(id, value, max, color) {{
    const canvas = document.getElementById(id);
    if (!canvas) {{
        console.error("Canvas not found:", id);
        return;
    }}
    if (typeof Chart === "undefined") {{
        console.error("Chart.js not loaded.");
        return;
    }}
    const ctx = canvas.getContext("2d");
    const existingChart = Chart.getChart(canvas);
    if (existingChart) {{
        existingChart.destroy();
    }}
    value = Math.max(0, Math.min(value, max));

    new Chart(ctx, {{
      type: "doughnut",
        data: {{
            datasets: [{{
                data: [
                    value,
                    max - value
                ],
                backgroundColor: [
                    color,
                    "#E5E7EB"
                ],
                borderWidth: 0
            }}]
        }},
        options: {{
            responsive: false,
            maintainAspectRatio: false,
            rotation: 270,
            circumference: 180,
            cutout: "72%",
            animation: {{
                duration: 1200
            }},
            plugins: {{
                legend: {{
                    display: false
                }},
                tooltip: {{
                    enabled: false
                }}
            }}
        }}
    }});
}}

// KPI Gauges
makeGauge(
    "gATS",
    atsScore,
    100,
    "#16A34A"
);
makeGauge(
    "gInteg",
    integScore,
    100,
    integScore >= 71
        ? "#16A34A"
        : integScore >= 41
        ? "#F59E0B"
        : "#DC2626"
);
makeGauge(
    "gCheat",
    cheatScore,
    100,
    cheatScore <= 30
        ? "#16A34A"
        : cheatScore <= 60
        ? "#F59E0B"
        : "#DC2626"
);
// BIG INTEGRITY GAUGE
(function () {{
    const canvas = document.getElementById("gIntegBig");
    if (!canvas) return;
    const ctx = canvas.getContext("2d");
    const oldChart = Chart.getChart(canvas);
    if (oldChart) {{
        oldChart.destroy();
    }}
    const color =
        integScore >= 71
            ? "#16A34A"
            : integScore >= 41
            ? "#F59E0B"
            : "#DC2626";
    new Chart(ctx, {{
        type: "doughnut",
        data: {{
            datasets: [{{
                data: [
                    integScore,
                    100 - integScore
                ],
                backgroundColor: [
                    color,
                    "#ECECEC"
                ],
                borderWidth: 0
            }}]
        }},
        options: {{
            responsive: false,
            maintainAspectRatio: false,
            rotation: 270,
            circumference: 180,
            cutout: "68%",
            plugins: {{
                legend: {{
                    display: false
                }},
                tooltip: {{
                    enabled: false
                }}
            }}
        }}
    }});
}})();

// Integrity Factor Bars
(function () {{
    const wrap = document.getElementById("factorBars");
    if (!wrap) return; // element not present in this layout - skip safely
    wrap.innerHTML = "";
    const colors = [
        "#16A34A",
        "#3B82F6",
        "#8B5CF6",
        "#F97316",
        "#FBBF24"
    ];
    rbLabels.forEach((label, i) => {{
        const score = rbScores[i] || 0;
        const max = rbMaxes[i] || 1;
        const percent = (score / max) * 100;
        wrap.innerHTML += `
<div class="factor-row">
<div class="factor-header">
<span>\${{label}}</span>
<span style="color:\${{colors[i % colors.length]}}">\${{score}}/\${{max}}</span>
</div>
<div class="factor-bar-bg">
<div class="factor-bar" style="width:\${{percent}}%;background:\${{colors[i % colors.length]}};"></div>
</div>
</div>
`;
    }});
}})();

// ── Integrity Timeline ────────────────────────────────────────────────────
(function() {{
  const list = document.getElementById('timelineList');
  if (!list) return;
  const iconMap = {{
    'phone':'📱', 'cell':'📱', 'looking':'👁️', 'gaze':'👁️',
    'head':'↩️', 'face':'👤', 'absent':'🚫', 'multiple':'👥',
    'camera':'📷', 'book':'📖', 'noise':'🔊', 'voice':'🔊'
  }};
  if (!logs || logs.length === 0) {{
    list.innerHTML = '<div style="text-align:center;color:var(--text-muted);font-size:0.75rem;padding:20px">No integrity events detected. Clean session!</div>';
    return;
  }}
  logs.forEach((log, idx) => {{
    const evLower = log.event.toLowerCase();
    let icon = '⚠️';
    for (const [k, v] of Object.entries(iconMap)) {{
      if (evLower.includes(k)) {{ icon = v; break; }}
    }}
    const confPct = Math.round((log.confidence||0)*100);
    const imgHtml = log.image_url
      ? `<span class="tl-img-toggle" onclick="toggleTlImg(${{idx}})">📷 View screenshot</span>
         <img class="tl-img" id="tlimg_${{idx}}" src="${{log.image_url}}" alt="Evidence" onclick="openModal('${{log.image_url}}','${{log.event}}','${{log.timestamp}} · Confidence: ${{confPct}}%')">`
      : '';
    list.innerHTML += `
      <div class="tl-item">
        <span class="tl-icon">${{icon}}</span>
        <span class="tl-ts">${{log.timestamp}}</span>
        <div class="tl-event">${{log.event}}${{imgHtml}}</div>
        <span class="tl-conf">${{confPct}}%</span>
      </div>`;
  }});
}})();

function toggleTlImg(idx) {{
  const el = document.getElementById('tlimg_'+idx);
  if (el) el.style.display = el.style.display === 'block' ? 'none' : 'block';
}}

// ── Evidence Snapshots - 5-up paged carousel ───────────────────────────────────────────
(function() {{
    const grid    = document.getElementById('snap3Grid');
    const noMsg   = document.getElementById('noSnapMsg');
    const vBtn    = document.getElementById('viewAllBtn');
    const prevBtn = document.getElementById('snapPrev');
    const nextBtn = document.getElementById('snapNext');

    if (!snapshots || snapshots.length === 0) {{
        if (grid) grid.style.display = 'none';
        if (noMsg) noMsg.style.display = 'block';
        if (vBtn)  vBtn.style.display = 'none';
        return;
  }}

  if (vBtn) vBtn.textContent = `View All (${{snapshots.length}})`;

  const PAGE = 6;
  let currentPage = 0;
  const totalPages = Math.ceil(snapshots.length / PAGE);

  function eventColor(event) {{
    const ev = (event || '').toLowerCase();
    if (ev.includes('phone') || ev.includes('multiple') || ev.includes('absent') || ev.includes('person') || ev.includes('object') || ev.includes('unauthori')) return '#DC2626'; // red
    if (ev.includes('gaze')  || ev.includes('looking')  || ev.includes('away')   || ev.includes('head') || ev.includes('turn')) return '#D97706'; // orange
    return '#B45309'; // yellow
  }}

  function renderPage(page) {{
    grid.innerHTML = '';
    const start = page * PAGE;
    const end   = Math.min(start + PAGE, snapshots.length);

    for (let i = start; i < end; i++) {{
      const snap = snapshots[i];
      const conf = Math.round((snap.confidence || 0) * 100);
      const card = document.createElement('div');
      card.className = 'snap-card' + (i === start ? ' active' : '');
      card.innerHTML = `
        <div style="position:relative">
          <img class="snap-thumb" src="${{snap.image_url}}" alt="${{snap.event}}" onerror="this.style.opacity='0.3'">
          <div style="position:absolute; bottom:4px; left:4px; background:rgba(0,0,0,0.65); color:white; font-size:0.55rem; font-weight:700; padding:1px 5px; border-radius:3px; font-family:monospace">${{snap.timestamp}}</div>
        </div>
        <div class="snap-card-body" style="padding:4px 6px; text-align:left">
          <div style="font-size:0.6rem; font-weight:700; color:${{eventColor(snap.event)}}; margin-bottom:1px; overflow:hidden; text-overflow:ellipsis; white-space:normal" title="${{snap.event}}">${{snap.event}}</div>
          <div style="font-size:0.56rem; color:#6B7280">Confidence: ${{conf}}%</div>
        </div>`;
      card.onclick = () => openModal(snap.image_url, snap.event, snap.timestamp + ' · Confidence: ' + conf + '%');
      grid.appendChild(card);
    }}

    // invisible placeholders to keep grid 3-col
    for (let i = end - start; i < PAGE; i++) {{
      const ph = document.createElement('div');
      ph.style.visibility = 'hidden';
      grid.appendChild(ph);
    }}

    prevBtn.classList.toggle('hidden', page === 0);
    nextBtn.classList.toggle('hidden', page >= totalPages - 1);
    renderDots();
  }}

  // Create dots wrap
  const dotsWrap = document.createElement('div');
  dotsWrap.className = 'carousel-dots';
  dotsWrap.style.marginTop = 'auto';
  dotsWrap.style.paddingTop = '14px';
  const carouselOuter = document.querySelector('.carousel-outer');
  if (carouselOuter) {{
    carouselOuter.parentNode.appendChild(dotsWrap);
  }}

  function renderDots() {{
    dotsWrap.innerHTML = '';
    if (totalPages <= 1) return;
    for (let i = 0; i < totalPages; i++) {{
      const dot = document.createElement('div');
      dot.className = 'carousel-dot' + (i === currentPage ? ' active' : '');
      dot.onclick = () => {{
        currentPage = i;
        renderPage(currentPage);
      }};
      dotsWrap.appendChild(dot);
    }}
  }}

  window.snapNav = function(dir) {{
    currentPage = Math.max(0, Math.min(totalPages - 1, currentPage + dir));
    renderPage(currentPage);
  }};

  renderPage(0);

}})();

// ── View-All popup ────────────────────────────────────────────────────────
window.openViewAll = function() {{
  const overlay = document.getElementById('vaOverlay');
  const vaGrid  = document.getElementById('vaGrid');
  const title   = document.getElementById('vaTitle');
  if (title) title.textContent = 'All Evidence Snapshots (' + snapshots.length + ')';
  vaGrid.innerHTML = '';

  snapshots.forEach(snap => {{
    if (!snap) return;
    const conf = Math.round((snap.confidence || 0) * 100);
    const ev   = (snap.event || '').toLowerCase();
    let cls = 'va-badge snap-badge-yellow';
    if (ev.includes('phone') || ev.includes('multiple') || ev.includes('absent') || ev.includes('person') || ev.includes('object') || ev.includes('unauthori')) cls = 'va-badge snap-badge-red';
    else if (ev.includes('gaze') || ev.includes('looking') || ev.includes('away') || ev.includes('head') || ev.includes('turn')) cls = 'va-badge snap-badge-orange';

    const card = document.createElement('div');
    card.className = 'va-card';
    card.innerHTML = `
        <img class="va-thumb" src="${{snap.image_url}}" alt="${{snap.event}}" 
           onerror="if(this.src.includes('/incidents/')) {{ this.src=this.src.replace('/incidents/', '/'); }} else {{ this.style.opacity='0.3'; }}"
           style="pointer-events:none;">
        <div class="va-body" style="pointer-events:none;">
        <div class="${{cls}}">${{snap.event}}</div>
        <div class="va-conf">Confidence: ${{conf}}%</div>
      </div>`;
    card.onclick = () => {{ const activeImg = card.querySelector('.va-thumb'); const imgSrc = activeImg ? activeImg.src : snap.image_url;
    console.log("[Report Dashboard] Opening snapshot modal:", imgSrc, snap.event);
    openModal(imgSrc, snap.event, snap.timestamp + ' - Confidence: ' + conf + '%');
       }};
    vaGrid.appendChild(card);
  }});

  overlay.classList.add('open');
  document.body.style.overflow = 'hidden';
}};

window.closeViewAll = function() {{
  document.getElementById('vaOverlay').classList.remove('open');
  document.body.style.overflow = '';
}};

// ── Skill Bars ────────────────────────────────────────────────────────────
(function() {{
  const wrap = document.getElementById('skillBars');
  if (!wrap) return; // element not present in this layout - skip safely

   const skillIcons = {{
    'technical': '<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#2563EB" stroke-width="2.5"><polyline points="16 18 22 12 16 6"/><polyline points="8 6 2 12 8 18"/></svg>',
    'problem': '<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#10B981" stroke-width="2.5"><path d="M12 22c5.522 0 10-4.477 10-10S17.522 2 12 2 2 6.477 2 12s4.478 10 10 10z"/><path d="M12 8v4l3 3"/></svg>',
    'communication': '<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#D97706" stroke-width="2.5"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg>',
    'confidence': '<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#7C3AED" stroke-width="2.5"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>'
  }};
  
  const skillBgs = {{
    'technical': '#EFF6FF',
    'problem': '#ECFDF5',
    'communication': '#FFFBEB',
    'confidence': '#F5F3FF'
  }};

  skillLabels.forEach((lbl, i) => {{
    const r = skillRatings[i] || 5;
    const pct = r * 10;

    const lLower = lbl.toLowerCase();
    let key = 'technical';
    if (lLower.includes('problem') || lLower.includes('solving')) key = 'problem';
    else if (lLower.includes('commun')) key = 'communication';
    else if (lLower.includes('confid')) key = 'confidence';
    
    const icon = skillIcons[key];
    const bg = skillBgs[key];
    
    let badgeClass = 'medium';
    let badgeText = 'Medium';
    if (r >= 8) {{
      badgeClass = 'high';
      badgeText = 'High';
    }} else if (r < 5) {{
      badgeClass = 'very-low';
      badgeText = 'Very Low';
    }}

    wrap.innerHTML += `
      <div style="display: flex; align-items: center; justify-content: space-between;">
        <!-- Left: Custom Square Icon -->
        <div style="width: 32px; height: 32px; border-radius: 8px; background:${{bg}}; display: flex; align-items: center; justify-content: center; flex-shrink: 0;">
          ${{icon}}
        </div>
        
        <!-- Middle: Title and Progress Bar -->
        <div style="flex: 1; margin-left: 12px; margin-right: 20px; display: flex; flex-direction: column; gap: 6px;">
          <span class="skill-title-text" style="font-size: 0.8rem; font-weight: 700; color: #1E293B; margin: 0;">${{lbl}}</span>
          <div style="height: 6px; background: #F1F5F9; border-radius: 10px; overflow: hidden; width: 100%;">
            <div style="height: 100%; width: ${{pct}}%; background: #2563EB; border-radius: 10px;"></div>
          </div>
        </div>
        
        <!-- Right: Stacked Score and Pill Badge -->
        <div style="display: flex; flex-direction: column; align-items: center; gap: 4px; flex-shrink: 0; min-width: 60px;">
          <span style="font-size: 0.8rem; font-weight: 700; color: #2563EB;">${{r}}/10</span>
          <span class="skill-rating-badge ${{badgeClass}}" style="font-size: 0.62rem; font-weight: 700; padding: 2px 8px; border-radius: 6px; white-space: nowrap;">${{badgeText}}</span>

        </div>
      </div>`;
  }});

   const overallScoreVal = (atsScore / 10).toFixed(1);
  const perfScoreEl = document.getElementById('lblOverallPerfScore');
  if (perfScoreEl) perfScoreEl.textContent = overallScoreVal;
  
  const perfLabelEl = document.getElementById('lblOverallPerfLabel');
  const perfDescEl = document.getElementById('lblOverallPerfDesc');
  if (perfLabelEl && perfDescEl) {{
      if (atsScore >= 80) {{
          perfLabelEl.textContent = 'Excellent Match';
          perfLabelEl.style.color = '#10B981';
          perfDescEl.textContent = 'Candidate demonstrates exceptional alignment with role requirements.';
      }} else if (atsScore >= 65) {{
          perfLabelEl.textContent = 'Good Match';
          perfLabelEl.style.color = '#3B82F6';
          perfDescEl.textContent = 'Candidate meets key competency criteria for this position.';
      }} else if (atsScore >= 50) {{
          perfLabelEl.textContent = 'Satisfactory';
          perfLabelEl.style.color = '#F59E0B';
          perfDescEl.textContent = 'Candidate meets basic requirements but has areas for improvement.';
      }} else {{
          perfLabelEl.textContent = 'Needs Significant Improvement';
          perfLabelEl.style.color = '#EF4444';
          perfDescEl.textContent = 'Performance is below expectations in all areas.';
      }}
  }}

}})();

// ── Modal ─────────────────────────────────────────────────────────────────
let activeSnapshotIndex = 0;

window.openModal = function(src, title, label) {{
  if (!snapshots || snapshots.length === 0) return;
  
  // Find which snapshot has this src or is matched
  const idx = snapshots.findIndex(s => s.image_url === src || src.endsWith(s.image_url.split('/').pop()));
  if (idx !== -1) {{
    activeSnapshotIndex = idx;
  }} else {{
    activeSnapshotIndex = 0;
  }}
  
  window.renderActiveModalSnapshot();
  document.getElementById('modalOverlay').classList.add('open');
document.body.style.overflow = 'hidden';
}};

window.renderActiveModalSnapshot = function() {{
  if (!snapshots || snapshots.length === 0) return;
  const snap = snapshots[activeSnapshotIndex];
  
  // Set images and headers
  const modalImg = document.getElementById('modalImg');
  modalImg.style.opacity = '1';
  modalImg.src = snap.image_url;
  
  // Render event title
  const eventEl = document.getElementById('modalTitle');
  eventEl.textContent = snap.event || 'Incident Flagged';
  
  // Format event label color based on severity/event type
  const ev = (snap.event || '').toLowerCase();
  if (ev.includes('phone') || ev.includes('multiple') || ev.includes('absent') || ev.includes('person') || ev.includes('object') || ev.includes('unauthori')) {{
    eventEl.style.color = '#EF4444'; // Red
  }} else {{
    eventEl.style.color = '#F97316'; // Orange
  }}

  // Format confidence
  const conf = Math.round((snap.confidence || 0) * 100);
  document.getElementById('modalConfidence').textContent = 'Confidence: ' + conf + '%';
  
  // Format timestamp
  document.getElementById('modalTimestamp').textContent = snap.timestamp || '00:00';
  
  // Render filmstrip of thumbnails
  const filmstrip = document.getElementById('modalFilmstrip');
  filmstrip.innerHTML = '';
  snapshots.forEach((s, idx) => {{
    const container = document.createElement('div');
    container.className = 'modal-thumb-container' + (idx === activeSnapshotIndex ? ' active' : '');
    
    // Create image
    const img = document.createElement('img');
    img.className = 'modal-thumb-img';
    img.src = s.image_url;
    img.alt = s.event;
    img.onerror = function() {{
      if (this.src.includes('/incidents/')) {{
        this.src = this.src.replace('/incidents/', '/');
      }} else {{
        this.style.opacity = '0.3';
      }}
    }};

    // Create timestamp overlay
    const timeLabel = document.createElement('span');
    timeLabel.className = 'modal-thumb-time';
    timeLabel.textContent = s.timestamp || '00:00';
    
    container.appendChild(img);
    container.appendChild(timeLabel);
    
    container.onclick = () => {{
      activeSnapshotIndex = idx;
      window.renderActiveModalSnapshot();
    }};
    filmstrip.appendChild(container);
  }});

  // Highlight active thumbnail and scroll into view if needed
  const activeThumb = filmstrip.children[activeSnapshotIndex];
  if (activeThumb) {{
    activeThumb.scrollIntoView({{ behavior: 'smooth', block: 'nearest', inline: 'center' }});
  }}

    // Update slide/page indicator (e.g. "4 / 7")
  const indicator = document.getElementById('modalIndicator');
  if (indicator) {{
    indicator.textContent = (activeSnapshotIndex + 1) + ' / ' + snapshots.length;
  }}

}};

window.modalNavigate = function(dir) {{
  if (!snapshots || snapshots.length === 0) return;
  activeSnapshotIndex = (activeSnapshotIndex + dir + snapshots.length) % snapshots.length;
  window.renderActiveModalSnapshot();
}};

window.closeModal = function() {{
  document.getElementById('modalOverlay').classList.remove('open');
if (!document.getElementById('vaOverlay').classList.contains('open')) {{
    document.body.style.overflow = '';
  }}
}};

document.addEventListener('keydown', e => {{ 
  if (e.key === 'Escape') {{ window.closeModal(); window.closeFullFeedback(); }}
  if (e.key === 'ArrowLeft' && document.getElementById('modalOverlay').classList.contains('open')) window.modalNavigate(-1);
  if (e.key === 'ArrowRight' && document.getElementById('modalOverlay').classList.contains('open')) window.modalNavigate(1);
}});

window.openFullFeedback = function() {{
  document.getElementById('fullFeedbackModal').classList.add('open');
  document.body.style.overflow = 'hidden';
}};

window.closeFullFeedback = function() {{
  document.getElementById('fullFeedbackModal').classList.remove('open');
  if (!document.getElementById('modalOverlay').classList.contains('open') && !document.getElementById('vaOverlay').classList.contains('open')) {{
    document.body.style.overflow = '';
  }}
}};

</script>

<!-- FULL AI FEEDBACK MODAL -->
<div class="ff-overlay" id="fullFeedbackModal">
  <div class="ff-box">
    <div class="ff-header">
      <div class="ff-title">Detailed AI Interview Evaluation</div>
      <div class="ff-close" onclick="closeFullFeedback()">&times;</div>
    </div>
    <div class="ff-body" id="fullFeedbackBody">
      {detailed_feedback}
    </div>
  </div>
</div>

</body>
</html>
"""
    return HTMLResponse(content=html_content)
