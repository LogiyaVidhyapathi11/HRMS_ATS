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
    Serves the HR candidate management dashboard.
    Designed with a highly-premium dark glassmorphic style.
    """
    html_content = """
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>HRMS AI Interview Dashboard</title>
        <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
        <style>
            :root {
                --bg-primary: #0a0f1d;
                --bg-secondary: rgba(17, 24, 39, 0.7);
                --accent-primary: #6366f1;
                --accent-secondary: #06b6d4;
                --text-primary: #f3f4f6;
                --text-secondary: #9ca3af;
                --card-border: rgba(255, 255, 255, 0.08);
                --success: #10b981;
                --warning: #f59e0b;
                --danger: #ef4444;
            }

            * {
                box-sizing: border-box;
                margin: 0;
                padding: 0;
            }

            body {
                font-family: 'Inter', sans-serif;
                background-color: var(--bg-primary);
                color: var(--text-primary);
                min-height: 100vh;
                background-image: 
                    radial-gradient(at 10% 20%, rgba(99, 102, 241, 0.15) 0px, transparent 50%),
                    radial-gradient(at 90% 10%, rgba(6, 182, 212, 0.1) 0px, transparent 50%);
                background-attachment: fixed;
                padding: 2.5rem 5rem;
            }

            h1, h2, h3, .brand {
                font-family: 'Outfit', sans-serif;
            }

            .container {
                max-width: 1200px;
                margin: 0 auto;
            }

            header {
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin-bottom: 3rem;
                padding-bottom: 1.5rem;
                border-bottom: 1px solid var(--card-border);
            }

            .brand-container {
                display: flex;
                align-items: center;
                gap: 0.75rem;
            }

            .logo-orb {
                width: 40px;
                height: 40px;
                background: linear-gradient(135deg, var(--accent-primary), var(--accent-secondary));
                border-radius: 12px;
                box-shadow: 0 0 20px rgba(99, 102, 241, 0.4);
                display: flex;
                align-items: center;
                justify-content: center;
                font-weight: 800;
                font-size: 1.25rem;
            }

            .brand h1 {
                font-size: 1.75rem;
                font-weight: 700;
                background: linear-gradient(to right, #ffffff, #d1d5db);
                -webkit-background-clip: text;
                -webkit-text-fill-color: transparent;
            }

            .brand p {
                font-size: 0.875rem;
                color: var(--text-secondary);
                margin-top: 0.25rem;
            }

            /* Stats Counter Grid */
            .stats-grid {
                display: grid;
                grid-template-columns: repeat(4, 1fr);
                gap: 1.5rem;
                margin-bottom: 3rem;
            }

            .stat-card {
                background: var(--bg-secondary);
                border: 1px solid var(--card-border);
                backdrop-filter: blur(12px);
                border-radius: 16px;
                padding: 1.5rem;
                display: flex;
                flex-direction: column;
                gap: 0.5rem;
                transition: transform 0.2s, border-color 0.2s;
            }

            .stat-card:hover {
                transform: translateY(-2px);
                border-color: rgba(99, 102, 241, 0.3);
            }

            .stat-label {
                font-size: 0.875rem;
                color: var(--text-secondary);
                font-weight: 500;
            }

            .stat-value {
                font-size: 2.25rem;
                font-weight: 700;
                color: #ffffff;
                font-family: 'Outfit', sans-serif;
            }

            /* Candidates Table Card */
            .table-card {
                background: var(--bg-secondary);
                border: 1px solid var(--card-border);
                backdrop-filter: blur(12px);
                border-radius: 20px;
                overflow: hidden;
                box-shadow: 0 10px 30px rgba(0,0,0,0.3);
            }

            .table-header {
                padding: 1.5rem 2rem;
                border-bottom: 1px solid var(--card-border);
                display: flex;
                justify-content: space-between;
                align-items: center;
            }

            .table-title {
                font-size: 1.25rem;
                font-weight: 600;
                color: #ffffff;
            }

            table {
                width: 100%;
                border-collapse: collapse;
                text-align: left;
            }

            th {
                padding: 1.25rem 2rem;
                font-size: 0.75rem;
                text-transform: uppercase;
                letter-spacing: 0.05em;
                color: var(--text-secondary);
                border-bottom: 1px solid var(--card-border);
                font-weight: 600;
            }

            td {
                padding: 1.5rem 2rem;
                font-size: 0.875rem;
                color: var(--text-primary);
                border-bottom: 1px solid var(--card-border);
                vertical-align: middle;
            }

            tr:last-child td {
                border-bottom: none;
            }

            tr:hover td {
                background-color: rgba(255, 255, 255, 0.02);
            }

            .candidate-info {
                display: flex;
                flex-direction: column;
                gap: 0.25rem;
            }

            .candidate-name {
                font-weight: 600;
                color: #ffffff;
                font-size: 1rem;
            }

            .candidate-email {
                color: var(--text-secondary);
                font-size: 0.8rem;
            }

            /* Badges */
            .badge {
                display: inline-flex;
                align-items: center;
                padding: 0.35rem 0.75rem;
                border-radius: 9999px;
                font-size: 0.75rem;
                font-weight: 600;
                letter-spacing: 0.02em;
            }

            .badge-scheduled {
                background-color: rgba(245, 158, 11, 0.15);
                color: var(--warning);
                border: 1px solid rgba(245, 158, 11, 0.3);
            }

            .badge-completed {
                background-color: rgba(16, 185, 129, 0.15);
                color: var(--success);
                border: 1px solid rgba(16, 185, 129, 0.3);
                box-shadow: 0 0 10px rgba(16, 185, 129, 0.1);
            }

            .score-pill {
                display: inline-flex;
                justify-content: center;
                align-items: center;
                width: 48px;
                height: 48px;
                border-radius: 12px;
                font-weight: 700;
                font-size: 1.1rem;
                font-family: 'Outfit', sans-serif;
            }

            .score-high {
                background: rgba(16, 185, 129, 0.12);
                color: var(--success);
                border: 1px solid rgba(16, 185, 129, 0.25);
            }

            .score-medium {
                background: rgba(245, 158, 11, 0.12);
                color: var(--warning);
                border: 1px solid rgba(245, 158, 11, 0.25);
            }

            .score-low {
                background: rgba(239, 68, 68, 0.12);
                color: var(--danger);
                border: 1px solid rgba(239, 68, 68, 0.25);
            }

            .btn {
                background: linear-gradient(135deg, var(--accent-primary), rgba(99, 102, 241, 0.8));
                color: #ffffff;
                border: none;
                padding: 0.6rem 1.2rem;
                border-radius: 10px;
                font-weight: 600;
                font-size: 0.825rem;
                cursor: pointer;
                transition: all 0.2s;
                text-decoration: none;
                display: inline-flex;
                align-items: center;
                box-shadow: 0 4px 15px rgba(99, 102, 241, 0.2);
            }

            .btn:hover {
                transform: translateY(-1px);
                box-shadow: 0 6px 20px rgba(99, 102, 241, 0.35);
            }

            .btn-disabled {
                background: rgba(255, 255, 255, 0.05);
                color: var(--text-secondary);
                cursor: not-allowed;
                box-shadow: none;
            }
            
            .btn-disabled:hover {
                transform: none;
                box-shadow: none;
            }

            /* Responsive */
            @media (max-width: 1024px) {
                body {
                    padding: 2rem;
                }
                .stats-grid {
                    grid-template-columns: repeat(2, 1fr);
                }
            }
        </style>
    </head>
    <body>
        <div class="container">
            <header>
                <div class="brand-container">
                    <div class="logo-orb">AB</div>
                    <div class="brand">
                        <h1>Candidate Evaluation Dashboard</h1>
                        <p>Adams Bridge AI HRMS Integration</p>
                    </div>
                </div>
            </header>

            <div class="stats-grid">
                <div class="stat-card">
                    <span class="stat-label">Total Candidates</span>
                    <span class="stat-value" id="stat-total">0</span>
                </div>
                <div class="stat-card">
                    <span class="stat-label">Completed Interviews</span>
                    <span class="stat-value" id="stat-completed">0</span>
                </div>
                <div class="stat-card">
                    <span class="stat-label">Scheduled Interviews</span>
                    <span class="stat-value" id="stat-scheduled">0</span>
                </div>
                <div class="stat-card">
                    <span class="stat-label">Average ATS Match Score</span>
                    <span class="stat-value" id="stat-avg-score">0%</span>
                </div>
            </div>

            <div class="table-card">
                <div class="table-header">
                    <h2 class="table-title">Recent Candidates</h2>
                </div>
                <table>
                    <thead>
                        <tr>
                            <th>Candidate Details</th>
                            <th>Status</th>
                            <th>ATS Match Score</th>
                            <th>Cheating Risk</th>
                            <th>Action</th>
                        </tr>
                    </thead>
                    <tbody id="candidates-body">
                        <!-- Loaded dynamically -->
                        <tr>
                            <td colspan="5" style="text-align: center; color: var(--text-secondary); padding: 3rem;">
                                Loading candidates database...
                            </td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>

        <script>
            async function loadCandidates() {
                try {
                    const response = await fetch('/api/candidates');
                    const candidates = await response.json();
                    
                    // Update stats
                    const total = candidates.length;
                    const completed = candidates.filter(c => c.status === 'completed').length;
                    const scheduled = candidates.filter(c => c.status === 'scheduled').length;
                    
                    let totalScore = 0;
                    let completedWithScore = 0;
                    candidates.forEach(c => {
                        if (c.status === 'completed' && c.ats_scorecard?.overall_score) {
                            totalScore += c.ats_scorecard.overall_score;
                            completedWithScore++;
                        }
                    });
                    const avgScore = completedWithScore > 0 ? Math.round(totalScore / completedWithScore) : 0;
                    
                    document.getElementById('stat-total').textContent = total;
                    document.getElementById('stat-completed').textContent = completed;
                    document.getElementById('stat-scheduled').textContent = scheduled;
                    document.getElementById('stat-avg-score').textContent = avgScore + '%';
                    
                    const tbody = document.getElementById('candidates-body');
                    if (candidates.length === 0) {
                        tbody.innerHTML = `
                            <tr>
                                <td colspan="5" style="text-align: center; color: var(--text-secondary); padding: 3rem;">
                                    No candidate records found. Upload Resumes to start scheduling!
                                </td>
                            </tr>
                        `;
                        return;
                    }
                    
                    tbody.innerHTML = '';
                    candidates.forEach(c => {
                        const statusBadge = c.status === 'completed' 
                            ? '<span class="badge badge-completed">Completed</span>' 
                            : '<span class="badge badge-scheduled">Scheduled</span>';
                            
                        let atsPill = '<span style="color: var(--text-secondary); font-style: italic;">N/A</span>';
                        let cheatPill = '<span style="color: var(--text-secondary); font-style: italic;">N/A</span>';
                        let actionBtn = `<button class="btn btn-disabled" disabled>No Report</button>`;
                        
                        if (c.status === 'completed') {
                            const score = c.ats_scorecard?.overall_score || 0;
                            const scoreClass = score >= 75 ? 'score-high' : (score >= 50 ? 'score-medium' : 'score-low');
                            atsPill = `<span class="score-pill ${scoreClass}">${score}</span>`;
                            
                            const cheatingScore = c.proctoring_analysis?.cheating_score || 0;
                            // Cheating score: low is green, high is red
                            const cheatClass = cheatingScore >= 50 ? 'score-low' : (cheatingScore >= 20 ? 'score-medium' : 'score-high');
                            cheatPill = `<span class="score-pill ${cheatClass}">${cheatingScore}%</span>`;
                            
                            actionBtn = `<a href="/api/scorecard/${c.candidate_email}" class="btn">View Scorecard</a>`;
                        }
                        
                        tbody.innerHTML += `
                            <tr>
                                <td>
                                    <div class="candidate-info">
                                        <span class="candidate-name">${c.candidate_name}</span>
                                        <span class="candidate-email">${c.candidate_email}</span>
                                    </div>
                                </td>
                                <td>${statusBadge}</td>
                                <td>${atsPill}</td>
                                <td>${cheatPill}</td>
                                <td>${actionBtn}</td>
                            </tr>
                        `;
                    });
                } catch (e) {
                    console.error("Error loading candidates:", e);
                    document.getElementById('candidates-body').innerHTML = `
                        <tr>
                            <td colspan="5" style="text-align: center; color: var(--danger); padding: 3rem;">
                                Failed to fetch candidates database. Make sure MongoDB is online.
                            </td>
                        </tr>
                    `;
                }
            }
            
            window.onload = loadCandidates;
        </script>
    </body>
    </html>
    """
    return HTMLResponse(content=html_content)

@router.get("/scorecard/{email}", response_class=HTMLResponse)
async def serve_scorecard(email: str):
    """
    Serves the detailed Candidate scorecard page with video player and timelines.
    Designed with a highly-premium dark glassmorphic style.
    """
    candidate = await candidates_collection.find_one({"candidate_email": email})
    if not candidate:
        raise HTTPException(status_code=404, detail="Candidate not found")
        
    if candidate.get("status") != "completed":
        raise HTTPException(status_code=400, detail="Interview is scheduled but feedback is not yet generated.")
        
    ats = candidate.get("ats_scorecard", {})
    proctor = candidate.get("proctoring_analysis", {})
    
    # Compile candidate details and paths safely
    c_name = candidate.get("candidate_name", "Unknown Candidate")
    overall = ats.get("overall_score", 0)
    recommendation = ats.get("recommendation", "N/A")
    rec_class = "rec-hire" if "Hire" in recommendation else "rec-reject"
    if recommendation == "Proceed with Caution":
        rec_class = "rec-caution"
        
    tech = ats.get("technical_rating", 0)
    comm = ats.get("communication_rating", 0)
    fit = ats.get("culture_fit_rating", 0)
    
    # Get recording file name
    rec_local_path = candidate.get("recording_local_path", "")
    video_src = ""
    if rec_local_path:
        filename = os.path.basename(rec_local_path)
        video_src = f"/static/recordings/{filename}"
        
    # Serialize json structures safely
    skills_json = json.dumps(ats.get("skills_assessment", []))
    strengths_json = json.dumps(ats.get("strengths", []))
    weaknesses_json = json.dumps(ats.get("weaknesses", []))
    logs_json = json.dumps(proctor.get("logs", []))

    html_content = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>ATS Scorecard - {c_name}</title>
        <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
        <style>
            :root {{
                --bg-primary: #0a0f1d;
                --bg-secondary: rgba(17, 24, 39, 0.75);
                --bg-card: rgba(30, 41, 59, 0.4);
                --accent-primary: #6366f1;
                --accent-secondary: #06b6d4;
                --text-primary: #f3f4f6;
                --text-secondary: #9ca3af;
                --card-border: rgba(255, 255, 255, 0.08);
                --success: #10b981;
                --warning: #f59e0b;
                --danger: #ef4444;
            }}

            * {{
                box-sizing: border-box;
                margin: 0;
                padding: 0;
            }}

            body {{
                font-family: 'Inter', sans-serif;
                background-color: var(--bg-primary);
                color: var(--text-primary);
                min-height: 100vh;
                background-image: 
                    radial-gradient(at 10% 10%, rgba(99, 102, 241, 0.12) 0px, transparent 50%),
                    radial-gradient(at 90% 90%, rgba(6, 182, 212, 0.08) 0px, transparent 50%);
                background-attachment: fixed;
                padding: 2.5rem 5rem;
            }}

            h1, h2, h3, .score-value, .brand {{
                font-family: 'Outfit', sans-serif;
            }}

            .container {{
                max-width: 1300px;
                margin: 0 auto;
            }}

            header {{
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin-bottom: 2.5rem;
                padding-bottom: 1.5rem;
                border-bottom: 1px solid var(--card-border);
            }}

            .back-link {{
                color: var(--text-secondary);
                text-decoration: none;
                font-size: 0.875rem;
                display: inline-flex;
                align-items: center;
                gap: 0.5rem;
                transition: color 0.2s;
            }}

            .back-link:hover {{
                color: #ffffff;
            }}

            .header-info {{
                display: flex;
                flex-direction: column;
                gap: 0.25rem;
            }}

            .header-info h1 {{
                font-size: 2rem;
                font-weight: 700;
            }}

            .candidate-meta {{
                display: flex;
                gap: 1.5rem;
                color: var(--text-secondary);
                font-size: 0.9rem;
            }}

            .recommendation-banner {{
                padding: 0.75rem 1.75rem;
                border-radius: 12px;
                font-weight: 700;
                font-size: 1.1rem;
                font-family: 'Outfit', sans-serif;
                letter-spacing: 0.02em;
            }}

            .rec-hire {{
                background: rgba(16, 185, 129, 0.15);
                color: var(--success);
                border: 1px solid rgba(16, 185, 129, 0.35);
                box-shadow: 0 0 15px rgba(16, 185, 129, 0.15);
            }}

            .rec-caution {{
                background: rgba(245, 158, 11, 0.15);
                color: var(--warning);
                border: 1px solid rgba(245, 158, 11, 0.35);
                box-shadow: 0 0 15px rgba(245, 158, 11, 0.15);
            }}

            .rec-reject {{
                background: rgba(239, 68, 68, 0.15);
                color: var(--danger);
                border: 1px solid rgba(239, 68, 68, 0.35);
                box-shadow: 0 0 15px rgba(239, 68, 68, 0.15);
            }}

            /* Core Gauges Grid */
            .gauges-grid {{
                display: grid;
                grid-template-columns: 2fr 1fr 1fr;
                gap: 1.5rem;
                margin-bottom: 2.5rem;
            }}

            .card {{
                background: var(--bg-secondary);
                border: 1px solid var(--card-border);
                border-radius: 20px;
                padding: 1.75rem;
                backdrop-filter: blur(12px);
                box-shadow: 0 10px 25px rgba(0,0,0,0.15);
            }}

            .card-title {{
                font-size: 1.1rem;
                color: var(--text-secondary);
                font-weight: 600;
                margin-bottom: 1.25rem;
                display: flex;
                justify-content: space-between;
                align-items: center;
            }}

            /* Radial Gauges */
            .gauge-container {{
                display: flex;
                align-items: center;
                justify-content: center;
                position: relative;
                height: 120px;
            }}

            .gauge-svg {{
                width: 120px;
                height: 120px;
                transform: rotate(-90deg);
            }}

            .gauge-bg-circle {{
                fill: none;
                stroke: rgba(255,255,255,0.05);
                stroke-width: 10;
            }}

            .gauge-val-circle {{
                fill: none;
                stroke-width: 10;
                stroke-linecap: round;
                transition: stroke-dasharray 1s ease-in-out;
            }}

            .gauge-text {{
                position: absolute;
                font-size: 1.75rem;
                font-weight: 700;
                color: #ffffff;
                font-family: 'Outfit', sans-serif;
            }}

            /* Star / Pill Ratings Grid */
            .ratings-vertical {{
                display: flex;
                flex-direction: column;
                gap: 0.75rem;
                width: 100%;
            }}

            .rating-bar-container {{
                display: flex;
                flex-direction: column;
                gap: 0.25rem;
            }}

            .rating-label-row {{
                display: flex;
                justify-content: space-between;
                font-size: 0.8rem;
                font-weight: 500;
            }}

            .rating-track-bg {{
                height: 8px;
                background: rgba(255, 255, 255, 0.05);
                border-radius: 999px;
                overflow: hidden;
            }}

            .rating-fill {{
                height: 100%;
                background: linear-gradient(to right, var(--accent-primary), var(--accent-secondary));
                border-radius: 999px;
            }}

            /* Multi-column Layout */
            .details-layout {{
                display: grid;
                grid-template-columns: 1.5fr 1fr;
                gap: 1.5rem;
            }}

            /* Detailed feedback styling */
            .feedback-report {{
                line-height: 1.7;
                font-size: 0.95rem;
                color: var(--text-primary);
                display: flex;
                flex-direction: column;
                gap: 1.5rem;
            }}

            .feedback-report h3 {{
                font-size: 1.15rem;
                color: #ffffff;
                margin-top: 0.5rem;
                border-bottom: 1px solid rgba(255,255,255,0.05);
                padding-bottom: 0.5rem;
            }}

            .feedback-report p {{
                color: #d1d5db;
            }}

            /* Skills Table */
            .skills-table {{
                width: 100%;
                border-collapse: collapse;
                margin-top: 1rem;
            }}

            .skills-table th {{
                padding: 0.75rem 1rem;
                font-size: 0.7rem;
                border-bottom: 1px solid var(--card-border);
            }}

            .skills-table td {{
                padding: 1rem;
                font-size: 0.85rem;
                border-bottom: 1px solid rgba(255,255,255,0.04);
            }}

            .skills-table tr:last-child td {{
                border-bottom: none;
            }}

            .skills-rating-pill {{
                display: inline-flex;
                justify-content: center;
                align-items: center;
                padding: 0.2rem 0.5rem;
                background: rgba(255,255,255,0.05);
                border-radius: 6px;
                font-weight: 700;
                font-size: 0.8rem;
                font-family: 'Outfit', sans-serif;
            }}

            /* Strengths / Weaknesses checklists */
            .checklist-box {{
                display: grid;
                grid-template-columns: 1fr 1fr;
                gap: 1.5rem;
            }}

            .checklist-list {{
                display: flex;
                flex-direction: column;
                gap: 0.75rem;
                list-style: none;
            }}

            .checklist-item {{
                display: flex;
                gap: 0.5rem;
                font-size: 0.85rem;
                line-height: 1.4;
            }}

            .icon-strength {{
                color: var(--success);
                font-weight: bold;
            }}

            .icon-weakness {{
                color: var(--danger);
                font-weight: bold;
            }}

            /* Video Player card */
            .video-container {{
                width: 100%;
                aspect-ratio: 16/9;
                border-radius: 12px;
                overflow: hidden;
                background: #000;
                border: 1px solid var(--card-border);
                box-shadow: 0 4px 15px rgba(0,0,0,0.4);
            }}

            .video-placeholder {{
                display: flex;
                align-items: center;
                justify-content: center;
                width: 100%;
                height: 100%;
                color: var(--text-secondary);
                font-size: 0.9rem;
            }}

            /* Proctoring Logs timeline */
            .timeline {{
                display: flex;
                flex-direction: column;
                gap: 1rem;
                margin-top: 1rem;
                max-height: 250px;
                overflow-y: auto;
                padding-right: 0.5rem;
            }}

            /* Custom scrollbar for timeline */
            .timeline::-webkit-scrollbar {{
                width: 4px;
            }}
            .timeline::-webkit-scrollbar-thumb {{
                background: rgba(255,255,255,0.1);
                border-radius: 4px;
            }}

            .timeline-item {{
                display: flex;
                gap: 1rem;
                border-left: 2px solid rgba(255,255,255,0.06);
                padding-left: 1rem;
                position: relative;
            }}

            .timeline-bullet {{
                position: absolute;
                left: -5px;
                top: 3px;
                width: 8px;
                height: 8px;
                border-radius: 50%;
            }}

            .bullet-high {{ background-color: var(--danger); box-shadow: 0 0 8px var(--danger); }}
            .bullet-medium {{ background-color: var(--warning); }}
            .bullet-low {{ background-color: var(--accent-secondary); }}

            .timeline-timestamp {{
                font-size: 0.8rem;
                font-weight: 700;
                color: #ffffff;
                background: rgba(255,255,255,0.05);
                padding: 0.1rem 0.4rem;
                border-radius: 4px;
                height: fit-content;
            }}

            .timeline-desc {{
                font-size: 0.8rem;
                color: var(--text-secondary);
                line-height: 1.4;
            }}

            /* Responsive */
            @media (max-width: 1024px) {{
                body {{
                    padding: 2rem;
                }}
                .gauges-grid {{
                    grid-template-columns: 1fr;
                }}
                .details-layout {{
                    grid-template-columns: 1fr;
                }}
            }}
            
        </style>
    </head>
    <body>
        <div class="container">
            <header>
                <div class="header-info">
                    <a href="/api/dashboard" class="back-link">← Back to Candidate Directory</a>
                    <h1 style="margin-top: 0.75rem;">{c_name}</h1>
                    <div class="candidate-meta">
                        <span>Email: {email}</span>
                        <span>Interview Completion: {candidate.get("updated_at").strftime("%Y-%m-%d %H:%M") if candidate.get("updated_at") else "N/A"}</span>
                    </div>
                </div>
                <div class="recommendation-banner {rec_class}">
                    Recommendation: {recommendation}
                </div>
            </header>

            <div class="gauges-grid">
                <div class="card" style="display: flex; align-items: center; justify-content: space-around; gap: 1rem;">
                    <!-- ATS radial progress -->
                    <div style="display: flex; flex-direction: column; align-items: center; gap: 0.5rem;">
                        <span class="stat-label" style="margin-bottom: 0.5rem;">ATS Match Score</span>
                        <div class="gauge-container">
                            <svg class="gauge-svg">
                                <circle class="gauge-bg-circle" cx="60" cy="60" r="50"></circle>
                                <circle class="gauge-val-circle" id="ats-gauge" cx="60" cy="60" r="50" stroke="url(#ats-grad)"></circle>
                                <defs>
                                    <linearGradient id="ats-grad" x1="0%" y1="0%" x2="100%" y2="100%">
                                        <stop offset="0%" stop-color="#6366f1" />
                                        <stop offset="100%" stop-color="#06b6d4" />
                                    </linearGradient>
                                </defs>
                            </svg>
                            <span class="gauge-text">{overall}%</span>
                        </div>
                    </div>
                    
                    <!-- Cheating radial progress -->
                    <div style="display: flex; flex-direction: column; align-items: center; gap: 0.5rem;">
                        <span class="stat-label" style="margin-bottom: 0.5rem;">Cheating Risk Score</span>
                        <div class="gauge-container">
                            <svg class="gauge-svg">
                                <circle class="gauge-bg-circle" cx="60" cy="60" r="50"></circle>
                                <circle class="gauge-val-circle" id="cheat-gauge" cx="60" cy="60" r="50" stroke="url(#cheat-grad)"></circle>
                                <defs>
                                    <linearGradient id="cheat-grad" x1="0%" y1="0%" x2="100%" y2="100%">
                                        <stop offset="0%" stop-color="#f59e0b" />
                                        <stop offset="100%" stop-color="#ef4444" />
                                    </linearGradient>
                                </defs>
                            </svg>
                            <span class="gauge-text">{proctor.get("cheating_score", 0)}%</span>
                        </div>
                    </div>
                </div>

                <!-- Sub-skills evaluations -->
                <div class="card">
                    <div class="card-title">Categorical Fit</div>
                    <div class="ratings-vertical">
                        <div class="rating-bar-container">
                            <div class="rating-label-row">
                                <span>Technical Rating</span>
                                <span>{tech}/10</span>
                            </div>
                            <div class="rating-track-bg">
                                <div class="rating-fill" style="width: {tech * 10}%;"></div>
                            </div>
                        </div>
                        <div class="rating-bar-container">
                            <div class="rating-label-row">
                                <span>Communication</span>
                                <span>{comm}/10</span>
                            </div>
                            <div class="rating-track-bg">
                                <div class="rating-fill" style="width: {comm * 10}%;"></div>
                            </div>
                        </div>
                        <div class="rating-bar-container">
                            <div class="rating-label-row">
                                <span>Culture Fit</span>
                                <span>{fit}/10</span>
                            </div>
                            <div class="rating-track-bg">
                                <div class="rating-fill" style="width: {fit * 10}%;"></div>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- Proctoring times -->
                <div class="card">
                    <div class="card-title">Proctoring Duration Stats</div>
                    <div class="ratings-vertical" style="gap: 0.5rem; font-size: 0.85rem;">
                        <div style="display:flex; justify-content:space-between;">
                            <span style="color: var(--text-secondary);">Looking Away:</span>
                            <span style="font-weight:600;">{proctor.get("gaze_away_seconds", 0)}s</span>
                        </div>
                        <div style="display:flex; justify-content:space-between;">
                            <span style="color: var(--text-secondary);">Candidate Absent:</span>
                            <span style="font-weight:600;">{proctor.get("no_face_seconds", 0)}s</span>
                        </div>
                        <div style="display:flex; justify-content:space-between;">
                            <span style="color: var(--text-secondary);">Multiple People:</span>
                            <span style="font-weight:600;">{proctor.get("multi_face_seconds", 0)}s</span>
                        </div>
                        <div style="display:flex; justify-content:space-between;">
                            <span style="color: var(--text-secondary);">Cell Phone:</span>
                            <span style="font-weight:600; color: {"var(--danger)" if proctor.get('phone_detected_seconds', 0) > 0 else 'inherit'};">
                                {proctor.get("phone_detected_seconds", 0)}s
                            </span>
                        </div>
                        <div style="display:flex; justify-content:space-between;">
                            <span style="color: var(--text-secondary);">Reference Book:</span>
                            <span style="font-weight:600;">{proctor.get("book_detected_seconds", 0)}s</span>
                        </div>
                    </div>
                </div>
            </div>

            <div class="details-layout">
                <!-- Left Column -->
                <div style="display: flex; flex-direction: column; gap: 1.5rem;">
                    <!-- Detailed Assessment Report -->
                    <div class="card">
                        <div class="card-title">Detailed Feedback Report</div>
                        <div class="feedback-report">
                            <p style="white-space: pre-wrap;">{ats.get("detailed_feedback", "No feedback details provided.")}</p>
                        </div>
                    </div>

                    <!-- Strengths and Weaknesses -->
                    <div class="card">
                        <div class="card-title">Key Observations</div>
                        <div class="checklist-box">
                            <div>
                                <h4 style="font-size:0.9rem; margin-bottom:0.75rem; color:var(--success);">Observed Strengths</h4>
                                <ul class="checklist-list" id="strengths-list"></ul>
                            </div>
                            <div>
                                <h4 style="font-size:0.9rem; margin-bottom:0.75rem; color:var(--warning);">Weaknesses & Growth Areas</h4>
                                <ul class="checklist-list" id="weaknesses-list"></ul>
                            </div>
                        </div>
                    </div>

                    <!-- Specific Skills Assessments -->
                    <div class="card">
                        <div class="card-title">Target Skills Matrix</div>
                        <table class="skills-table">
                            <thead>
                                <tr>
                                    <th>Skill</th>
                                    <th>Rating</th>
                                    <th>Evaluation Commentary</th>
                                </tr>
                            </thead>
                            <tbody id="skills-list-body">
                                <!-- Populated dynamically -->
                            </tbody>
                        </table>
                    </div>
                </div>

                <!-- Right Column -->
                <div style="display: flex; flex-direction: column; gap: 1.5rem;">
                    <!-- Video Recording Player -->
                    <div class="card">
                        <div class="card-title">Session Video Recording</div>
                        <div class="video-container">
                            {"<video src='" + video_src + "' controls style='width:100%; height:100%; object-fit:contain;'></video>" if video_src else "<div class='video-placeholder'>Video recording not archived locally.</div>"}
                        </div>
                    </div>

                    <!-- Proctoring incidents timeline -->
                    <div class="card">
                        <div class="card-title">Integrity Violations Timeline</div>
                        <div class="timeline" id="timeline-list">
                            <!-- Populated dynamically -->
                        </div>
                    </div>

                    <!-- Integrity check commentary -->
                    <div class="card">
                        <div class="card-title">Integrity Verification Commentary</div>
                        <p style="font-size:0.875rem; line-height:1.6; color:#d1d5db; white-space: pre-wrap;">
                            {ats.get("integrity_check_commentary", "No integrity check commentary generated.")}
                        </p>
                    </div>
                </div>
            </div>
        </div>

        <script>
            // Data bindings
            const overallScore = {overall};
            const cheatScore = {proctor.get("cheating_score", 0)};
            const skills = {skills_json};
            const strengths = {strengths_json};
            const weaknesses = {weaknesses_json};
            const proctorLogs = {logs_json};

            // Set radial gauges stroke-dasharray properties
            const circleRadius = 50;
            const circumference = 2 * Math.PI * circleRadius;
            
            function setGauge(id, val) {{
                const fill = document.getElementById(id);
                if (fill) {{
                    fill.style.strokeDasharray = `${{circumference}} ${{circumference}}`;
                    const offset = circumference - (val / 100) * circumference;
                    fill.style.strokeDashoffset = offset;
                }}
            }}

            setGauge('ats-gauge', overallScore);
            setGauge('cheat-gauge', cheatScore);

            // Populates strengths
            const strengthsUl = document.getElementById('strengths-list');
            if (strengths.length === 0) {{
                strengthsUl.innerHTML = '<li class="checklist-item"><span class="icon-strength">✓</span> None listed.</li>';
            }} else {{
                strengths.forEach(s => {{
                    strengthsUl.innerHTML += `<li class="checklist-item"><span class="icon-strength">✓</span> <span>${{s}}</span></li>`;
                }});
            }}

            // Populates weaknesses
            const weaknessesUl = document.getElementById('weaknesses-list');
            if (weaknesses.length === 0) {{
                weaknessesUl.innerHTML = '<li class="checklist-item"><span class="icon-weakness">⚠</span> None listed.</li>';
            }} else {{
                weaknesses.forEach(w => {{
                    weaknessesUl.innerHTML += `<li class="checklist-item"><span class="icon-weakness">⚠</span> <span>${{w}}</span></li>`;
                }});
            }}

            // Populates skills assessment matrix
            const skillsBody = document.getElementById('skills-list-body');
            if (skills.length === 0) {{
                skillsBody.innerHTML = '<tr><td colspan="3" style="text-align:center; color:var(--text-secondary);">No individual skills evaluated.</td></tr>';
            }} else {{
                skills.forEach(item => {{
                    skillsBody.innerHTML += `
                        <tr>
                            <td style="font-weight:600;">${{item.skill}}</td>
                            <td><span class="skills-rating-pill">${{item.rating}}/10</span></td>
                            <td style="color:var(--text-secondary);">${{item.comment}}</td>
                        </tr>
                    `;
                }});
            }}

            // Populates proctoring logs timeline
            const timelineDiv = document.getElementById('timeline-list');
            if (proctorLogs.length === 0) {{
                timelineDiv.innerHTML = '<div style="color:var(--text-secondary); font-size:0.85rem; padding: 1rem; text-align:center;">No integrity flags or alerts triggered. Session was completely clean!</div>';
            }} else {{
                proctorLogs.forEach(log => {{
                    const bulletClass = log.severity === 'high' ? 'bullet-high' : (log.severity === 'medium' ? 'bullet-medium' : 'bullet-low');
                    timelineDiv.innerHTML += `
                        <div class="timeline-item">
                            <span class="timeline-bullet ${{bulletClass}}"></span>
                            <span class="timeline-timestamp">${{log.timestamp}}</span>
                            <span class="timeline-desc">${{log.event}}</span>
                        </div>
                    `;
                }});
            }}
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
    str_html  = "".join(f'<div style="display:flex;align-items:flex-start;gap:6px;margin-bottom:6px"><span style="color:#0B6623;margin-top:1px">✓</span><span style="font-size:0.78rem;color:#374151">{s}</span></div>' for s in strengths_s[:6])
    conc_html = "".join(f'<div style="display:flex;align-items:flex-start;gap:6px;margin-bottom:6px"><span style="color:#DC2626;margin-top:1px">⚠</span><span style="font-size:0.78rem;color:#374151">{c}</span></div>' for c in concerns_s[:6])

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
  .grid-4 {{ display:grid;grid-template-columns:1.6fr 2.2fr 2fr 1.2fr;gap:14px;margin-bottom:14px; }}
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
  .va-badge {{ display:inline-block;font-size:0.62rem;font-weight:700;padding:2px 7px;border-radius:4px;margin-bottom:3px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:100%; }}
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
  .rec-section {{ margin-bottom:6px; }}
  .rec-section-title {{ font-size:0.68rem;font-weight:700;text-transform:uppercase;color:#6B7280;margin-bottom:6px; }}
  
  /* AI RECOMMENDATION CARD */
  .reco-card {{ border-radius:10px;padding:14px;margin-bottom:10px; }}
  .reco-badge {{ font-size:1rem;font-weight:800;letter-spacing:0.03em;margin-bottom:2px; }}
  .reco-sub {{ font-size:0.7rem;color:rgba(255,255,255,0.75);margin-bottom:8px; }}
  .reco-reason {{ font-size:0.72rem;color:rgba(255,255,255,0.95);line-height:1.5; }}
  .conf-bar-wrap {{ margin-top:10px; }}
  .conf-bar-label {{ display:flex;justify-content:space-between;font-size:0.68rem;color:rgba(255,255,255,0.75);margin-bottom:4px; }}
  .conf-bar-bg {{ height:5px;background:rgba(255,255,255,0.2);border-radius:10px; }}
  .conf-bar-fill {{ height:100%;border-radius:10px;background:#FFFFFF; }}

  /* PRINT STYLES */
  @media print {{
    .sidebar,.topbar-actions {{ display:none!important; }}
    .main {{ margin-left:0!important; }}
    .topbar {{ position:static!important; }}
    body {{ background:white!important;color:black!important; }}
    .card,.kpi-card {{ border:1px solid #ddd!important;background:#f9f9f9!important; }}
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
    <div class="nav-item active">👤 Candidates</div>
    <div class="nav-item">📁 Reports</div>
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
      <div class="card">
        <div style="display:flex;align-items:center;gap:8px;margin-bottom:18px">
          <div class="card-title">Candidate Performance</div>
        </div>
        <div id="skillBars" style="display:flex;flex-direction:column;gap:12px;padding:2px 4px"></div>
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

        <!-- TECHNICAL EVALUATION -->
      <div class="card" style="display:flex;flex-direction:column;gap:0">
        <div style="display:flex;align-items:center;gap:7px;margin-bottom:10px">
          <div class="card-title" style="margin-bottom:0">AI Interview Feedback</div>
        </div>
        <div style="background:linear-gradient(135deg,#EFF6FF 0%,#F5F3FF 100%);border-left:3px solid #2563EB;border-radius:0 8px 8px 0;padding:10px 12px;position:relative;overflow:hidden">
          <div style="display:flex;align-items:center;gap:6px;margin-bottom:6px">
            <span style="font-size:1.1rem;color:#2563EB;font-weight:900;line-height:1">&ldquo;</span>
            <span style="font-size:0.68rem;font-weight:700;color:#2563EB;text-transform:uppercase;letter-spacing:0.06em">Overall Summary</span>
          </div>
          <div style="font-size:0.71rem;color:#1E3A5F;line-height:1.55;overflow-y:auto;max-height:150px;padding-right:4px;text-align:justify">
            {detailed_feedback if detailed_feedback else "No detailed evaluation feedback provided for this candidate."}
          </div>
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
        <div style="display:flex;gap:20px;margin-top:12px;align-items:flex-start">
          <div class="rec-section" style="flex:1">
            <div class="rec-section-title" style="color:#0B6623;font-weight:700">✓ Strengths</div>
            {str_html if str_html else '<div style="font-size:0.75rem;color:var(--text-muted)">No strengths noted</div>'}
          </div>
          <div class="rec-section" style="flex:1">
            <div class="rec-section-title" style="color:#DC2626;font-weight:700">⚠ Concerns</div>
            {conc_html if conc_html else '<div style="font-size:0.75rem;color:var(--text-muted)">No concerns flagged</div>'}
          </div>
        </div>
      </div>

      <!-- AI RECOMMENDATION -->
      <div class="card" style="justify-content:space-between">
        <div>
          <div class="card-title" style="margin-bottom:10px">AI Recommendation</div>
          <div class="reco-card" style="background:{rec_bg};border:1px solid {rec_text_color}33;padding:10px 12px;margin-bottom:10px;border-radius:8px">
            <div class="reco-badge" style="color:{rec_text_color};font-size:0.85rem;font-weight:800">{rec_icon} {recommendation.upper()}</div>
            <div class="reco-sub" style="font-size:0.65rem;color:rgba(255,255,255,0.8);margin-bottom:0">{rec_sub}</div>
          </div>
          <div style="background:linear-gradient(135deg,#EFF6FF 0%,#F5F3FF 100%);border-left:3px solid #2563EB;border-radius:0 8px 8px 0;padding:9px 11px;margin-bottom:10px">
            <div style="display:flex;align-items:center;gap:5px;margin-bottom:5px">
              <span style="font-size:1rem;color:#2563EB;font-weight:900;line-height:1">&ldquo;</span>
              <span style="font-size:0.65rem;font-weight:700;color:#2563EB;text-transform:uppercase;letter-spacing:0.06em">Analysis Reason</span>
            </div>
            <div style="font-size:0.68rem;color:#1E3A5F;line-height:1.55;text-align:justify;max-height:130px;overflow-y:auto;padding-right:4px">{integrity_commentary if integrity_commentary else (detailed_feedback if detailed_feedback else "No analysis commentary available.")}</div>
          </div>
        </div>
        <div style="margin-top:auto">
          <div style="display:flex;justify-content:space-between;font-size:0.65rem;color:#4B5563;margin-bottom:4px;font-weight:600">
            <span>AI Analysis Confidence</span>
            <span style="color:#2563EB;font-weight:700">{ai_conf}%</span>
          </div>
          <div style="height:6px;background:#F3F4F6;border-radius:10px;overflow:hidden;border:1px solid #E5E7EB">
            <div style="height:100%;width:{min(ai_conf,100)}%;background:linear-gradient(90deg,#2563EB,#4F46E5);border-radius:10px"></div>
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

  const PAGE = 3;
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
  skillLabels.forEach((lbl, i) => {{
    const r = skillRatings[i] || 5;
    const pct = r * 10;
    wrap.innerHTML += `
      <div class="skill-bar-row">
        <div class="skill-bar-header">
          <span style="color:var(--text);font-size:0.72rem">${{lbl}}</span>
          <span style="color:#2563EB;font-weight:600;font-size:0.72rem">${{r}}/10</span>
        </div>
        <div class="skill-bar-bg"><div class="skill-bar" style="width:${{pct}}%"></div></div>
      </div>`;
  }});
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
  if (e.key === 'Escape') window.closeModal(); 
  if (e.key === 'ArrowLeft' && document.getElementById('modalOverlay').classList.contains('open')) window.modalNavigate(-1);
  if (e.key === 'ArrowRight' && document.getElementById('modalOverlay').classList.contains('open')) window.modalNavigate(1);
}});
</script>
</body>
</html>
"""
    return HTMLResponse(content=html_content)
