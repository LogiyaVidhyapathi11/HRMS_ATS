from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.responses import Response
from fastapi.routing import APIRoute
import os
from app.api import interview_routes
from app.services.scheduler_service import start_schedule

app = FastAPI(
    title = "AI Interview Platform", 
    version = "1.0.0"
)

# ─── Static recordings directory resolution ─────────────────────────────────
# Enforce one single location: teams-bot/public/recordings
current_dir = os.path.dirname(os.path.abspath(__file__))
backend_root = os.path.abspath(os.path.join(current_dir, "..", ".."))

recordings_dir = os.path.abspath(os.path.join(backend_root, "teams-bot", "public", "recordings"))
os.makedirs(os.path.join(recordings_dir, "incidents"), exist_ok = True)

print("="*80)
print("recordings_dir =", recordings_dir)
print("Exists =", os.path.exists(recordings_dir))
print("Incidents =", os.path.exists(os.path.join(recordings_dir, "incidents")))

print("Files:")
print(os.listdir(os.path.join(recordings_dir, "incidents"))[:5])
print("="*80)

# Mount static files to serve candidate recording video and incident screenshots
app.mount("/static/recordings", StaticFiles(directory=recordings_dir), name="recordings")
print(f"Mounted recordings directory: {recordings_dir}")

print("\n========== ROUTES ==========")
for r in app.routes:
    print(type(r).__name__, r.path)
print("============================\n")

app.include_router(interview_routes.router, prefix = "/api")

@app.on_event("startup")
def startup_event():
    print("Starting AI Interview Platform")
    start_schedule()
    print("Scheduler Started")










