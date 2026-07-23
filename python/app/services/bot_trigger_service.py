import requests
from app.core.config import settings

def trigger_bot_join(join_url, candidate_name, candidate_email, questions, end_time, meeting_id = None, event_id = None):

    print("Triggering Node Teams Bot")

    payload = {
        "meetingUrl": join_url, 
        "candidateName": candidate_name, 
        "candidateEmail": candidate_email, 
        "questions": questions, 
        "endTime": end_time.isoformat() if hasattr(end_time, 'isoformat') else end_time, 
        "meetingId": meeting_id, # Used after call ends to fetch the full session recording
        "event_id": event_id
    }

    print(f"[Trigger] Payload: {payload}")

    response = requests.post(
        f"{settings.BOT_SERVICE_URL}/bot-test/join-meeting-test", 
        json = payload
    )

    print("Bot service response:", response.text)

    if "application/json" in response.headers.get("content-type", ""):
        return response.json()
    else:
        return {"message": response.text}
