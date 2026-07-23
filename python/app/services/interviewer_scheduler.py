import requests
from datetime import datetime
from app.core.config import settings
from app.services.scheduler_service import schedule_bot_join
from app.services.auth_service import get_graph_token
import base64
import urllib.parse

def schedule_interview(
        candidate_name: str,
        candidate_email: str, 
        start_time: datetime, 
        end_time: datetime, 
        questions: list
):
    """
        Schedules an interview by:
        1. Creating a Teams-enabled event in a secondary calendar.
        2. Generating a backend redirect link for 10-minute expiry validation.
        3. Updating the event body to include the redirect link.
        4. Scheduling the bot join.
    """
    token = get_graph_token()
    organizer = settings.TEAMS_ORGANIZER_EMAIL

    headers = {
        "Authorization": f"Bearer {token}", 
        "Content-Type": "application/json"
    }

    # Step 0: Resolve Organizer Email to User ID (GUID)
    # Graph APIs are much more reliable using IDs than emails.
    # ──────────────────────────────────────────────────────────────
    user_id = organizer
    try:
        user_res = requests.get(f"https://graph.microsoft.com/v1.0/users/{organizer}", headers=headers)
        if user_res.status_code == 200:
            user_id = user_res.json().get("id", organizer)
            print(f"[Scheduler] Resolved organizer {organizer} to ID: {user_id}")
    except Exception as e:
        print(f"[Scheduler] Warning: Failed to resolve user ID: {e}")

    # Step 1: Write directly to the primary Calendar for immediate visibility.

    # calendars_url = f"https://graph.microsoft.com/v1.0/users/{user_id}/calendars"
    # calendars_response = requests.get(calendars_url, headers=headers)

    # # ... (rest of calendar lookup logic remains the same, using user_id)
    # calendar_id = None
    # if calendars_response.status_code == 200:
    #     for cal in calendars_response.json().get("value", []):
    #         if cal.get("name") == "AI Interviews":
    #             calendar_id = cal.get("id")
    #             break

    # if not calendar_id:
    #     new_cal_response = requests.post(calendars_url, headers = headers, json = {"name": "AI Interviews"})

    #     if new_cal_response.status_code == 201:
    #         calendar_id = new_cal_response.json().get("id")

    # if calendar_id:
    #     event_url = f"https://graph.microsoft.com/v1.0/users/{user_id}/calendars/{calendar_id}/events"
    # else:
    event_url = f"https://graph.microsoft.com/v1.0/users/{user_id}/events"

    # Step 2: Try to create onlineMeeting directly (for auto-recording)

    meeting_id = None
    join_url = None

    try: 
        online_meeting_url = f"https://graph.microsoft.com/v1.0/users/{user_id}/onlineMeetings"

        # With auto-recording
        payload = {
            "subject": f"AI Interview - {candidate_name}", 

            "startDateTime": start_time.strftime("%Y-%m-%dT%H:%M:%S.0000000Z"), 

            "endDateTime": end_time.strftime("%Y-%m-%dT%H:%M:%S.0000000Z"), 

            "recordAutomatically": True, 
        }
        
        r1 = requests.post(online_meeting_url, headers=headers, json=payload)

        if r1.status_code in [200, 201]:
            d = r1.json()
            meeting_id = d["id"]
            join_url   = d["joinWebUrl"]

            print(f"[Scheduler] ✅ Direct onlineMeeting created — id: {meeting_id}")

        else:
            print(f"[Scheduler] ⚠️ Direct onlineMeeting failed ({r1.status_code}). Falling back to Calendar.")

    except Exception as e:
        print(f"[Scheduler] ⚠️ Direct creation exception: {e}.")
  
    
    # Step 3: Create the Calendar Event 

    event_payload = {
        "subject": f"Interview Invitation: {candidate_name} - Adamsbridge", 

        "start": {
            "dateTime": start_time.isoformat(), "timeZone": "UTC"
        }, 

        "end": {
            "dateTime": end_time.isoformat(), 
            "timeZone": "UTC"
        }, 

        # This makes it a PROPER native Teams meeting so the candidate gets a popup!
        "isOnlineMeeting": True, 
        "onlineMeetingProvider": "teamsForBusiness",

        # No attendees - just creating the meeting link
        "attendees": [
            {
                "emailAddress": {
                    "address": candidate_email, 
                    "name": candidate_name
                }, 
                "type": "required"
            }
        ], 

        "showAs": "free", 
    }

    if join_url:
        event_payload["onlineMeetingUrl"] = join_url

    event_response = requests.post(event_url, headers = headers, json = event_payload)

    # print("Secondary Calendar Event Status:", event_response.status_code)
    # print("Calendar Event Response:", response.text[:500])

    if event_response.status_code != 201:
       print(f"[Scheduler] MS Graph Calendar Event Creation Failed! Status: {event_response.status_code}, Response: {event_response.text}")
       if not join_url:
        raise Exception(f"Failed to create meeting: {event_response.text}")
        event_id = None
    else:
        event_data = event_response.json()
        event_id   = event_data["id"]
        
        if not join_url:
            join_url = event_data.get("onlineMeeting", {}).get("joinUrl", "")
            print(f"[Scheduler] ✅ Calendar event created — joinUrl: {join_url}")

            # Discovery Attempt: Try to find the meeting ID and enable recording
            try:
                # Filter requires the URL to be single-quoted
                filter_url = join_url.replace("'", "''") 
                lookup_url = f"https://graph.microsoft.com/v1.0/users/{user_id}/onlineMeetings?$filter=JoinWebUrl eq '{filter_url}'"
                lookup_res = requests.get(lookup_url, headers=headers)
                
                if lookup_res.status_code == 200:
                    meetings = lookup_res.json().get("value", [])
                    if meetings:
                        meeting_id = meetings[0]["id"]
                        print(f"[Scheduler] 🔍 Discovered meeting ID: {meeting_id}")
                        # Try to enable recording on this discovered meeting
                        patch_res = requests.patch(
                            f"https://graph.microsoft.com/v1.0/users/{user_id}/onlineMeetings/{meeting_id}",
                            headers=headers,
                            json={"recordAutomatically": True}
                        )
                        if patch_res.status_code == 200:
                            print("[Scheduler] ✨ Auto-recording ENABLED on discovered meeting.")
            except Exception as e:
                print(f"[Scheduler] ⚠️  Discovery/Patch failed: {e}")

    # ──────────────────────────────────────────────────────────────
    # Step 4: Expiry Redirect Link
    # ──────────────────────────────────────────────────────────────
    encoded_url = base64.b64encode(join_url.encode('utf-8')).decode('utf-8')
    safe_encoded_url = urllib.parse.quote(encoded_url) # URL-encode the base64 to prevent broken links
    start_timestamp = int(start_time.timestamp())
    redirect_link = f"{settings.BACKEND_URL}/api/join-redirect?url={safe_encoded_url}&start_time={start_timestamp}"

    # Update event body with a professional "Join" button
    if event_id:
        # Professional HTML body with Adamsbridge branding
        button_html = f"""
        <div style = "font-family: 'Segoe UI, Arial, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden;">
            <div style="background-color: #004a99; color: #ffffff; padding: 20px; text-align: center;">
                <h1 style = "margin: 0; font-size: 20px;">Interview Invitation</h1>
                <p style = "margin: 5px 0 0 0; font-size: 14px; opacity: 0.9;">Adamsbridge Global LLC</p>
            </div>
            <div style = "padding: 30px; background-color: #ffffff; color: #333333; line-height: 1.6;">
                <p style = "margin-top: 0;">Dear <strong>{candidate_name}</strong>,</p>
                <p> Thank you for your interest in joining <strong>Adamsbridge</strong>. We are pleased to invite you to an AI-led interview session for your application. </p>

                <div style = "background-color: #f8f9fa; border-left: 4px solid #004a99; padding: 15px; margin: 20px 0;">
                    <p style = "margin: 0; font-size: 14px;"><strong> Format: </strong> AI-Led Video Interview </p>
                    <p style = "margin: 5px 0 0 0; font-size: 14px;"><strong>Duration:</strong> 1 hour</p>
                </div>

                <p>To begin your session, please click the button below. Please ensure you are in a quiet environment with a stable internet connection.</p>

                <div style = "text-align: center; margin: 30px 0;">
                    <a href = "{redirect_link}" style = "background-color: #004a99; color: #ffffff; padding: 14px 32px; border-radius: 6px; text-decoration: none; font-weight: bold; font-size: 16px; display: inline-block; box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);"> Start Your Interview</a>
                </div>

                <p style = "font-size: 12px; color: #777777; margin-bottom: 0;">
                    <em><strong>Security Note: </strong> This secure join link is valid for 10 minutes from your scheduled start time.</em>
                </p>
            </div>
            <div style = "background-color: #f1f1f1; color: #888888; padding: 15px; text-align: center; font-size: 12px;">
                &copy; {datetime.now().year} Adamsbridge. All rights reserved.
            </div>
        </div>
        """

        patch_payload = {
            "location": {"displayName": "Adamsbridge Virtual Interview Room", "locationUri": redirect_link}, 
            "body": {
                "contentType": "html", 
                "content": button_html
            }
        }

        requests.patch(f"{event_url}/{event_id}", headers = headers, json = patch_payload)

    # Step 5: Schedule Bot

    schedule_bot_join(
        join_url,
        start_time, 
        candidate_name, 
        candidate_email, 
        questions, 
        end_time, 
        meeting_id, 
        event_id
    )

    print("Scheduling bot join")
    print("Join URL:", join_url)
    print("Start Time:", start_time)

    return {
        "meeting_link": redirect_link,
        "direct_join_url": join_url,
        "meeting_id": meeting_id, 
        "event_id": event_id, 
        "eventId": event_id, 
        "candidate_email": candidate_email,
        "start_time": start_time.isoformat(),
        "end_time": end_time.isoformat()
    }