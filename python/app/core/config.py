import os
from pathlib import Path
from pydantic_settings import BaseSettings
from functools import lru_cache

class Settings(BaseSettings):
    CLIENT_ID: str
    CLIENT_SECRET: str
    TENANT_ID: str

    # Microsoft Graph
    GRAPH_BASE_URL: str 
    GRAPH_SCOPE: str 

    # Teams / Calendar
    TEAMS_ORGANIZER_EMAIL: str
    DEFAULT_TIMEZONE: str = "UTC"

    # Bot Service
    BOT_SERVICE_URL: str
    BACKEND_URL: str 
    VERTEX_URL: str

    # Database
    MONGO_URI: str = "mongodb+srv://logiyavidhyapathi_db_user:Adams%40123%24@cluster0.jp7yfum.mongodb.net/ai_interview_db?retryWrites=true&w=majority"
    MONGO_DB_NAME: str = "ai_interview_db"

    # AI
    GEMINI_API_KEY: str

    # Scheduler
    scheduler_timezone: str = "UTC"

    # App
    ENV: str = "development"
    API_PREFIX: str = "/api"
    APP_NAME: str = "AI Interview Platform"

    class Config:
        env_file = ".env"
        case_sensitive = False

@lru_cache()
def get_settings():
    return Settings()

settings = get_settings()