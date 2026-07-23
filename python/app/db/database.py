from motor.motor_asyncio import AsyncIOMotorClient
from app.core.config import settings

# Initialize Motor Client
client = AsyncIOMotorClient(settings.MONGO_URI)

# Define reference to our specific Database
db = client[settings.MONGO_DB_NAME]

# Export specific collections for ease of access
candidates_collection = db.candidates