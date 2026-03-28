from motor.motor_asyncio import AsyncIOMotorClient
from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    mongodb_url: str = "mongodb://mongodb:27017"
    database_name: str = "voc_system"

    class Config:
        env_file = ".env"


settings = Settings()

client: AsyncIOMotorClient = None


async def connect_to_mongo():
    global client
    client = AsyncIOMotorClient(settings.mongodb_url)


async def close_mongo_connection():
    global client
    if client:
        client.close()


def get_database():
    return client[settings.database_name]


def get_collection(name: str):
    db = get_database()
    return db[name]
