from pydantic import Field
from pydantic_settings import BaseSettings
import os

class Settings(BaseSettings):
    BASE_DIR: str = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    DATABASE_URL: str = Field(..., env="DATABASE_URL")
    DOCUMENTS_DIR: str = Field(..., env="DOCUMENTS_DIR")
    TEMP_DIR: str = os.path.join(os.getcwd(), "temp")
    OUTPUT_DIR: str = os.path.join(os.getcwd(), "outputs")

    YPAERO_BASE_URL: str
    YPAERO_API_TOKEN: str
    DATABASE_URL: str

    class Config:
        env_file = ".env"

settings = Settings()
