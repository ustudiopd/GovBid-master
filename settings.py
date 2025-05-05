import os
from dotenv import load_dotenv
from dataclasses import dataclass

# .env 파일 로드
load_dotenv()

@dataclass
class Settings:
    """애플리케이션 설정"""
    # Azure AD (Microsoft Graph API)
    AZURE_TENANT_ID: str = os.getenv("AZURE_TENANT_ID", "")
    AZURE_CLIENT_ID: str = os.getenv("AZURE_CLIENT_ID", "")
    AZURE_CLIENT_SECRET: str = os.getenv("AZURE_CLIENT_SECRET", "")
    MAILBOX: str = os.getenv("MAILBOX", "")
    
    # Dropbox
    DROPBOX_APP_KEY: str = os.getenv("DROPBOX_APP_KEY", "")
    DROPBOX_APP_SECRET: str = os.getenv("DROPBOX_APP_SECRET", "")
    DROPBOX_ACCESS_TOKEN: str = os.getenv("DROPBOX_ACCESS_TOKEN", "")
    DROPBOX_REFRESH_TOKEN: str = os.getenv("DROPBOX_REFRESH_TOKEN", "")
    DROPBOX_SHARED_FOLDER_ID: str = os.getenv("DROPBOX_SHARED_FOLDER_ID", "")
    DROPBOX_SHARED_FOLDER_NAME: str = os.getenv("DROPBOX_SHARED_FOLDER_NAME", "입찰 2025")
    
    # ChatGPT
    CHATGPT_API_KEY: str = os.getenv("CHATGPT_API_KEY", "")
    GPT_MODEL: str = os.getenv("CHATGPT_MODEL", "gpt-4.1-mini")

# 단일 settings 인스턴스 생성
settings = Settings() 