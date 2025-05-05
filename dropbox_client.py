# dropbox_client.py
import os
import json
from dotenv import dotenv_values
import dropbox

# .env 파일에서 설정값을 로드하고 키를 소문자로 변환하여 반환합니다.
def load_config():
    config = dotenv_values(".env")
    return {k.lower(): v for k, v in config.items() if v is not None}

cfg = load_config()
APP_KEY       = cfg.get("dropbox_app_key")      or os.getenv("DROPBOX_APP_KEY")
APP_SECRET    = cfg.get("dropbox_app_secret")   or os.getenv("DROPBOX_APP_SECRET")
ACCESS_TOKEN  = cfg.get("dropbox_access_token") or os.getenv("DROPBOX_ACCESS_TOKEN")
REFRESH_TOKEN = cfg.get("dropbox_refresh_token")or os.getenv("DROPBOX_REFRESH_TOKEN")

if not all([APP_KEY, APP_SECRET, ACCESS_TOKEN, REFRESH_TOKEN]):
    raise RuntimeError("Dropbox OAuth 설정이 올바르게 되어 있지 않습니다. 확인 필요")

# Dropbox 클라이언트 객체를 생성하고 인증을 확인하여 반환합니다.
def get_dbx():
    dbx = dropbox.Dropbox(
        oauth2_access_token=ACCESS_TOKEN,
        oauth2_refresh_token=REFRESH_TOKEN,
        app_key=APP_KEY,
        app_secret=APP_SECRET,
    )
    dbx.check_user()
    return dbx

# 경로 정규화 유틸리티
def _normalize_path(p: str) -> str:
    """경로 앞뒤 공백 제거 후 슬래시로 시작하도록 정규화합니다."""
    p = p.strip()
    return p if p.startswith("/") else f"/{p}"

def list_folder(path: str) -> list[str]:
    """Dropbox에서 지정 폴더의 항목 이름 목록을 반환합니다."""
    dbx = get_dbx()
    p = _normalize_path(path)
    res = dbx.files_list_folder(p)
    return [entry.name for entry in res.entries]

def download_json(path: str) -> list[dict]:
    """Dropbox에서 JSON 파일을 다운로드하여 파싱한 후 반환합니다."""
    dbx = get_dbx()
    p = _normalize_path(path)
    _, res = dbx.files_download(p)
    return json.loads(res.content.decode("utf-8"))

def download_file(remote_path: str, local_path: str) -> None:
    """Dropbox에서 파일을 다운로드하여 로컬에 저장합니다."""
    dbx = get_dbx()
    p = _normalize_path(remote_path)
    _, res = dbx.files_download(p)
    with open(local_path, "wb") as f:
        f.write(res.content)

def upload_json(remote_path: str, data: dict) -> None:
    """딕셔너리를 JSON으로 덤프해 Dropbox에 업로드합니다."""
    dbx = get_dbx()
    p = _normalize_path(remote_path)
    content = json.dumps(data, ensure_ascii=False, indent=2)
    dbx.files_upload(content.encode("utf-8"), p, mode=dropbox.files.WriteMode.overwrite)