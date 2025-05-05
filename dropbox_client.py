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

def upload_file(remote_path: str, local_path: str) -> None:
    """로컬 파일을 Dropbox에 업로드합니다."""
    dbx = get_dbx()
    p = _normalize_path(remote_path)
    
    # 디렉토리 경로 자동 생성
    folder_path = os.path.dirname(p)
    try:
        # 폴더 존재 여부 확인
        try:
            dbx.files_get_metadata(folder_path)
            print(f"폴더가 이미 존재합니다: {folder_path}")
        except Exception:
            # 폴더가 없으면 생성
            print(f"폴더 생성 시도: {folder_path}")
            try:
                result = dbx.files_create_folder_v2(folder_path)
                print(f"폴더 생성 완료: {result.metadata.path_display}")
            except dropbox.exceptions.ApiError as e:
                if isinstance(e.error, dropbox.files.CreateFolderError) and e.error.is_path() and e.error.get_path().is_conflict():
                    print(f"폴더가 이미 존재합니다 (충돌): {folder_path}")
                else:
                    print(f"폴더 생성 중 API 오류: {e}")
    except Exception as e:
        print(f"폴더 확인/생성 중 오류 (무시됨): {e}")
    
    # 파일 업로드
    print(f"파일 업로드 시작: {p}")
    try:
        with open(local_path, "rb") as f:
            result = dbx.files_upload(f.read(), p, mode=dropbox.files.WriteMode.overwrite)
        print(f"파일 업로드 완료: {result.path_display}, 크기: {result.size} 바이트")
        return result.path_display
    except Exception as e:
        print(f"파일 업로드 오류: {e}")
        raise