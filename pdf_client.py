# pdf_client.py
# PDF 문서에서 서식 파일을 찾고 추출하는 모듈

import os
import json
import logging
from typing import List, Dict, Any, Callable, Optional
import re
import tempfile
import shutil
from dotenv import load_dotenv
from PyPDF2 import PdfReader, PdfWriter
from openai import OpenAI

# Dropbox 클라이언트 임포트
from dropbox_client import upload_file, upload_json

# 로깅 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# .env 파일에서 API 키 로드
load_dotenv()
CHATGPT_API_KEY = os.getenv("CHATGPT_API_KEY")
GPT_MODEL = os.getenv("CHATGPT_MODEL", "gpt-4.1-mini")
DROPBOX_SHARED_FOLDER_ID = os.getenv("DROPBOX_SHARED_FOLDER_ID")
DROPBOX_SHARED_FOLDER_NAME = os.getenv("DROPBOX_SHARED_FOLDER_NAME", "입찰 2025")

# 디버그 모드 - 콘솔에 자세한 정보 출력
DEBUG = False

def extract_text_from_pdf(path: str) -> str:
    """PDF 파일에서 텍스트 추출"""
    try:
        reader = PdfReader(path)
        text_parts = []
        
        for i, page in enumerate(reader.pages):
            try:
                text = page.extract_text() or ""
                if text.strip():
                    text_parts.append(f"--- PAGE {i+1} ---\n{text}")
                else:
                    text_parts.append(f"--- PAGE {i+1} ---\n[Page {i+1} has no extractable text]")
            except Exception as e:
                logger.warning(f"페이지 {i+1} 텍스트 추출 오류: {e}")
                text_parts.append(f"--- PAGE {i+1} ---\n[Error: {str(e)}]")
        
        return "\n".join(text_parts)
    except Exception as e:
        logger.error(f"PDF 파일 텍스트 추출 오류: {e}")
        return f"[Error extracting text: {str(e)}]"

def find_dropbox_folder(start_path):
    """지정된 경로에서 드롭박스 폴더를 찾아 경로 반환"""
    # 드롭박스 폴더명 (일반적인 패턴)
    dropbox_names = ["Dropbox", "드롭박스"]
    
    # 현재 디렉토리가 드롭박스인지 확인
    current_dir = os.path.basename(start_path)
    if current_dir in dropbox_names:
        return start_path
    
    # 일반적인 드롭박스 설치 경로 확인 (우선 순위 높음)
    potential_paths = [
        os.path.join(os.environ.get("USERPROFILE", ""), "Dropbox"),
        os.path.expanduser("~/Dropbox"),
        "D:/Dropbox",
        "C:/Dropbox",
        os.path.join(os.environ.get("USERPROFILE", ""), "Documents", "Dropbox"),
        os.path.expanduser("~/Documents/Dropbox"),
        "D:/Documents/Dropbox",
        "C:/Documents/Dropbox",
        "D:/문서/Dropbox",
        "C:/문서/Dropbox"
    ]
    
    for path in potential_paths:
        if os.path.exists(path) and os.path.isdir(path):
            # 입찰 폴더 확인
            bid_folder = os.path.join(path, "입찰 2025")
            if os.path.exists(bid_folder) and os.path.isdir(bid_folder):
                return path
    
    # 상위 디렉토리 탐색 (최대 5단계)
    path = start_path
    for _ in range(5):
        parent = os.path.dirname(path)
        if not parent or parent == path:
            break
            
        parent_name = os.path.basename(parent)
        if parent_name in dropbox_names:
            return parent
            
        path = parent
    
    # 못 찾으면 None 반환
    return None

def create_dropbox_forms_dir(pdf_path, base_folder_name=None):
    """
    PDF 파일이 위치한 경로를 기반으로 적절한 Dropbox 서식 폴더 경로를 생성
    
    Args:
        pdf_path: PDF 파일 경로
        base_folder_name: 기본 폴더명 (기본값은 환경변수 또는 "입찰 2025")
        
    Returns:
        (성공여부, 드롭박스 서식폴더 경로)
    """
    # 환경 변수에서 기본 폴더명 사용 (인자로 지정된 값이 우선)
    if base_folder_name is None:
        base_folder_name = DROPBOX_SHARED_FOLDER_NAME
    
    # 환경 변수에 ID가 있으면 해당 공유 폴더 사용
    if DROPBOX_SHARED_FOLDER_ID:
        logger.info(f"환경 변수에서 공유 폴더 정보 사용: ID={DROPBOX_SHARED_FOLDER_ID}, 이름={base_folder_name}")
        
        # PDF 파일명에서 대상 폴더명 추출
        pdf_name = os.path.basename(pdf_path)
        folder_name = os.path.splitext(pdf_name)[0]
        
        # 임시 폴더 경로 생성 (로컬에서 작업용)
        temp_dir = tempfile.gettempdir()
        target_folder = os.path.join(temp_dir, base_folder_name, folder_name)
        forms_dir = os.path.join(target_folder, "서식")
        
        try:
            os.makedirs(forms_dir, exist_ok=True)
            logger.info(f"서식 폴더 생성 (환경 변수 기반): {forms_dir}")
            return True, forms_dir
        except Exception as e:
            logger.error(f"서식 폴더 생성 실패: {e}")
            return False, None
    
    # 이하 기존 코드 (환경 변수가 없을 경우)
    # PDF 경로의 디렉토리
    pdf_dir = os.path.dirname(pdf_path)
    
    # 먼저 드롭박스 루트 폴더 찾기
    dropbox_folder = find_dropbox_folder(pdf_dir)
    if not dropbox_folder:
        logger.warning(f"드롭박스 폴더를 찾을 수 없습니다: {pdf_dir}")
        return False, None
    
    # PDF 파일이 입찰 폴더 내에 있는지 확인
    bid_folder = os.path.join(dropbox_folder, base_folder_name)
    if not os.path.exists(bid_folder):
        logger.warning(f"입찰 폴더가 없습니다: {bid_folder}")
        try:
            os.makedirs(bid_folder, exist_ok=True)
            logger.info(f"입찰 폴더 생성: {bid_folder}")
        except Exception as e:
            logger.error(f"입찰 폴더 생성 실패: {e}")
            return False, None
    
    # PDF 경로에서 드롭박스 루트 이후의 상대 경로 찾기
    if pdf_dir.startswith(dropbox_folder):
        rel_path = os.path.relpath(pdf_dir, dropbox_folder)
        parts = rel_path.split(os.sep)
        
        # 첫 번째 부분이 base_folder_name과 같은지 확인
        if parts and parts[0] == base_folder_name:
            # 이미 입찰 폴더 아래에 있는 경우
            target_folder = os.path.join(dropbox_folder, rel_path)
        else:
            # 드롭박스 내부이지만 입찰 폴더 외부인 경우
            # 마지막 디렉토리 이름을 사용
            if len(parts) > 0:
                last_dir = parts[-1]
                if last_dir:
                    target_folder = os.path.join(bid_folder, last_dir)
                else:
                    # 최후의 방어 로직: PDF 파일명을 폴더명으로 사용
                    pdf_name = os.path.basename(pdf_path)
                    folder_name = os.path.splitext(pdf_name)[0]
                    target_folder = os.path.join(bid_folder, folder_name)
            else:
                # 단순한 경우 드롭박스 바로 아래
                pdf_name = os.path.basename(pdf_path)
                folder_name = os.path.splitext(pdf_name)[0]
                target_folder = os.path.join(bid_folder, folder_name)
    else:
        # 드롭박스 외부인 경우: PDF 파일명을 폴더명으로 사용
        pdf_name = os.path.basename(pdf_path)
        folder_name = os.path.splitext(pdf_name)[0]
        target_folder = os.path.join(bid_folder, folder_name)
    
    # 최종 서식 폴더 생성
    forms_dir = os.path.join(target_folder, "서식")
    try:
        os.makedirs(forms_dir, exist_ok=True)
        logger.info(f"서식 폴더 생성: {forms_dir}")
        return True, forms_dir
    except Exception as e:
        logger.error(f"서식 폴더 생성 실패: {e}")
        return False, None

def analyze_form_templates(
    pdf_paths: List[str], 
    progress_callback: Optional[Callable[[int], None]] = None,
    log_callback: Optional[Callable[[str], None]] = None,
    folder_name: Optional[str] = None  # 명시적 폴더명 인자 추가
) -> Dict[str, Any]:
    """
    PDF 파일 목록에서 서식 페이지를 찾아 분석하는 함수
    
    Args:
        pdf_paths: PDF 파일 경로 목록
        progress_callback: 진행률 콜백 함수 (0-100)
        log_callback: 로그 메시지 콜백 함수
        folder_name: 명시적 폴더명 지정
        
    Returns:
        분석 결과 딕셔너리 (doc, forms 키를 포함)
    """
    if not CHATGPT_API_KEY:
        log_msg = "CHATGPT_API_KEY가 설정되지 않았습니다. .env 파일을 확인하세요."
        if log_callback:
            log_callback(log_msg)
        raise ValueError(log_msg)
    
    # 최신 openai 방식
    client = OpenAI(api_key=CHATGPT_API_KEY)
    
    # 임시 폴더 생성 (중간 처리용)
    output_dir = tempfile.mkdtemp()
    log_msg = f"임시 처리 폴더 생성: {output_dir}"
    logger.info(log_msg)
    if log_callback:
        log_callback(log_msg)
    
    # 드롭박스 폴더 찾기 (첫 번째 PDF 파일 기준) - 수정 필요
    forms_dir = None
    dropbox_target_folder = None
    if pdf_paths:
        # 폴더명이 지정되었으면 해당 폴더와 관련된 처리
        if folder_name and DROPBOX_SHARED_FOLDER_NAME:
            dropbox_target_folder = f"{DROPBOX_SHARED_FOLDER_NAME}/{folder_name}"
            if log_callback:
                log_callback(f"서식 저장 폴더: {dropbox_target_folder}/서식")
        
        # 기존 방식대로 처리
        success, dropbox_forms_dir = create_dropbox_forms_dir(pdf_paths[0])
        if success:
            forms_dir = dropbox_forms_dir
            log_msg = f"드롭박스 서식 폴더 생성: {forms_dir}"
            logger.info(log_msg)
            if log_callback:
                log_callback(log_msg)
    
    # 임시 서식 폴더 생성 (백업용)
    temp_forms_dir = os.path.join(output_dir, "서식")
    os.makedirs(temp_forms_dir, exist_ok=True)
    
    try:
        # 분석된 PDF 파일 정보 기록
        analyzed_files = []
        for path in pdf_paths:
            filename = os.path.basename(path)
            analyzed_files.append({
                "filename": filename,
                "path": path,
                "size": os.path.getsize(path) if os.path.exists(path) else 0,
                "timestamp": os.path.getmtime(path) if os.path.exists(path) else 0
            })
        
        # PDF 텍스트 추출
        all_texts = []
        for i, path in enumerate(pdf_paths):
            filename = os.path.basename(path)
            log_msg = f"PDF 텍스트 추출 중: {filename}"
            logger.info(log_msg)
            if log_callback:
                log_callback(log_msg)
            
            text = extract_text_from_pdf(path)
            all_texts.append(f"\n=== FILE: {filename} ===\n{text}")
            
            # 진행률 업데이트 (텍스트 추출 단계: 0-30%)
            if progress_callback:
                progress_callback(i * 30 // len(pdf_paths))
        
        # 모든 PDF 텍스트 결합
        combined_text = "\n\n".join(all_texts)
        
        # 프롬프트 준비
        system_prompt = "당신은 대한민국 공공 입찰 서류의 '제출용 서식(양식)' 페이지를 정확히 식별하는 전문가입니다.\n"\
        "여러 개의 PDF 문서를 분석하여, **입찰 참여자가 작성·제출해야 하는 모든 '서식' 페이지**를 찾아내세요.\n"\
        "'서식'이란 AcroForm이든 스캔본이든 상관없이 \"입찰참가신청서\", \"청렴계약 이행각서\", \"별지 제1호 서식\" 등 제출용 양식을 말합니다.\n\n"\
        "**출력 형식** (JSON 배열, 순수 JSON만):\n"\
        "[\n"\
        "  {\n"\
        "    \"doc\": \"문서 파일명.pdf\",\n"\
        "    \"forms\": [\n"\
        "      {\n"\
        "        \"page\": 페이지 번호,\n"\
        "        \"title\": \"서식 정확한 제목\",\n"\
        "        \"filename\": \"12p_입찰참가신청서.pdf\",\n"\
        "        \"requires_input\": true\n"\
        "      },\n"\
        "      ...\n"\
        "    ]\n"\
        "  },\n"\
        "  ...\n"\
        "]\n"\
        "forms가 없으면 빈 배열 (\"forms\": [])로 반환\n"\
        "추가 설명, 주석, 텍스트는 절대 포함 금지"
        
        # 사용자 프롬프트 구성
        user_text = f"아래는 {len(pdf_paths)}개 PDF의 텍스트 추출 내용입니다. 문서별로 구분하기 위해 다음 포맷으로 섹션을 나누었습니다:\n\n{combined_text}\n\n위 각 섹션을 분석해, 제출용 '서식' 페이지를 모두 찾아 JSON으로 반환해주세요."
        
        # 진행률 업데이트 (API 호출 준비: 30%)
        if progress_callback:
            progress_callback(30)
        
        # OpenAI API 호출
        log_msg = "OpenAI API 호출 중..."
        logger.info(log_msg)
        if log_callback:
            log_callback(log_msg)
        
        response = client.chat.completions.create(
            model=GPT_MODEL,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_text}
            ]
        )
        
        # 진행률 업데이트 (API 호출 완료: 60%)
        if progress_callback:
            progress_callback(60)
        
        # API 응답 처리
        content = response.choices[0].message.content
        if DEBUG:
            logger.info(f"API 응답: {content}")
        
        # JSON 추출 (텍스트에서 JSON 부분만 추출)
        try:
            # JSON 시작/끝 위치 찾기 (배열이나 객체 형태)
            json_start = min(content.find("["), content.find("{")) if content.find("[") >= 0 and content.find("{") >= 0 else max(content.find("["), content.find("{"))
            json_end = max(content.rfind("]"), content.rfind("}")) + 1
            
            if json_start >= 0 and json_end > json_start:
                json_content = content[json_start:json_end]
                json_result = json.loads(json_content)
                
                # 배열인 경우 첫 번째 항목을 기본 문서로 처리하고 나머지는 통합
                if isinstance(json_result, list):
                    # 결과가 배열 형태인 경우, 모든 서식을 통합
                    all_forms = []
                    doc_name = ""
                    
                    for doc_result in json_result:
                        if doc_result.get("forms"):
                            all_forms.extend(doc_result.get("forms", []))
                        if not doc_name and doc_result.get("doc"):
                            doc_name = doc_result.get("doc")
                    
                    # 통합 결과 생성
                    result = {
                        "doc": doc_name or os.path.basename(pdf_paths[0]) if pdf_paths else "",
                        "forms": all_forms,
                        "analyzed_files": analyzed_files,
                        "multi_doc_result": json_result  # 원본 다중 문서 결과 보존
                    }
                else:
                    # 결과가 객체 형태인 경우, 그대로 사용
                    result = json_result
                    
                    # 기본 필드 보장
                    if not result.get("doc") and pdf_paths:
                        result["doc"] = os.path.basename(pdf_paths[0])
                    
                    # 분석한 파일 목록 추가
                    result["analyzed_files"] = analyzed_files
                
                # 서식 PDF 파일 생성 시작 부분에 추가 (result["forms"] 처리 전)
                if result.get("forms"):
                    # 진행률 업데이트 (PDF 생성 시작: 70%)
                    if progress_callback:
                        progress_callback(70)
                    
                    # Dropbox 경로 설정
                    if dropbox_target_folder:
                        dropbox_forms_path = f"{dropbox_target_folder}/서식"
                    else:
                        # 첫 번째 PDF 파일명에서 폴더명 추출
                        pdf_name = os.path.basename(pdf_paths[0]) if pdf_paths else ""
                        temp_folder_name = os.path.splitext(pdf_name)[0]
                        dropbox_forms_path = f"{DROPBOX_SHARED_FOLDER_NAME}/{temp_folder_name}/서식"
                    
                    # 각 서식별 PDF 파일 생성
                    successful_forms = []
                    for i, form in enumerate(result["forms"]):
                        page = form.get("page")
                        if page is None:
                            continue
                        
                        # 파일명 생성
                        filename = form.get("filename", f"{page}p_서식.pdf")
                        filename = re.sub(r'[\\/*?:"<>|]', "", filename)
                        
                        # 임시 경로와 최종 경로 설정 - 항상 드롭박스 폴더 우선 사용
                        if forms_dir:
                            # 드롭박스 폴더가 있으면 최종 파일은 드롭박스에 저장
                            final_output_path = os.path.join(forms_dir, filename)
                            # 임시 파일은 더 이상 필요 없음 (직접 최종 경로에 저장)
                            temp_output_path = final_output_path
                        else:
                            # 드롭박스 폴더가 없으면 임시 폴더에 저장
                            temp_output_path = os.path.join(temp_forms_dir, filename)
                            final_output_path = temp_output_path
                        
                        log_msg = f"서식 파일 생성 중: {filename}"
                        logger.info(log_msg)
                        if log_callback:
                            log_callback(log_msg)
                        
                        # 서식 파일이 있는 PDF 및 페이지 번호 찾기
                        target_doc = form.get("doc")
                        
                        # 서식이 발견된 원본 문서 확인
                        target_paths = [p for p in pdf_paths if os.path.basename(p) == target_doc] if target_doc else pdf_paths
                        
                        # 대상 문서가 지정되지 않았으면 모든 PDF 확인
                        if not target_paths:
                            target_paths = pdf_paths
                            
                        # 해당 페이지가 있는 PDF 찾기
                        for pdf_path in target_paths:
                            try:
                                reader = PdfReader(pdf_path)
                                if page <= len(reader.pages):
                                    # 0-기반 인덱스로 변환
                                    page_idx = page - 1
                                    
                                    # 단일 페이지 추출
                                    writer = PdfWriter()
                                    writer.add_page(reader.pages[page_idx])
                                    
                                    # 파일로 저장 (최종 경로에 직접 저장)
                                    with open(final_output_path, "wb") as out_file:
                                        writer.write(out_file)
                                    
                                    log_msg = f"서식 파일 저장 완료: {final_output_path}"
                                    logger.info(log_msg)
                                    if log_callback:
                                        log_callback(log_msg)
                                    
                                    # 원본 파일 경로 추가
                                    form["source_pdf"] = os.path.basename(pdf_path)
                                    form["output_path"] = final_output_path  # 이제 출력 경로는 최종 경로와 동일
                                    form["final_path"] = final_output_path  # 최종 경로도 설정
                                    form["dropbox_path"] = f"{dropbox_forms_path}/{filename}"  # Dropbox 경로 추가
                                    successful_forms.append(form)
                                    break
                            except Exception as e:
                                error_msg = f"서식 추출 오류 (페이지 {page}): {e}"
                                logger.error(error_msg)
                                if log_callback:
                                    log_callback(error_msg)
                        
                        # 진행률 업데이트 (PDF 생성 진행: 70-90%)
                        if progress_callback:
                            progress_callback(70 + (i * 20) // len(result["forms"]))
                    
                    # 결과 업데이트
                    result["forms"] = successful_forms
                    result["forms_generated"] = len(successful_forms)
                    
                    # 분석 결과 JSON 파일 저장 (항상 서식 파일과 같은 위치에 저장)
                    result_json_path = ""
                    if forms_dir:
                        result_json_path = os.path.join(forms_dir, "서식분석결과.json")
                    else:
                        result_json_path = os.path.join(temp_forms_dir, "서식분석결과.json")
                    
                    try:
                        with open(result_json_path, "w", encoding="utf-8") as f:
                            json.dump(result, f, ensure_ascii=False, indent=2)
                        log_msg = f"분석 결과 JSON 저장 완료: {result_json_path}"
                        logger.info(log_msg)
                        if log_callback:
                            log_callback(log_msg)
                    except Exception as e:
                        error_msg = f"분석 결과 JSON 저장 오류: {e}"
                        logger.error(error_msg)
                        if log_callback:
                            log_callback(error_msg)
                
                # 결과 반환 전에 Dropbox 업로드 코드 추가
                # 분석 완료 후 서식 파일 및 결과 JSON을 Dropbox에 업로드
                try:
                    if result.get("forms"):
                        log_msg = f"Dropbox 업로드 시작: {dropbox_forms_path}"
                        logger.info(log_msg)
                        if log_callback:
                            log_callback(log_msg)
                            
                        # 서식 파일 업로드
                        for form in result["forms"]:
                            output_path = form.get("output_path")
                            if output_path and os.path.exists(output_path):
                                dropbox_path = form.get("dropbox_path")
                                if dropbox_path:
                                    try:
                                        upload_file(dropbox_path, output_path)
                                        log_msg = f"Dropbox 업로드 완료: {os.path.basename(output_path)}"
                                        logger.info(log_msg)
                                        if log_callback:
                                            log_callback(log_msg)
                                    except Exception as e:
                                        log_msg = f"Dropbox 업로드 실패: {e}"
                                        logger.error(log_msg)
                                        if log_callback:
                                            log_callback(log_msg)
                        
                        # 결과 JSON 업로드
                        json_path = f"{dropbox_forms_path}/서식분석결과.json"
                        try:
                            upload_json(json_path, result)
                            log_msg = f"결과 JSON Dropbox 업로드 완료: {json_path}"
                            logger.info(log_msg)
                            if log_callback:
                                log_callback(log_msg)
                        except Exception as e:
                            log_msg = f"결과 JSON 업로드 실패: {e}"
                            logger.error(log_msg)
                            if log_callback:
                                log_callback(log_msg)
                except Exception as e:
                    log_msg = f"Dropbox 업로드 중 오류: {e}"
                    logger.error(log_msg)
                    if log_callback:
                        log_callback(log_msg)
                
                # 진행률 업데이트 (완료: 95%)
                if progress_callback:
                    progress_callback(95)
                
                # 최종 저장 위치 기록
                result["forms_dir"] = forms_dir or temp_forms_dir
                return result
        except Exception as e:
            error_msg = f"JSON 파싱 오류: {e}"
            logger.error(error_msg)
            if log_callback:
                log_callback(error_msg)
        
        # JSON 파싱 실패 시 백업 처리
        try:
            # 정규식으로 페이지, 제목 추출
            result = {
                "doc": os.path.basename(pdf_paths[0]) if pdf_paths else "", 
                "forms": [],
                "analyzed_files": analyzed_files
            }
            
            log_msg = "JSON 파싱 실패, 백업 처리 시작..."
            logger.info(log_msg)
            if log_callback:
                log_callback(log_msg)
            
            # 다양한 패턴으로 검색 시도
            patterns = [
                r'page[^\d]*(\d+).*?title[^\w가-힣]*([\w가-힣]+)',  # 기본 패턴
                r'"page"[^\d]*(\d+).*?"title"[^\w가-힣]*"([\w가-힣]+)"',  # JSON 스타일
                r'서식.*?페이지.*?(\d+).*?제목.*?[\'"]([^\'"]+)[\'"]',  # 한글 설명 패턴
                r'별지\s*제\s*(\d+)\s*호',  # 별지 서식 패턴
                r'서식\s*제\s*(\d+)\s*호',  # 서식 번호 패턴
                r'(입찰참가신청서|청렴계약\s*이행각서|입찰인감증명서|가격제안서|견적서)'  # 일반적인 서식명
            ]
            
            # 각 패턴별로 매칭 시도
            forms_found = []
            
            for pattern in patterns:
                matches = re.findall(pattern, content, re.IGNORECASE | re.DOTALL)
                if matches:
                    for match in matches:
                        if isinstance(match, tuple) and len(match) >= 2:
                            # 페이지 번호와 서식명이 있는 경우
                            try:
                                page = int(match[0])
                                title = match[1].strip()
                                forms_found.append({
                                    "page": page,
                                    "title": title,
                                    "filename": f"{page}p_{title}.pdf",
                                    "requires_input": True
                                })
                            except:
                                pass
                        elif isinstance(match, str):
                            # 서식명만 있는 경우 (페이지는 불명확)
                            forms_found.append({
                                "title": match.strip(),
                                "requires_input": True
                            })
            
            # 중복 제거 및 페이지 번호별 정렬
            if forms_found:
                seen_titles = set()
                unique_forms = []
                
                for form in forms_found:
                    title = form.get("title", "")
                    if title and title not in seen_titles:
                        seen_titles.add(title)
                        unique_forms.append(form)
                
                # 페이지 번호가 있는 항목들 우선 정렬
                result["forms"] = sorted(
                    [f for f in unique_forms if "page" in f],
                    key=lambda x: x.get("page", 999)
                ) + [f for f in unique_forms if "page" not in f]
                
                # 백업 처리로 서식 폴더에 저장
                if forms_dir:
                    backup_forms_dir = forms_dir
                else:
                    backup_forms_dir = temp_forms_dir
                
                log_msg = f"백업 처리: 서식 폴더 사용 {backup_forms_dir}"
                logger.info(log_msg)
                if log_callback:
                    log_callback(log_msg)
                
                # 페이지 번호가 있는 서식만 처리
                for form in [f for f in result["forms"] if "page" in f]:
                    page = form.get("page")
                    if page is None:
                        continue
                        
                    # 파일명 생성
                    filename = form.get("filename", f"{page}p_서식.pdf")
                    filename = re.sub(r'[\\/*?:"<>|]', "", filename)
                    final_output_path = os.path.join(backup_forms_dir, filename)
                    
                    # 첫 번째 PDF에서 해당 페이지 추출 시도
                    if pdf_paths:
                        try:
                            reader = PdfReader(pdf_paths[0])
                            if page <= len(reader.pages):
                                # 0-기반 인덱스로 변환
                                page_idx = page - 1
                                
                                # 단일 페이지 추출
                                writer = PdfWriter()
                                writer.add_page(reader.pages[page_idx])
                                
                                # 최종 위치에 저장
                                with open(final_output_path, "wb") as out_file:
                                    writer.write(out_file)
                                
                                log_msg = f"백업 처리: 서식 파일 저장 {final_output_path}"
                                logger.info(log_msg)
                                if log_callback:
                                    log_callback(log_msg)
                                    
                                # 경로 정보 추가    
                                form["final_path"] = final_output_path
                        except Exception as e:
                            error_msg = f"백업 처리 서식 추출 오류 (페이지 {page}): {e}"
                            logger.error(error_msg)
                            if log_callback:
                                log_callback(error_msg)
                
                # 분석 결과 JSON 파일도 저장
                result_json_path = os.path.join(backup_forms_dir, "서식분석결과.json")
                try:
                    with open(result_json_path, "w", encoding="utf-8") as f:
                        json.dump(result, f, ensure_ascii=False, indent=2)
                    log_msg = f"백업 처리: 분석 결과 JSON 저장 완료 {result_json_path}"
                    logger.info(log_msg)
                    if log_callback:
                        log_callback(log_msg)
                except Exception as e:
                    error_msg = f"백업 처리 JSON 저장 오류: {e}"
                    logger.error(error_msg)
                    if log_callback:
                        log_callback(error_msg)
            
            # 최종 저장 위치 기록
            result["forms_dir"] = forms_dir or temp_forms_dir
            
            # 진행률 업데이트 (완료: 100%)
            if progress_callback:
                progress_callback(100)
            
            return result
        except Exception as e:
            error_msg = f"백업 처리 오류: {e}"
            logger.error(error_msg)
            if log_callback:
                log_callback(error_msg)
            return {
                "doc": "", 
                "forms": [], 
                "error": str(e),
                "analyzed_files": analyzed_files,
                "forms_dir": forms_dir or temp_forms_dir
            }
    
    except Exception as e:
        error_msg = f"서식 분석 오류: {e}"
        logger.error(error_msg)
        if log_callback:
            log_callback(error_msg)
        return {
            "doc": "", 
            "forms": [], 
            "error": str(e),
            "analyzed_files": [os.path.basename(p) for p in pdf_paths],
            "forms_dir": forms_dir or temp_forms_dir
        }
    finally:
        # 임시 폴더는 삭제하지 않음 - 임시 폴더가 실제 저장 위치일 수 있음
        # 임시 폴더 경로는 결과에 포함되어 있으므로, 호출자가 적절히 처리해야 함
        pass

def save_form_templates(
    result: Dict[str, Any], 
    destination_folder: str,
    dropbox_prefix: str = "입찰 2025"
) -> int:
    """
    서식 분석 결과를 Dropbox에 저장
    
    Args:
        result: 서식 분석 결과 딕셔너리
        destination_folder: Dropbox 대상 폴더명
        dropbox_prefix: Dropbox 기본 경로
        
    Returns:
        저장된 서식 파일 수
    """
    try:
        logger.info(f"서식 저장 시작: {destination_folder}")
        
        # 결과 JSON 파일 저장
        upload_json(f"{dropbox_prefix}/{destination_folder}/서식분석결과.json", result)
        
        # 서식 파일이 있으면 PDF 업로드
        successful = 0
        for form in result.get("forms", []):
            output_path = form.get("output_path")
            if not output_path or not os.path.exists(output_path):
                continue
                
            filename = os.path.basename(output_path)
            remote_path = f"{dropbox_prefix}/{destination_folder}/서식/{filename}"
            
            try:
                # 파일 업로드
                upload_file(remote_path, output_path)
                successful += 1
            except Exception as e:
                logger.error(f"서식 파일 업로드 오류 ({filename}): {e}")
        
        logger.info(f"서식 저장 완료: {successful}/{len(result.get('forms', []))} 파일")
        return successful
    
    except Exception as e:
        logger.error(f"서식 저장 오류: {e}")
        return 0

# 테스트용 코드
if __name__ == "__main__":
    # 테스트 PDF 파일
    test_pdf = "example.pdf"
    
    if os.path.exists(test_pdf):
        print(f"PDF 분석 중: {test_pdf}")
        result = analyze_form_templates([test_pdf])
        print(json.dumps(result, indent=2, ensure_ascii=False))
    else:
        print(f"테스트 PDF 파일이 없습니다: {test_pdf}")