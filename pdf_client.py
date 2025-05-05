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
import openai

# Dropbox 클라이언트 임포트
from dropbox_client import upload_file, upload_json

# 로깅 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# .env 파일에서 API 키 로드
load_dotenv()
CHATGPT_API_KEY = os.getenv("CHATGPT_API_KEY")
GPT_MODEL = os.getenv("CHATGPT_MODEL", "gpt-4.1-mini")

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

def analyze_form_templates(
    pdf_paths: List[str], 
    progress_callback: Optional[Callable[[int], None]] = None,
    log_callback: Optional[Callable[[str], None]] = None
) -> Dict[str, Any]:
    """
    PDF 파일 목록에서 서식 페이지를 찾아 분석하는 함수
    
    Args:
        pdf_paths: PDF 파일 경로 목록
        progress_callback: 진행률 콜백 함수 (0-100)
        log_callback: 로그 메시지 콜백 함수
        
    Returns:
        분석 결과 딕셔너리 (doc, forms 키를 포함)
    """
    if not CHATGPT_API_KEY:
        log_msg = "CHATGPT_API_KEY가 설정되지 않았습니다. .env 파일을 확인하세요."
        if log_callback:
            log_callback(log_msg)
        raise ValueError(log_msg)
    
    # OpenAI API 설정
    openai.api_key = CHATGPT_API_KEY
    
    # 임시 폴더 생성 (서식 PDF 저장용)
    output_dir = tempfile.mkdtemp()
    log_msg = f"임시 출력 폴더 생성: {output_dir}"
    logger.info(log_msg)
    if log_callback:
        log_callback(log_msg)
    
    # 원본 PDF가 있는 폴더 확인 (첫 번째 파일 기준)
    original_pdf_dir = None
    if pdf_paths and len(pdf_paths) > 0:
        original_pdf_dir = os.path.dirname(pdf_paths[0])
        # 모든 PDF가 같은 폴더에 있는지 확인
        same_dir = all(os.path.dirname(p) == original_pdf_dir for p in pdf_paths)
        if not same_dir:
            log_msg = "분석 대상 PDF 파일들이 서로 다른 폴더에 위치해 있습니다."
            logger.warning(log_msg)
            if log_callback:
                log_callback(log_msg)
    
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
            
        response = openai.ChatCompletion.create(
            model=GPT_MODEL,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_text}
            ],
            temperature=0.1,
            max_tokens=1000
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
                
                # 원본 PDF 폴더에 서식 폴더 생성
                if original_pdf_dir and result.get("forms"):
                    # 드롭박스 경로 찾기 시도 - 원본 PDF 경로 상위 폴더 탐색
                    dropbox_path = find_dropbox_folder(original_pdf_dir)
                    
                    # 최종 저장 경로 설정
                    save_dir = original_pdf_dir
                    if dropbox_path:
                        # 드롭박스 폴더를 찾았다면 해당 경로 사용
                        log_msg = f"드롭박스 폴더 발견: {dropbox_path}"
                        logger.info(log_msg)
                        if log_callback:
                            log_callback(log_msg)
                        
                        # 원본 경로에서 상대 경로 추출 (드롭박스 폴더 이후 부분)
                        relative_path = os.path.relpath(original_pdf_dir, dropbox_path)
                        if relative_path == ".":
                            relative_path = ""
                        
                        # 드롭박스 루트 폴더의 "입찰 2025" 확인
                        bid_folder = os.path.join(dropbox_path, "입찰 2025")
                        if os.path.exists(bid_folder) and os.path.isdir(bid_folder):
                            # 원본이 입찰 2025 폴더 내에 있는지 확인
                            if "입찰 2025" in original_pdf_dir:
                                # 기존 경로 그대로 사용
                                save_dir = original_pdf_dir
                            else:
                                # 입찰 2025 폴더의 하위 폴더로 저장
                                save_dir = bid_folder
                                for part in relative_path.split(os.sep):
                                    if part and part != "입찰 2025":
                                        save_dir = os.path.join(save_dir, part)
                        else:
                            # 입찰 2025 폴더가 없으면 원본 경로에 저장
                            save_dir = original_pdf_dir
                    
                    # 원본 PDF가 있는 폴더에 "서식" 하위 폴더 생성
                    forms_dir = os.path.join(save_dir, "서식")
                    os.makedirs(forms_dir, exist_ok=True)
                    
                    log_msg = f"서식 폴더 생성: {forms_dir}"
                    logger.info(log_msg)
                    if log_callback:
                        log_callback(log_msg)
                    
                    # 임시 폴더에도 "서식" 폴더 생성 (함수 내 처리용)
                    forms_output_dir = os.path.join(output_dir, "서식")
                    os.makedirs(forms_output_dir, exist_ok=True)
                    
                    # 진행률 업데이트 (PDF 생성 시작: 70%)
                    if progress_callback:
                        progress_callback(70)
                    
                    # 각 서식별 PDF 파일 생성
                    successful_forms = []
                    for i, form in enumerate(result["forms"]):
                        page = form.get("page")
                        if page is None:
                            continue
                        
                        # 파일명 생성
                        filename = form.get("filename", f"{page}p_서식.pdf")
                        filename = re.sub(r'[\\/*?:"<>|]', "", filename)
                        
                        # 임시 폴더 내 출력 경로
                        temp_output_path = os.path.join(forms_output_dir, filename)
                        # 최종 저장 경로 (드롭박스 폴더 내 서식 폴더)
                        final_output_path = os.path.join(forms_dir, filename)
                        
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
                                    
                                    # 임시 파일로 저장
                                    with open(temp_output_path, "wb") as out_file:
                                        writer.write(out_file)
                                    
                                    # 최종 위치에도 저장
                                    with open(final_output_path, "wb") as out_file:
                                        writer.write(out_file)
                                    
                                    log_msg = f"서식 파일 저장 완료: {final_output_path}"
                                    logger.info(log_msg)
                                    if log_callback:
                                        log_callback(log_msg)
                                    
                                    # 원본 파일 경로 추가
                                    form["source_pdf"] = os.path.basename(pdf_path)
                                    form["output_path"] = temp_output_path  # 임시 경로 기록
                                    form["final_path"] = final_output_path  # 최종 경로 기록
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
                    
                    # 분석 결과 JSON 파일도 원본 폴더에 저장
                    # 서식분석결과.json 파일을 서식 폴더에 저장
                    result_json_path = os.path.join(forms_dir, "서식분석결과.json")
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
                
                # 진행률 업데이트 (완료: 95%)
                if progress_callback:
                    progress_callback(95)
                
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
                
                # 원본 PDF 폴더에 서식 생성 시도
                if original_pdf_dir and result["forms"]:
                    # 드롭박스 경로 찾기 시도
                    dropbox_path = find_dropbox_folder(original_pdf_dir)
                    
                    # 최종 저장 경로 설정
                    save_dir = original_pdf_dir
                    if dropbox_path:
                        # 원본 경로에서 상대 경로 추출 (드롭박스 폴더 이후 부분)
                        relative_path = os.path.relpath(original_pdf_dir, dropbox_path)
                        if relative_path == ".":
                            relative_path = ""
                        
                        # 드롭박스 루트 폴더의 "입찰 2025" 확인
                        bid_folder = os.path.join(dropbox_path, "입찰 2025")
                        if os.path.exists(bid_folder) and os.path.isdir(bid_folder):
                            # 원본이 입찰 2025 폴더 내에 있는지 확인
                            if "입찰 2025" in original_pdf_dir:
                                # 기존 경로 그대로 사용
                                save_dir = original_pdf_dir
                            else:
                                # 입찰 2025 폴더의 하위 폴더로 저장
                                save_dir = bid_folder
                                for part in relative_path.split(os.sep):
                                    if part and part != "입찰 2025":
                                        save_dir = os.path.join(save_dir, part)
                        else:
                            # 입찰 2025 폴더가 없으면 원본 경로에 저장
                            save_dir = original_pdf_dir
                    
                    forms_dir = os.path.join(save_dir, "서식")
                    os.makedirs(forms_dir, exist_ok=True)
                    
                    log_msg = f"백업 처리: 서식 폴더 생성 {forms_dir}"
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
                        final_output_path = os.path.join(forms_dir, filename)
                        
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
                    result_json_path = os.path.join(forms_dir, "서식분석결과.json")
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
                "analyzed_files": analyzed_files
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
            "analyzed_files": [os.path.basename(p) for p in pdf_paths]
        }
    finally:
        # 임시 폴더 삭제 - 이미 최종 위치에 파일을 저장했으므로 임시 폴더는 삭제
        try:
            shutil.rmtree(output_dir)
        except:
            pass

def find_dropbox_folder(start_path):
    """지정된 경로에서 드롭박스 폴더를 찾아 경로 반환"""
    # 드롭박스 폴더명 (일반적인 패턴)
    dropbox_names = ["Dropbox", "드롭박스"]
    
    # 현재 디렉토리가 드롭박스인지 확인
    current_dir = os.path.basename(start_path)
    if current_dir in dropbox_names:
        return start_path
    
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
    
    # 일반적인 드롭박스 설치 경로 확인
    potential_paths = [
        os.path.expanduser("~/Dropbox"),
        os.path.expanduser("~/Documents/Dropbox"),
        os.path.join(os.environ.get("USERPROFILE", ""), "Dropbox"),
        os.path.join(os.environ.get("USERPROFILE", ""), "Documents", "Dropbox"),
        "D:/Dropbox",
        "C:/Dropbox",
        "D:/문서/Dropbox",
        "C:/문서/Dropbox",
        "D:/Documents/Dropbox",
        "C:/Documents/Dropbox"
    ]
    
    for path in potential_paths:
        if os.path.exists(path) and os.path.isdir(path):
            # 입찰 폴더 확인
            bid_folder = os.path.join(path, "입찰 2025")
            if os.path.exists(bid_folder) and os.path.isdir(bid_folder):
                return path
    
    # 못 찾으면 None 반환
    return None

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