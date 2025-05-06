import json
import logging
from typing import List, Dict, Any
import os
import tempfile
from PyQt5.QtWidgets import QMessageBox, QProgressDialog, QApplication
from PyQt5.QtCore import Qt
import openai

from PyPDF2 import PdfReader

# 루트 설정 파일 임포트로 변경
from settings import settings

logger = logging.getLogger(__name__)

# 1) OpenAI API 키 및 모델 설정
MODEL = settings.GPT_MODEL  # e.g. "gpt-4.1-mini"
API_KEY = settings.CHATGPT_API_KEY

# 2) 시스템 프롬프트 (분석 포맷 안내)
SYSTEM_PROMPT = """
당신은 대한민국 정부 입찰문서를 분석해 아래 여섯 가지 항목을 정확히 추출하여 JSON 형태로 반환하는 전문가입니다.

1. 공고 정보 (announcement_info)  
   - 등록마감  
   - 공고명  
   - 추정가격  
2. 프로젝트 성격 1줄 요약 (project_summary)  
3. 입찰 관련사항 요약 (bid_summary)  
4. 제안요청서에 담긴 핵심 사항 (rfp_core_items)  
5. 입찰 제출 서류 (submission_documents)  
6. 제안요청서 목차 (rfp_table_of_contents)  

출력은 반드시 **순수 JSON**으로만, 키는 다음과 같이 사용하세요:
- "announcement_info"      : 객체 (object)  
  - "등록마감": 문자열  
  - "공고명": 문자열  
  - "추정가격": 문자열  
- "project_summary"        : 문자열  
- "bid_summary"            : 문자열 배열  
- "rfp_core_items"         : 문자열 배열  
- "submission_documents"   : 문자열 배열  
- "rfp_table_of_contents"  : 문자열 배열  
"""

def extract_text_from_pdf(path: str) -> str:
    """
    PyPDF2를 사용해 PDF 전체 페이지의 텍스트를 추출하여 하나의 문자열로 반환합니다.
    """
    try:
        reader = PdfReader(path)
    except Exception as e:
        logger.error(f"Failed to open PDF {path}: {e}")
        return ""
    text_parts = []
    for i, page in enumerate(reader.pages):
        try:
            text = page.extract_text() or ""
        except Exception as e:
            logger.warning(f"Failed to extract text from {path} page {i}: {e}")
            text = ""
        text_parts.append(f"===PAGE {i+1}===\n{text}")
    return "\n".join(text_parts)

def analyze_pdfs(pdf_paths, parent=None):
    try:
        # Create a temporary file to store the PDF paths
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.json') as temp_file:
            json.dump(pdf_paths, temp_file)
            temp_file_path = temp_file.name

        # Create progress dialog
        progress = QProgressDialog("PDF 분석 중...", "취소", 0, 100, parent)
        progress.setWindowModality(Qt.WindowModal)
        progress.show()

        # Initialize OpenAI client
        client = openai.OpenAI(api_key=settings.OPENAI_API_KEY)

        # Read the PDF paths from the temporary file
        with open(temp_file_path, 'r') as f:
            pdf_paths = json.load(f)

        # Process each PDF
        for i, pdf_path in enumerate(pdf_paths):
            if progress.wasCanceled():
                break

            # Update progress
            progress.setValue(int((i / len(pdf_paths)) * 100))
            QApplication.processEvents()

            # Process PDF here
            # Add your PDF processing logic

        # Clean up
        os.unlink(temp_file_path)
        progress.setValue(100)

        return True

    except Exception as e:
        QMessageBox.critical(parent, "오류", f"PDF 분석 중 오류가 발생했습니다: {str(e)}")
        return False

def clean_gpt_response(content: str) -> str:
    """
    ChatGPT가 반환한 응답에서 마크다운 코드 블록 구문 등을 제거합니다.
    
    예) ```json { ... } ``` -> { ... }
    """
    import re
    
    # 마크다운 코드 블록 제거
    content = re.sub(r'^```(\w+)?\s*', '', content, flags=re.MULTILINE)
    content = re.sub(r'\s*```$', '', content, flags=re.MULTILINE)
    
    # 추가 공백 및 줄바꿈 정리
    content = content.strip()
    
    return content 