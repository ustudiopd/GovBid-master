import json
import logging
from typing import List, Dict, Any

import openai
from PyPDF2 import PdfReader

# 루트 설정 파일 임포트로 변경
from settings import settings

logger = logging.getLogger(__name__)

# 1) OpenAI API 키 및 모델 설정
openai.api_key = settings.CHATGPT_API_KEY
MODEL = settings.GPT_MODEL  # e.g. "gpt-4.1-mini"

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

def analyze_pdfs(pdf_paths: List[str]) -> Dict[str, Any]:
    """
    주어진 PDF 경로 리스트를 읽어 ChatGPT API에 분석 요청,
    반환된 JSON 문자열을 파싱하여 dict로 반환합니다.
    """
    # 1) PDF 텍스트 합치기
    docs = []
    for p in pdf_paths:
        docs.append(f"\n===DOCUMENT: {p}===\n" + extract_text_from_pdf(p))
    full_text = "\n".join(docs)

    # 2) 사용자 프롬프트: 문서 텍스트 포함
    user_prompt = f"""
다음은 PDF 문서의 페이지별 텍스트입니다.  
각 페이지 앞에 "===PAGE n===" 로 구분되어 있습니다.  

===DOCUMENT_TEXT===
{full_text}

위 텍스트를 읽고, 여섯 가지 항목을 추출하여 순수 JSON으로만 답해주세요.
"""
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": user_prompt},
    ]

    # 3) ChatGPT 호출
    try:
        resp = openai.ChatCompletion.create(
            model=MODEL,
            messages=messages,
            temperature=0.0,
        )
    except Exception as e:
        logger.error(f"OpenAI ChatCompletion API call failed: {e}")
        raise

    # 4) 응답에서 content 추출
    content = resp.choices[0].message.content
    logger.debug(f"GPT response content: {content}")

    # 4-1) 마크다운 코드 블록 제거 (```json ... ```)
    content = clean_gpt_response(content)

    # 5) JSON 파싱
    try:
        result = json.loads(content)
    except json.JSONDecodeError as e:
        logger.error(f"Failed to parse JSON from GPT response: {e}")
        # 파싱 에러 시 원본을 함께 던집니다.
        raise ValueError(f"Invalid JSON response: {content}") from e

    return result

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