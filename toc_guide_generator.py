import os
import json
import tempfile
from PyQt5.QtWidgets import QMessageBox, QProgressDialog, QApplication
from dropbox_client import list_folder, download_file, upload_json
from gpt_client import analyze_pdfs
import glob
import openai
from settings import settings

MODEL = settings.GPT_MODEL  # e.g. "gpt-4.1-mini"

class TocGuideGenerator:
    """목차 가이드 생성 클래스"""
    
    @staticmethod
    def generate_guide(folder, parent=None):
        """
        지정된 폴더의 PDF 파일을 분석하여 목차 가이드 생성
        
        Args:
            folder: 분석할 폴더명
            parent: 부모 위젯 (QMessageBox 표시용)
            
        Returns:
            성공 여부 (boolean)
        """
        try:
            # 폴더 내 PDF 파일 목록 가져오기
            files = list_folder(f"입찰 2025/{folder}")
            pdfs = [f for f in files if f.lower().endswith(".pdf")]
            
            if not pdfs:
                QMessageBox.warning(parent, "PDF 없음", f"{folder} 폴더에 PDF 파일이 없습니다.")
                return False
            
            # 진행 상태 대화상자 생성
            progress = QProgressDialog("목차 가이드 생성 중...", "취소", 0, 100, parent)
            progress.setWindowTitle("목차 가이드 생성")
            progress.setModal(True)
            progress.show()
            
            # 작업 진행률 업데이트 함수
            def update_progress(value):
                progress.setValue(value)
                QApplication.processEvents()
                return not progress.wasCanceled()
            
            # 임시 폴더 생성 및 PDF 다운로드
            temp_dir = tempfile.mkdtemp()
            paths = []
            
            # 다운로드 진행 상태 표시
            progress.setLabelText("PDF 파일 다운로드 중...")
            for i, pdf in enumerate(pdfs):
                progress.setValue(int(i / len(pdfs) * 20))
                QApplication.processEvents()
                if progress.wasCanceled():
                    return False
                    
                local = os.path.join(temp_dir, pdf)
                download_file(f"입찰 2025/{folder}/{pdf}", local)
                paths.append(local)
            
            # 목차 가이드 생성 프롬프트
            prompt = TocGuideGenerator.build_prompt(paths)
            
            # 분석 진행 상태 표시
            progress.setLabelText("PDF 내용 분석 중...")
            progress.setValue(20)
            QApplication.processEvents()
                
            try:
                # GPT API 호출하여 목차 가이드 생성
                guide_data = analyze_pdfs(paths, prompt)
                
                # 결과를 JSON으로 저장
                guide_path = os.path.join(temp_dir, "목차가이드.json")
                with open(guide_path, "w", encoding="utf-8") as f:
                    json.dump(guide_data, f, ensure_ascii=False, indent=2)
                
                # Dropbox에 업로드
                upload_json(f"입찰 2025/{folder}/목차가이드.json", guide_data)
                
                # 완료
                progress.setValue(100)
                QMessageBox.information(parent, "완료", f"{folder} 목차 가이드가 생성되었습니다.")
                return True
                
            except Exception as e:
                QMessageBox.critical(parent, "분석 오류", f"목차 가이드 생성 중 오류 발생:\n{str(e)}")
                return False
                
        except Exception as e:
            QMessageBox.critical(parent, "오류", f"목차 가이드 생성 중 오류 발생:\n{str(e)}")
            return False 

    @staticmethod
    def build_prompt(pdf_files):
        docs_list = '\n'.join(pdf_files)
        # 최종 조정된 한국어 프롬프트
        return f"""
당신은 입찰 제안서 작성 전문가 어시스턴트입니다. 현재 작업 디렉터리에 있는 모든 PDF 파일은 하나의 입찰 제안서를 작성하기 위한 안내 문서입니다. 이 문서들 안에 제안서 작성요령 항목이 있습니다. 이 문서를 종합 분석하여, 이 입찰이 요구하는 제안서 목차와 작성 가이드라인을 자동 추출하세요:
{docs_list}

각 PDF 문서에 대해:
1. **목차(Table of Contents) 추출**
   - 제안서 구조를 정의하는 주요 장(chapter) 제목과 모든 소항목(sub-section)을 문서에 나타난 그대로 캡처합니다.
   - 장 및 소항목에 해당하는 페이지 번호를 정확히 기록합니다.
2. **작성 가이드라인 추출**
   - "작성 방법", "제안서 구성", "작성 가이드" 등 제목의 섹션이나 단락을 찾아 원문 그대로 캡처합니다.
   - 이 가이드라인은 제안서 각 섹션을 어떻게 작성해야 하는지 지시하는 내용입니다.
3. **메타데이터 기록**
   - `documents_processed`: 처리된 모든 PDF 파일명을 배열로 나열합니다.
   - `source_references`: 추출된 각 장, 소항목, 가이드라인마다 아래 형식으로 기록합니다:
     {{
       "document": "<파일명>",
       "page": <페이지번호>,
       "location": "<heading table|paragraph|ocr_table>",
       "description": "추출 항목에 대한 간단 설명"
     }}
4. 필요 시 OCR을 사용해 스캔된 표나 이미지로 된 목차를 처리합니다.
5. **최종 출력은 순수 JSON**이어야 하며, 아래 스키마를 엄격히 준수해야 합니다:
```json
{{
  "documents_processed": ["file1.pdf", "file2.pdf"],
  "source_references": [...],
  "table_of_contents": [
    {{
      "title": "Ⅰ. 사업 개요",
      "page": 2,
      "subsections": [
        {{"title": "1. 사업개요", "page": 2}},
        {{"title": "2. 사업목적", "page": 3}}
      ]
    }}
    // 추가 장
  ],
  "writing_guidelines": [
    {{"text": "제안서 구성은 다음과 같이 작성해야 합니다...", "source": {{"document":"file.pdf","page":5}}}}
    // 추가 가이드라인
  ]
}}
```
- `documents_processed`, `source_references`, `table_of_contents`, `writing_guidelines` 네 가지 최상위 키가 반드시 포함되어야 합니다.
- 파일명은 glob을 사용해 동적으로 검색하며, 하드코딩하지 마세요.
- 각 소항목별 페이지 정보를 정확히 기록해야 합니다.
- JSON에 추가 필드나 주석을 포함하지 마십시오.
"""

def extract_toc_and_guidelines():
    pdf_files = glob.glob('*.pdf')
    if not pdf_files:
        raise FileNotFoundError("No PDF files found in current directory.")

    prompt = TocGuideGenerator.build_prompt(pdf_files)
    messages = [
        {"role": "system", "content": "You are ChatGPT, a large language model."},
        {"role": "user", "content": prompt}
    ]

    response = openai.ChatCompletion.create(
        model=MODEL,
        messages=messages,
        temperature=0
    )

    content = response.choices[0].message.content.strip()
    try:
        result = json.loads(content)
    except json.JSONDecodeError:
        raise ValueError(f"Invalid JSON returned: {content}")
    print(json.dumps(result, ensure_ascii=False, indent=2))

if __name__ == '__main__':
    extract_toc_and_guidelines() 