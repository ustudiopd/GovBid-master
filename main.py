# main.py
import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QTableWidget, QTableWidgetItem, QPushButton, QMessageBox,
    QHeaderView, QToolTip, QFileDialog, QHBoxLayout
)
from PyQt5.QtGui import QCursor
from PyQt5.QtCore import Qt
from dropbox_client import list_folder, download_json
from detail_dialog import DetailDialog
from analyzer import Analyzer
from typing import List, Dict, Any
import glob
import openai
import json

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("입찰 2025 폴더 리스트")
        # 윈도우 가로 1500px로 초기 크기 설정
        self.resize(1500, 600)
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout()
        central.setLayout(layout)

        # 상단 버튼 레이아웃
        top_layout = QHBoxLayout()
        
        self.load_button = QPushButton("데이터 로드")
        self.load_button.clicked.connect(self.load_data)
        top_layout.addWidget(self.load_button)
        
        # 전체화면 전환 버튼 추가
        self.fullscreen_button = QPushButton("전체화면")
        self.fullscreen_button.clicked.connect(self.toggle_fullscreen)
        top_layout.addWidget(self.fullscreen_button)
        
        # 상단 레이아웃을 메인 레이아웃에 추가
        layout.addLayout(top_layout)

        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels([
            "No", "등록마감", "공고명",
            "추정가격", "PDF", "분석"
        ])
        # 인터랙티브 모드: 사용자가 마우스로 열 너비 조절 가능
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.sectionResized.connect(self.show_section_width)
        layout.addWidget(self.table)
        # 초기 열 너비 설정 (전체 가로 1250px 기준)
        self.table.setColumnWidth(0, 30)
        self.table.setColumnWidth(1, 170)
        self.table.setColumnWidth(2, 900)
        self.table.setColumnWidth(3, 200)
        self.table.setColumnWidth(4, 45)
        self.table.setColumnWidth(5, 45)
        # 공고명 클릭 시 analysis.json 상세 보기
        self.table.cellClicked.connect(self.on_cell_clicked)

    def load_data(self):
        try:
            folders = list_folder("입찰 2025")
        except Exception as e:
            QMessageBox.warning(self, "Dropbox 에러", f"폴더 리스트를 가져오는 중 오류 발생:\n{e}")
            folders = []

        try:
            data = download_json("입찰 2025/smpp.json")
        except Exception as e:
            QMessageBox.critical(self, "Dropbox JSON 에러", f"JSON 다운로드 중 오류 발생:\n{e}")
            return

        entries = [
            item for item in data
            if item.get("folder_name") in folders
        ]
        # 클릭 이벤트에서 참조할 용도
        self.entries = entries

        self.table.setRowCount(len(entries) * 2)
        for idx, item in enumerate(entries):
            info = item.get("announcement_info", {})
            # PDF 및 analysis_status 매핑
            pdf_text = "Yes" if item.get("has_pdfs") else "No"
            status = item.get("analysis_status", "")
            if status == "completed":
                analysis_text = "분석완료"
            elif status == "pending":
                analysis_text = "분석가능"
            else:
                analysis_text = status or ""
            # 첫 줄: 기본 정보
            basic = [
                str(item.get("no", "")),
                info.get("등록마감", ""),
                info.get("공고명", ""),
                info.get("추정가격", ""),
                pdf_text,
                analysis_text,
            ]
            row0 = idx * 2
            # 기본 정보 및 버튼 생성
            self.table.setItem(row0, 0, QTableWidgetItem(basic[0]))
            self.table.setItem(row0, 1, QTableWidgetItem(basic[1]))
            title_btn = QPushButton(basic[2])
            title_btn.setEnabled(status == "completed")
            title_btn.clicked.connect(lambda _, i=idx: self.show_analysis_detail(i))
            self.table.setCellWidget(row0, 2, title_btn)
            self.table.setItem(row0, 3, QTableWidgetItem(basic[3]))
            self.table.setItem(row0, 4, QTableWidgetItem(basic[4]))
            status_btn = QPushButton(basic[5])
            if status == "pending":
                status_btn.clicked.connect(lambda _, i=idx: self.start_analysis(i))
            else:
                status_btn.clicked.connect(lambda _, i=idx: self.show_analysis_detail(i))
            self.table.setCellWidget(row0, 5, status_btn)
            # 두번째 줄: 입찰내용 요약 (컬럼 헤더 없음, 전체 열 span)
            summary = info.get("입찰내용 요약", "")
            row1 = row0 + 1
            self.table.setSpan(row1, 0, 1, 6)
            self.table.setItem(row1, 0, QTableWidgetItem(summary))
            # 두 번째 줄(요약) 행 높이를 기본 높이의 2배로 설정
            default_h = self.table.verticalHeader().defaultSectionSize()
            self.table.setRowHeight(row1, default_h * 2)

        # 테이블 내용에 맞춰 자동 조정하지 않고 초기 열 너비 유지
        self.table.setColumnWidth(0, 30)
        self.table.setColumnWidth(1, 170)
        self.table.setColumnWidth(2, 900)
        self.table.setColumnWidth(3, 200)
        self.table.setColumnWidth(4, 45)
        self.table.setColumnWidth(5, 45)

    def show_section_width(self, index, old_size, new_size):
        # 열 크기 변경 시 마우스 위치에 픽셀 크기 툴팁 표시
        QToolTip.showText(QCursor.pos(), f"{new_size}px")

    def on_cell_clicked(self, row, col):
        # 공고명(컬럼 인덱스 2)의 첫 번째 줄 클릭 시 분석 상세 열기
        if col == 2 and row % 2 == 0:
            idx = row // 2
            entry = getattr(self, 'entries', [])[idx]
            # completed 상태만 상세 보기
            if entry.get("analysis_status") != "completed":
                return
            self.show_analysis_detail(idx)

    def start_analysis(self, idx):
        """PDF 파일 분석 시작"""
        entry = self.entries[idx]
        folder = entry.get("folder_name")
        
        # Analyzer 클래스를 사용하여 분석 수행
        if Analyzer.analyze_folder(folder, self):
            # 분석 성공 시 데이터 다시 로드
            self.load_data()

    def show_analysis_detail(self, idx):
        """세부 대화상자 표시"""
        entry = self.entries[idx]
        folder = entry.get("folder_name")
        
        # DetailDialog 인스턴스 생성 및 표시
        detail_dialog = DetailDialog(self, entry, folder)
        detail_dialog.exec_()

    def find_dropbox_folder(self):
        """로컬 Dropbox 폴더 위치 탐색"""
        # 일반적인 Dropbox 설치 경로
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
        
        # 특정 경로 확인
        for path in potential_paths:
            if os.path.exists(path) and os.path.isdir(path):
                print(f"Dropbox 폴더 발견: {path}")
                
                # 입찰 폴더 확인
                bid_folder = os.path.join(path, "입찰 2025")
                if os.path.exists(bid_folder):
                    return path
        
        # 폴더가 없으면 사용자에게 물어보기
        folder = QFileDialog.getExistingDirectory(
            self, "Dropbox 폴더 선택", "", 
            QFileDialog.ShowDirsOnly | QFileDialog.DontResolveSymlinks
        )
        
        if folder and os.path.isdir(folder):
            return folder
            
        return None

    def toggle_fullscreen(self):
        """전체화면 전환"""
        if self.isFullScreen():
            self.showNormal()
            self.fullscreen_button.setText("전체화면")
        else:
            self.showFullScreen()
            self.fullscreen_button.setText("일반화면")
            
    def keyPressEvent(self, event):
        """키 이벤트 처리"""
        if event.key() == Qt.Key_F11:
            self.toggle_fullscreen()
        else:
            super().keyPressEvent(event)

def build_prompt(pdf_files):
    docs_list = '\n'.join(pdf_files)
    return f"""
You are an expert proposal-writing assistant.
Your task is to analyze all the following PDF files in the current working directory:
{docs_list}

For each PDF:
1. Extract the Table of Contents: detect main chapter titles and their exact sub-sections, capturing for each sub-section its page number as well.
2. Extract Writing Guidelines: any prose instructions on proposal structure; capture them verbatim and record the source reference for each guideline.
3. Use OCR if needed to read any scanned tables or images containing headings or ToC.
4. Record metadata:
   - documents_processed: list of all PDF filenames
   - source_references: for every extracted chapter, sub-section, or guideline, record {{
       document: <filename>,
       page: <page#>,
       location: <"heading table" | "paragraph" | "ocr_table">,
       description: <short desc>
     }}

5. Output must be pure JSON matching this schema exactly:

```json
{{
  "documents_processed": ["..."],
  "source_references": [
    {{"document": "...", "page": 6, "location": "heading table", "description": "제안서 작성방법 – 제안서 구성 표"}},
    ...
  ],
  "table_of_contents": [
    {{
      "title": "Ⅰ. 사업 개요",
      "page": 2,
      "subsections": [
        {{"title": "1. 사업개요", "page": 2}},
        {{"title": "2. 사업목적", "page": 3}},
        ...
      ]
    }},
    ...
  ],
  "writing_guidelines": [
    {{"text": "제안 업체는 상기 제안 내용에 대한 내용을 명확하고 간결하게 작성하여야 한다.", "source": {"document": "...", "page": 7}}},
    ...
  ]
}}
```

- documents_processed: all PDF filenames.
- source_references: include document, page, location, description for each item.
- table_of_contents: each chapter with subsections including their titles and page numbers.
- writing_guidelines: each guideline text plus source reference.

Do not hard-code filenames; dynamically discover PDFs with glob. Ensure JSON is valid with no extra fields or comments.
"""

def extract_toc_and_guidelines():
    pdf_files = glob.glob('*.pdf')
    if not pdf_files:
        raise FileNotFoundError("No PDF files found in current directory.")

    prompt = build_prompt(pdf_files)
    messages = [
        {"role": "system", "content": "You are ChatGPT, a large language model."},
        {"role": "user", "content": prompt}
    ]

    response = openai.ChatCompletion.create(
        model="gpt-4.1-mini",
        messages=messages,
        temperature=0
    )

    content = response.choices[0].message.content.strip()
    try:
        result = json.loads(content)
    except json.JSONDecodeError:
        raise ValueError(f"Invalid JSON returned: {content}")
    # Validate expected keys
    for key in ["documents_processed", "source_references", "table_of_contents", "writing_guidelines"]:
        if key not in result:
            raise KeyError(f"Missing expected key in JSON: {key}")
    # Output formatted JSON
    print(json.dumps(result, ensure_ascii=False, indent=2))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())