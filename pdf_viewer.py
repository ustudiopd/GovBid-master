import sys
import fitz  # PyMuPDF
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, 
    QHBoxLayout, QPushButton, QFileDialog, QLabel,
    QScrollArea, QFrame, QListWidget, QListWidgetItem, QSizePolicy, QSplitter, QToolTip, QTextEdit, QLineEdit, QCheckBox, QComboBox, QTextBrowser, QSlider
)
from PyQt5.QtGui import QPixmap, QImage, QIcon, QFont, QColor
from PyQt5.QtCore import Qt, QSize, QPoint
import requests
import os
from dotenv import load_dotenv

# .env 파일 로드
load_dotenv()

class PDFViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF 뷰어")
        self.setGeometry(100, 100, 1400, 900)
        
        # GPT API KEY 및 모델
        self.gpt_api_key = os.getenv("CHATGPT_API_KEY", "")
        self.gpt_model = os.getenv("CHATGPT_MODEL", "gpt-4.1-mini")
        self.gpt_model_list = ["gpt-4.1-mini", "gpt-o4-mini"]
        
        # 메인 위젯 설정
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)
        main_layout.setSpacing(0)  # 간격 최소화
        main_layout.setContentsMargins(0, 0, 0, 0)
        
        # QSplitter로 썸네일 / PDF 본문 / 챗봇 영역 분리
        self.splitter = QSplitter(Qt.Horizontal)
        main_layout.addWidget(self.splitter)
        
        # ── 1) 썸네일 패널 추가 ─────────────────────────
        self.thumbnail_widget = QWidget()
        thumbnail_layout = QVBoxLayout(self.thumbnail_widget)
        thumbnail_layout.setContentsMargins(0, 0, 0, 0)
        thumbnail_layout.setSpacing(0)
        
        # 썸네일 목록
        self.thumbnail_list = QListWidget()
        self.thumbnail_list.setMaximumWidth(250)
        self.thumbnail_list.setSpacing(4)
        self.thumbnail_list.itemClicked.connect(self.thumbnail_clicked)
        self.thumbnail_list.setStyleSheet("QListWidget{border:none; background:#f8f8f8;} QListWidget::item{margin:0;}")
        thumbnail_layout.addWidget(self.thumbnail_list)
        
        # 썸네일 영역 고정 너비 설정으로 빈 공간 제거
        self.thumbnail_widget.setMaximumWidth(250)
        self.thumbnail_widget.setMinimumWidth(250)
        self.thumbnail_widget.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Expanding)
        self.splitter.addWidget(self.thumbnail_widget)
        
        # ── 2) 중앙: PDF 본문 영역 ─────────────────────────
        pdf_widget = QWidget()
        pdf_layout = QVBoxLayout(pdf_widget)
        pdf_layout.setContentsMargins(0, 0, 0, 0)
        pdf_layout.setSpacing(0)
        
        # 확대/축소 관련 변수
        self.zoom_percent = 100
        self.fit_to_width = False

        # 상단 버튼 레이아웃 (파일열기/이전/다음/페이지정보)
        button_layout = QHBoxLayout()
        button_layout.setContentsMargins(0, 0, 0, 0)
        button_layout.setSpacing(4)
        open_button = QPushButton("PDF 파일 열기")
        open_button.clicked.connect(self.open_pdf)
        button_layout.addWidget(open_button)
        self.prev_button = QPushButton("이전 페이지")
        self.prev_button.clicked.connect(self.prev_page)
        self.prev_button.setEnabled(False)
        button_layout.addWidget(self.prev_button)
        self.next_button = QPushButton("다음 페이지")
        self.next_button.clicked.connect(self.next_page)
        self.next_button.setEnabled(False)
        button_layout.addWidget(self.next_button)
        self.page_label = QLabel("페이지: 0/0")
        button_layout.addWidget(self.page_label)

        # --- 확대/축소 슬라이더 및 화면맞춤 버튼 추가 ---
        button_layout.addSpacing(20)
        self.zoom_slider = QSlider(Qt.Horizontal)
        self.zoom_slider.setMinimum(50)
        self.zoom_slider.setMaximum(300)
        self.zoom_slider.setValue(100)
        self.zoom_slider.setTickInterval(10)
        self.zoom_slider.setSingleStep(10)
        self.zoom_slider.setFixedWidth(120)
        self.zoom_slider.valueChanged.connect(self.on_zoom_slider)
        button_layout.addWidget(QLabel("확대/축소:"))
        button_layout.addWidget(self.zoom_slider)
        self.zoom_label = QLabel("100%")
        button_layout.addWidget(self.zoom_label)
        self.fit_btn = QPushButton("화면에 맞춤")
        self.fit_btn.clicked.connect(self.on_fit_to_width)
        button_layout.addWidget(self.fit_btn)

        # PDF 본문용 스크롤 영역
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        pdf_layout.addLayout(button_layout)
        pdf_layout.addWidget(scroll_area)
        # PDF 표시 레이블
        self.pdf_label = QLabel()
        self.pdf_label.setAlignment(Qt.AlignCenter)
        scroll_area.setWidget(self.pdf_label)
        # 스플리터에 중앙 패널 추가
        self.splitter.addWidget(pdf_widget)
        
        # ── 3) 오른쪽: 챗봇 메시지창 (세로 배치, 구분선 포함) ─────────
        chat_frame = QFrame()
        chat_frame.setFrameShape(QFrame.StyledPanel)
        chat_layout = QVBoxLayout(chat_frame)
        chat_layout.setContentsMargins(8, 8, 8, 8)
        chat_layout.setSpacing(8)

        # 1) 출력창
        self.chat_output = QTextBrowser()
        self.chat_output.setOpenExternalLinks(True)
        self.chat_output.setMinimumHeight(180)
        chat_layout.addWidget(self.chat_output)

        # 2) 구분선
        line1 = QFrame()
        line1.setFrameShape(QFrame.HLine)
        line1.setFrameShadow(QFrame.Sunken)
        chat_layout.addWidget(line1)

        # 3) 토글(이 페이지만), 모델명 표시
        option_layout = QHBoxLayout()
        self.page_only_checkbox = QCheckBox("이 페이지만")
        self.page_only_checkbox.setChecked(True)
        option_layout.addWidget(self.page_only_checkbox)
        # 모델명 라벨
        self.model_label = QLabel(f"모델: {self.gpt_model}")
        self.model_label.setStyleSheet("font-weight:bold;color:#0057b8;")
        option_layout.addWidget(self.model_label)
        option_layout.addStretch(1)
        option_widget = QWidget()
        option_widget.setLayout(option_layout)
        chat_layout.addWidget(option_widget)

        # 4) 구분선
        line2 = QFrame()
        line2.setFrameShape(QFrame.HLine)
        line2.setFrameShadow(QFrame.Sunken)
        chat_layout.addWidget(line2)

        # 5) 질문 입력창 (크게)
        self.chat_input = QTextEdit()
        self.chat_input.setPlaceholderText("질문을 입력하세요...")
        self.chat_input.setMinimumHeight(40)  # 1/3로 줄임
        self.chat_input.setStyleSheet("border:2px solid #000;")
        chat_layout.addWidget(self.chat_input)

        # 6) 구분선
        line3 = QFrame()
        line3.setFrameShape(QFrame.HLine)
        line3.setFrameShadow(QFrame.Sunken)
        chat_layout.addWidget(line3)

        # 7) 질문하기 버튼 (가운데 정렬)
        btn_layout = QHBoxLayout()
        btn_layout.addStretch(1)
        self.send_btn = QPushButton("질문하기")
        self.send_btn.clicked.connect(self.ask_gpt)
        btn_layout.addWidget(self.send_btn)
        btn_layout.addStretch(1)
        btn_widget = QWidget()
        btn_widget.setLayout(btn_layout)
        chat_layout.addWidget(btn_widget)

        self.splitter.addWidget(chat_frame)
        self.splitter.setSizes([250, 800, 350])
        
        # PDF 관련 변수 초기화
        self.current_doc = None
        self.current_page = 0
        self.total_pages = 0
    
    def show_splitter_tooltip(self, pos, index):
        # 썸네일 패널의 현재 width를 툴팁으로 표시
        width = self.thumbnail_widget.width()
        global_pos = self.mapToGlobal(QPoint(self.thumbnail_widget.width()//2, 0))
        QToolTip.showText(global_pos, f"{width}px", self.thumbnail_widget)
    
    def open_pdf(self):
        """PDF 파일 열기"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "PDF 파일 선택", "", "PDF Files (*.pdf)"
        )
        
        if file_path:
            try:
                # 이전 문서가 있으면 닫기
                if self.current_doc:
                    self.current_doc.close()
                
                # 새 문서 열기
                self.current_doc = fitz.open(file_path)
                self.total_pages = len(self.current_doc)
                self.current_page = 0
                
                # 썸네일 생성
                self.create_thumbnails()
                
                # 페이지 표시
                self.display_page()
                
                # 버튼 상태 업데이트
                self.update_buttons()
                
            except Exception as e:
                self.pdf_label.setText(f"PDF 파일을 열 수 없습니다: {str(e)}")
    
    def create_thumbnails(self):
        """썸네일 목록 생성"""
        self.thumbnail_list.clear()
        thumb_w = 200
        for page_num in range(self.total_pages):
            try:
                # PDF 페이지에서 이미지 생성 (확대 비율 조절 가능)
                page = self.current_doc[page_num]
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
                pixmap = QPixmap.fromImage(img)

                # ▶ 컨테이너 위젯에 썸네일+번호 함께 담기
                item_widget = QWidget()
                item_layout = QVBoxLayout(item_widget)
                item_layout.setContentsMargins(2, 2, 2, 2)
                item_layout.setSpacing(4)

                # 1) 썸네일 라벨: 200px로 고정
                thumb_label = QLabel()
                thumb_label.setAlignment(Qt.AlignCenter)
                thumb_pix = pixmap.scaledToWidth(thumb_w, Qt.SmoothTransformation)
                thumb_label.setPixmap(thumb_pix)
                thumb_label.setFixedWidth(thumb_w)
                item_layout.addWidget(thumb_label)

                # 2) 페이지 번호 라벨: 썸네일 바로 아래 중앙 정렬
                page_label = QLabel(str(page_num + 1))
                page_label.setAlignment(Qt.AlignCenter)
                font = QFont()
                font.setPointSize(8)
                page_label.setFont(font)
                page_label.setStyleSheet("color: #555;")
                item_layout.addWidget(page_label)

                # ▶ QListWidgetItem 생성 및 위젯 바인딩
                item = QListWidgetItem()
                item.setSizeHint(item_widget.sizeHint())
                self.thumbnail_list.addItem(item)
                self.thumbnail_list.setItemWidget(item, item_widget)

            except Exception as e:
                print(f"썸네일 생성 오류 (페이지 {page_num + 1}): {str(e)}")
    
    def thumbnail_clicked(self, item):
        """썸네일 클릭 시 해당 페이지로 이동"""
        page_num = self.thumbnail_list.row(item)
        if 0 <= page_num < self.total_pages:
            self.current_page = page_num
            self.display_page()
            self.update_buttons()
    
    def display_page(self):
        """현재 페이지 표시"""
        if not self.current_doc:
            return
        try:
            page = self.current_doc[self.current_page]
            # 확대/축소 적용
            if self.fit_to_width:
                # 뷰어 영역의 가로 크기에 맞춤
                scroll_area = self.pdf_label.parentWidget().parentWidget()
                area_width = scroll_area.viewport().width()
                # PDF 원본 크기 기준(2배 확대 기준 2*612=1224)
                base_width = 612
                zoom = area_width / base_width
                matrix = fitz.Matrix(zoom, zoom)
                self.zoom_percent = int(zoom * 100)
                self.zoom_slider.setValue(self.zoom_percent)
                self.zoom_label.setText(f"{self.zoom_percent}%")
            else:
                zoom = self.zoom_percent / 100
                matrix = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=matrix)
            img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
            pixmap = QPixmap.fromImage(img)
            self.pdf_label.setPixmap(pixmap)
            self.page_label.setText(f"페이지: {self.current_page + 1}/{self.total_pages}")
            self.thumbnail_list.setCurrentRow(self.current_page)
        except Exception as e:
            self.pdf_label.setText(f"페이지를 표시할 수 없습니다: {str(e)}")
    
    def prev_page(self):
        """이전 페이지로 이동"""
        if self.current_page > 0:
            self.current_page -= 1
            self.display_page()
            self.update_buttons()
    
    def next_page(self):
        """다음 페이지로 이동"""
        if self.current_page < self.total_pages - 1:
            self.current_page += 1
            self.display_page()
            self.update_buttons()
    
    def update_buttons(self):
        """페이지 이동 버튼 상태 업데이트"""
        self.prev_button.setEnabled(self.current_page > 0)
        self.next_button.setEnabled(self.current_page < self.total_pages - 1)
    
    def closeEvent(self, event):
        """프로그램 종료 시 문서 닫기"""
        if self.current_doc:
            self.current_doc.close()
        event.accept()

    def ask_gpt(self):
        question = self.chat_input.toPlainText().strip()
        if not question:
            return
        self.chat_output.append(f"<b>질문:</b> {question}")
        QApplication.processEvents()
        # PDF 텍스트 추출
        if self.page_only_checkbox.isChecked():
            context = self.extract_page_text(self.current_page)
        else:
            context = self.extract_all_text()
        # GPT 호출 (env의 모델만 사용)
        answer = ask_gpt_api(question, context, self.gpt_api_key, self.gpt_model)
        self.chat_output.append("<b>GPT:</b>")
        self.chat_output.setMarkdown(answer)
        self.chat_input.clear()

    def extract_page_text(self, page_num):
        if self.current_doc is None:
            return ""
        try:
            page = self.current_doc[page_num]
            return page.get_text()
        except Exception:
            return ""

    def extract_all_text(self):
        if self.current_doc is None:
            return ""
        texts = []
        for i in range(self.total_pages):
            try:
                texts.append(self.current_doc[i].get_text())
            except Exception:
                continue
        return "\n".join(texts)

    def on_zoom_slider(self, value):
        self.zoom_percent = value
        self.fit_to_width = False
        self.zoom_label.setText(f"{value}%")
        self.display_page()

    def on_fit_to_width(self):
        self.fit_to_width = True
        self.display_page()

# --- GPT API 호출 함수 (API키, 모델 인자로 받음) ---
def ask_gpt_api(question, context, api_key, model):
    if not api_key:
        return "[OpenAI API 키를 .env에 입력하세요]"
    url = "https://api.openai.com/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    data = {
        "model": model,
        "messages": [
            {"role": "system", "content": "아래 context를 참고해서 사용자의 질문에 답변해줘."},
            {"role": "user", "content": f"context: {context}\n\n질문: {question}"}
        ],
        "max_tokens": 2048,
        "temperature": 0.7
    }
    try:
        resp = requests.post(url, headers=headers, json=data, timeout=30)
        resp.raise_for_status()
        return resp.json()["choices"][0]["message"]["content"].strip()
    except Exception as e:
        return f"[GPT 호출 오류] {e}"

if __name__ == "__main__":
    app = QApplication(sys.argv)
    viewer = PDFViewer()
    viewer.show()
    sys.exit(app.exec_()) 