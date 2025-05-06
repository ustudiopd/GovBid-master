import sys
import io
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QGraphicsView, QGraphicsScene, QGraphicsTextItem,
    QFileDialog, QAction, QToolBar, QComboBox, QLabel, QMessageBox, QProgressDialog,
    QInputDialog, QLineEdit
)
from PyQt5.QtGui import QPainter, QImage, QFont, QPixmap, QColor
from PyQt5.QtCore import Qt, QRectF, QObject, pyqtSignal
from PyPDF2 import PdfReader
import tempfile
import json
import openai
import shutil
import re
import threading
from dotenv import load_dotenv
from openai import OpenAI

# PDF2Image 라이브러리 사용 (poppler 대체)
try:
    from pdf2image import convert_from_path
    PDF_RENDERER_AVAILABLE = True
except ImportError:
    PDF_RENDERER_AVAILABLE = False


class PdfFormEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF 폼 편집기")
        self.resize(800, 800)
        
        # 그래픽스 뷰 및 씬 설정
        self.scene = QGraphicsScene()
        self.view = QGraphicsView(self.scene)
        self.view.setRenderHint(QPainter.Antialiasing)
        self.view.setRenderHint(QPainter.TextAntialiasing)
        self.view.setRenderHint(QPainter.SmoothPixmapTransform)
        self.setCentralWidget(self.view)
        
        # PDF 관련 변수 초기화
        self.pdf_path = None
        self.pdf_pages = []
        self.page_index = 0
        self.page_images = []
        self.pdf_dpi = 300  # 기본 PDF 렌더링 해상도 (DPI)
        
        # UI 초기화
        self.init_ui()
        
        # 상태 표시줄 설정
        self.statusBar().showMessage("PDF 파일을 여십시오")

    def init_ui(self):
        # 툴바 생성
        tb = QToolBar("메인 툴바")
        self.addToolBar(tb)
        
        # 툴바 액션 추가
        actions = [
            ("PDF 열기", self.open_pdf, "PDF 파일을 엽니다"),
            ("이전 페이지", self.prev_page, "이전 페이지로 이동합니다"),
            ("다음 페이지", self.next_page, "다음 페이지로 이동합니다"),
            ("텍스트 상자 추가", self.add_textbox, "편집 가능한 텍스트 상자를 추가합니다"),
            ("PNG로 저장", self.save_png, "현재 페이지를 PNG 이미지로 저장합니다"),
            ("서식파일 내보내기", self.extract_forms, "ChatGPT를 이용해 서식 페이지를 찾아 저장합니다")
        ]
        
        for name, slot, tooltip in actions:
            act = QAction(name, self)
            act.setToolTip(tooltip)
            act.triggered.connect(slot)
            tb.addAction(act)
        
        # 툴바에 구분선 추가
        tb.addSeparator()
        
        # 폰트 크기 콤보박스 추가
        tb.addWidget(QLabel("폰트 크기:"))
        self.font_size = QComboBox()
        for sz in [8, 10, 12, 14, 16, 18, 20, 24, 28, 32, 36]:
            self.font_size.addItem(str(sz))
        self.font_size.setCurrentIndex(3)  # 기본값 14 설정
        tb.addWidget(self.font_size)
        
        # 페이지 정보 표시
        tb.addSeparator()
        self.page_info = QLabel("페이지: 0/0")
        tb.addWidget(self.page_info)

    def open_pdf(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "PDF 파일 열기", "", "PDF 파일 (*.pdf)"
        )
        if not path:
            return
            
        try:
            # PDF 파일 열기
            self.pdf_path = path
            reader = PdfReader(path)
            self.pdf_pages = reader.pages
            self.page_images = []
            
            # 페이지 이미지 초기화
            for _ in range(len(self.pdf_pages)):
                self.page_images.append(None)
                
            # pdf2image 라이브러리가 없을 경우 경고
            if not PDF_RENDERER_AVAILABLE:
                QMessageBox.warning(self, "알림", 
                                  "pdf2image 라이브러리가 설치되지 않았습니다.\n"
                                  "PDF 렌더링이 제한됩니다.\n\n"
                                  "pip install pdf2image 명령으로 설치하세요.")
                
            self.page_index = 0
            self.render_page()
            self.statusBar().showMessage(f"PDF 파일 열림: {path}")
            
            # 콜백이 설정되어 있다면 호출
            if hasattr(self, 'on_pdf_loaded_callback') and self.on_pdf_loaded_callback:
                self.on_pdf_loaded_callback(path, len(self.pdf_pages))
                
        except Exception as e:
            QMessageBox.critical(self, "오류", f"PDF 파일을 여는 중 오류 발생: {str(e)}")

    def render_page(self):
        if not self.pdf_pages:
            return
            
        # 씬 초기화
        self.scene.clear()
        
        # 페이지 렌더링
        if PDF_RENDERER_AVAILABLE and not self.page_images[self.page_index]:
            try:
                # pdf2image로 현재 페이지만 렌더링
                images = convert_from_path(
                    self.pdf_path,
                    first_page=self.page_index + 1,
                    last_page=self.page_index + 1,
                    dpi=self.pdf_dpi  # 높은 해상도 (300 DPI)
                )
                if images:
                    # PIL Image를 QImage로 변환
                    pil_image = images[0]
                    buffer = io.BytesIO()
                    pil_image.save(buffer, format='PNG')
                    q_img = QImage()
                    q_img.loadFromData(buffer.getvalue())
                    self.page_images[self.page_index] = q_img
            except Exception as e:
                # 렌더링 실패 시 빈 이미지 생성
                print(f"PDF 렌더링 오류: {e}")
                image = QImage(600, 800, QImage.Format_RGB32)
                image.fill(Qt.white)
                self.page_images[self.page_index] = image
        elif not self.page_images[self.page_index]:
            # PDF2Image가 없거나 변환 실패한 경우 빈 이미지 생성
            image = QImage(600, 800, QImage.Format_RGB32)
            image.fill(Qt.white)
            self.page_images[self.page_index] = image
        
        # 이미지를 QGraphicsScene에 추가
        pixmap = QPixmap.fromImage(self.page_images[self.page_index])
        self.scene.addPixmap(pixmap)
        
        # PDF2Image가 없는 경우 경고 텍스트 표시
        if not PDF_RENDERER_AVAILABLE:
            text_item = QGraphicsTextItem(f"PDF 페이지 {self.page_index + 1}\n\n(더 나은 렌더링을 위해 pdf2image 설치 필요)")
            text_item.setPos(50, 50)
            self.scene.addItem(text_item)
        
        # 뷰 크기 조정
        self.view.setSceneRect(QRectF(0, 0, self.page_images[self.page_index].width(), 
                                       self.page_images[self.page_index].height()))
        self.view.fitInView(self.scene.sceneRect(), Qt.KeepAspectRatio)
        
        # 페이지 정보 업데이트
        total_pages = len(self.pdf_pages)
        self.page_info.setText(f"페이지: {self.page_index + 1}/{total_pages}")
        
        # 콜백이 설정되어 있다면 호출
        if hasattr(self, 'on_page_changed_callback') and self.on_page_changed_callback:
            self.on_page_changed_callback(self.page_index, total_pages)

    def prev_page(self):
        if self.pdf_pages and self.page_index > 0:
            self.page_index -= 1
            self.render_page()
            self.statusBar().showMessage(f"페이지 {self.page_index + 1} 로 이동")

    def next_page(self):
        if self.pdf_pages and self.page_index < len(self.pdf_pages) - 1:
            self.page_index += 1
            self.render_page()
            self.statusBar().showMessage(f"페이지 {self.page_index + 1} 로 이동")

    def add_textbox(self):
        if not self.pdf_pages:
            QMessageBox.warning(self, "경고", "먼저 PDF 파일을 열어주세요.")
            return
            
        # 텍스트 아이템 생성
        text_item = QGraphicsTextItem("더블클릭하여 편집")
        
        # 폰트 설정
        size = int(self.font_size.currentText())
        font = QFont(QApplication.font().family(), size)
        text_item.setFont(font)
        
        # 상호작용 설정
        text_item.setTextInteractionFlags(Qt.TextEditorInteraction)
        text_item.setFlag(QGraphicsTextItem.ItemIsMovable)
        text_item.setFlag(QGraphicsTextItem.ItemIsSelectable)
        
        # 장면에 추가
        self.scene.addItem(text_item)
        
        # 뷰포트 중앙에 배치
        viewport_rect = self.view.viewport().rect()
        viewport_center = self.view.mapToScene(
            viewport_rect.center()
        )
        text_item.setPos(viewport_center)
        
        self.statusBar().showMessage("텍스트 상자가 추가되었습니다. 더블클릭하여 편집하세요.")
        
        # 콜백이 설정되어 있다면 호출
        if hasattr(self, 'on_textbox_added_callback') and self.on_textbox_added_callback:
            self.on_textbox_added_callback(text_item)
            
        return text_item

    def save_png(self):
        if not self.pdf_pages:
            QMessageBox.warning(self, "경고", "먼저 PDF 파일을 열어주세요.")
            return
            
        file_name, _ = QFileDialog.getSaveFileName(
            self, "PNG로 저장", f"page_{self.page_index + 1}.png", "PNG 파일 (*.png)"
        )
        
        if not file_name:
            return
            
        try:
            # 씬 크기의 이미지 생성
            scene_rect = self.scene.sceneRect().toRect()
            image = QImage(scene_rect.size(), QImage.Format_ARGB32)
            image.fill(Qt.white)
            
            # 페인터로 씬 렌더링
            painter = QPainter(image)
            self.scene.render(painter)
            painter.end()
            
            # 이미지 저장
            image.save(file_name)
            self.statusBar().showMessage(f"이미지 저장됨: {file_name}")
            
            # 콜백이 설정되어 있다면 호출
            if hasattr(self, 'on_image_saved_callback') and self.on_image_saved_callback:
                self.on_image_saved_callback(file_name)
                
            return file_name
        except Exception as e:
            QMessageBox.critical(self, "오류", f"이미지 저장 중 오류 발생: {str(e)}")
            return None

    def resizeEvent(self, event):
        # 창 크기 변경 시 PDF 페이지 크기 조정
        if self.pdf_pages:
            self.view.fitInView(self.scene.sceneRect(), Qt.KeepAspectRatio)
        super().resizeEvent(event)
    
    # API 메서드 - 콜백 설정
    def set_on_pdf_loaded_callback(self, callback):
        """PDF 로드 완료 시 호출될 콜백 설정"""
        self.on_pdf_loaded_callback = callback
    
    def set_on_page_changed_callback(self, callback):
        """페이지 변경 시 호출될 콜백 설정"""
        self.on_page_changed_callback = callback
    
    def set_on_textbox_added_callback(self, callback):
        """텍스트 상자 추가 시 호출될 콜백 설정"""
        self.on_textbox_added_callback = callback
    
    def set_on_image_saved_callback(self, callback):
        """이미지 저장 시 호출될 콜백 설정"""
        self.on_image_saved_callback = callback
    
    # API 메서드 - 프로그래밍 방식으로 기능 호출
    def load_pdf(self, path):
        """프로그래밍 방식으로 PDF 파일 로드"""
        if not os.path.exists(path):
            raise FileNotFoundError(f"PDF 파일을 찾을 수 없습니다: {path}")
            
        try:
            self.pdf_path = path
            reader = PdfReader(path)
            self.pdf_pages = reader.pages
            self.page_images = []
            
            for _ in range(len(self.pdf_pages)):
                self.page_images.append(None)
                
            self.page_index = 0
            self.render_page()
            self.statusBar().showMessage(f"PDF 파일 열림: {path}")
            
            if hasattr(self, 'on_pdf_loaded_callback') and self.on_pdf_loaded_callback:
                self.on_pdf_loaded_callback(path, len(self.pdf_pages))
                
            return True
        except Exception as e:
            self.statusBar().showMessage(f"PDF 로드 오류: {str(e)}")
            return False
    
    def go_to_page(self, page_index):
        """특정 페이지로 이동"""
        if not self.pdf_pages:
            return False
            
        if 0 <= page_index < len(self.pdf_pages):
            self.page_index = page_index
            self.render_page()
            self.statusBar().showMessage(f"페이지 {self.page_index + 1} 로 이동")
            return True
        return False
    
    def get_current_page_index(self):
        """현재 페이지 인덱스 반환"""
        return self.page_index
    
    def get_total_pages(self):
        """총 페이지 수 반환"""
        return len(self.pdf_pages) if self.pdf_pages else 0
    
    def add_textbox_at(self, x, y, text="텍스트 입력", font_size=None):
        """지정된 위치에 텍스트 상자 추가"""
        if not self.pdf_pages:
            return None
            
        text_item = QGraphicsTextItem(text)
        
        # 폰트 설정
        if font_size is None:
            font_size = int(self.font_size.currentText())
        font = QFont(QApplication.font().family(), font_size)
        text_item.setFont(font)
        
        # 상호작용 설정
        text_item.setTextInteractionFlags(Qt.TextEditorInteraction)
        text_item.setFlag(QGraphicsTextItem.ItemIsMovable)
        text_item.setFlag(QGraphicsTextItem.ItemIsSelectable)
        
        # 장면에 추가
        self.scene.addItem(text_item)
        text_item.setPos(x, y)
        
        # 콜백 호출
        if hasattr(self, 'on_textbox_added_callback') and self.on_textbox_added_callback:
            self.on_textbox_added_callback(text_item)
            
        return text_item
    
    def export_current_page(self, output_path=None):
        """현재 페이지를 이미지로 내보내기"""
        if not self.pdf_pages:
            return None
            
        if output_path is None:
            # 임시 파일 생성
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
                output_path = tmp.name
        
        try:
            # 씬 크기의 이미지 생성
            scene_rect = self.scene.sceneRect().toRect()
            image = QImage(scene_rect.size(), QImage.Format_ARGB32)
            image.fill(Qt.white)
            
            # 페인터로 씬 렌더링
            painter = QPainter(image)
            self.scene.render(painter)
            painter.end()
            
            # 이미지 저장
            image.save(output_path)
            
            # 콜백 호출
            if hasattr(self, 'on_image_saved_callback') and self.on_image_saved_callback:
                self.on_image_saved_callback(output_path)
                
            return output_path
        except Exception as e:
            self.statusBar().showMessage(f"이미지 내보내기 오류: {str(e)}")
            return None

    def set_pdf_dpi(self, dpi):
        """PDF 렌더링 해상도(DPI) 설정
        
        매개변수:
            dpi (int): PDF 렌더링 해상도, 기본값 300
        """
        if isinstance(dpi, int) and dpi > 0:
            self.pdf_dpi = dpi
            # 이미지 캐시 초기화
            self.page_images = [None] * len(self.pdf_pages) if self.pdf_pages else []
            # 현재 페이지 다시 렌더링
            if self.pdf_pages:
                self.render_page()
            return True
        return False
    
    def get_pdf_dpi(self):
        """현재 PDF 렌더링 해상도(DPI) 반환"""
        return self.pdf_dpi

    def extract_forms(self):
        """ChatGPT를 이용해 서식 페이지를 찾아내고 개별 PDF로 저장합니다."""
        if not self.pdf_path:
            QMessageBox.warning(self, "경고", "먼저 PDF 파일을 열어주세요.")
            return

        # .env 파일에서 API 키 로드
        load_dotenv()
        api_key = os.getenv("CHATGPT_API_KEY")
        
        if not api_key:
            # API 키가 없으면 사용자에게 입력 요청
            api_key, ok = QInputDialog.getText(
                self, "OpenAI API 키", 
                ".env 파일에 CHATGPT_API_KEY가 없습니다.\nAPI 키를 입력하세요:",
                QLineEdit.Password
            )
            if not ok or not api_key:
                QMessageBox.warning(self, "경고", "API 키가 필요합니다.")
                return
        
        # 진행 대화상자 표시
        progress = QProgressDialog("서식 페이지 분석 중...", "취소", 0, 100, self)
        progress.setWindowTitle("분석 중")
        progress.setMinimumDuration(0)
        progress.setValue(0)
        progress.show()

        # 서식 분석 워커 시작
        self.form_extractor = FormExtractor(self.pdf_path, api_key)
        self.form_extractor.progress_updated.connect(progress.setValue)
        self.form_extractor.analysis_complete.connect(self._handle_form_analysis_result)
        self.form_extractor.start()

    def _handle_form_analysis_result(self, result):
        """서식 분석 결과 처리"""
        if not result or not result.get('forms'):
            QMessageBox.information(self, "알림", "서식 페이지를 찾을 수 없습니다.")
            return
        
        # 저장 폴더 선택
        save_dir = QFileDialog.getExistingDirectory(
            self, "서식 파일 저장 폴더 선택", os.path.dirname(self.pdf_path)
        )
        if not save_dir:
            return
        
        # 서식 폴더 생성
        forms_dir = os.path.join(save_dir, "서식")
        os.makedirs(forms_dir, exist_ok=True)
        
        # 결과 로그 파일 생성
        with open(os.path.join(save_dir, "서식분석결과.json"), "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        # 각 서식 페이지 추출 및 저장
        successful = 0
        for form in result.get('forms', []):
            page = form.get('page')
            if page is None:
                continue
                
            # 0부터 시작하는 페이지 인덱스로 조정
            page_idx = page - 1
            if page_idx < 0 or page_idx >= len(self.pdf_pages):
                continue
                
            # 파일명 생성 (특수문자 제거)
            filename = form.get('filename', f"{page}p_서식.pdf")
            filename = re.sub(r'[\\/*?:"<>|]', "", filename)
            output_path = os.path.join(forms_dir, filename)
            
            try:
                # PDF로 단일 페이지 저장
                self._extract_pdf_page(page_idx, output_path)
                successful += 1
            except Exception as e:
                print(f"페이지 {page} 저장 중 오류: {e}")
        
        # 완료 메시지
        QMessageBox.information(
            self, "완료", 
            f"서식 페이지 추출 완료: {successful}개 파일이 '{forms_dir}'에 저장되었습니다."
        )
    
    def _extract_pdf_page(self, page_idx, output_path):
        """단일 페이지 PDF로 저장"""
        from PyPDF2 import PdfWriter, PdfReader
        
        # 원본 PDF 열기
        reader = PdfReader(self.pdf_path)
        writer = PdfWriter()
        
        # 페이지 추가
        writer.add_page(reader.pages[page_idx])
        
        # 새 PDF로 저장
        with open(output_path, "wb") as out_file:
            writer.write(out_file)


class FormExtractor(QObject):
    """ChatGPT API를 이용한 서식 페이지 추출 워커"""
    progress_updated = pyqtSignal(int)
    analysis_complete = pyqtSignal(dict)
    
    def __init__(self, pdf_path, api_key):
        super().__init__()
        self.pdf_path = pdf_path
        self.api_key = api_key
        self.thread = None
        
    def start(self):
        """분석 스레드 시작"""
        self.thread = threading.Thread(target=self._run_analysis)
        self.thread.daemon = True
        self.thread.start()
        
    def _run_analysis(self):
        """PDF 분석 실행"""
        try:
            # 진행률 업데이트
            self.progress_updated.emit(10)
            
            # PDF 텍스트 추출
            reader = PdfReader(self.pdf_path)
            text_content = ""
            for i, page in enumerate(reader.pages):
                text_content += f"\n--- PAGE {i+1} ---\n"
                text_content += page.extract_text() or f"[Page {i+1} has no extractable text]"
                # 진행률 업데이트 (페이지당)
                self.progress_updated.emit(10 + (i * 40) // len(reader.pages))
            
            # API 요청 준비
            self.progress_updated.emit(50)
            
            # 프롬프트 준비
            prompt = """
You are a PDF forms extractor for public procurement proposals.  
Given multiple attached PDF documents of various bids, identify for each document all pages that are "제출용 서식" (i.e. 개별 양식 페이지 which the bidder must 작성·제출해야 하는 서류).  

For each document, output a JSON object with:
+- "doc": the original PDF filename (without path)  
+- "forms": an array of form entries, each with:
    - "page": page number (integer)
    - "title": the exact form title in Korean (e.g. "입찰참가신청서")
    - "filename": suggested filename in the format "{page}p_{제목}.pdf"
    - "requires_input": true if the form contains blank fields the user must 입력해야 하는 경우

Return a JSON array of these document objects. Do **not** include any extra commentary.
            """
            
            filename = os.path.basename(self.pdf_path)
            
            # OpenAI API 호출
            client = OpenAI(api_key=self.api_key)
            self.progress_updated.emit(60)
            response = client.chat.completions.create(
                model="gpt-4.1-mini",
                messages=[
                    {"role": "system", "content": prompt},
                    {"role": "user", "content": f"Analyze this PDF content from file '{filename}':\n\n{text_content}"}
                ],
                temperature=0.2,
                max_tokens=1500
            )
            self.progress_updated.emit(90)
            
            # 결과 파싱
            content = response.choices[0].message.content
            
            # JSON 추출 (텍스트에서 JSON 부분만 추출)
            try:
                json_start = content.find("[")
                json_end = content.rfind("]") + 1
                if json_start >= 0 and json_end > json_start:
                    json_content = content[json_start:json_end]
                    result = json.loads(json_content)
                    
                    # 배열 내 첫 번째 객체 반환 (단일 PDF 분석)
                    if isinstance(result, list) and len(result) > 0:
                        self.progress_updated.emit(100)
                        self.analysis_complete.emit(result[0])
                        return
            except Exception as e:
                print(f"JSON 파싱 오류: {e}")
                
            # 직접 JSON 파싱에 실패한 경우 전체 반환
            try:
                result = {"doc": filename, "forms": []}
                
                # 정규식으로 페이지 번호와 제목 추출 시도
                matches = re.findall(r'page[^\d]*(\d+).*title[^\w가-힣]*([\w가-힣]+)', content)
                for page_str, title in matches:
                    page = int(page_str)
                    result["forms"].append({
                        "page": page,
                        "title": title,
                        "filename": f"{page}p_{title}.pdf",
                        "requires_input": True
                    })
                
                self.progress_updated.emit(100)
                self.analysis_complete.emit(result)
            except Exception as e:
                print(f"백업 파싱 오류: {e}")
                self.progress_updated.emit(100)
                self.analysis_complete.emit({})
                
        except Exception as e:
            print(f"분석 오류: {e}")
            self.progress_updated.emit(100)
            self.analysis_complete.emit({})


# 독립 실행 시 테스트용 코드
if __name__ == "__main__":
    app = QApplication(sys.argv)
    editor = PdfFormEditor()
    editor.show()
    sys.exit(app.exec_()) 