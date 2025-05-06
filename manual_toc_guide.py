from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTextEdit, 
    QPushButton, QLabel, QApplication, QMessageBox,
    QProgressDialog
)
from PyQt5.QtCore import Qt
import os
import json
import tempfile
from dropbox_client import list_folder, download_file, upload_file, upload_json
from dotenv import load_dotenv
from PyPDF2 import PdfReader, PdfWriter
import re

class ManualTocGuideDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("목차 가이드 생성")
        self.setMinimumSize(800, 600)
        
        # .env 파일 로드
        load_dotenv()
        self.local_base_path = os.getenv('LOCAL_BID_FOLDER', 'C:/Users/ustudiogram/U-Studio Dropbox/양승철/입찰 2025')
        
        # 부모 창에서 폴더 정보 가져오기
        self.folder = getattr(parent, 'folder', None)
        
        # PDF 파일 목록 가져오기
        self.pdf_files = []
        if self.folder:
            try:
                files = list_folder(f"입찰 2025/{self.folder}")
                self.pdf_files = [f for f in files if f.lower().endswith('.pdf')]
            except Exception as e:
                QMessageBox.warning(self, "PDF 목록 오류", f"PDF 파일 목록을 가져오는 중 오류 발생:\n{e}")
        
        # 프롬프트 텍스트
        self.prompt_text = f"""첨부된 PDF를 종합분석하여, 이 입찰이 요구하는 입찰제안서의 **상세 목차**(대·중·소항목까지)와 각 목차별로 반드시 지켜야 할 **작성 가이드**(유의사항·표현방식·분량 제한 등)를 한글로 깔끔하게 정리해 주세요.  
문서 내 "제안서 작성 안내" 또는 "제안 요청사항" 섹션을 참고하고, 페이지 번호는 생략해 주세요.

분석할 PDF 파일 목록:
{chr(10).join(f'- {pdf}' for pdf in self.pdf_files)}"""
        
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout()
        
        # PDF 파일 목록 표시
        if self.pdf_files:
            pdf_label = QLabel("분석할 PDF 파일 목록:")
            pdf_list = QTextEdit()
            pdf_list.setPlainText(chr(10).join(self.pdf_files))
            pdf_list.setReadOnly(True)
            pdf_list.setMaximumHeight(100)
            layout.addWidget(pdf_label)
            layout.addWidget(pdf_list)
        
        # 프롬프트 섹션
        prompt_label = QLabel("ChatGPT에 복사할 프롬프트:")
        self.prompt_edit = QTextEdit()
        self.prompt_edit.setPlainText(self.prompt_text)
        self.prompt_edit.setReadOnly(True)
        
        # 버튼 레이아웃
        button_layout = QHBoxLayout()
        
        copy_button = QPushButton("프롬프트 복사")
        copy_button.clicked.connect(self.copy_prompt)
        button_layout.addWidget(copy_button)
        
        if self.folder:
            open_folder_button = QPushButton("로컬 폴더 열기")
            open_folder_button.clicked.connect(self.open_local_folder)
            button_layout.addWidget(open_folder_button)
            
            auto_analyze_button = QPushButton("자동 분석")
            auto_analyze_button.clicked.connect(self.auto_analyze)
            button_layout.addWidget(auto_analyze_button)
        
        prompt_layout = QVBoxLayout()
        prompt_layout.addWidget(prompt_label)
        prompt_layout.addWidget(self.prompt_edit)
        prompt_layout.addLayout(button_layout)
        
        # 결과 입력 섹션
        result_label = QLabel("ChatGPT 결과를 여기에 붙여넣기:")
        self.result_edit = QTextEdit()
        
        # 저장/취소 버튼 섹션
        save_button_layout = QHBoxLayout()
        save_button = QPushButton("저장")
        save_button.clicked.connect(self.save_result)
        cancel_button = QPushButton("취소")
        cancel_button.clicked.connect(self.reject)
        
        save_button_layout.addWidget(save_button)
        save_button_layout.addWidget(cancel_button)
        
        # 레이아웃 조립
        layout.addLayout(prompt_layout)
        layout.addWidget(result_label)
        layout.addWidget(self.result_edit)
        layout.addLayout(save_button_layout)
        
        self.setLayout(layout)
    
    def copy_prompt(self):
        """프롬프트 복사"""
        prompt = """첨부된 pdf를 종합 분석하여, 이 입찰이 요구하는 입찰제안서의 **상세 목차**(대·중·소항목까지)와 각 목차별로 반드시 지켜야 할 **작성 가이드**(유의사항·표현방식·분량 제한 등)를 한글로 깔끔하게 정리해 주세요. 문서 내 "제안서 작성 안내" 또는 "제안 요청사항" 섹션을 참고하고, 페이지 번호는 생략해 주세요."""
        clipboard = QApplication.clipboard()
        clipboard.setText(prompt)
    
    def open_local_folder(self):
        """로컬 폴더 열기"""
        if not self.folder:
            return
            
        folder_path = os.path.join(self.local_base_path, self.folder)
        if os.path.exists(folder_path):
            os.startfile(folder_path)
        else:
            QMessageBox.warning(self, "폴더 없음", f"로컬 폴더를 찾을 수 없습니다:\n{folder_path}")
    
    def auto_analyze(self):
        """PDF 자동 분석"""
        if not self.folder or not self.pdf_files:
            return
            
        try:
            # 진행 상태 대화상자 생성
            progress = QProgressDialog("PDF 분석 중...", "취소", 0, 100, self)
            progress.setWindowTitle("목차 가이드 생성")
            progress.setModal(True)
            progress.show()
            
            # 임시 폴더 생성
            temp_dir = tempfile.mkdtemp()
            local_paths = []
            
            # PDF 파일 다운로드
            progress.setLabelText("PDF 파일 다운로드 중...")
            for i, pdf in enumerate(self.pdf_files):
                progress.setValue(int(i / len(self.pdf_files) * 10))
                if progress.wasCanceled():
                    return
                local_path = os.path.join(temp_dir, pdf)
                download_file(f"입찰 2025/{self.folder}/{pdf}", local_path)
                local_paths.append(local_path)
            
            # PDF 내용 추출 및 목차/가이드 페이지 추출
            progress.setLabelText("PDF 내용 분석 중...")
            progress.setValue(20)
            
            all_text = ""
            toc_writer = PdfWriter()
            keywords = ["목차", "작성 가이드", "제안서 작성 안내"]
            found_pages = 0
            for pdf_path in local_paths:
                reader = PdfReader(pdf_path)
                for i, page in enumerate(reader.pages):
                    text = page.extract_text() or ""
                    all_text += text + "\n\n"
                    if any(keyword in text for keyword in keywords):
                        toc_writer.add_page(page)
                        found_pages += 1
            
            # 목차.pdf 저장
            folder_path = os.path.join(self.local_base_path, self.folder)
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)
            toc_pdf_path = os.path.join(folder_path, "목차.pdf")
            if found_pages > 0:
                with open(toc_pdf_path, "wb") as f:
                    toc_writer.write(f)
            
            # ChatGPT API 호출 (openai 최신 방식)
            progress.setLabelText("ChatGPT 분석 중...")
            progress.setValue(40)
            import openai
            from openai import OpenAI
            api_key = os.getenv("CHATGPT_API_KEY")
            if not api_key:
                raise ValueError(".env 파일에 CHATGPT_API_KEY가 없습니다.")
            client = OpenAI(api_key=api_key)
            response = client.chat.completions.create(
                model="gpt-4",
                messages=[
                    {"role": "system", "content": "당신은 입찰 제안서 작성 전문가입니다."},
                    {"role": "user", "content": f"{self.prompt_text}\n\nPDF 내용:\n{all_text[:4000]}"}
                ]
            )
            result = response.choices[0].message.content
            
            # 결과를 JSON으로 변환
            progress.setLabelText("결과 저장 중...")
            progress.setValue(80)
            toc_match = re.search(r'\[목차\](.*?)(?=\[작성 가이드\]|$)', result, re.DOTALL)
            guide_match = re.search(r'\[작성 가이드\](.*?)$', result, re.DOTALL)
            toc_content = toc_match.group(1).strip() if toc_match else ""
            guide_content = guide_match.group(1).strip() if guide_match else ""
            json_data = {
                "toc": toc_content,
                "guide": guide_content,
                "source_files": self.pdf_files
            }
            # JSON 파일 저장
            json_path = os.path.join(folder_path, "목차가이드.json")
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(json_data, f, ensure_ascii=False, indent=2)
            # 텍스트 파일 저장
            txt_path = os.path.join(folder_path, "목차가이드.txt")
            with open(txt_path, "w", encoding="utf-8") as f:
                f.write(result)
            # Dropbox에 업로드
            upload_json(f"입찰 2025/{self.folder}/목차가이드.json", json_data)
            progress.setValue(100)
            QMessageBox.information(self, "완료", 
                f"목차 가이드 생성이 완료되었습니다.\n"
                f"저장 위치: {folder_path}\n"
                f"- 목차가이드.txt\n"
                f"- 목차가이드.json\n"
                f"- 목차.pdf (목차/가이드 관련 페이지)\n")
            self.result_edit.setPlainText(result)
        except Exception as e:
            QMessageBox.critical(self, "분석 오류", f"PDF 분석 중 오류가 발생했습니다:\n{str(e)}")
    
    def save_result(self):
        """결과 저장"""
        result = self.result_edit.toPlainText()
        if not result:
            QMessageBox.warning(self, "저장 실패", "저장할 내용이 없습니다.")
            return
            
        try:
            # 로컬 폴더 경로에 저장
            folder_path = os.path.join(self.local_base_path, self.folder)
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)
                
            file_path = os.path.join(folder_path, "목차가이드.txt")
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(result)
                
            QMessageBox.information(self, "저장 완료", f"목차가이드.txt 파일이 저장되었습니다.\n저장 위치: {folder_path}")
            self.accept()
            
        except Exception as e:
            QMessageBox.critical(self, "저장 실패", f"파일 저장 중 오류가 발생했습니다:\n{str(e)}") 