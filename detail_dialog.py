import os
import re
import json
import tempfile
import shutil
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QTreeWidget, QTreeWidgetItem,
    QPushButton, QMessageBox, QApplication, QTextEdit,
    QHBoxLayout
)
from PyQt5.QtCore import Qt
from dropbox_client import list_folder, download_json, download_file, upload_json, upload_file
from toc_guide_generator import TocGuideGenerator
from manual_toc_guide import ManualTocGuideDialog

class DetailDialog(QDialog):
    def __init__(self, parent=None, entry=None, folder=None):
        """
        세부 정보를 표시하는 대화상자
        
        Args:
            parent: 부모 위젯
            entry: 항목 데이터 (딕셔너리)
            folder: 폴더명
        """
        super().__init__(parent)
        self.entry = entry
        self.folder = folder
        
        # 부모 창에서 entries 목록 참조 (분석 후 업데이트용)
        self.parent_entries = getattr(parent, 'entries', [])
        
        title = entry.get('announcement_info', {}).get('공고명', '분석 상세')
        self.setWindowTitle(title)
        self.resize(800, 600)
        
        # 레이아웃 설정
        self.layout = QVBoxLayout(self)
        
        # 상단 버튼 레이아웃
        top_layout = QHBoxLayout()
        
        # 목차 가이드 생성 버튼 추가
        self.toc_guide_btn = QPushButton("목차가이드 생성")
        self.toc_guide_btn.clicked.connect(self.generate_toc_guide)
        top_layout.addWidget(self.toc_guide_btn)
        
        # 전체화면 전환 버튼 추가
        self.fullscreen_button = QPushButton("전체화면")
        self.fullscreen_button.clicked.connect(self.toggle_fullscreen)
        top_layout.addWidget(self.fullscreen_button)
        
        # 상단 레이아웃을 메인 레이아웃에 추가
        self.layout.addLayout(top_layout)
        
        # 트리위젯 생성
        self.tree = QTreeWidget()
        self.tree.setColumnCount(2)
        self.tree.setHeaderLabels(['Key', 'Value'])
        self.layout.addWidget(self.tree)
        
        # 분석 데이터 불러오기 및 표시
        self.load_analysis_data()
        
        # 키 이벤트 핸들러 설정
        self.tree.keyPressEvent = self.handle_key_press
    
    def toggle_fullscreen(self):
        """전체화면 전환"""
        if self.isFullScreen():
            self.showNormal()
            self.fullscreen_button.setText("전체화면")
        else:
            self.showFullScreen()
            self.fullscreen_button.setText("일반화면")
    
    def handle_key_press(self, event):
        """키 이벤트 처리"""
        if event.key() == Qt.Key_Space:
            # 스페이스바를 누르면 모든 항목 펼치기/접기 토글
            root = self.tree.invisibleRootItem()
            is_expanded = False
            
            # 현재 상태 확인 (첫 번째 항목의 상태로 판단)
            if root.childCount() > 0:
                is_expanded = root.child(0).isExpanded()
            
            # 모든 항목의 상태 변경
            for i in range(root.childCount()):
                item = root.child(i)
                item.setExpanded(not is_expanded)
                
                # 재귀적으로 모든 하위 항목도 변경
                self._toggle_all_children(item, not is_expanded)
        elif event.key() == Qt.Key_F11:
            self.toggle_fullscreen()
        else:
            # 다른 키는 기본 이벤트 핸들러로 전달
            QTreeWidget.keyPressEvent(self.tree, event)
    
    def _toggle_all_children(self, item, expanded):
        """재귀적으로 모든 하위 항목의 펼침/접기 상태 변경"""
        for i in range(item.childCount()):
            child = item.child(i)
            child.setExpanded(expanded)
            self._toggle_all_children(child, expanded)
    
    def load_analysis_data(self):
        """분석 데이터 로드 및 트리 위젯에 표시"""
        try:
            # 현재 폴더의 JSON 파일 목록 가져오기
            folder_contents = list_folder(f"입찰 2025/{self.folder}")
            json_files = [f for f in folder_contents if f.lower().endswith('.json')]
            
            # 트리위젯 생성
            self.tree.clear()
            root = self.tree.invisibleRootItem()
            
            # 서식 파일 리스트 노드 추가
            forms_node = QTreeWidgetItem(["서식파일", ""])
            
            # 서식 폴더 확인 및 파일 리스트 불러오기
            has_forms_folder = False
            try:
                # 서식 폴더가 있는지 확인
                has_forms_folder = "서식" in folder_contents
                
                if has_forms_folder:
                    # 서식 폴더 내 파일 목록 가져오기
                    forms_files = list_folder(f"입찰 2025/{self.folder}/서식")
                    pdf_files = [f for f in forms_files if f.lower().endswith('.pdf')]
                    
                    if pdf_files:
                        for pdf_file in pdf_files:
                            form_item = QTreeWidgetItem([pdf_file, ""])
                            # PDF 파일과 페이지 번호 추출
                            match = re.match(r'(\d+)p_(.+)\.pdf', pdf_file)
                            if match:
                                page_num, form_name = match.groups()
                                form_item.setText(1, f"페이지 {page_num} - {form_name}")
                            forms_node.addChild(form_item)
                    else:
                        forms_node.addChild(QTreeWidgetItem(["", "서식 PDF 파일 없음"]))
                else:
                    forms_node.addChild(QTreeWidgetItem(["", "서식 폴더 없음"]))
            except Exception as e:
                forms_node.addChild(QTreeWidgetItem(["오류", str(e)]))
            
            # 서식파일 노드를 기본적으로 펼쳐서 보여줌
            forms_node.setExpanded(True)
            root.addChild(forms_node)
            
            # JSON 파일 노드 추가
            json_node = QTreeWidgetItem(["JSON 파일", ""])
            root.addChild(json_node)
            
            # 각 JSON 파일 내용 로드 및 표시
            for json_file in json_files:
                try:
                    # JSON 파일 다운로드
                    json_data = download_json(f"입찰 2025/{self.folder}/{json_file}")
                    
                    # JSON 파일 노드 생성
                    file_node = QTreeWidgetItem([json_file, ""])
                    json_node.addChild(file_node)
                    
                    # JSON 데이터를 트리에 추가
                    self._add_json_to_tree(file_node, json_data)
                    
                except Exception as e:
                    error_item = QTreeWidgetItem([json_file, f"로드 오류: {str(e)}"])
                    json_node.addChild(error_item)
            
            # JSON 노드도 기본적으로 펼쳐서 보여줌
            json_node.setExpanded(True)
            
            # 서식 추출 버튼 추가 (서식 폴더가 없을 경우)
            if not has_forms_folder:
                extract_btn = QPushButton("서식파일 분석하기")
                extract_btn.clicked.connect(self.extract_form_templates)
                self.layout.addWidget(extract_btn)
                
        except Exception as e:
            QMessageBox.warning(self, '폴더 접근 에러', f'폴더 내용을 가져오는 중 오류 발생:\n{e}')
    
    def _add_json_to_tree(self, parent_node, data, key=None):
        """JSON 데이터를 트리 위젯에 추가하는 재귀 함수"""
        if isinstance(data, dict):
            for k, v in data.items():
                child = QTreeWidgetItem([str(k), ""])
                parent_node.addChild(child)
                self._add_json_to_tree(child, v, k)
        elif isinstance(data, list):
            for i, v in enumerate(data):
                child = QTreeWidgetItem([f"[{i}]", ""])
                parent_node.addChild(child)
                self._add_json_to_tree(child, v, key)
        else:
            if key:
                parent_node.setText(1, str(data))
            else:
                parent_node.setText(0, str(data))
    
    def extract_form_templates(self):
        """서식 분석 기능 호출 - 모든 PDF를 분석하여 서식 찾기"""
        try:
            # 해당 폴더의 PDF 파일 목록 가져오기
            files = list_folder(f"입찰 2025/{self.folder}")
            pdfs = [f for f in files if f.lower().endswith(".pdf")]
            
            if not pdfs:
                QMessageBox.warning(self, "PDF 없음", f"{self.folder} 폴더에 PDF 파일이 없습니다.")
                return
                
            # 로그 대화상자 생성
            log_dialog = QDialog(self)
            log_dialog.setWindowTitle("서식 분석 로그")
            log_dialog.resize(700, 400)
            log_layout = QVBoxLayout(log_dialog)
            
            log_text = QTextEdit()
            log_text.setReadOnly(True)
            log_layout.addWidget(log_text)
            
            # 로그 콜백 함수
            def log_callback(message):
                log_text.append(message)
                log_text.moveCursor(log_text.textCursor().End)
                QApplication.processEvents()  # UI 업데이트 처리
            
            # 최초 로그 메시지
            log_callback(f"서식 분석 시작: {self.folder} ({len(pdfs)}개 PDF 파일)")
            log_dialog.show()
            
            # pdf_client 모듈 사용
            from pdf_client import analyze_form_templates
            
            # 분석 실행 (임시 폴더에 PDF 다운로드 후 분석)
            temp_dir = tempfile.mkdtemp()
            local_paths = []

            # PDF 파일 다운로드
            for i, pdf in enumerate(pdfs):
                log_callback(f"PDF 다운로드 중: {pdf}")
                local_path = os.path.join(temp_dir, pdf)
                download_file(f"입찰 2025/{self.folder}/{pdf}", local_path)
                local_paths.append(local_path)
                QApplication.processEvents()  # UI 업데이트
            
            # 서식 분석 실행
            log_callback("서식 페이지 분석 중...")
            
            # 분석 및 결과 저장 (프로그레스바 없이 로그 콜백만 사용)
            result = analyze_form_templates(
                local_paths, 
                progress_callback=None,  # 프로그레스바 콜백 제거
                log_callback=log_callback,
                folder_name=self.folder  # 현재 폴더명 전달
            )
                
            if not result or not result.get('forms'):
                log_callback("서식 페이지를 찾을 수 없습니다.")
                
                # 결과 없음으로 JSON 저장
                # PDF 파일 목록 확인
                analyzed_files = result.get('analyzed_files', []) if result else []
                if not analyzed_files:
                    analyzed_files = [{"filename": pdf} for pdf in pdfs]
                
                result_json = {
                    "doc": self.folder, 
                    "forms": [], 
                    "message": "서식 페이지를 찾을 수 없습니다.",
                    "analyzed_files": analyzed_files
                }
                
                # 결과 저장 - pdf_client에서 이미 저장한 경우 생략
                json_saved = False
                for path in local_paths:
                    forms_dir = os.path.dirname(path)
                    result_path = os.path.join(forms_dir, "서식분석결과.json")
                    if os.path.exists(result_path):
                        json_saved = True
                        break
                
                if not json_saved:
                    # 공고명 폴더에 저장
                    try:
                        with open(os.path.join(temp_dir, "서식분석결과.json"), "w", encoding="utf-8") as f:
                            json.dump(result_json, f, ensure_ascii=False, indent=2)
                        log_callback(f"서식분석결과.json 파일 저장: {temp_dir}")
                        # Dropbox에 업로드
                        upload_json(f"입찰 2025/{self.folder}/서식분석결과.json", result_json)
                        log_callback(f"서식분석결과.json 파일 Dropbox 업로드 완료")
                        QMessageBox.information(self, "알림", 
                            "서식 페이지를 찾을 수 없습니다.\n서식분석결과.json 파일이 Dropbox에 저장되었습니다.")
                    except Exception as e:
                        error_msg = f"JSON 저장 오류: {e}"
                        log_callback(error_msg)
                        # Dropbox에 업로드
                        upload_json(f"입찰 2025/{self.folder}/서식분석결과.json", result_json)
                        log_callback(f"서식분석결과.json 파일 Dropbox 업로드 완료")
                        QMessageBox.information(self, "알림", 
                            "서식 페이지를 찾을 수 없습니다.\n서식분석결과.json 파일이 Dropbox에 저장되었습니다.")
                
                # 세부창 다시 로드 (새로운 데이터 표시)
                self.load_analysis_data()
                
                # 로그 대화상자는 열어둠 (사용자가 닫을 수 있음)
                return
                
            # 서식 파일 생성 완료 확인
            forms_saved = False
            for form in result.get('forms', []):
                if form.get('final_path') and os.path.exists(form.get('final_path')):
                    forms_saved = True
                    break
            
            if forms_saved:
                # 이미 pdf_client.py에서 서식 파일 저장 완료
                forms_dir = os.path.dirname(result['forms'][0].get('final_path'))
                log_callback(f"서식 파일 저장 완료: {len(result.get('forms', []))}개 파일")
                QMessageBox.information(self, "완료", 
                    f"서식 페이지 분석 완료: {len(result.get('forms', []))}개 서식 PDF가 '{forms_dir}'에 저장되었습니다.")
            else:
                # 서식 파일이 저장되지 않은 경우 (백업 처리)
                try:
                    # 원본 PDF 폴더에 저장
                    if local_paths:
                        original_dir = os.path.dirname(local_paths[0])
                        forms_dir = os.path.join(original_dir, "서식")
                        os.makedirs(forms_dir, exist_ok=True)
                        log_callback(f"서식 폴더 생성: {forms_dir}")
                        
                        # 각 서식 파일 추출 및 저장
                        saved_count = 0
                        for form in result.get('forms', []):
                            page = form.get('page')
                            if page is None:
                                continue
                                
                            output_path = form.get('output_path')
                            if output_path and os.path.exists(output_path):
                                # 이미 생성된 파일 복사
                                filename = os.path.basename(output_path)
                                dest_path = os.path.join(forms_dir, filename)
                                shutil.copy2(output_path, dest_path)
                                log_callback(f"서식 파일 복사: {filename}")
                                saved_count += 1
                            else:
                                # 페이지 추출 시도 (필요 시 PyPDF2 임포트)
                                try:
                                    from PyPDF2 import PdfReader, PdfWriter
                                    for pdf_path in local_paths:
                                        try:
                                            reader = PdfReader(pdf_path)
                                            if page <= len(reader.pages):
                                                # 파일명 생성
                                                filename = form.get('filename', f"{page}p_서식.pdf")
                                                filename = re.sub(r'[\\/*?:"<>|]', "", filename)
                                                dest_path = os.path.join(forms_dir, filename)
                                                
                                                # 0-기반 인덱스로 변환
                                                page_idx = page - 1
                                                
                                                # 단일 페이지 추출
                                                writer = PdfWriter()
                                                writer.add_page(reader.pages[page_idx])
                                                
                                                # 파일로 저장
                                                with open(dest_path, "wb") as out_file:
                                                    writer.write(out_file)
                                                
                                                log_callback(f"서식 파일 생성: {filename}")
                                                saved_count += 1
                                                break
                                        except Exception as e:
                                            error_msg = f"서식 추출 오류 (페이지 {page}): {e}"
                                            log_callback(error_msg)
                                except ImportError:
                                    log_callback("PyPDF2 라이브러리를 찾을 수 없습니다. pip install PyPDF2로 설치하세요.")
                        
                        # 결과 JSON 파일 저장
                        result_path = os.path.join(temp_dir, "서식분석결과.json")
                        with open(result_path, "w", encoding="utf-8") as f:
                            json.dump(result, f, ensure_ascii=False, indent=2)
                        
                        log_callback(f"서식분석결과.json 파일 저장: {result_path}")
                        
                        # Dropbox에 업로드
                        upload_json(f"입찰 2025/{self.folder}/서식분석결과.json", result)
                        log_callback(f"서식분석결과.json 파일 Dropbox 업로드 완료")
                        
                        # 완료 메시지 표시
                        QMessageBox.information(self, "완료", 
                            f"서식 페이지 분석 완료: {saved_count}개 서식 PDF가 '{forms_dir}'에 저장되었습니다.")
                    else:
                        # Dropbox API 사용
                        forms_dir = f"입찰 2025/{self.folder}/서식"
                        saved_count = 0
                        log_callback(f"Dropbox 폴더 생성: {forms_dir}")
                        
                        # 각 서식 파일을 Dropbox에 업로드
                        for form in result.get('forms', []):
                            output_path = form.get('output_path')
                            if not output_path or not os.path.exists(output_path):
                                continue
                                
                            filename = os.path.basename(output_path)
                            remote_path = f"{forms_dir}/{filename}"
                            
                            try:
                                # 파일 업로드
                                upload_file(remote_path, output_path)
                                log_callback(f"서식 파일 업로드: {filename}")
                                saved_count += 1
                            except Exception as e:
                                error_msg = f"파일 업로드 오류: {e}"
                                log_callback(error_msg)
                        
                        # 결과 JSON 파일 저장
                        upload_json(f"입찰 2025/{self.folder}/서식/서식분석결과.json", result)
                        log_callback(f"서식분석결과.json 파일 업로드 완료")
                        
                        # 완료 메시지 표시
                        QMessageBox.information(self, "완료", 
                            f"서식 페이지 분석 완료: {saved_count}개 서식 PDF가 Dropbox에 저장되었습니다.")
                    
                except Exception as e:
                    error_msg = f"서식 파일 저장 중 오류 발생: {str(e)}"
                    log_callback(error_msg)
                    QMessageBox.warning(self, "저장 오류", error_msg)
                    # 로컬 경로 비상 대책 안내
                    temp_forms_dir = os.path.join(temp_dir, "서식")
                    if os.path.exists(temp_forms_dir) and os.listdir(temp_forms_dir):
                        log_callback(f"임시 저장 위치: {temp_forms_dir}")
                        QMessageBox.information(self, "임시 저장 위치", 
                            f"서식 파일이 다음 임시 폴더에 저장되어 있습니다:\n{temp_forms_dir}\n"
                            f"이 폴더의 내용을 수동으로 복사하세요.")
            
            # 세부창 데이터 다시 로드
            self.load_analysis_data()
            
            # 로그 대화상자는 열어둠 (사용자가 닫을 수 있음)
                
        except Exception as e:
            error_msg = f"서식 분석 오류: {str(e)}"
            QMessageBox.critical(self, "서식 분석 오류", error_msg) 
    
    def generate_toc_guide(self):
        """목차 가이드 생성"""
        # 수동 목차 가이드 생성 창 열기
        dialog = ManualTocGuideDialog(self)
        if dialog.exec_():
            # 저장 성공 시 데이터 다시 로드
            self.load_analysis_data() 