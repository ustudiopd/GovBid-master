# main.py
import sys
import json
import traceback
import re
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QTableWidget, QTableWidgetItem, QPushButton, QMessageBox,
    QHeaderView, QToolTip, QDialog, QTreeWidget, QTreeWidgetItem,
    QProgressDialog, QFileDialog, QTextEdit
)
from PyQt5.QtGui import QCursor
from dropbox_client import list_folder, download_json, download_file, upload_json
# 루트 기반 gpt_client 임포트로 변경
from gpt_client import analyze_pdfs
import tempfile, os
from pdf_editor import PdfFormEditor
import shutil

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

        self.load_button = QPushButton("데이터 로드")
        self.load_button.clicked.connect(self.load_data)
        layout.addWidget(self.load_button)

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
            folder = entry.get('folder_name', '')
            try:
                analysis = download_json(f"입찰 2025/{folder}/analysis.json")
            except Exception as e:
                QMessageBox.warning(self, '분석 파일 에러', f'Analysis JSON 로드 실패:\n{e}')
                return

            dlg = QDialog(self)
            title = entry.get('announcement_info', {}).get('공고명', '분석 상세')
            dlg.setWindowTitle(title)
            dlg_layout = QVBoxLayout(dlg)
            tree = QTreeWidget()
            tree.setColumnCount(2)
            tree.setHeaderLabels(['Key', 'Value'])

            def add_items(parent, key, value):
                if isinstance(value, dict):
                    node = QTreeWidgetItem([str(key), ''])
                    parent.addChild(node)
                    for k, v in value.items():
                        add_items(node, k, v)
                elif isinstance(value, list):
                    node = QTreeWidgetItem([str(key), ''])
                    parent.addChild(node)
                    for i, v in enumerate(value):
                        add_items(node, f'[{i}]', v)
                else:
                    node = QTreeWidgetItem([str(key), str(value)])
                    parent.addChild(node)

            root = tree.invisibleRootItem()
            if isinstance(analysis, dict):
                for k, v in analysis.items():
                    add_items(root, k, v)
            elif isinstance(analysis, list):
                for i, v in enumerate(analysis):
                    add_items(root, f'[{i}]', v)

            dlg_layout.addWidget(tree)
            dlg.resize(600, 400)
            dlg.exec_()

    def start_analysis(self, idx):
        entry = self.entries[idx]
        folder = entry.get("folder_name")
        try:
            files = list_folder(f"입찰 2025/{folder}")
            pdfs = [f for f in files if f.lower().endswith(".pdf")]
            if not pdfs:
                QMessageBox.warning(self, "PDF 없음", f"{folder} 폴더에 PDF 파일이 없습니다.")
                return
            
            temp_dir = tempfile.mkdtemp()
            paths = []
            for pdf in pdfs:
                local = os.path.join(temp_dir, pdf)
                download_file(f"입찰 2025/{folder}/{pdf}", local)
                paths.append(local)
                
            try:
                analysis = analyze_pdfs(paths)
            except ValueError as e:
                # JSON 파싱 에러 상세 표시
                error_msg = str(e)
                QMessageBox.critical(self, "GPT 응답 파싱 오류", f"API 응답을 파싱할 수 없습니다:\n{error_msg}")
                return
            except Exception as e:
                # 기타 분석 에러 상세 표시
                error_msg = traceback.format_exc()
                QMessageBox.critical(self, "PDF 분석 오류", f"분석 중 에러 발생:\n{str(e)}\n\n{error_msg}")
                return
                
            # 분석 결과 업로드
            upload_json(f"입찰 2025/{folder}/analysis.json", analysis)
            # smpp.json 업데이트
            smpp = download_json("입찰 2025/smpp.json")
            for item in smpp:
                if item.get("folder_name") == folder:
                    info = item.get("announcement_info", {})
                    ann = analysis.get("announcement_info", {})
                    info["등록마감"] = ann.get("등록마감", info.get("등록마감"))
                    info["공고명"] = ann.get("공고명", info.get("공고명"))
                    info["추정가격"] = ann.get("추정가격", info.get("추정가격"))
                    info["입찰내용 요약"] = analysis.get("project_summary", info.get("입찰내용 요약"))
                    item["analysis_status"] = "completed"
                    break
            upload_json("입찰 2025/smpp.json", smpp)
            QMessageBox.information(self, "분석 완료", f"{folder} 분석이 완료되었습니다.")
            self.load_data()
        except Exception as e:
            QMessageBox.critical(self, "분석 에러", str(e))

    def show_analysis_detail(self, idx):
        entry = self.entries[idx]
        folder = entry.get("folder_name")
        # Dropbox에서 analysis.json 불러오기, 실패 시 로컬 예제 사용
        try:
            analysis = download_json(f"입찰 2025/{folder}/analysis.json")
        except Exception:
            try:
                with open("alalysis.json", "r", encoding="utf-8") as f:
                    analysis = json.load(f)
            except Exception as e:
                QMessageBox.warning(self, '분석 파일 에러', f'Analysis JSON 로드 실패:\n{e}')
                return

        dlg = QDialog(self)
        title = entry.get('announcement_info', {}).get('공고명', '분석 상세')
        dlg.setWindowTitle(title)
        dlg_layout = QVBoxLayout(dlg)
        
        # 트리위젯 생성
        tree = QTreeWidget()
        tree.setColumnCount(2)
        tree.setHeaderLabels(['Key', 'Value'])
        
        # 서식 파일 리스트 노드 추가
        forms_node = QTreeWidgetItem(["서식파일", ""])
        
        # 서식 폴더 확인 및 파일 리스트 불러오기
        try:
            # 서식 폴더가 있는지 확인
            folder_contents = list_folder(f"입찰 2025/{folder}")
            has_forms_folder = "서식" in folder_contents
            
            if has_forms_folder:
                # 서식 폴더 내 파일 목록 가져오기
                forms_files = list_folder(f"입찰 2025/{folder}/서식")
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
        
        def add_items(parent, key, value):
            if isinstance(value, dict):
                node = QTreeWidgetItem([str(key), ''])
                parent.addChild(node)
                for k, v in value.items():
                    add_items(node, k, v)
            elif isinstance(value, list):
                node = QTreeWidgetItem([str(key), ''])
                parent.addChild(node)
                for i, v in enumerate(value):
                    add_items(node, f'[{i}]', v)
            else:
                node = QTreeWidgetItem([str(key), str(value)])
                parent.addChild(node)
        
        root = tree.invisibleRootItem()
        # 서식 파일 노드 먼저 추가
        root.addChild(forms_node)
        
        # 분석 결과 노드 추가
        analysis_node = QTreeWidgetItem(["분석결과", ""])
        root.addChild(analysis_node)
        
        if isinstance(analysis, dict):
            for k, v in analysis.items():
                add_items(analysis_node, k, v)
        elif isinstance(analysis, list):
            for i, v in enumerate(analysis):
                add_items(analysis_node, f'[{i}]', v)
        
        # 서식파일 노드를 기본적으로 펼쳐서 보여줌
        forms_node.setExpanded(True)
        
        dlg_layout.addWidget(tree)
        
        # 서식 추출 버튼 추가 (서식 폴더가 없을 경우)
        if not has_forms_folder:
            extract_btn = QPushButton("서식파일 분석하기")
            extract_btn.clicked.connect(lambda: self.extract_form_templates(folder))
            dlg_layout.addWidget(extract_btn)
        
        dlg.resize(800, 600)
        dlg.exec_()
    
    def extract_form_templates(self, folder):
        """서식 분석 기능 호출 - 모든 PDF를 분석하여 서식 찾기"""
        try:
            # 해당 폴더의 PDF 파일 목록 가져오기
            files = list_folder(f"입찰 2025/{folder}")
            pdfs = [f for f in files if f.lower().endswith(".pdf")]
            
            if not pdfs:
                QMessageBox.warning(self, "PDF 없음", f"{folder} 폴더에 PDF 파일이 없습니다.")
                return
                
            # 진행 대화상자 생성
            progress = QProgressDialog("서식파일 분석 준비 중...", "취소", 0, 100, self)
            progress.setWindowTitle("서식 분석")
            progress.setMinimumDuration(0)
            progress.setValue(5)
            progress.show()
            
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
            log_callback(f"서식 분석 시작: {folder} ({len(pdfs)}개 PDF 파일)")
            log_dialog.show()
            
            # pdf_client 모듈 사용
            from pdf_client import analyze_form_templates
            
            # 분석 실행 (임시 폴더에 PDF 다운로드 후 분석)
            temp_dir = tempfile.mkdtemp()
            local_paths = []

            # 진행률 업데이트
            progress.setLabelText("PDF 파일 다운로드 중...")
            progress_step = 30 / len(pdfs)
            
            for i, pdf in enumerate(pdfs):
                log_callback(f"PDF 다운로드 중: {pdf}")
                local_path = os.path.join(temp_dir, pdf)
                download_file(f"입찰 2025/{folder}/{pdf}", local_path)
                local_paths.append(local_path)
                progress.setValue(10 + int(i * progress_step))
                
                # 취소 처리
                if progress.wasCanceled():
                    log_dialog.close()
                    return
            
            # 서식 분석 실행
            progress.setLabelText("서식 페이지 분석 중...")
            progress.setValue(40)
            
            # 분석 및 결과 저장
            result = analyze_form_templates(
                local_paths, 
                progress_callback=lambda v: progress.setValue(40 + int(v * 0.5)),
                log_callback=log_callback
            )
            
            if progress.wasCanceled():
                log_dialog.close()
                return
                
            if not result or not result.get('forms'):
                log_callback("서식 페이지를 찾을 수 없습니다.")
                
                # 결과 없음으로 JSON 저장
                # PDF 파일 목록 확인
                analyzed_files = result.get('analyzed_files', []) if result else []
                if not analyzed_files:
                    analyzed_files = [{"filename": pdf} for pdf in pdfs]
                
                result_json = {
                    "doc": folder, 
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
                    # 첫 번째 PDF 위치에 저장
                    if local_paths:
                        save_dir = os.path.dirname(local_paths[0])
                        try:
                            with open(os.path.join(save_dir, "서식분석결과.json"), "w", encoding="utf-8") as f:
                                json.dump(result_json, f, ensure_ascii=False, indent=2)
                            log_callback(f"서식분석결과.json 파일 저장: {save_dir}")
                            QMessageBox.information(self, "알림", 
                                f"서식 페이지를 찾을 수 없습니다.\n서식분석결과.json 파일이 '{save_dir}'에 저장되었습니다.")
                        except Exception as e:
                            error_msg = f"JSON 저장 오류: {e}"
                            log_callback(error_msg)
                            # Dropbox에 업로드
                            upload_json(f"입찰 2025/{folder}/서식분석결과.json", result_json)
                            log_callback(f"서식분석결과.json 파일 Dropbox 업로드 완료")
                            QMessageBox.information(self, "알림", 
                                "서식 페이지를 찾을 수 없습니다.\n서식분석결과.json 파일이 Dropbox에 저장되었습니다.")
                    else:
                        # Dropbox에 업로드
                        upload_json(f"입찰 2025/{folder}/서식분석결과.json", result_json)
                        log_callback(f"서식분석결과.json 파일 Dropbox 업로드 완료")
                        QMessageBox.information(self, "알림", 
                            "서식 페이지를 찾을 수 없습니다.\n서식분석결과.json 파일이 Dropbox에 저장되었습니다.")
                else:
                    QMessageBox.information(self, "알림", "서식 페이지를 찾을 수 없습니다.")
                
                # 세부창 다시 열기 (새로운 데이터 표시)
                for i, entry in enumerate(self.entries):
                    if entry.get("folder_name") == folder:
                        log_callback("분석 결과 세부창 업데이트")
                        self.show_analysis_detail(i)
                        break
                        
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
                                # 페이지 추출 시도
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
                        
                        # 결과 JSON 파일 저장
                        result_path = os.path.join(forms_dir, "서식분석결과.json")
                        with open(result_path, "w", encoding="utf-8") as f:
                            json.dump(result, f, ensure_ascii=False, indent=2)
                        
                        log_callback(f"서식분석결과.json 파일 저장: {result_path}")
                        
                        # 완료 메시지 표시
                        QMessageBox.information(self, "완료", 
                            f"서식 페이지 분석 완료: {saved_count}개 서식 PDF가 '{forms_dir}'에 저장되었습니다.")
                    else:
                        # Dropbox API 사용
                        forms_dir = f"입찰 2025/{folder}/서식"
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
                        upload_json(f"입찰 2025/{folder}/서식/서식분석결과.json", result)
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
            
            # 세부창 다시 열기 (새로운 데이터 표시)
            for i, entry in enumerate(self.entries):
                if entry.get("folder_name") == folder:
                    log_callback("분석 결과 세부창 업데이트")
                    self.show_analysis_detail(i)
                    break
            
            # 로그 대화상자는 열어둠 (사용자가 닫을 수 있음)
                
        except Exception as e:
            error_msg = f"서식 분석 오류: {str(e)}"
            try:
                # 로그 대화상자가 있으면 로그 추가
                log_callback(error_msg)
            except:
                pass
            QMessageBox.critical(self, "서식 분석 오류", error_msg)
            
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
        
        # 특정 경로 확인 (프로젝트에 맞게 변경)
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

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_()) 