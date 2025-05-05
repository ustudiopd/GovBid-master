# main.py
import sys
import json
import traceback
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QTableWidget, QTableWidgetItem, QPushButton, QMessageBox,
    QHeaderView, QToolTip, QDialog, QTreeWidget, QTreeWidgetItem
)
from PyQt5.QtGui import QCursor
from dropbox_client import list_folder, download_json, download_file, upload_json
# 루트 기반 gpt_client 임포트로 변경
from gpt_client import analyze_pdfs
import tempfile, os

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
        dlg.resize(800, 600)
        dlg.exec_()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_()) 