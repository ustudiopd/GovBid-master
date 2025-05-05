import os
import tempfile
import traceback
from PyQt5.QtWidgets import QMessageBox, QProgressDialog, QApplication
from dropbox_client import list_folder, download_file, upload_json, download_json
from gpt_client import analyze_pdfs

class Analyzer:
    """PDF 분석 관리 클래스"""
    
    @staticmethod
    def analyze_folder(folder, parent=None):
        """
        지정된 폴더의 PDF 파일을 분석
        
        Args:
            folder: 분석할 폴더명
            parent: 부모 위젯 (QMessageBox 표시용)
            
        Returns:
            성공 여부 (boolean)
        """
        try:
            files = list_folder(f"입찰 2025/{folder}")
            pdfs = [f for f in files if f.lower().endswith(".pdf")]
            if not pdfs:
                QMessageBox.warning(parent, "PDF 없음", f"{folder} 폴더에 PDF 파일이 없습니다.")
                return False
            
            # 진행 상태 대화상자 생성
            progress = QProgressDialog("PDF 분석 중...", "취소", 0, 100, parent)
            progress.setWindowTitle("PDF 분석")
            progress.setModal(True)
            progress.show()
            
            # 작업 진행률 업데이트 함수
            def update_progress(value):
                progress.setValue(value)
                QApplication.processEvents()
                # 사용자가 취소 버튼을 눌렀는지 확인
                return not progress.wasCanceled()
            
            temp_dir = tempfile.mkdtemp()
            paths = []
            
            # 다운로드 진행 상태 표시
            progress.setLabelText("PDF 파일 다운로드 중...")
            for i, pdf in enumerate(pdfs):
                progress.setValue(int(i / len(pdfs) * 20))  # 다운로드는 20%까지
                QApplication.processEvents()
                if progress.wasCanceled():
                    return False  # 사용자가 취소함
                    
                local = os.path.join(temp_dir, pdf)
                download_file(f"입찰 2025/{folder}/{pdf}", local)
                paths.append(local)
            
            # 분석 진행 상태 표시
            progress.setLabelText("PDF 내용 분석 중...")
            progress.setValue(20)  # 다운로드 완료, 분석 시작
            QApplication.processEvents()
                
            try:
                analysis = analyze_pdfs(paths)
            except ValueError as e:
                # JSON 파싱 에러 상세 표시
                error_msg = str(e)
                QMessageBox.critical(parent, "GPT 응답 파싱 오류", f"API 응답을 파싱할 수 없습니다:\n{error_msg}")
                return False
            except Exception as e:
                # 기타 분석 에러 상세 표시
                error_msg = traceback.format_exc()
                QMessageBox.critical(parent, "PDF 분석 오류", f"분석 중 에러 발생:\n{str(e)}\n\n{error_msg}")
                return False
            
            # 업로드 진행 상태 표시    
            progress.setLabelText("분석 결과 업로드 중...")
            progress.setValue(80)  # 분석 완료, 업로드 시작
            QApplication.processEvents()
                
            # 분석 결과 업로드
            upload_json(f"입찰 2025/{folder}/analysis.json", analysis)
            
            # smpp.json 업데이트
            progress.setLabelText("메타데이터 업데이트 중...")
            progress.setValue(90)  # 업로드 완료, 메타데이터 업데이트 시작
            
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
            
            # 완료
            progress.setValue(100)
            QMessageBox.information(parent, "분석 완료", f"{folder} 분석이 완료되었습니다.")
            return True
        except Exception as e:
            QMessageBox.critical(parent, "분석 에러", str(e))
            return False 