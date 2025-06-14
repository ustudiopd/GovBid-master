# 버전 변경 기록

## 버전 1.0.0 (2025-05-05)

### 주요 변경사항
- 초기 릴리스 버전
- PyQt5 기반 GUI 인터페이스 구현
- Dropbox 연동 시스템 구현
- PDF 문서 분석 기능 구현
- 서식 추출 기능 구현

### 세부 변경사항

#### 1. PDF 서식 추출 기능
- ChatGPT API를 사용하여 서식 페이지 자동 식별
- 300dpi 고해상도 PDF 렌더링 지원
- 서식을 개별 PDF 파일로 추출

#### 2. 서식 분석 전용 모듈 구현
- `pdf_client.py` 모듈 개발
- 여러 PDF 파일에서 서식 찾기 기능
- 서식 분석 결과를 JSON으로 저장

#### 3. Dropbox 연동 개선
- API 업로드 실패 문제 해결
- 로컬 Dropbox 폴더에 직접 파일 저장 기능 추가
- 자동 Dropbox 폴더 감지 기능 구현

#### 4. UI 개선
- 서식 파일 리스트 트리 뷰 구현
- 분석 상세 대화상자 구현
- 진행 상태 표시 대화상자 추가

### 문제 해결
- .env 파일에서 CHATGPT_API_KEY 로드 문제 해결
- 다중 따옴표 문자열 구문 오류 수정
- GPT API 토큰 한도 조정 (50000→30000→5000→1000)
- Dropbox API 업로드 실패 문제 해결

### 테스트 결과
- 인천글로벌캠퍼스 홍보동영상 제작 입찰 문서에서 9개의 서식(입찰참가신청서, 위임장 등) 추출 완료

## 버전 0.9.0 (2025-04-15) - 베타 릴리스

### 주요 변경사항
- PDF 서식 추출 기능 초기 구현
- PDF 편집기 구현
- ChatGPT API 연동 구현
- Dropbox API 연동 구현

### 세부 변경사항

#### 1. PDF 편집기 구현
- `pdf_editor.py` 모듈 개발
- PDF 파일 열기 및 페이지 이동 기능
- 텍스트 상자 추가 및 편집 기능
- 현재 페이지 PNG로 저장 기능

#### 2. ChatGPT API 연동
- `gpt_client.py` 모듈 개발
- OpenAI API 호출 및 응답 처리
- 입찰 문서 분석 프롬프트 설계

#### 3. Dropbox API 연동
- `dropbox_client.py` 모듈 개발
- 폴더 목록 조회, 파일 다운로드/업로드 기능
- OAuth 인증 처리

### 문제 해결
- PyQt5 위젯 레이아웃 문제 해결
- PDF 텍스트 추출 인코딩 문제 해결
- API 키 관리 문제 해결

## 버전 0.5.0 (2025-03-20) - 알파 릴리스

### 주요 변경사항
- 메인 애플리케이션 구조 설계
- PyQt5 기반 GUI 인터페이스 초기 구현
- 입찰 폴더 관리 기능 구현

### 세부 변경사항

#### 1. 메인 애플리케이션 구조 설계
- `main.py` 모듈 개발
- 입찰 폴더 목록 표시 기능
- 입찰 정보 테이블 구현

#### 2. 애플리케이션 설정 관리
- `settings.py` 모듈 개발
- .env 파일을 통한 설정 관리

### 문제 해결
- PyQt5 설치 이슈 해결
- 개발 환경 구성 문제 해결 