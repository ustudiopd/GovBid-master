# 입찰 문서 관리 시스템 명세서

## 1. 시스템 개요

본 시스템은 대한민국 공공 입찰 문서를 관리하고 분석하기 위한 데스크톱 애플리케이션입니다. 주요 기능으로는 입찰 폴더 관리, PDF 문서 분석, 서식 파일 자동 추출 등이 있습니다.

## 2. 시스템 구조

### 2.1 주요 모듈

- **main.py**: 메인 애플리케이션 (PyQt5 기반 GUI 인터페이스)
- **pdf_client.py**: PDF 서식 분석 및 추출 모듈
- **pdf_editor.py**: PDF 뷰어 및 텍스트 편집 기능
- **gpt_client.py**: ChatGPT API 연동 모듈
- **dropbox_client.py**: Dropbox API 연동 모듈
- **settings.py**: 애플리케이션 설정 관리

### 2.2 기술 스택

- **언어**: Python 3.8+
- **GUI 프레임워크**: PyQt5
- **PDF 처리**: PyPDF2, pdf2image
- **API 연동**: 
  - OpenAI ChatGPT API (GPT-4.1-mini)
  - Dropbox API

## 3. 주요 기능

### 3.1 입찰 폴더 관리

- 입찰 2025 폴더 내 하위 폴더 리스트 표시
- 입찰 정보 테이블 형태로 보기 (공고명, 등록마감, 추정가격 등)
- 입찰 내용 요약 정보 표시

### 3.2 PDF 문서 분석

- ChatGPT API를 활용한 PDF 문서 내용 분석
- 분석 결과 JSON 형식으로 저장
- 주요 추출 정보:
  - 공고 정보 (등록마감, 공고명, 추정가격)
  - 프로젝트 성격 요약
  - 입찰 관련사항 요약
  - 제안요청서 핵심 사항
  - 입찰 제출 서류
  - 제안요청서 목차

### 3.3 서식 파일 자동 추출

- PDF 문서에서 서식 페이지 자동 식별
- 식별된 서식을 개별 PDF 파일로 저장
- 고해상도(300 DPI) PDF 렌더링 지원
- 서식 파일 목록 트리 형태로 표시
- 분석 결과 JSON 형식으로 저장

### 3.4 PDF 편집기

- PDF 파일 열기 및 페이지 이동
- 텍스트 상자 추가 및 편집
- 고해상도 PDF 렌더링
- 현재 페이지 PNG로 저장

## 4. 데이터 저장 및 관리

### 4.1 로컬 Dropbox 폴더 연동

- Dropbox 폴더 자동 감지 기능
- 로컬 Dropbox 폴더에 직접 파일 저장:
  - 서식 PDF를 공고 폴더 내 "서식" 서브폴더에 저장
  - 분석 결과를 "서식분석결과.json" 파일로 저장

### 4.2 Dropbox API 연동

- Dropbox API를 통한 원격 파일 관리
- 폴더 목록 조회, 파일 다운로드/업로드 기능
- API 오류 시 로컬 저장 방식으로 대체

## 5. 환경 설정

### 5.1 API 키 관리

- .env 파일을 통한 설정 관리:
  - CHATGPT_API_KEY: OpenAI API 키
  - CHATGPT_MODEL: 사용할 GPT 모델 (기본: gpt-4.1-mini)
  - DROPBOX_APP_KEY: Dropbox 앱 키
  - DROPBOX_APP_SECRET: Dropbox 앱 비밀키
  - DROPBOX_ACCESS_TOKEN: Dropbox 액세스 토큰
  - DROPBOX_REFRESH_TOKEN: Dropbox 리프레시 토큰

### 5.2 시스템 요구사항

- **운영체제**: Windows 10+ / macOS / Linux
- **Python**: 3.8 이상
- **필수 라이브러리**:
  - PyQt5
  - PyPDF2
  - pdf2image (선택사항, 고품질 PDF 렌더링용)
  - openai
  - dropbox
  - python-dotenv

## 6. 제한사항 및 고려사항

- ChatGPT API 토큰 한도 조정 필요 (현재 1000 토큰 사용)
- PDF 텍스트 추출 시 스캔된 문서의 경우 텍스트 인식 제한
- 로컬 Dropbox 폴더를 찾지 못할 경우 사용자에게 선택 요청
- pdf2image 라이브러리가 없을 경우 PDF 렌더링 제한 