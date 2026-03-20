---
description: Mindmap 프로젝트 개발·실행·배포 워크플로우
---

// turbo-all

## 환경 설정

1. 가상환경 생성 및 활성화
```bash
python -m venv .venv
.venv\Scripts\activate
```

2. 의존성 설치
```bash
pip install -r requirements.txt
```

3. 환경변수 설정
```bash
copy .env.example .env
```
- `.env` 파일을 열어 사용할 AI 모델의 API 키를 입력한다:
  - `GEMINI_API_KEY` — Google AI Studio에서 발급
  - `OPENAI_API_KEY` — platform.openai.com에서 발급
  - `ANTHROPIC_API_KEY` — console.anthropic.com에서 발급

## 개발 실행

4. Streamlit 앱 실행
```bash
streamlit run app.py
```

5. 브라우저에서 확인
- http://localhost:8501 에서 앱이 정상 작동하는지 확인한다.
- 사이드바에서 AI 모델 선택 및 API 키 입력 상태를 확인한다.

## 기능 테스트 체크리스트

6. 아래 항목을 순서대로 테스트한다:
- [ ] 텍스트 직접 입력 → 마인드맵 생성
- [ ] PDF 파일 업로드 → 텍스트 추출 → 마인드맵 생성
- [ ] 이미지 파일 업로드 → OCR 또는 Vision → 마인드맵 생성
- [ ] TXT 파일 업로드 → 마인드맵 생성
- [ ] 한국어/English 언어 전환 정상 동작
- [ ] AI 모델 전환 (Gemini ↔ GPT ↔ Claude) 정상 동작
- [ ] .mmd / .md 다운로드 버튼 정상 동작

## Git 커밋 & 푸시

7. 변경사항 커밋
```bash
git add .
git commit -m "설명 메시지"
git push
```
