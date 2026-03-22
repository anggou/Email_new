현재 프로젝트 코드를 CLAUDE.md의 코딩 규칙에 따라 검토하세요:

1. `main_window.py`, `outlook_manager.py`, `ai_processor.py`를 읽는다.
2. 아래 항목을 체크한다:
   - [ ] QFont 사용 시 `pointSize() <= 0` 가드 처리 여부
   - [ ] 폰트 스타일시트가 `pt` 단위 사용 여부 (`px` 금지)
   - [ ] API 호출이 `try-except`로 감싸져 있는지
   - [ ] 장시간 작업이 QThread로 분리됐는지
   - [ ] API 키가 하드코딩되지 않았는지
3. 문제가 있는 항목은 즉시 수정한다.
4. 검토 결과를 요약하여 보고한다.
