# Email AI Summarizer - Claude Code 설정

## 프로젝트 개요
Outlook 이메일을 날짜별로 조회하고, Gemini AI로 요약 및 To-Do를 자동 추출하는 데스크탑 앱.

## 기술 스택
- **UI**: PySide6 (Qt)
- **이메일 연동**: pywin32 (Outlook COM)
- **AI**: Google Gemini API (google-genai)
- **Python**: 3.9+

## 파일 구조
```
main.py              # 진입점 - QApplication 생성 및 MainWindow 실행
main_window.py       # UI 전체 (MainWindow, 각종 Dialog, Thread 클래스)
outlook_manager.py   # Outlook COM 연동 (이메일 조회)
ai_processor.py      # Gemini API 호출 및 응답 파싱
.env                 # API 키 관리 (Git 제외)
requirements.txt     # 의존성
```

---

## UI 구조 (화면 기준)

앱은 하나의 윈도우 안에서 3개 화면(Page)이 전환되는 구조다.
수정 요청 시 아래 화면 구성을 기준으로 어디를 바꿔야 하는지 파악한다.

---

### [화면 전환 흐름]
```
Page 1 (계정 선택)
  → [다음 화면으로 이동] → Page 2

Page 2 (메일 조회 + 이메일 TODO)
  → [전체 TODO 관리 →]     → Page 3
  ← [← 계정 다시 선택하기] → Page 1

Page 3 (전체 TODO 관리)
  ← [← 메일 조회로 돌아가기] → Page 2
```

---

### [Page 1] 계정 선택 화면
> `init_page1()` / `stacked_widget` index 0

```
┌─────────────────────────────────────┐
│   환영합니다!                        │
│   조회할 이메일 계정을 선택해주세요.  │
│                                     │
│   접속할 이메일: [콤보박스▼] [다음→] │
│   Gemini API 키: [입력박스        ]  │
└─────────────────────────────────────┘
```

| UI 요소 | 변수명 | 역할 |
|---------|--------|------|
| 계정 콤보박스 | `account_combo` | Outlook 계정 목록 선택 |
| API 키 입력 | `api_key_combo` | Gemini API 키 입력 (편집 가능 콤보) |
| 다음 버튼 | `next_btn` | Page 2로 이동 |

---

### [Page 2] 메일 조회 + 이메일 TODO 화면
> `init_page2()` / `stacked_widget` index 1

```
┌──────────────────────────────────────────────────────────────────┐
│ [← 계정선택] [접속정보] [상태] [날짜▼] [조회] [새메일] [전체TODO→] │  ← 상단 바
├───────────────────────────┬──────────────────────────────────────┤
│ 메일 목록                  │  [전체선택] [완료/해제] [삭제] [전달] │  ← Todo 탭 헤더
│ □ [발신자] 제목            │  ─────────────────────────────────  │
│ □ [발신자] 제목            │  • Todo 항목 1                      │  ← 탭1: To-Do 리스트
│ □ [발신자] 제목            │  • Todo 항목 2                      │
│ ...                       │  • Todo 항목 3                      │
│───────────────────────────│──────────────────────────────────────│
│ 원본 메일 본문             │  (휴지통 탭)                         │  ← 탭2: 휴지통
│ (읽기전용)                 │  삭제된 Todo 항목들                  │
│                           │                                      │
│ [메일 AI 일괄 요약하기]    │                                      │
└───────────────────────────┴──────────────────────────────────────┘
```

**상단 바 요소**
| UI 요소 | 변수명 | 역할 |
|---------|--------|------|
| 계정 선택으로 | `back_btn` | Page 1 이동 |
| 접속 정보 라벨 | `current_account_label` | 현재 계정 표시 |
| 상태 라벨 | `status_label` | 조회/분석 진행 상태 표시 |
| 날짜 선택 | `date_picker` | 날짜 변경 시 자동 조회 |
| 해당 날짜 조회 | `refresh_btn` | 선택 날짜 메일 조회 |
| 새로 고침 | `fetch_new_btn` | 새 메일만 추가 조회 |
| 전체 TODO 관리 → | `goto_page3_btn` | Page 3 이동 (보라) |

**좌측 - 메일 영역**
| UI 요소 | 변수명 | 역할 |
|---------|--------|------|
| 전체 선택 체크박스 | `email_select_all_cb` | 메일 목록 전체 선택 |
| 메일 목록 | `email_list_widget` | 날짜별 메일 목록, 클릭 시 본문 표시 |
| 원본 본문 | `email_viewer` | 선택한 메일 본문 (읽기전용) |
| AI 요약 버튼 | `analyze_btn` | 체크된 메일 일괄 AI 분석 |

**우측 - To-Do 탭 (탭1)**
| UI 요소 | 변수명 | 역할 |
|---------|--------|------|
| 전체 선택 | `active_select_all_cb` | Todo 전체 선택 |
| 완료/해제 버튼 | `complete_selected_btn` | 체크 항목 완료 처리 |
| 삭제 버튼 | `delete_selected_btn` | 체크 항목 휴지통으로 |
| 전체 TODO로 보내기 | `forward_to_p3_btn` | 체크 항목 Page3로 전달 (보라) |
| Todo 목록 | `active_todo_list` | active+completed 항목, 더블클릭=상세 팝업 |

**우측 - 휴지통 탭 (탭2)**
| UI 요소 | 변수명 | 역할 |
|---------|--------|------|
| 전체 선택 | `trash_select_all_cb` | 휴지통 전체 선택 |
| 복구 버튼 | `restore_selected_btn` | 체크 항목 active로 복구 |
| 영구 삭제 | `permanent_delete_btn` | 체크 항목 완전 삭제 |
| 휴지통 목록 | `trash_todo_list` | deleted 항목 |

**팝업 - AI 분석 중 로딩 다이얼로그**
| UI 요소 | 변수명 | 역할 |
|---------|--------|------|
| 진행 상태 텍스트 | `loading_label` | "n/N 분석 중..." 표시 |
| 취소 버튼 | `loading_cancel_btn` | 분석 중단 |

---

### [Page 3] 전체 TODO 관리 화면
> `init_page3()` / `stacked_widget` index 2

```
┌──────────────────────────────────────────────────────────┐
│ [← 돌아가기]  전체 TODO 관리  [우선순위▼] [정렬▼]         │  ← 상단 바
├──────────────────────────────────────────────────────────┤
│  [ 할 일 ]  [ 완료됨 ]  [ 휴지통 ]                        │  ← 탭 헤더
│──────────────────────────────────────────────────────────│
│  [전체선택] [완료처리] [삭제] [날짜/우선순위 편집]          │  ← 할 일 탭 버튼
│  [높음] Todo 텍스트 | 출처: 메일제목 | 마감: 2026-03-25   │
│  [보통] Todo 텍스트 | 출처: 메일제목 | 마감: 없음          │
│  [낮음] Todo 텍스트 | 출처: 메일제목 | 마감: 2026-04-01   │
└──────────────────────────────────────────────────────────┘
```

**상단 바 요소**
| UI 요소 | 변수명 | 역할 |
|---------|--------|------|
| 돌아가기 | `p3_back_btn` | Page 2 이동 |
| 우선순위 필터 | `p3_priority_filter` | 전체/높음/보통/낮음 필터 |
| 정렬 콤보 | `p3_sort_combo` | 등록순/마감일순/우선순위순 |

**탭1: 할 일**
| UI 요소 | 변수명 | 역할 |
|---------|--------|------|
| 전체 선택 | `p3_active_select_all_cb` | 전체 선택 |
| 완료 처리 | `p3_complete_btn` | 체크 항목 → 완료됨 탭으로 |
| 삭제 | `p3_delete_active_btn` | 체크 항목 → 휴지통 탭으로 |
| 날짜/우선순위 편집 | `p3_edit_btn` | 체크 항목 편집 다이얼로그 오픈 |
| 할 일 목록 | `p3_active_list` | forwarded+active 항목, 더블클릭=편집 팝업 |

**탭2: 완료됨**
| UI 요소 | 변수명 | 역할 |
|---------|--------|------|
| 전체 선택 | `p3_completed_select_all_cb` | 전체 선택 |
| 미완료로 되돌리기 | `p3_uncomplete_btn` | 체크 항목 → 할 일 탭으로 |
| 삭제 | `p3_delete_completed_btn` | 체크 항목 → 휴지통으로 |
| 완료 목록 | `p3_completed_list` | forwarded+completed 항목 (취소선+회색) |

**탭3: 휴지통**
| UI 요소 | 변수명 | 역할 |
|---------|--------|------|
| 전체 선택 | `p3_trash_select_all_cb` | 전체 선택 |
| 복구 | `p3_restore_btn` | 체크 항목 → 할 일 탭으로 |
| 영구 삭제 | `p3_perm_delete_btn` | 체크 항목 완전 삭제 |
| 휴지통 목록 | `p3_trash_list` | forwarded+deleted 항목 |

---

### [팝업] TodoDetailDialog — Todo 상세 (Page2 전용)
> Page2 `active_todo_list` 또는 `trash_todo_list` 더블클릭 시 오픈

```
┌─────────────────────────────┐
│ [출처 메일 제목]             │
│ 제목 텍스트 (파란색)         │
│ [지시사항 / 행동할 일]       │
│ Todo 텍스트 (읽기전용)       │
│ [원본 메일 요약]             │
│ 요약 텍스트 (읽기전용)       │
│                             │
│ [완료 처리]  [삭제(휴지통)]  │  ← active/completed 상태
│ (또는)                      │
│ [휴지통에서 복구] [영구삭제] │  ← deleted 상태
└─────────────────────────────┘
```

---

### [팝업] TodoEditDialog — 날짜/우선순위 편집 (Page3 전용)
> Page3 더블클릭 또는 "날짜/우선순위 편집" 버튼 클릭 시 오픈

```
┌──────────────────────────────┐
│ [할 일]                      │
│ Todo 텍스트 (회색 배경)       │
│                              │
│ 우선순위: [높음/보통/낮음 ▼]  │
│ □ 마감일 지정: [날짜선택 ▼]   │
│                              │
│              [저장]  [취소]  │
└──────────────────────────────┘
```

---

### Todo 항목 색상 규칙
| 상태/조건 | 색상 | 적용 위치 |
|-----------|------|-----------|
| active (일반) | 검정 | Page2 |
| active + forwarded | 보라 `#7B1FA2` | Page2 |
| completed | 파랑 + 취소선 | Page2 |
| deleted | 회색 | Page2, Page3 |
| 우선순위 높음 | 빨강 `#D32F2F` | Page3 |
| 우선순위 보통 | 주황 `#E65100` | Page3 |
| 우선순위 낮음 | 파랑 `#1565C0` | Page3 |

---

### Todo 데이터 구조 (self.all_todos 각 항목)
```python
{
    "id":            str,   # uuid4, 고유 식별자
    "text":          str,   # Todo 텍스트
    "summary":       str,   # Gemini AI 요약 (출처 메일)
    "email_subject": str,   # 출처 이메일 제목
    "status":        str,   # "active" | "completed" | "deleted"
    "forwarded":     bool,  # True = Page3 전체 TODO로 전달됨
    "due_date":      str,   # "YYYY-MM-DD" 또는 "" (미설정)
    "priority":      str,   # "높음" | "보통" | "낮음"
}
```

### 수정 요청 시 참고 가이드
| 수정 내용 | 수정 위치 |
|-----------|-----------|
| Page1 화면 변경 | `init_page1()` |
| Page2 레이아웃/버튼 | `init_page2()` |
| Page3 레이아웃/버튼 | `init_page3()` |
| Page2 Todo 동작 | `complete_selected_active`, `delete_selected_active` 등 |
| Page3 Todo 동작 | `p3_complete_selected`, `p3_delete_selected_active` 등 `p3_*` 메서드 |
| Todo 상세 팝업 | `TodoDetailDialog` 클래스 |
| 날짜/우선순위 팝업 | `TodoEditDialog` 클래스 |
| AI 분석 로직 | `ai_processor.py` → `AIProcessor.analyze_email()` |
| 메일 조회 로직 | `outlook_manager.py` → `OutlookManager` |
| Todo 생성 구조 | `on_analyze_success()` |
| Page3 필터/정렬 | `_get_filtered_sorted()` |
| Page3 목록 렌더링 | `refresh_page3()` |

---

## Rules (.agents/rules/coding-standards.md 반영)

1. 코드는 **Python + PySide6** 를 사용한다.
2. 코드 생성 후 반드시 실행하여 정상 동작 여부를 확인한다.
   - 에러 발생 시 원인 분석 → 수정 → 재실행을 반복한다.
   - 모든 에러가 해결된 후에만 작업을 종료한다.
3. API 키는 반드시 `.env` 파일로 관리하며 Git에 커밋하지 않는다.
4. 외부 API 호출(Gemini)은 반드시 `try-except`로 감싼다.
5. 에러 발생 시 `QMessageBox.critical()`로 사용자에게 명확히 표시한다.
6. QFont 사용 시 `pointSize() <= 0` 이면 반드시 `setPointSize(10)`을 먼저 호출한다.
7. UI 스타일은 `pt` 단위 사용 (px 단위 사용 금지 - Qt DPI 스케일링 이슈).

## Workflow (.agents/workflows/python-dev.md 반영)

### 환경 설정
```bash
pip install -r requirements.txt
copy .env.example .env   # API 키 입력
```

### 실행
```bash
python main.py
```

### 기능 테스트 체크리스트
- [ ] Page1: 계정 선택 및 API 키 입력 정상 동작
- [ ] Page2: 날짜 선택 후 이메일 조회 정상 동작
- [ ] 이메일 체크박스 선택 및 전체 선택 동작
- [ ] AI 일괄 요약 분석 및 To-Do 생성
- [ ] 완료/삭제/복구/영구삭제 기능
- [ ] 분석 중 취소 버튼 동작

### Git 커밋
```bash
git add .
git commit -m "설명 메시지"
```

## Skills (skills/ 반영)

### outlook-email-fetching
Outlook COM으로 날짜별 이메일 목록을 가져오는 패턴 → `outlook_manager.py` 참고

### gemini-email-analysis
Gemini API로 이메일 본문을 분석하여 요약 + To-Do 추출하는 패턴 → `ai_processor.py` 참고

### pyside6-thread-pattern
UI 블로킹 없이 백그라운드 작업을 처리하는 QThread 패턴:
```python
class WorkThread(QThread):
    finished = Signal(list)
    error = Signal(str)

    def run(self):
        try:
            result = do_work()
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(str(e))
```
