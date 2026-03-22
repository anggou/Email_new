---
name: notion-sync
description: >
  Email AI Summarizer 앱의 분석 결과(요약, To-Do)를 Notion 데이터베이스에 저장/동기화하는 기능을 구현하는 스킬.
  사용자가 "노션에 저장", "Notion에 동기화", "노션 연동", "To-Do를 Notion으로", "send to Notion", "sync to Notion",
  "save to Notion", "Notion MCP 연결", "노션 MCP" 등을 언급할 때 반드시 이 스킬을 사용하세요.
  Notion MCP 서버 설정, API 키 발급, notion_sync.py 모듈 생성, UI 버튼 추가까지 전 과정을 안내합니다.
---

# Notion Sync 스킬

이 스킬은 Email AI Summarizer 앱과 Notion을 연결하는 전체 과정을 담당합니다.
`ai_processor.py`가 생성하는 `summary` / `todos` 데이터를 Notion 데이터베이스 페이지로 자동 저장합니다.

---

## 구현 순서

### 1단계 — Notion MCP 서버 설정

`settings.local.json`에 Notion MCP 서버를 추가합니다.

```json
{
  "mcpServers": {
    "notionApi": {
      "command": "npx",
      "args": ["-y", "@notionhq/notion-mcp-server"],
      "env": {
        "OPENAPI_MCP_HEADERS": "{\"Authorization\": \"Bearer YOUR_NOTION_API_KEY\", \"Notion-Version\": \"2022-06-28\"}"
      }
    }
  }
}
```

> **확인**: 설정 후 `/reload-mcp` 또는 Claude Code 재시작이 필요합니다.
> API 키는 `.env` 파일에도 함께 저장하여 `notion_sync.py`에서도 사용합니다.

---

### 2단계 — Notion API 키 및 데이터베이스 준비

1. **Notion Integration 생성**
   - [https://www.notion.so/my-integrations](https://www.notion.so/my-integrations) 접속
   - "새 통합 만들기" 클릭 → 이름 입력 (예: `Email Summarizer`) → 제출
   - **Internal Integration Token** 복사 (형식: `secret_xxxxxx`)

2. **Notion 데이터베이스 생성 및 연결**
   - Notion에서 새 페이지 생성 → `/database` 입력 → "Full page" 데이터베이스 선택
   - 오른쪽 상단 `...` 메뉴 → "연결 추가" → 방금 만든 Integration 선택
   - 브라우저 URL에서 데이터베이스 ID 복사 (형식: `https://notion.so/workspace/[DATABASE_ID]?v=...`)

3. **데이터베이스 속성 설정** (권장 구조)

   | 속성명 | 타입 |
   |--------|------|
   | 제목 (이메일 제목) | Title |
   | 요약 | Text |
   | To-Do 항목 | Text |
   | 저장 날짜 | Date |
   | 상태 | Select (active / completed / deleted) |

4. **`.env` 파일에 추가**
   ```
   NOTION_API_KEY=secret_xxxxxx
   NOTION_DATABASE_ID=xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
   ```

---

### 3단계 — `notion_sync.py` 모듈 생성

아래 코드를 프로젝트 루트에 `notion_sync.py`로 생성합니다.

```python
import os
import logging
import requests
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

logger = logging.getLogger(__name__)

NOTION_API_URL = "https://api.notion.com/v1/pages"
NOTION_VERSION = "2022-06-28"
MAX_TEXT_LENGTH = 2000  # Notion rich_text 속성 최대 길이


class NotionSync:
    """Notion API를 통해 이메일 분석 결과를 Notion 데이터베이스에 저장합니다."""

    def __init__(self):
        self.api_key = os.getenv("NOTION_API_KEY")
        self.database_id = os.getenv("NOTION_DATABASE_ID")

        if not self.api_key:
            raise ValueError("NOTION_API_KEY가 .env 파일에 설정되지 않았습니다.")
        if not self.database_id:
            raise ValueError("NOTION_DATABASE_ID가 .env 파일에 설정되지 않았습니다.")

    def _truncate(self, text: str, limit: int = MAX_TEXT_LENGTH) -> str:
        """Notion API 제한에 맞게 텍스트를 자릅니다. 잘린 경우 말줄임표를 추가합니다."""
        if len(text) <= limit:
            return text
        return text[: limit - 3] + "..."

    def sync_todo(self, email_subject: str, summary: str, todos: list, status: str = "active") -> dict:
        """
        단일 이메일의 분석 결과를 Notion 데이터베이스 페이지로 생성합니다.

        Args:
            email_subject: 이메일 제목
            summary: Gemini AI가 생성한 요약 텍스트
            todos: To-Do 항목 문자열 리스트
            status: 항목 상태 (active / completed / deleted)

        Returns:
            생성된 Notion 페이지 정보 dict

        Raises:
            Exception: API 호출 실패 시
        """
        todos_text = "\n".join(f"• {t}" for t in todos) if todos else "할 일 없음"

        payload = {
            "parent": {"database_id": self.database_id},
            "properties": {
                "제목": {
                    "title": [{"text": {"content": self._truncate(email_subject, 100)}}]
                },
                "요약": {
                    "rich_text": [{"text": {"content": self._truncate(summary)}}]
                },
                "To-Do 항목": {
                    "rich_text": [{"text": {"content": self._truncate(todos_text)}}]
                },
                "저장 날짜": {
                    "date": {"start": datetime.now().strftime("%Y-%m-%d")}
                },
                "상태": {
                    "select": {"name": status}
                },
            },
        }

        try:
            response = requests.post(
                NOTION_API_URL,
                headers={
                    "Authorization": f"Bearer {self.api_key}",
                    "Notion-Version": "2022-06-28",
                    "Content-Type": "application/json",
                },
                json=payload,
                timeout=10,
            )
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            logger.error(f"Notion API 오류: {e}")
            raise Exception(f"Notion 저장 실패: {e}")

    def sync_all_todos(self, todos_list: list) -> tuple[int, list]:
        """
        all_todos 리스트 전체(또는 선택 항목)를 Notion에 일괄 저장합니다.

        Args:
            todos_list: main_window.py의 all_todos 형식 리스트
                        각 항목: {'id', 'text', 'summary', 'email_subject', 'status'}

        Returns:
            (성공 수, 실패한 항목 subject 리스트)
        """
        success_count = 0
        failed = []

        # email_subject 기준으로 그룹핑하여 하나의 페이지에 모든 todos 저장
        from collections import defaultdict
        grouped = defaultdict(lambda: {"summary": "", "todos": [], "status": "active"})

        for todo in todos_list:
            subject = todo.get("email_subject", "제목 없음")
            grouped[subject]["summary"] = todo.get("summary", "")
            grouped[subject]["todos"].append(todo.get("text", ""))
            grouped[subject]["status"] = todo.get("status", "active")

        for subject, data in grouped.items():
            try:
                self.sync_todo(
                    email_subject=subject,
                    summary=data["summary"],
                    todos=data["todos"],
                    status=data["status"],
                )
                success_count += 1
            except Exception as e:
                logger.warning(f"'{subject}' 저장 실패: {e}")
                failed.append(subject)

        return success_count, failed
```

---

### 4단계 — UI에 "Notion에 저장" 버튼 추가

`main_window.py`의 `init_page2()` 메서드에서 To-Do 탭 영역에 버튼을 추가합니다.

**삽입 위치**: `init_page2()` 내에서 `self.delete_selected_btn`을 추가하는 줄 바로 다음에 추가합니다.

```python
# 기존 코드 (위치 참고용):
#   active_header_layout.addWidget(self.delete_selected_btn)
# ↓ 아래에 추가

self.notion_sync_btn = QPushButton("Notion에 저장")
self.notion_sync_btn.setStyleSheet(
    "background-color: #000000; color: white; font-weight: bold; padding: 0 12pt;"
)
self.notion_sync_btn.clicked.connect(self.sync_to_notion)
active_header_layout.addWidget(self.notion_sync_btn)
```

**`MainWindow` 클래스에 슬롯 메서드 추가**:

```python
def sync_to_notion(self):
    """체크된 To-Do 항목 또는 전체 활성 항목을 Notion에 저장합니다."""
    # 체크된 항목만 수집, 없으면 전체 active 항목
    checked_todos = []
    for i in range(self.active_todo_list.count()):
        item = self.active_todo_list.item(i)
        if item.checkState() == Qt.Checked:
            checked_todos.append(item.data(Qt.UserRole))

    target_todos = checked_todos if checked_todos else [
        t for t in self.all_todos if t["status"] == "active"
    ]

    if not target_todos:
        QMessageBox.information(self, "저장할 항목 없음", "Notion에 저장할 To-Do 항목이 없습니다.")
        return

    try:
        from notion_sync import NotionSync
        syncer = NotionSync()
        success, failed = syncer.sync_all_todos(target_todos)

        msg = f"Notion 저장 완료: {success}개 그룹"
        if failed:
            msg += f"\n실패: {', '.join(failed)}"
        QMessageBox.information(self, "Notion 동기화", msg)

    except ImportError:
        QMessageBox.critical(self, "모듈 없음", "notion_sync.py 파일이 없습니다.\nrequirements.txt에 requests를 추가하고 notion_sync.py를 생성하세요.")
    except ValueError as e:
        QMessageBox.critical(self, "설정 오류", f".env 파일을 확인하세요.\n\n{e}")
    except Exception as e:
        QMessageBox.critical(self, "Notion 오류", f"저장 중 오류가 발생했습니다.\n\n{e}")
```

---

### 5단계 — `requirements.txt` 업데이트

```
requests
```

를 `requirements.txt`에 추가합니다 (아직 없다면).

---

## Notion MCP 도구 활용 (Claude Code에서 직접 조작)

Claude Code에 Notion MCP가 연결된 경우, 코딩 없이 직접 조작도 가능합니다:

```
# 데이터베이스 목록 조회
notion:list_databases

# 특정 데이터베이스에 페이지 생성
notion:create_page  database_id=<ID>  properties={...}

# 페이지 검색
notion:search  query="이메일 제목"
```

---

## 데이터 구조 참고

`main_window.py`의 `all_todos` 각 항목:
```python
{
    "id": "uuid-string",
    "text": "To-Do 항목 텍스트",
    "summary": "Gemini AI 요약 텍스트",
    "email_subject": "원본 이메일 제목",
    "status": "active" | "completed" | "deleted"
}
```

`notion_sync.py`는 `email_subject` 기준으로 그룹핑하여 하나의 Notion 페이지에 같은 이메일의 모든 To-Do를 묶어서 저장합니다.

---

## 문제 해결

| 증상 | 원인 | 해결 |
|------|------|------|
| `401 Unauthorized` | API 키 오류 | `.env`의 `NOTION_API_KEY` 확인 |
| `404 Not Found` | 데이터베이스 ID 오류 또는 Integration 미연결 | DB ID 재확인, Integration 연결 여부 확인 |
| `400 Bad Request` | 속성명 불일치 | Notion DB의 실제 컬럼명과 코드 일치 여부 확인 |
| MCP 도구 없음 | MCP 서버 미설정 | `settings.local.json` 설정 후 재시작 |
