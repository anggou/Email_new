import logging
import re
import requests
from datetime import datetime

logger = logging.getLogger(__name__)

NOTION_VERSION = "2022-06-28"
MAX_TEXT_LENGTH = 2000
DB_TITLE = "Email AI Summarizer - TODO"
DB_DESCRIPTION = "⚠️ 이 데이터베이스는 Email AI Summarizer 앱과 연동됩니다.\n상태(Status) 컬럼 변경만 앱에 자동 반영됩니다. 텍스트 수정은 앱에 반영되지 않습니다."

STATUS_MAP = {
    "active":    "active",
    "completed": "completed",
    "complete":  "completed",
    "완료":      "completed",
    "deleted":   "deleted",
    "삭제":      "deleted",
}


def _extract_page_id(url_or_id: str) -> str:
    cleaned = url_or_id.strip()
    match = re.search(r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}", cleaned, re.I)
    if match:
        return match.group(0).replace("-", "")
    match = re.search(r"[0-9a-f]{32}", cleaned, re.I)
    if match:
        return match.group(0)
    raise ValueError(f"올바른 Notion 페이지 URL 또는 ID가 아닙니다: {url_or_id}")


class NotionSync:
    def __init__(self, api_key: str, parent_page_id: str):
        if not api_key:
            raise ValueError("Notion API 키가 입력되지 않았습니다.")
        if not parent_page_id:
            raise ValueError("Notion 페이지 URL/ID가 입력되지 않았습니다.")
        self.api_key = api_key.strip()
        self.parent_page_id = _extract_page_id(parent_page_id)
        self._headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Notion-Version": NOTION_VERSION,
            "Content-Type": "application/json",
        }
        self._db_id = None

    def _req(self, method: str, url: str, **kwargs) -> dict:
        resp = requests.request(method, url, headers=self._headers, timeout=10, **kwargs)
        if not resp.ok:
            raise Exception(f"HTTP {resp.status_code}: {resp.text}")
        return resp.json()

    def update_page_status(self, page_id: str, new_status: str):
        """기존 Notion 페이지의 상태 필드만 업데이트합니다."""
        url = f"https://api.notion.com/v1/pages/{page_id}"
        self._req("PATCH", url, json={"properties": {"상태": {"select": {"name": new_status}}}})

    def archive_page(self, page_id: str):
        """Notion 페이지를 아카이브(삭제)합니다."""
        url = f"https://api.notion.com/v1/pages/{page_id}"
        self._req("PATCH", url, json={"archived": True})

    def archive_pages(self, page_ids: list) -> list:
        """여러 Notion 페이지를 아카이브합니다. 실패 메시지 목록 반환."""
        errors = []
        for page_id in page_ids:
            try:
                self.archive_page(page_id)
            except Exception as e:
                logger.warning(f"Notion 페이지 아카이브 실패 ({page_id}): {e}")
                errors.append(str(e))
        return errors

    def _truncate(self, text, limit: int = MAX_TEXT_LENGTH) -> str:
        if not text:
            return ""
        if isinstance(text, list):
            text = " ".join(str(i) for i in text)
        text = str(text)
        return text if len(text) <= limit else text[: limit - 3] + "..."

    # ── DB 자동 생성 / 조회 ─────────────────────────────────────────────────
    def _find_existing_db(self):
        try:
            data = self._req("GET", f"https://api.notion.com/v1/blocks/{self.parent_page_id}/children")
            for block in data.get("results", []):
                if block.get("type") == "child_database":
                    title = block.get("child_database", {}).get("title", "")
                    if title == DB_TITLE:
                        db_id = block["id"].replace("-", "")
                        db_info = self._req("GET", f"https://api.notion.com/v1/databases/{db_id}")
                        if "할 일" in db_info.get("properties", {}):
                            return db_id
        except Exception:
            pass
        return None

    def _create_db(self) -> str:
        payload = {
            "parent": {"type": "page_id", "page_id": self.parent_page_id},
            "title": [{"type": "text", "text": {"content": DB_TITLE}}],
            "description": [{"type": "text", "text": {"content": DB_DESCRIPTION}}],
            "properties": {
                "이메일 제목": {"title": {}},
                "할 일":      {"rich_text": {}},
                "요약":       {"rich_text": {}},
                "저장 날짜":  {"date": {}},
                "상태": {
                    "select": {
                        "options": [
                            {"name": "active",    "color": "blue"},
                            {"name": "completed", "color": "green"},
                            {"name": "deleted",   "color": "red"},
                        ]
                    }
                },
            },
        }
        data = self._req("POST", "https://api.notion.com/v1/databases", json=payload)
        return data["id"].replace("-", "")

    def get_or_create_db(self) -> str:
        if self._db_id:
            return self._db_id
        db_id = self._find_existing_db()
        if not db_id:
            db_id = self._create_db()
        self._db_id = db_id
        return db_id

    # ── 항목 저장 ───────────────────────────────────────────────────────────
    def _save_page(self, db_id: str, email_subject: str, summary: str,
                   todo_text: str, status: str) -> dict:
        payload = {
            "parent": {"database_id": db_id},
            "properties": {
                "이메일 제목": {"title": [{"text": {"content": self._truncate(email_subject, 100)}}]},
                "할 일":      {"rich_text": [{"text": {"content": self._truncate(todo_text)}}]},
                "요약":       {"rich_text": [{"text": {"content": self._truncate(summary)}}]},
                "저장 날짜":  {"date": {"start": datetime.now().strftime("%Y-%m-%d")}},
                "상태":       {"select": {"name": status}},
            },
        }
        return self._req("POST", "https://api.notion.com/v1/pages", json=payload)

    # ── 공개 API ────────────────────────────────────────────────────────────
    def sync_all_todos(self, todos_list: list) -> tuple:
        """
        각 TODO를 개별 행으로 저장.
        반환: (성공수, 오류목록, {todo_app_id: notion_page_id}, db_id)
        """
        db_id = self.get_or_create_db()
        success, failed, id_map = 0, [], {}

        for todo in todos_list:
            subject  = todo.get("email_subject") or "제목 없음"
            summary  = todo.get("summary", "")
            text     = todo.get("text", "")
            status   = todo.get("status", "active")
            todo_id  = todo.get("id", "")
            try:
                resp = self._save_page(db_id, subject, summary, text, status)
                notion_page_id = resp["id"].replace("-", "")
                if todo_id:
                    id_map[todo_id] = notion_page_id
                success += 1
            except Exception as e:
                logger.warning(f"'{subject}' 저장 실패: {e}")
                failed.append(str(e))

        return success, failed, id_map, db_id

    def fetch_status_changes(self, db_id: str, notion_id_to_todo_id: dict) -> dict:
        """
        Notion DB를 조회하여 상태 변경 및 삭제(아카이브) 항목을 반환.
        notion_id_to_todo_id: {notion_page_id: todo_app_id}
        반환: {todo_app_id: new_app_status}  (삭제된 경우 "deleted")
        """
        changes = {}
        try:
            seen_ids = set()
            cursor = None
            while True:
                body = {"page_size": 100}
                if cursor:
                    body["start_cursor"] = cursor
                data = self._req(
                    "POST",
                    f"https://api.notion.com/v1/databases/{db_id}/query",
                    json=body,
                )
                for page in data.get("results", []):
                    page_id = page["id"].replace("-", "")
                    seen_ids.add(page_id)
                    if page_id not in notion_id_to_todo_id:
                        continue
                    select = (page.get("properties", {}).get("상태", {}).get("select")) or {}
                    status_name = select.get("name", "active").lower()
                    app_status = STATUS_MAP.get(status_name, "active")
                    todo_id = notion_id_to_todo_id[page_id]
                    changes[todo_id] = app_status
                if not data.get("has_more"):
                    break
                cursor = data.get("next_cursor")

            # Notion에서 아카이브/삭제된 항목 감지
            for notion_id, todo_id in notion_id_to_todo_id.items():
                if notion_id not in seen_ids:
                    changes[todo_id] = "deleted"
        except Exception as e:
            logger.warning(f"Notion 상태 조회 실패: {e}")
        return changes
