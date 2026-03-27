import multiprocessing
multiprocessing.freeze_support()

import dash
from dash import dcc, html, Input, Output, State, ALL, MATCH, ctx, no_update, callback
import dash_bootstrap_components as dbc
import json
import os
import uuid
from datetime import date, datetime

import logging
import traceback
import webbrowser
import threading
import sys
from firebase_client import FirebaseClient
from dotenv import load_dotenv

logger = logging.getLogger(__name__)

# PyInstaller exe 실행 시 번들된 .env 파일 위치(sys._MEIPASS) 기준으로 로드
if getattr(sys, 'frozen', False):
    _basedir = sys._MEIPASS
else:
    _basedir = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(_basedir, '.env'))

fb_client = FirebaseClient(
    api_key=os.getenv("FIREBASE_API_KEY", ""),
    project_id=os.getenv("FIREBASE_PROJECT_ID", ""),
)


def fb_load_settings(uid: str, id_token: str) -> dict:
    """Firestore에서 프로필 + API 키 불러오기. 없으면 빈 dict 반환."""
    try:
        profile  = fb_client.get_data(uid, id_token, "profile", "main") or {}
        api_keys = fb_client.get_data(uid, id_token, "settings", "keys") or {}
        return {"profile": profile, "api_keys": api_keys}
    except Exception:
        return {"profile": {}, "api_keys": {}}


def fb_save_profile(uid: str, id_token: str, profile: dict):
    try:
        fb_client.save_data(uid, id_token, "profile", "main", profile)
    except Exception as e:
        logger.warning(f"Firestore 프로필 저장 실패: {e}")


def fb_save_keys(uid: str, id_token: str, gemini_key: str, notion_key: str, notion_db: str):
    try:
        fb_client.save_data(uid, id_token, "settings", "keys", {
            "gemini_key": gemini_key,
            "notion_key": notion_key,
            "notion_db":  notion_db,
        })
    except Exception as e:
        logger.warning(f"Firestore 키 저장 실패: {e}")

# ── Threading-based analysis state ───────────────────────────────────────────
import threading as _threading
_analysis = {
    "running": False,
    "cancel": _threading.Event(),
    "progress": 0,
    "text": "",
    "status": "idle",   # idle | running | done | error | cancelled
    "todos": [],
    "errors": [],
    "new_todo_count": 0,
    "total": 0,
}

_ai_reply = {
    "status": "idle",   # idle | running | done | error
    "text": "",         # 생성된 답장 또는 에러 메시지
    "entry_id": "",
}

app = dash.Dash(
    __name__,
    external_stylesheets=[dbc.themes.BOOTSTRAP, dbc.icons.BOOTSTRAP],
    suppress_callback_exceptions=True,
    title="Email AI Summarizer",
    update_title=None,
)
app.index_string = '''<!DOCTYPE html>
<html>
    <head>
        {%metas%}
        <title>{%title%}</title>
        <link rel="icon" type="image/png" href="/assets/icon.png">
        {%css%}
    </head>
    <body>
        {%app_entry%}
        <footer>
            {%config%}
            {%scripts%}
            {%renderer%}
        </footer>
    </body>
</html>'''
server = app.server

# ── Profile helpers ───────────────────────────────────────────────────────────
def build_user_context(profile):
    if not profile or not profile.get("name"):
        return ""
    lines = [
        f"[내 정보]",
        f"- 이름: {profile.get('name','')}",
        f"- 이메일: {profile.get('email','')}",
        f"- 역할/직책: {profile.get('role','')}",
    ]
    if profile.get("projects"):
        lines.append(f"- 참여 프로젝트: {', '.join(profile['projects'])}")
    lines.append("[조직 구조]")
    if profile.get("superiors"):
        lines.append(f"- 상사: {', '.join(profile['superiors'])}")
    if profile.get("peers"):
        lines.append(f"- 동료: {', '.join(profile['peers'])}")
    if profile.get("subordinates"):
        lines.append(f"- 부하직원: {', '.join(profile['subordinates'])}")
    if profile.get("clients"):
        lines.append(f"- 주요 고객/외부인: {', '.join(profile['clients'])}")
    return "\n".join(lines)

def split_comma(s):
    return [x.strip() for x in s.split(",") if x.strip()] if s else []

def _archive_notion_bg(notion_key, notion_db, page_ids):
    """Notion 아카이브를 백그라운드 스레드에서 실행."""
    try:
        from notion_sync import NotionSync
        NotionSync(notion_key, notion_db).archive_pages(page_ids)
    except Exception as e:
        logger.warning(f"Notion 백그라운드 아카이브 실패: {e}")

def _pf_list_widget(field, label, placeholder):
    """프로필 항목 추가/삭제 위젯"""
    return html.Div([
        dbc.Label(label, className="small fw-semibold"),
        html.Div(
            id=f"pf-list-{field}",
            className="border rounded p-1 mb-1 bg-white",
            style={"minHeight": "34px", "maxHeight": "90px", "overflowY": "auto"},
        ),
        dbc.InputGroup([
            dbc.Input(
                id={"type": "pf-input", "field": field},
                placeholder=placeholder, size="sm",
            ),
            dbc.Button("+ 추가", id={"type": "pf-add-btn", "field": field},
                       color="primary", outline=True, size="sm"),
        ], className="mb-2"),
    ])

# ── Page layouts ──────────────────────────────────────────────────────────────
def page0_layout():
    return dbc.Container([
        dbc.Row([
            dbc.Col([
                dbc.Card([
                    dbc.CardBody([
                        html.H3("이메일 요약 앱", className="text-center mb-3 fw-bold"),
                        html.P("계정에 로그인하거나 새로 가입하세요.", className="text-center text-muted small mb-4"),
                        dbc.Input(id="auth-email", type="email", placeholder="이메일", className="mb-2"),
                        dbc.Input(id="auth-password", type="password", placeholder="비밀번호", className="mb-3"),
                        dbc.Button("로그인", id="btn-login", color="primary", className="w-100 mb-2 fw-bold"),
                        dbc.Button("회원가입", id="btn-signup", color="secondary", outline=True, className="w-100"),
                        html.Div(id="auth-error", className="text-danger mt-3 small text-center"),
                    ])
                ], className="card-clean mt-5 shadow-sm"),
            ], md={"size": 6, "offset": 3}, lg={"size": 4, "offset": 4}),
        ]),
    ], fluid=True, className="py-4")

def page1_layout():
    return dbc.Container([
        dbc.Row([
            dbc.Col([
                dbc.Card([
                    dbc.CardBody([
                        html.H4("Email AI Summarizer", className="text-center mb-1 fw-bold"),
                        html.P("Outlook 이메일을 AI로 요약하고 Todo를 자동 추출합니다.",
                               className="text-center text-muted small mb-4"),

                        dbc.Label("이메일 계정", className="fw-semibold"),
                        dcc.Dropdown(id="account-dropdown", placeholder="계정 선택…", className="mb-3"),

                        dbc.Label([
                            "Gemini API 키",
                            html.A(" (키 발급/관리)", href="https://aistudio.google.com/apikey",
                                   target="_blank", className="ms-1 small text-primary"),
                        ], className="fw-semibold"),
                        dbc.InputGroup([
                            dbc.Input(id="api-key-input", type="password", placeholder="API 키 입력…"),
                            dbc.Button("표시", id="btn-toggle-api-key", color="secondary", outline=True, size="sm"),
                        ], className="mb-3"),

                        dbc.Button("Notion 연동 사용 ▾ (선택)", id="btn-toggle-notion",
                                   color="secondary", outline=True, size="sm",
                                   className="mb-2 w-100"),
                        dbc.Collapse(id="notion-collapse", is_open=False, children=[
                            dbc.Card(dbc.CardBody([
                                dbc.Label([
                                    "Notion API 키",
                                    html.A(" (키 발급/관리)", href="https://www.notion.so/my-integrations",
                                           target="_blank", className="ms-1 small text-primary"),
                                ], className="small fw-semibold"),
                                dbc.InputGroup([
                                    dbc.Input(id="notion-key-input", type="password", placeholder="secret_xxxxxx"),
                                    dbc.Button("표시", id="btn-toggle-notion-key", color="secondary", outline=True, size="sm"),
                                ], className="mb-1"),
                                dbc.Input(id="notion-db-input", placeholder="Notion 페이지 URL (DB가 생성될 위치)",
                                          size="sm"),
                            ]), className="bg-light border-0 mb-2"),
                        ]),

                        dbc.Button("내 정보 입력 ▾", id="btn-toggle-profile",
                                   color="secondary", outline=True, size="sm",
                                   className="mb-2 w-100"),

                        dbc.Collapse(id="profile-collapse", is_open=False, children=[
                            dbc.Card(dbc.CardBody([
                                dbc.Row([
                                    dbc.Col([dbc.Label("이름", className="small fw-semibold"),
                                             dbc.Input(id="profile-name", placeholder="홍길동", size="sm")], md=6),
                                    dbc.Col([dbc.Label("이메일 (자동)", className="small fw-semibold"),
                                             dbc.Input(id="profile-email", disabled=True, size="sm")], md=6),
                                ], className="mb-2"),
                                dbc.Label("역할 / 직책", className="small fw-semibold"),
                                dbc.Textarea(id="profile-role", placeholder="예) 선박 AI 엔지니어", rows=2,
                                             className="mb-2", style={"fontSize": "var(--fs-body)"}),
                                _pf_list_widget("projects", "참여 프로젝트", "프로젝트명 입력 후 추가"),
                                _pf_list_widget("superiors", "상사", "이름 입력 후 추가"),
                                _pf_list_widget("peers", "동료", "이름 입력 후 추가"),
                                _pf_list_widget("subordinates", "부하직원", "이름 입력 후 추가"),
                                _pf_list_widget("clients", "주요 고객/외부인", "이름 입력 후 추가"),
                                dbc.Button("저장", id="btn-save-profile", color="success",
                                           size="sm", className="me-2"),
                                html.Span(id="profile-save-status", className="text-success small"),
                            ]), className="bg-light border-0"),
                        ]),

                        html.Hr(className="my-3"),
                        dbc.Button("다음 →", id="btn-next-page2", color="primary", className="w-100 fw-bold"),
                        html.Div(id="page1-error", className="text-danger mt-2 small"),
                    ])
                ], className="card-clean mt-4"),
                html.Div("[ Page 1 — 계정 선택 / API 키 / 프로필 입력 ]",
                         className="text-center text-muted mt-2",
                         style={"fontSize": "var(--fs-xs)", "opacity": "0.5"}),
            ], md={"size": 6, "offset": 3}, lg={"size": 4, "offset": 4}),
        ]),
    ], fluid=True, className="py-4")


def page2_layout():
    return html.Div([
        # ── Header bar ──────────────────────────────────────────────────────
        html.Div([
            dbc.Container([
                dbc.Row([
                    dbc.Col(dbc.Button("← 계정선택", id="btn-back-page1",
                                       color="secondary", outline=True, size="sm"), width="auto"),
                    dbc.Col(html.Span(id="p2-account-info",
                                      className="fw-semibold small align-self-center"), width="auto"),
                    dbc.Col(html.Span(id="p2-status",
                                      className="text-muted small align-self-center"), width="auto"),
                    dbc.Col([
                        dbc.InputGroup([
                            dcc.DatePickerSingle(
                                id="p2-date-picker",
                                date=str(date.today()),
                                display_format="YYYY-MM-DD",
                                style={"zIndex": 9999},
                            ),
                            dbc.Button("조회", id="btn-refresh", color="primary", size="sm"),
                        ]),
                    ], width="auto"),
                    dbc.Col(dbc.Button([
                                "전체 TODO 관리 ",
                                dbc.Badge(id="p3-todo-count-badge", children="", color="light",
                                          text_color="dark", className="me-1",
                                          style={"fontSize": "0.7rem", "display": "none"}),
                                "→",
                            ], id="btn-goto-page3",
                                       color="secondary", size="sm",
                                       style={"backgroundColor": "#7B1FA2", "borderColor": "#7B1FA2",
                                              "color": "white"}),
                            className="ms-auto", width="auto"),
                ], align="center", className="g-2"),
            ], fluid=True),
        ], className="page-header"),

        # ── Main content ────────────────────────────────────────────────────
        dbc.Container([
            dbc.Row([
                # Left: Email list + viewer
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader(html.Span("메일 목록", className="fw-semibold small")),
                        dbc.CardBody([
                            html.Div([
                                dbc.Checkbox(id="email-select-all-cb",
                                              label=html.Span(["전체선택", html.Span(id="email-select-count", className="text-muted ms-1")]),
                                              label_class_name="small ms-1"),
                            ], className="px-2 py-1 border-bottom"),
                            html.Div(id="email-list-container",
                                     className="email-list",
                                     children=[html.Span("이메일을 조회하세요.", className="text-muted small p-2 d-block")]),
                        ], className="p-0"),
                    ], className="card-clean mb-2"),
                    dbc.Card([
                        dbc.CardHeader(html.Span("원본 메일 본문", className="fw-semibold small")),
                        dbc.CardBody([
                            html.Div(id="email-viewer", className="email-viewer",
                                     children="메일을 클릭하면 본문이 표시됩니다."),
                        ], className="p-2"),
                    ], className="card-clean mb-2"),
                    dbc.Button("메일 AI 일괄 요약하기", id="btn-analyze",
                               color="primary", size="sm", className="w-100"),
                ], md=5),

                # Right: Todo tabs
                dbc.Col([
                    dbc.Card([
                        dbc.CardBody([
                            dbc.Tabs([
                                dbc.Tab(
                                    html.Div([
                                        dbc.Row([
                                            dbc.Col(
                                                dbc.Checkbox(id="todo-active-select-all",
                                                             label=html.Span(["전체선택", html.Span(id="p2-active-select-count", className="text-muted ms-1")]),
                                                             label_class_name="small ms-1",
                                                             value=False),
                                                width="auto",
                                            ),
                                            dbc.Col([
                                                dbc.Button("삭제", id="btn-todo-delete", size="sm",
                                                           color="danger", outline=True, className="me-1"),
                                                dbc.Button("전체 TODO로 전달", id="btn-todo-forward",
                                                           size="sm",
                                                           style={"backgroundColor": "#7B1FA2",
                                                                  "borderColor": "#7B1FA2",
                                                                  "color": "#fff"}),
                                            ], className="ms-auto", width="auto"),
                                        ], align="center", className="mb-2 g-1"),
                                        html.Div(id="todo-list-active", className="todo-list-container"),
                                    ]),
                                    label="임시 TO-DO", tab_id="tab-active",
                                ),
                                dbc.Tab(
                                    html.Div([
                                        dbc.Row([
                                            dbc.Col(
                                                dbc.Checkbox(id="todo-trash-select-all",
                                                             label=html.Span(["전체선택", html.Span(id="p2-trash-select-count", className="text-muted ms-1")]),
                                                             label_class_name="small ms-1",
                                                             value=False),
                                                width="auto",
                                            ),
                                            dbc.Col([
                                                dbc.Button("복구", id="btn-todo-restore", size="sm",
                                                           color="secondary", outline=True, className="me-1"),
                                                dbc.Button("영구삭제", id="btn-todo-perm-delete", size="sm",
                                                           color="danger", outline=True),
                                            ], className="ms-auto", width="auto"),
                                        ], align="center", className="mb-2 g-1"),
                                        html.Div(id="todo-list-trash", className="todo-list-container"),
                                    ]),
                                    label="휴지통", tab_id="tab-trash",
                                ),
                            ], id="todo-tabs", active_tab="tab-active"),
                        ], className="p-2"),
                    ], className="card-clean"),
                ], md=7),
            ]),
        ], fluid=True),

        # ── Analyze modal ────────────────────────────────────────────────────
        dbc.Modal([
            dbc.ModalHeader(dbc.ModalTitle(html.Span(id="analyze-modal-title", children="AI 분석 중…"))),
            dbc.ModalBody([
                html.Div(id="analyze-progress-wrap", children=[
                    dbc.Progress(id="analyze-progress-bar", value=0, striped=True, animated=True,
                                 className="mb-3"),
                    html.P(id="analyze-progress-text", className="text-center text-muted small mb-3"),
                    dbc.Button("취소", id="btn-analyze-cancel", color="danger", outline=True,
                               size="sm", className="d-block mx-auto"),
                ]),
                html.Div(id="analyze-result-wrap", style={"display": "none"}, children=[
                    html.Div(id="analyze-result-text", className="text-center mb-3"),
                    dbc.Button("확인", id="btn-analyze-confirm", color="primary",
                               size="sm", className="d-block mx-auto px-4"),
                ]),
            ]),
        ], id="analyze-modal", is_open=False, backdrop="static", centered=True),

        html.Div("[ Page 2 — 메일 조회 / AI 요약 / 이메일 TODO ]",
                 style={"textAlign": "center", "fontSize": "var(--fs-xs)",
                        "opacity": "0.5", "padding": "6px 0", "color": "#aaa"}),

        # ── Toast notification ────────────────────────────────────────────────
        dbc.Toast(
            id="p2-toast",
            header="알림",
            is_open=False,
            dismissable=True,
            duration=3000,
            style={"position": "fixed", "bottom": 16, "left": 16, "zIndex": 9999,
                   "maxWidth": "500px", "whiteSpace": "pre-wrap"},
        ),
    ])


def page3_layout():
    return html.Div([
        # ── Header bar ──────────────────────────────────────────────────────
        html.Div([
            dbc.Container([
                dbc.Row([
                    dbc.Col(dbc.Button("← 메일 조회", id="btn-back-page2",
                                       color="secondary", outline=True, size="sm"), width="auto"),
                    dbc.Col(html.H6("전체 TODO 관리", className="mb-0 fw-bold align-self-center"),
                            width="auto"),
                    dbc.Col([
                        dcc.Dropdown(
                            id="p3-priority-filter",
                            options=[
                                {"label": "전체", "value": "all"},
                                {"label": "높음", "value": "높음"},
                                {"label": "보통", "value": "보통"},
                                {"label": "낮음", "value": "낮음"},
                            ],
                            value="all",
                            clearable=False,
                            style={"minWidth": 100, "fontSize": "var(--fs-body)"},
                        ),
                    ], width="auto"),
                    dbc.Col([
                        dcc.Dropdown(
                            id="p3-sort-combo",
                            options=[
                                {"label": "등록순", "value": "default"},
                                {"label": "마감일순", "value": "due_date"},
                                {"label": "우선순위순", "value": "priority"},
                            ],
                            value="default",
                            clearable=False,
                            style={"minWidth": 110, "fontSize": "var(--fs-body)"},
                        ),
                    ], width="auto"),
                ], align="center", className="g-2"),
            ], fluid=True),
        ], className="page-header"),

        # ── Main content ────────────────────────────────────────────────────
        dbc.Container([
            dbc.Card([
                dbc.CardBody([
                    dbc.Tabs([
                        dbc.Tab(
                            html.Div([
                                dbc.Row([
                                    dbc.Col(dbc.Checkbox(id="p3-active-select-all",
                                                          label=html.Span(["전체선택", html.Span(id="p3-active-select-count", className="text-muted ms-1")]),
                                                          label_class_name="small ms-1"), width="auto"),
                                    dbc.Col([
                                        dbc.Button("완료 처리", id="p3-btn-complete", size="sm",
                                                   color="success", outline=True, className="me-1"),
                                        dbc.Button("삭제", id="p3-btn-delete", size="sm",
                                                   color="danger", outline=True, className="me-1"),
                                        dbc.Button("편집", id="p3-btn-edit", size="sm",
                                                   color="primary", outline=True, className="me-1"),
                                        dbc.Button(
                                            html.Img(src="/assets/notion.png",
                                                     style={"width": "18px", "height": "18px"}),
                                            id="p3-btn-notion-sync", size="sm",
                                            style={"backgroundColor": "transparent", "borderColor": "#ccc",
                                                   "padding": "3px 7px"},
                                            className="me-1",
                                        ),
                                        html.Span(id="p3-notion-sync-status",
                                                  className="text-muted small align-self-center",
                                                  style={"fontSize": "var(--fs-meta)"}),
                                    ], className="ms-auto", width="auto"),
                                ], align="center", className="mb-2 g-1"),
                                html.Div(id="p3-active-list", className="todo-list-container"),
                            ]),
                            label="할 일", tab_id="p3-tab-active",
                        ),
                        dbc.Tab(
                            html.Div([
                                dbc.Row([
                                    dbc.Col(dbc.Checkbox(id="p3-completed-select-all",
                                                          label=html.Span(["전체선택", html.Span(id="p3-completed-select-count", className="text-muted ms-1")]),
                                                          label_class_name="small ms-1"), width="auto"),
                                    dbc.Col([
                                        dbc.Button("미완료로", id="p3-btn-uncomplete", size="sm",
                                                   color="secondary", outline=True, className="me-1"),
                                        dbc.Button("삭제", id="p3-btn-del-completed", size="sm",
                                                   color="danger", outline=True),
                                    ], className="ms-auto", width="auto"),
                                ], align="center", className="mb-2 g-1"),
                                html.Div(id="p3-completed-list", className="todo-list-container"),
                            ]),
                            label="완료됨", tab_id="p3-tab-completed",
                        ),
                        dbc.Tab(
                            html.Div([
                                dbc.Row([
                                    dbc.Col(dbc.Checkbox(id="p3-trash-select-all",
                                                          label=html.Span(["전체선택", html.Span(id="p3-trash-select-count", className="text-muted ms-1")]),
                                                          label_class_name="small ms-1"), width="auto"),
                                    dbc.Col([
                                        dbc.Button("복구", id="p3-btn-restore", size="sm",
                                                   color="success", outline=True, className="me-1"),
                                        dbc.Button("영구 삭제", id="p3-btn-perm-delete", size="sm",
                                                   color="danger", outline=True),
                                    ], className="ms-auto", width="auto"),
                                ], align="center", className="mb-2 g-1"),
                                html.Div(id="p3-trash-list", className="todo-list-container"),
                            ]),
                            label="휴지통", tab_id="p3-tab-trash",
                        ),
                    ], id="p3-tabs", active_tab="p3-tab-active"),
                ]),
            ], className="card-clean"),
        ], fluid=True),

        # ── Edit modal ────────────────────────────────────────────────────────
        dbc.Modal([
            dbc.ModalHeader(dbc.ModalTitle("Todo 편집")),
            dbc.ModalBody([
                dbc.Label("할 일", className="small fw-semibold"),
                dbc.Textarea(id="edit-todo-text", rows=3,
                              style={"background": "#f5f5f5", "fontSize": "var(--fs-body)"},
                              className="mb-3"),
                dbc.Label("우선순위", className="small fw-semibold"),
                dbc.RadioItems(
                    id="edit-todo-priority",
                    options=[
                        {"label": "높음", "value": "높음"},
                        {"label": "보통", "value": "보통"},
                        {"label": "낮음", "value": "낮음"},
                    ],
                    value="보통",
                    inline=True,
                    className="mb-3",
                ),
                dbc.Checkbox(id="edit-due-date-cb", label="마감일 지정", className="mb-1"),
                dcc.DatePickerSingle(id="edit-due-date", display_format="YYYY-MM-DD",
                                     style={"display": "none"}),
            ]),
            dbc.ModalFooter([
                dbc.Button("저장", id="btn-edit-save", color="primary", className="me-2"),
                dbc.Button("취소", id="btn-edit-cancel", color="secondary", outline=True),
            ]),
        ], id="edit-modal", is_open=False, centered=True),

        dcc.Store(id="edit-target-id"),

        # ── AI 답장 모달 ────────────────────────────────────────────────────────
        dbc.Modal([
            dbc.ModalHeader(dbc.ModalTitle(html.Span(id="ai-reply-modal-title", children="AI 답장 생성 중…"))),
            dbc.ModalBody([
                html.Div(id="ai-reply-loading-wrap", children=[
                    dbc.Spinner(size="sm", color="primary"),
                    html.Span("Gemini AI가 답장을 작성하고 있습니다…", className="text-muted small ms-2"),
                ]),
                html.Div(id="ai-reply-result-wrap", style={"display": "none"}, children=[
                    dbc.Label("생성된 답장 초안 (수정 후 Outlook에서 열기)", className="small fw-semibold mb-1"),
                    dbc.Textarea(id="ai-reply-text", rows=12,
                                 style={"fontSize": "var(--fs-body)", "fontFamily": "inherit"}),
                ]),
            ]),
            dbc.ModalFooter([
                dbc.Button("Outlook에서 열기", id="btn-ai-reply-open-outlook",
                           color="primary", className="me-2", style={"display": "none"}),
                dbc.Button("닫기", id="btn-ai-reply-close", color="secondary", outline=True),
            ]),
        ], id="ai-reply-modal", is_open=False, backdrop="static", centered=True, size="lg"),
        dcc.Store(id="store-ai-reply-entry-id", data=""),
        dcc.Store(id="store-ai-reply-done", data=0),
        dcc.Interval(id="ai-reply-interval", interval=500, n_intervals=0, disabled=True),

        html.Div("[ Page 3 — 전체 TODO 관리 / Notion 동기화 ]",
                 style={"textAlign": "center", "fontSize": "var(--fs-xs)",
                        "opacity": "0.5", "padding": "6px 0", "color": "#aaa"}),

        dbc.Toast(
            id="p3-toast",
            header="알림",
            is_open=False,
            dismissable=True,
            duration=3000,
            color="primary",
            style={"position": "fixed", "bottom": 16, "left": 16, "zIndex": 9999},
        ),
    ])


# ── Main app layout ───────────────────────────────────────────────────────────
app.layout = html.Div([
    # Stores
    dcc.Store(id="store-page", data=0),
    dcc.Store(id="store-auth-token", storage_type="local", data=""),
    dcc.Store(id="store-uid", storage_type="local", data=""),
    dcc.Store(id="store-account", data=""),
    dcc.Store(id="store-api-key", data=""),
    dcc.Store(id="store-user-profile", data={}),
    dcc.Store(id="store-emails", data=[]),
    dcc.Store(id="store-todos", storage_type="local", data=[]),
    dcc.Store(id="store-selected-email-idx", data=None),
    dcc.Store(id="store-email-checked", data=[]),
    dcc.Store(id="store-todo-checked-p2", data=[]),
    dcc.Store(id="store-highlighted-emails", data=[]),
    dcc.Store(id="store-new-entry-ids", data=[]),   # 현재 세션에서 NEW 표시할 entry_id 목록
    dcc.Store(id="store-seen-by-date", data={}),    # {date: [entry_ids]} 이전 조회에서 본 메일
    dcc.Store(id="store-todo-trash-checked-p2", data=[]),
    dcc.Store(id="store-todo-checked-p3-active", data=[]),
    dcc.Store(id="store-todo-checked-p3-completed", data=[]),
    dcc.Store(id="store-todo-checked-p3-trash", data=[]),
    dcc.Store(id="store-analyze-done", data=0),
    dcc.Interval(id="analyze-interval", interval=500, n_intervals=0, disabled=True),
    dcc.Store(id="store-profile-lists", data={"projects":[],"superiors":[],"peers":[],"subordinates":[],"clients":[]}),
    dcc.Store(id="store-todos-p3", storage_type="local", data=[]),
    dcc.Store(id="store-notion-enabled", data=False),
    dcc.Store(id="store-notion-key", data=""),
    dcc.Store(id="store-notion-db", data=""),
    dcc.Store(id="store-notion-db-id", data=""),
    dcc.Store(id="store-perm-deleted-ids", storage_type="local", data=[]),
    dcc.Store(id="store-notion-archive-queue", storage_type="local", data=[]),
    dcc.Interval(id="notion-poll-interval", interval=60*1000, n_intervals=0, disabled=False),

    # 로그인 로딩 테두리 오버레이
    html.Div(id="login-loading-border"),

    # Pages
    html.Div(id="page0-container", children=page0_layout()),
    html.Div(id="page1-container", style={"display": "none"}, children=page1_layout()),
    html.Div(id="page2-container", style={"display": "none"}, children=page2_layout()),
    html.Div(id="page3-container", style={"display": "none"}, children=page3_layout()),
])


# ════════════════════════════════════════════════════════════════════════════
# CALLBACKS
# ════════════════════════════════════════════════════════════════════════════

# ── Page visibility ───────────────────────────────────────────────────────────
@app.callback(
    Output("page0-container", "style"),
    Output("page1-container", "style"),
    Output("page2-container", "style"),
    Output("page3-container", "style"),
    Output("login-loading-border", "className"),
    Input("store-page", "data"),
)
def toggle_pages(page):
    show = {"display": "block"}
    hide = {"display": "none"}
    # 페이지 전환 완료 시 로딩 테두리 숨김
    border_class = "" if page != 0 else ""
    return (
        show if page == 0 else hide,
        show if page == 1 else hide,
        show if page == 2 else hide,
        show if page == 3 else hide,
        "",  # 페이지 전환되면 active 클래스 제거
    )


# ── 로그인 버튼 클릭 시 로딩 테두리 표시 ─────────────────────────────────────
@app.callback(
    Output("login-loading-border", "className", allow_duplicate=True),
    Input("btn-login", "n_clicks"),
    Input("btn-signup", "n_clicks"),
    prevent_initial_call=True,
)
def show_login_loading(n_login, n_signup):
    return "active"


# ── Auth Logic ────────────────────────────────────────────────────────────────
@app.callback(
    Output("store-auth-token", "data"),
    Output("store-uid", "data"),
    Output("store-page", "data", allow_duplicate=True),
    Output("auth-error", "children"),
    Output("profile-name", "value", allow_duplicate=True),
    Output("profile-role", "value", allow_duplicate=True),
    Output("store-profile-lists", "data", allow_duplicate=True),
    Output("api-key-input", "value", allow_duplicate=True),
    Output("notion-key-input", "value", allow_duplicate=True),
    Output("notion-db-input", "value", allow_duplicate=True),
    Input("btn-login", "n_clicks"),
    Input("btn-signup", "n_clicks"),
    State("auth-email", "value"),
    State("auth-password", "value"),
    prevent_initial_call=True,
)
def handle_auth(n_login, n_signup, email, password):
    triggered = ctx.triggered_id
    empty = (no_update,) * 6  # profile/key 필드 no_update
    if not email or not password:
        return no_update, no_update, no_update, "이메일과 비밀번호를 모두 입력하세요.", *empty

    try:
        if triggered == "btn-login":
            res = fb_client.sign_in(email, password)
        elif triggered == "btn-signup":
            res = fb_client.sign_up(email, password)
        else:
            return no_update, no_update, no_update, "", *empty

        id_token = res.get("idToken")
        uid      = res.get("localId")

        # Firestore에서 저장된 데이터 불러오기
        saved    = fb_load_settings(uid, id_token)
        profile  = saved.get("profile", {})
        api_keys = saved.get("api_keys", {})

        pf_lists = {
            "projects":     profile.get("projects", []),
            "superiors":    profile.get("superiors", []),
            "peers":        profile.get("peers", []),
            "subordinates": profile.get("subordinates", []),
            "clients":      profile.get("clients", []),
        }

        return (
            id_token, uid, 1, "",
            profile.get("name", ""),
            profile.get("role", ""),
            pf_lists,
            api_keys.get("gemini_key", ""),
            api_keys.get("notion_key", ""),
            api_keys.get("notion_db",  ""),
        )
    except Exception as e:
        return no_update, no_update, no_update, str(e), *empty


# ── Page 1: Load accounts on startup ─────────────────────────────────────────
@app.callback(
    Output("account-dropdown", "options"),
    Output("account-dropdown", "value"),
    Input("store-page", "data"),
)
def load_accounts(page):
    if page != 1:
        return no_update, no_update
    try:
        from outlook_manager import OutlookManager
        mgr = OutlookManager()
        accounts = mgr.get_accounts()
        options = [{"label": a, "value": a} for a in accounts]
        val = accounts[0] if accounts else None
        return options, val
    except Exception as e:
        return [{"label": f"오류: {e}", "value": ""}], ""


# ── Page 1: Toggle API key visibility ────────────────────────────────────────
@app.callback(
    Output("api-key-input", "type"),
    Output("btn-toggle-api-key", "children"),
    Input("btn-toggle-api-key", "n_clicks"),
    State("api-key-input", "type"),
    prevent_initial_call=True,
)
def toggle_api_key_visibility(n, current_type):
    if current_type == "password":
        return "text", "숨기기"
    return "password", "표시"


@app.callback(
    Output("notion-key-input", "type"),
    Output("btn-toggle-notion-key", "children"),
    Input("btn-toggle-notion-key", "n_clicks"),
    State("notion-key-input", "type"),
    prevent_initial_call=True,
)
def toggle_notion_key_visibility(n, current_type):
    if current_type == "password":
        return "text", "숨기기"
    return "password", "표시"


# ── Page 1: Toggle Notion collapse ───────────────────────────────────────────
@app.callback(
    Output("notion-collapse", "is_open"),
    Output("btn-toggle-notion", "children"),
    Input("btn-toggle-notion", "n_clicks"),
    State("notion-collapse", "is_open"),
    prevent_initial_call=True,
)
def toggle_notion_collapse(n, is_open):
    if not is_open:
        return True, "Notion 연동 사용 ▴ (선택)"
    return False, "Notion 연동 사용 ▾ (선택)"


# ── Page 1: Toggle profile collapse ───────────────────────────────────────────
@app.callback(
    Output("profile-collapse", "is_open"),
    Input("btn-toggle-profile", "n_clicks"),
    State("profile-collapse", "is_open"),
    prevent_initial_call=True,
)
def toggle_profile(n, is_open):
    return not is_open


# ── Page 1: Load profile when account changes ────────────────────────────────
@app.callback(
    Output("profile-email", "value"),
    Output("profile-name", "value"),
    Output("profile-role", "value"),
    Output("store-profile-lists", "data"),
    Output("api-key-input", "value"),
    Output("notion-key-input", "value"),
    Output("notion-db-input", "value"),
    Output("store-todos", "data", allow_duplicate=True),
    Output("store-todos-p3", "data", allow_duplicate=True),
    Output("store-user-profile", "data", allow_duplicate=True),
    Output("store-seen-by-date", "data", allow_duplicate=True),
    Input("account-dropdown", "value"),
    State("store-uid", "data"),
    State("store-auth-token", "data"),
    State("store-todos", "data"),
    State("store-todos-p3", "data"),
    prevent_initial_call=True,
)
def load_profile(account, uid, id_token, current_todos, current_todos_p3):
    p = {}
    keys = {}
    todos_data = no_update
    todos_p3_data = no_update
    seen_by_date_data = no_update
    if uid and id_token and account:
        safe_acc = account.replace(".", "_")
        try:
            cloud_p = fb_client.get_data(uid, id_token, "profiles", safe_acc)
            if cloud_p:
                p = cloud_p
            else:
                # 구버전 마이그레이션: profile/main에서 이 계정 데이터 로드
                cloud_p_main = fb_client.get_data(uid, id_token, "profile", "main")
                if cloud_p_main and cloud_p_main.get("email") == account:
                    p = cloud_p_main
                    # 새 경로로 자동 마이그레이션
                    try:
                        fb_client.save_data(uid, id_token, "profiles", safe_acc, p)
                    except Exception:
                        pass

            cloud_keys = fb_client.get_data(uid, id_token, "keys", safe_acc)
            if cloud_keys:
                keys = cloud_keys
            else:
                # 구버전 마이그레이션: settings/keys에서 로드
                cloud_keys_main = fb_client.get_data(uid, id_token, "settings", "keys")
                if cloud_keys_main:
                    keys = cloud_keys_main
                    # 새 경로로 자동 마이그레이션
                    try:
                        fb_client.save_data(uid, id_token, "keys", safe_acc, {
                            "gemini": cloud_keys_main.get("gemini_key", ""),
                            "notion_key": cloud_keys_main.get("notion_key", ""),
                            "notion_db":  cloud_keys_main.get("notion_db", ""),
                        })
                    except Exception:
                        pass
            # 로컬에 데이터가 없을 때만 Firebase에서 로드 (로컬 우선)
            if not current_todos:
                cloud_todos = fb_client.get_data(uid, id_token, "todos", safe_acc)
                if cloud_todos and "todos" in cloud_todos: todos_data = cloud_todos["todos"]
            if not current_todos_p3:
                cloud_todos_p3 = fb_client.get_data(uid, id_token, "todos-p3", safe_acc)
                if cloud_todos_p3 and "todos" in cloud_todos_p3: todos_p3_data = cloud_todos_p3["todos"]
            cloud_seen = fb_client.get_data(uid, id_token, "seen-by-date", safe_acc)
            if cloud_seen and "data" in cloud_seen:
                seen_by_date_data = cloud_seen["data"]
        except Exception as e:
            print(f"Firebase 로드 실패: {e}")

    if not p:
        p = {}

    return (
        account or "",
        p.get("name", ""),
        p.get("role", ""),
        {
            "projects":     p.get("projects", []),
            "superiors":    p.get("superiors", []),
            "peers":        p.get("peers", []),
            "subordinates": p.get("subordinates", []),
            "clients":      p.get("clients", []),
        },
        keys.get("gemini", ""),
        keys.get("notion_key", ""),
        keys.get("notion_db", ""),
        todos_data,
        todos_p3_data,
        p if p else no_update,
        seen_by_date_data,
    )


# ── Page 1: Save profile ─────────────────────────────────────────────────────
@app.callback(
    Output("profile-save-status", "children"),
    Output("store-user-profile", "data", allow_duplicate=True),
    Input("btn-save-profile", "n_clicks"),
    State("account-dropdown", "value"),
    State("profile-name", "value"),
    State("profile-role", "value"),
    State("store-profile-lists", "data"),
    State("store-uid", "data"),
    State("store-auth-token", "data"),
    prevent_initial_call=True,
)
def save_profile(n, account, name, role, lists_data, uid, id_token):
    if not account:
        return "계정을 먼저 선택하세요.", no_update
    lists_data = lists_data or {}
    profile_data = {
        "name": name or "",
        "email": account,
        "role": role or "",
        "projects":     lists_data.get("projects", []),
        "superiors":    lists_data.get("superiors", []),
        "peers":        lists_data.get("peers", []),
        "subordinates": lists_data.get("subordinates", []),
        "clients":      lists_data.get("clients", []),
    }

    if uid and id_token:
        safe_acc = account.replace(".", "_")
        try:
            fb_client.save_data(uid, id_token, "profiles", safe_acc, profile_data)
        except Exception as e:
            logger.warning(f"Firestore profiles 저장 실패: {e}")

    return "저장됨 ✓", profile_data


# ── Page 1: Profile list add / delete / render ───────────────────────────────
_PF_FIELDS = ["projects", "superiors", "peers", "subordinates", "clients"]

@app.callback(
    Output("store-profile-lists", "data", allow_duplicate=True),
    Output({"type": "pf-input", "field": ALL}, "value"),
    Input({"type": "pf-add-btn", "field": ALL}, "n_clicks"),
    Input({"type": "pf-input", "field": ALL}, "n_submit"),
    State({"type": "pf-input", "field": ALL}, "value"),
    State("store-profile-lists", "data"),
    prevent_initial_call=True,
)
def add_profile_item(n_clicks_list, n_submit_list, input_values, lists_data):
    triggered = ctx.triggered_id
    if not triggered:
        return no_update, no_update
    field = triggered.get("field") if isinstance(triggered, dict) else None
    if not field or field not in _PF_FIELDS:
        return no_update, no_update
    field_idx = _PF_FIELDS.index(field)
    value = (input_values or [None] * len(_PF_FIELDS))[field_idx]
    if not value or not value.strip():
        return no_update, no_update
    lists_data = dict(lists_data or {f: [] for f in _PF_FIELDS})
    current = list(lists_data.get(field, []))
    if value.strip() not in current:
        current.append(value.strip())
    lists_data[field] = current
    new_inputs = list(input_values or [""] * len(_PF_FIELDS))
    new_inputs[field_idx] = ""
    return lists_data, new_inputs


@app.callback(
    Output("store-profile-lists", "data", allow_duplicate=True),
    Input({"type": "pf-del", "field": ALL, "index": ALL}, "n_clicks"),
    State("store-profile-lists", "data"),
    prevent_initial_call=True,
)
def delete_profile_item(n_clicks_list, lists_data):
    triggered = ctx.triggered_id
    if not triggered or not isinstance(triggered, dict):
        return no_update
    if not any(n for n in n_clicks_list if n):
        return no_update
    field = triggered.get("field")
    index = triggered.get("index")
    if field not in _PF_FIELDS:
        return no_update
    lists_data = dict(lists_data or {f: [] for f in _PF_FIELDS})
    current = list(lists_data.get(field, []))
    if 0 <= index < len(current):
        current.pop(index)
    lists_data[field] = current
    return lists_data


def _render_pf_list(field, items):
    if not items:
        return html.Span("항목 없음", className="text-muted", style={"fontSize": "var(--fs-meta)"})
    return [
        html.Span([
            dbc.Badge(item, color="secondary", className="fw-normal me-1",
                      style={"fontSize": "var(--fs-badge)"}),
            html.Span("×", id={"type": "pf-del", "field": field, "index": i},
                      n_clicks=0,
                      style={"cursor": "pointer", "color": "#999", "fontSize": "var(--fs-body)",
                             "marginRight": "6px", "userSelect": "none"},
                      className="pf-del-x"),
        ], className="d-inline-flex align-items-center mb-1")
        for i, item in enumerate(items)
    ]


@app.callback(
    Output("pf-list-projects", "children"),
    Output("pf-list-superiors", "children"),
    Output("pf-list-peers", "children"),
    Output("pf-list-subordinates", "children"),
    Output("pf-list-clients", "children"),
    Input("store-profile-lists", "data"),
)
def render_profile_lists(data):
    data = data or {f: [] for f in _PF_FIELDS}
    return (
        _render_pf_list("projects",     data.get("projects", [])),
        _render_pf_list("superiors",    data.get("superiors", [])),
        _render_pf_list("peers",        data.get("peers", [])),
        _render_pf_list("subordinates", data.get("subordinates", [])),
        _render_pf_list("clients",      data.get("clients", [])),
    )


# ── Page 1: Go to Page 2 ──────────────────────────────────────────────────────
@app.callback(
    Output("store-page", "data"),
    Output("store-account", "data"),
    Output("store-api-key", "data"),
    Output("store-user-profile", "data"),
    Output("store-notion-key", "data"),
    Output("store-notion-db", "data"),
    Output("store-notion-enabled", "data"),
    Output("page1-error", "children"),
    Input("btn-next-page2", "n_clicks"),
    State("account-dropdown", "value"),
    State("api-key-input", "value"),
    State("notion-key-input", "value"),
    State("notion-db-input", "value"),
    State("notion-collapse", "is_open"),
    State("store-uid", "data"),
    State("store-auth-token", "data"),
    State("store-user-profile", "data"),
    prevent_initial_call=True,
)
def go_to_page2(n, account, api_key, notion_key, notion_db, notion_open, uid, id_token, user_profile):
    if not account:
        return no_update, no_update, no_update, no_update, no_update, no_update, no_update, "계정을 선택해주세요."
    if not api_key or api_key.strip() == "":
        return no_update, no_update, no_update, no_update, no_update, no_update, no_update, "Gemini API 키를 입력해주세요."

    notion_enabled = bool(notion_open and notion_key and notion_key.strip())

    if uid and id_token:
        safe_acc = account.replace(".", "_")
        try:
            fb_client.save_data(uid, id_token, "keys", safe_acc, {
                "gemini": api_key.strip(),
                "notion_key": (notion_key or "").strip(),
                "notion_db":  (notion_db or "").strip(),
            })
        except Exception as e:
            logger.warning(f"Firestore 키 저장 실패: {e}")

    return 2, account, api_key.strip(), user_profile or {}, (notion_key or "").strip(), (notion_db or "").strip(), notion_enabled, ""


# ── Page 2: Update account info label ────────────────────────────────────────
@app.callback(
    Output("p2-account-info", "children"),
    Input("store-account", "data"),
)
def update_account_info(account):
    return f"계정: {account}" if account else ""


# ── Page 2: Fetch emails ──────────────────────────────────────────────────────
@app.callback(
    Output("store-emails", "data"),
    Output("p2-status", "children"),
    Output("store-new-entry-ids", "data"),
    Output("store-seen-by-date", "data"),
    Input("btn-refresh", "n_clicks"),
    State("store-account", "data"),
    State("p2-date-picker", "date"),
    State("store-uid", "data"),
    State("store-auth-token", "data"),
    State("store-seen-by-date", "data"),
    State("store-new-entry-ids", "data"),
    prevent_initial_call=True,
)
def fetch_emails(n, account, date_str, uid, id_token, seen_by_date, existing_new_ids):
    seen_by_date = seen_by_date or {}
    if not account:
        return [], "계정을 선택하세요.", (existing_new_ids or []), seen_by_date
    try:
        from outlook_manager import OutlookManager
        mgr = OutlookManager()
        target_date = datetime.strptime(date_str, "%Y-%m-%d").date() if date_str else None

        # Outlook에서 직접 가져오기
        emails = mgr.get_emails_by_date(account_email=account, target_date=target_date, limit=100)

        # 신규 메일 판별: 이번 날짜에서 이전에 못 본 entry_id만 NEW
        seen_ids = set(seen_by_date.get(date_str, []))
        truly_new = [
            e["entry_id"] for e in emails
            if e.get("entry_id") and e["entry_id"] not in seen_ids
        ]
        # 기존 NEW 목록에 신규만 추가 (다른 날짜의 NEW는 유지)
        combined_new = list(set(existing_new_ids or []) | set(truly_new))
        # 이번 조회한 모든 entry_id를 seen에 추가
        updated_seen = dict(seen_by_date)
        updated_seen[date_str] = list(seen_ids | set(e.get("entry_id", "") for e in emails))

        # Firebase에 백그라운드로 저장 (본문은 5000자 제한 - Firestore 1MB 한도 대비)
        if uid and id_token and emails and date_str:
            def _save_emails_to_cloud():
                try:
                    safe_acc = account.replace(".", "_")
                    doc_id = f"{safe_acc}_{date_str}"
                    emails_to_save = [
                        {**e, "body": e.get("body", "")[:5000]}
                        for e in emails
                    ]
                    fb_client.save_data(uid, id_token, "emails", doc_id, {
                        "account": account,
                        "date": date_str,
                        "emails": emails_to_save,
                    })
                except Exception as e:
                    logger.warning(f"이메일 클라우드 저장 실패: {e}")
            threading.Thread(target=_save_emails_to_cloud, daemon=True).start()

        return emails, f"조회 완료: 총 {len(emails)}개 메일", combined_new, updated_seen
    except Exception as e:
        return [], f"오류: {e}", (existing_new_ids or []), seen_by_date


# ── Page 2: Render email list ─────────────────────────────────────────────────
@app.callback(
    Output("email-list-container", "children"),
    Input("store-emails", "data"),
    Input("store-analyze-done", "data"),
    Input("store-highlighted-emails", "data"),
    Input("store-new-entry-ids", "data"),
    State("store-email-checked", "data"),
)
def render_email_list(emails, _, highlighted, new_entry_ids, checked_indices):
    if not emails:
        return html.Span("이메일을 조회하세요.", className="text-muted small p-2 d-block")

    highlighted_set = set(highlighted or [])
    new_set = set(new_entry_ids or [])
    checked_set = set(checked_indices or [])
    items = []
    for i, email in enumerate(emails):
        subject = email.get("subject", "제목 없음")
        sender = email.get("sender", "알 수 없음")
        entry_id = email.get("entry_id", "")
        is_highlighted = i in highlighted_set
        is_new = entry_id in new_set

        highlight_style = {
            "borderLeft": "3px solid #D32F2F",
            "backgroundColor": "#fff5f5",
        } if is_highlighted else {}

        # NEW 아이콘 (new.png)
        new_icon = html.Img(
            src="/assets/new.png",
            style={"height": "16px", "marginRight": "4px", "verticalAlign": "middle"},
        ) if is_new else None

        text_color = {"color": "#D32F2F"} if (is_new or is_highlighted) else {}
        subject_color = {"color": "#D32F2F"} if (is_new or is_highlighted) else {"color": "#6c757d"}

        items.append(
            html.Div([
                dbc.Row([
                    dbc.Col(
                        dbc.Checkbox(
                            id={"type": "email-checkbox", "index": i},
                            value=i in checked_set,
                        ),
                        width="auto",
                        className="pe-1",
                    ),
                    dbc.Col(
                        html.Div(
                            [
                                html.Span(sender, className="fw-semibold small d-block",
                                          style=text_color),
                                html.Span([
                                    *([new_icon] if new_icon else []),
                                    subject[:55] + ("…" if len(subject) > 55 else ""),
                                ], className="small", style=subject_color),
                            ],
                            id={"type": "email-row", "index": i},
                            n_clicks=0,
                            style={"cursor": "pointer"},
                        ),
                    ),
                ], className="g-0 align-items-center"),
            ], className="email-item px-2 py-1", style=highlight_style)
        )
    return html.Div(items)


# ── Page 2: Email row click → show body ──────────────────────────────────────
@app.callback(
    Output("store-selected-email-idx", "data"),
    Input({"type": "email-row", "index": ALL}, "n_clicks"),
    prevent_initial_call=True,
)
def email_row_clicked(n_clicks_list):
    triggered = ctx.triggered_id
    if triggered and isinstance(triggered, dict):
        return triggered["index"]
    return no_update


@app.callback(
    Output("email-viewer", "children"),
    Input("store-selected-email-idx", "data"),
    State("store-emails", "data"),
)
def show_email_body(idx, emails):
    if idx is None or not emails or idx >= len(emails):
        return "메일을 클릭하면 본문이 표시됩니다."
    email = emails[idx]
    body = email.get("body", "")
    return body[:2000] + ("…" if len(body) > 2000 else "")


# ── Page 2: Checkbox state → store ───────────────────────────────────────────
@app.callback(
    Output("store-email-checked", "data"),
    Input({"type": "email-checkbox", "index": ALL}, "value"),
    prevent_initial_call=True,
)
def update_email_checked(values):
    return [i for i, v in enumerate(values) if v]


# ── Page 2: 체크박스 선택 시 NEW 상태 해제 ────────────────────────────────────
# store-email-checked → store-new-entry-ids 체인으로 순서 보장
@app.callback(
    Output("store-new-entry-ids", "data", allow_duplicate=True),
    Input("store-email-checked", "data"),
    State("store-emails", "data"),
    State("store-new-entry-ids", "data"),
    prevent_initial_call=True,
)
def clear_new_on_check(checked_indices, emails, new_entry_ids):
    new_set = set(new_entry_ids or [])
    for i in (checked_indices or []):
        if i < len(emails or []):
            entry_id = emails[i].get("entry_id", "")
            new_set.discard(entry_id)
    return list(new_set)


# ── Page 2: Select all emails ─────────────────────────────────────────────────
@app.callback(
    Output({"type": "email-checkbox", "index": ALL}, "value"),
    Input("email-select-all-cb", "value"),
    State("store-emails", "data"),
    prevent_initial_call=True,
)
def select_all_emails(checked, emails):
    return [bool(checked)] * len(ctx.outputs_list)


# ── Page 2: AI Analysis (threading-based) ────────────────────────────────────
def _run_analysis_thread(checked_indices, emails, user_profile, api_key, existing_todos):
    """별도 스레드에서 AI 분석 실행"""
    global _analysis
    _analysis["status"] = "running"
    _analysis["progress"] = 0
    _analysis["text"] = "준비 중…"
    _analysis["errors"] = []
    _analysis["todos"] = list(existing_todos or [])

    try:
        from ai_processor import AIProcessor
        processor = AIProcessor(override_api_key=api_key)
    except Exception as e:
        _analysis["status"] = "error"
        _analysis["text"] = f"API 연결 실패: {e}"
        return

    user_context = build_user_context(user_profile) if user_profile else ""
    total = len(checked_indices)
    _analysis["total"] = total
    results = list(existing_todos or [])
    errors = []

    for i, idx in enumerate(checked_indices):
        if _analysis["cancel"].is_set():
            _analysis["status"] = "cancelled"
            return

        if idx >= len(emails):
            continue
        email = emails[idx]
        _analysis["progress"] = int((i / total) * 100)
        _analysis["text"] = f"{i + 1}/{total} 분석 중… {email.get('subject', '')[:30]}"

        try:
            result = processor.analyze_email(
                email_body=email.get("body", ""),
                user_context=user_context,
                to_recipients=email.get("to_recipients", ""),
                cc_recipients=email.get("cc_recipients", ""),
            )
        except Exception as e:
            err_detail = f"[{email.get('subject','?')[:20]}] {e}"
            errors.append(err_detail)
            continue

        # 분석 성공한 메일 Outlook 읽음 처리
        entry_id = email.get("entry_id", "")
        if entry_id:
            try:
                from outlook_manager import OutlookManager
                OutlookManager().mark_as_read(entry_id)
            except Exception:
                pass

        summary = result.get("summary", "")
        mail_type = result.get("mail_type", "ACTION")
        priority = result.get("priority", "보통")
        raw_todos = result.get("todos", [])

        if mail_type == "IGNORE":
            continue

        for t in raw_todos:
            if isinstance(t, dict):
                text = t.get("text", "")
            else:
                text = str(t)
            if not text:
                continue
            results.append({
                "id": str(uuid.uuid4()),
                "text": text,
                "summary": summary,
                "email_subject": email.get("subject", ""),
                "entry_id": email.get("entry_id", ""),
                "status": "active",
                "forwarded": False,
                "due_date": "",
                "priority": priority,
                "mail_type": mail_type,
            })

    _analysis["todos"] = results
    _analysis["errors"] = errors
    _analysis["new_todo_count"] = len(results) - len(existing_todos or [])
    _analysis["progress"] = 100
    _analysis["text"] = f"완료! {total}개 분석됨"
    _analysis["status"] = "done"


@app.callback(
    Output("analyze-modal", "is_open"),
    Output("btn-analyze", "disabled"),
    Output("analyze-interval", "disabled"),
    Output("analyze-progress-wrap", "style", allow_duplicate=True),
    Output("analyze-result-wrap", "style", allow_duplicate=True),
    Output("analyze-modal-title", "children", allow_duplicate=True),
    Input("btn-analyze", "n_clicks"),
    State("store-email-checked", "data"),
    State("store-emails", "data"),
    State("store-user-profile", "data"),
    State("store-api-key", "data"),
    State("store-todos", "data"),
    prevent_initial_call=True,
)
def start_analyze(n_clicks, checked_indices, emails, user_profile, api_key, todos):
    global _analysis
    hidden = {"display": "none"}
    visible = {}

    if not checked_indices or not emails:
        return no_update, no_update, True, visible, hidden, "알림"

    _analysis["cancel"].clear()
    _analysis["status"] = "idle"

    t = _threading.Thread(
        target=_run_analysis_thread,
        args=(checked_indices, emails, user_profile, api_key, todos),
        daemon=True,
    )
    t.start()

    return True, True, False, visible, hidden, "AI 분석 중…"


@app.callback(
    Output("analyze-progress-text", "children"),
    Output("analyze-progress-bar", "value"),
    Output("analyze-progress-wrap", "style"),
    Output("analyze-result-wrap", "style"),
    Output("analyze-result-text", "children"),
    Output("analyze-modal-title", "children"),
    Output("store-todos", "data"),
    Output("store-analyze-done", "data"),
    Output("btn-analyze", "disabled", allow_duplicate=True),
    Output("analyze-interval", "disabled", allow_duplicate=True),
    Output("p2-toast", "children"),
    Output("p2-toast", "is_open"),
    Input("analyze-interval", "n_intervals"),
    State("store-todos", "data"),
    State("store-analyze-done", "data"),
    prevent_initial_call=True,
)
def poll_analysis_progress(n, todos, done_count):
    global _analysis
    hidden = {"display": "none"}
    visible = {}
    status = _analysis["status"]

    if status == "running":
        return (
            _analysis["text"], _analysis["progress"],
            visible, hidden,
            no_update, "AI 분석 중…",
            no_update, no_update,
            no_update, False,
            no_update, no_update,
        )

    if status == "done":
        new_todos = _analysis["todos"]
        errors = _analysis["errors"]
        new_count = _analysis["new_todo_count"]
        total = _analysis["total"]
        done = (done_count or 0) + 1
        _analysis["status"] = "idle"

        result_body = html.Div([
            html.P("분석이 완료되었습니다.", className="fw-bold mb-3"),
            html.P(f"분석한 메일: {total}개", className="mb-1 text-muted small"),
            html.P(f"생성된 TODO: {new_count}개", className="mb-0 text-muted small"),
        ])
        toast_msg = no_update
        toast_open = no_update
        if errors:
            toast_msg = f"오류 {len(errors)}건 발생"
            toast_open = True

        return (
            "완료!", 100,
            hidden, visible,
            result_body, "분석 완료",
            new_todos, done,
            False, True,
            toast_msg, toast_open,
        )

    if status in ("error", "cancelled"):
        msg = _analysis["text"]
        _analysis["status"] = "idle"
        return (
            msg, 0,
            hidden, visible,
            msg, "오류" if status == "error" else "취소됨",
            no_update, no_update,
            False, True,
            msg, True,
        )

    return (
        no_update, no_update,
        no_update, no_update,
        no_update, no_update,
        no_update, no_update,
        no_update, no_update,
        no_update, no_update,
    )


@app.callback(
    Output("analyze-modal", "is_open", allow_duplicate=True),
    Output("analyze-interval", "disabled", allow_duplicate=True),
    Output("analyze-progress-bar", "value", allow_duplicate=True),
    Output("analyze-progress-text", "children", allow_duplicate=True),
    Input("btn-analyze-cancel", "n_clicks"),
    prevent_initial_call=True,
)
def cancel_analyze(n):
    global _analysis
    _analysis["cancel"].set()
    _analysis["status"] = "idle"
    return False, True, 0, ""


# ── Select-all count labels ───────────────────────────────────────────────────
@app.callback(
    Output("email-select-count", "children"),
    Input("store-email-checked", "data"),
    State("store-emails", "data"),
)
def update_email_count(checked, emails):
    n = len(checked or [])
    total = len(emails or [])
    return f"({n}/{total})"


@app.callback(
    Output("p2-active-select-count", "children"),
    Input("store-todo-checked-p2", "data"),
    Input("store-todos", "data"),
)
def update_p2_active_count(checked, todos):
    n = len(checked or [])
    total = sum(1 for t in (todos or []) if t.get("status") in ("active", "completed"))
    return f"({n}/{total})"


@app.callback(
    Output("p3-todo-count-badge", "children"),
    Output("p3-todo-count-badge", "style"),
    Input("store-todos-p3", "data"),
)
def update_p3_badge(todos):
    count = sum(1 for t in (todos or []) if t.get("status") == "active")
    if count > 0:
        return str(count), {"fontSize": "0.7rem", "display": "inline-block"}
    return "", {"fontSize": "0.7rem", "display": "none"}


@app.callback(
    Output("p2-trash-select-count", "children"),
    Input("store-todo-trash-checked-p2", "data"),
    Input("store-todos", "data"),
)
def update_p2_trash_count(checked, todos):
    n = len(checked or [])
    total = sum(1 for t in (todos or []) if t.get("status") == "deleted")
    return f"({n}/{total})"


@app.callback(
    Output("p3-active-select-count", "children"),
    Input("store-todo-checked-p3-active", "data"),
    Input("store-todos-p3", "data"),
)
def update_p3_active_count(checked, todos):
    n = len(checked or [])
    total = sum(1 for t in (todos or []) if t.get("status") == "active")
    return f"({n}/{total})"


@app.callback(
    Output("p3-completed-select-count", "children"),
    Input("store-todo-checked-p3-completed", "data"),
    Input("store-todos-p3", "data"),
)
def update_p3_completed_count(checked, todos):
    n = len(checked or [])
    total = sum(1 for t in (todos or []) if t.get("status") == "completed")
    return f"({n}/{total})"


@app.callback(
    Output("p3-trash-select-count", "children"),
    Input("store-todo-checked-p3-trash", "data"),
    Input("store-todos-p3", "data"),
)
def update_p3_trash_count(checked, todos):
    n = len(checked or [])
    total = sum(1 for t in (todos or []) if t.get("status") == "deleted")
    return f"({n}/{total})"


# ── Todo collapse toggles ─────────────────────────────────────────────────────
@app.callback(
    Output({"type": "todo-collapse-p2", "index": MATCH, "st": MATCH}, "is_open"),
    Input({"type": "todo-toggle-p2", "index": MATCH, "st": MATCH}, "n_clicks"),
    State({"type": "todo-collapse-p2", "index": MATCH, "st": MATCH}, "is_open"),
    prevent_initial_call=True,
)
def toggle_todo_p2(n, is_open):
    return not is_open


@app.callback(
    Output({"type": "todo-collapse-p3", "index": MATCH, "st": MATCH}, "is_open"),
    Input({"type": "todo-toggle-p3", "index": MATCH, "st": MATCH}, "n_clicks"),
    State({"type": "todo-collapse-p3", "index": MATCH, "st": MATCH}, "is_open"),
    prevent_initial_call=True,
)
def toggle_todo_p3(n, is_open):
    return not is_open


# ── Page 3: Open original email in Outlook ────────────────────────────────────
@app.callback(
    Output("p3-toast", "children", allow_duplicate=True),
    Output("p3-toast", "is_open", allow_duplicate=True),
    Output("p3-toast", "color", allow_duplicate=True),
    Input({"type": "btn-open-email-p3", "index": ALL}, "n_clicks"),
    State("store-todos-p3", "data"),
    prevent_initial_call=True,
)
def open_original_email_p3(n_clicks_list, todos):
    if not any(n for n in (n_clicks_list or []) if n):
        return no_update, no_update, no_update
    triggered = ctx.triggered_id
    if not triggered:
        return no_update, no_update, no_update
    idx = triggered["index"]
    if not todos or idx >= len(todos):
        return "메일 정보를 찾을 수 없습니다.", True, "danger"
    entry_id = todos[idx].get("entry_id", "")
    if not entry_id:
        return "이 TODO는 원본 메일 정보가 없습니다. (재분석 필요)", True, "warning"
    try:
        from outlook_manager import OutlookManager
        mgr = OutlookManager()
        mgr.open_email_by_entry_id(entry_id)
        return no_update, no_update, no_update
    except Exception as e:
        return str(e), True, "danger"


# ── Page 2: Render Todo lists ─────────────────────────────────────────────────
def _render_todo_item_p2(todo, idx, tab="active"):
    status = todo.get("status", "active")
    forwarded = todo.get("forwarded", False)
    text = todo.get("text", "")
    subject = todo.get("email_subject", "")
    summary = todo.get("summary", "")
    priority = todo.get("priority", "보통")
    cb_type = "todo-cb-active-p2" if tab == "active" else "todo-cb-trash-p2"

    if status == "deleted":
        text_style = {"textDecoration": "line-through", "color": "#bbb"}
    elif status == "completed":
        text_style = {"textDecoration": "line-through", "color": "#1976D2"}
    elif forwarded:
        text_style = {"color": "#7B1FA2"}
    else:
        text_style = {"color": "#212121"}

    priority_colors = {"높음": "danger", "보통": "warning", "낮음": "primary"}
    priority_badge = dbc.Badge(
        priority,
        color=priority_colors.get(priority, "secondary"),
        className="me-1",
        style={"fontSize": "var(--fs-badge)"},
    )

    return html.Div([
        dbc.Row([
            dbc.Col(
                dbc.Checkbox(id={"type": cb_type, "index": idx}, value=False),
                width="auto",
            ),
            dbc.Col([
                html.Div([
                    priority_badge,
                    html.Span(text[:55] + ("…" if len(text) > 55 else ""),
                               className="small", style=text_style),
                ], className="d-flex align-items-center flex-wrap"),
                html.Span(f"출처: {subject[:30]}", className="text-muted",
                           style={"fontSize": "var(--fs-meta)"}),
            ]),
            dbc.Col(
                dbc.Button(
                    html.I(className="bi bi-chevron-down"),
                    id={"type": "todo-toggle-p2", "index": idx, "st": status},
                    size="sm", color="link", className="p-0 text-muted",
                    style={"fontSize": "var(--fs-meta)", "lineHeight": "1"},
                ),
                width="auto",
            ),
        ], className="g-0 align-items-center"),
        dbc.Collapse(
            html.Div([
                html.Div("내용 요약", className="text-muted mb-1",
                         style={"fontSize": "var(--fs-xs)", "fontWeight": "600"}),
                html.Div(summary or "요약 없음", className="small",
                         style={"whiteSpace": "pre-wrap", "color": "#555"}),
            ], className="pt-1 pb-2 ps-4"),
            id={"type": "todo-collapse-p2", "index": idx, "st": status},
            is_open=False,
        ),
    ], className="todo-item px-2 pt-1 border-bottom")


@app.callback(
    Output("todo-list-active", "children"),
    Input("store-todos", "data"),
    Input("store-analyze-done", "data"),
)
def render_active_todos(todos, _):
    if not todos:
        return html.Span("Todo 없음", className="text-muted small p-3 d-block")
    items = [
        _render_todo_item_p2(t, i, tab="active")
        for i, t in enumerate(todos)
        if t.get("status") in ("active", "completed")
    ]
    return items if items else html.Span("Todo 없음", className="text-muted small p-3 d-block")


@app.callback(
    Output("todo-list-trash", "children"),
    Input("store-todos", "data"),
)
def render_trash_todos(todos):
    if not todos:
        return html.Span("휴지통 비어있음", className="text-muted small p-3 d-block")
    items = [
        _render_todo_item_p2(t, i, tab="trash")
        for i, t in enumerate(todos)
        if t.get("status") == "deleted"
    ]
    return items if items else html.Span("휴지통 비어있음", className="text-muted small p-3 d-block")


# ── Page 2: Todo checkbox state ───────────────────────────────────────────────
@app.callback(
    Output("store-todo-checked-p2", "data"),
    Input({"type": "todo-cb-active-p2", "index": ALL}, "value"),
    State("store-todos", "data"),
    prevent_initial_call=True,
)
def update_todo_checked_p2(values, todos):
    if not todos:
        return []
    inputs = ctx.inputs_list[0]
    checked = []
    for item, val in zip(inputs, values):
        if val:
            checked.append(item["id"]["index"])
    print(f"[DEBUG] checked-p2 updated: {checked}, values={values}", flush=True)
    return checked


# ── Page 2: Todo 선택 → 연관 이메일 강조 ─────────────────────────────────────
@app.callback(
    Output("store-highlighted-emails", "data"),
    Input("store-todo-checked-p2", "data"),
    State("store-todos", "data"),
    State("store-emails", "data"),
)
def highlight_related_emails(checked, todos, emails):
    if not checked or not todos or not emails:
        return []
    # 체크된 TODO들의 email_subject 수집
    target_subjects = set()
    for idx in checked:
        i = int(idx)
        if 0 <= i < len(todos):
            subj = todos[i].get("email_subject", "")
            if subj:
                target_subjects.add(subj)
    # 해당 subject와 일치하는 이메일 인덱스 수집
    highlighted = [
        i for i, email in enumerate(emails)
        if email.get("subject", "") in target_subjects
    ]
    return highlighted


@app.callback(
    Output("store-todo-trash-checked-p2", "data"),
    Input({"type": "todo-cb-trash-p2", "index": ALL}, "value"),
    State("store-todos", "data"),
    prevent_initial_call=True,
)
def update_todo_trash_checked_p2(values, todos):
    if not todos:
        return []
    inputs = ctx.inputs_list[0]
    checked = []
    for item, val in zip(inputs, values):
        if val:
            checked.append(item["id"]["index"])
    return checked


# ── Page 2: Todo select all ───────────────────────────────────────────────────
@app.callback(
    Output({"type": "todo-cb-active-p2", "index": ALL}, "value"),
    Input("todo-active-select-all", "value"),
    prevent_initial_call=True,
)
def select_all_active_p2(checked):
    outputs = ctx.outputs_list
    return [bool(checked)] * len(outputs)


@app.callback(
    Output({"type": "todo-cb-trash-p2", "index": ALL}, "value"),
    Input("todo-trash-select-all", "value"),
    prevent_initial_call=True,
)
def select_all_trash_p2(checked):
    outputs = ctx.outputs_list
    return [bool(checked)] * len(outputs)


# ── Page 2: Todo actions ──────────────────────────────────────────────────────
@app.callback(
    Output("store-todos", "data", allow_duplicate=True),
    Output("store-todos-p3", "data", allow_duplicate=True),
    Output("p2-toast", "children", allow_duplicate=True),
    Output("p2-toast", "is_open", allow_duplicate=True),
    Output("store-page", "data", allow_duplicate=True),
    Input("btn-todo-delete", "n_clicks"),
    Input("btn-todo-forward", "n_clicks"),
    State("store-todo-checked-p2", "data"),
    State("store-todos", "data"),
    State("store-todos-p3", "data"),
    prevent_initial_call=True,
)
def todo_actions_p2(n_delete, n_forward, checked, todos, todos_p3):
    try:
        if not todos:
            return no_update, no_update, "항목이 없습니다.", True, no_update

        triggered = ctx.triggered_id
        todos = [dict(t) for t in todos]

        # 전달 버튼은 체크 없으면 전체 active 항목 전달
        if triggered == "btn-todo-forward":
            import copy
            todos_p3 = list(todos_p3 or [])
            # checked 인덱스를 정수로 변환
            checked_ints = [int(c) for c in checked] if checked else []
            target_indices = set(checked_ints) if checked_ints else {i for i, t in enumerate(todos) if t.get("status") == "active"}
            if not target_indices:
                return no_update, no_update, "전달할 항목이 없습니다.", True, no_update
            # store-todos-p3에 독립 복사본 추가 (중복 제외)
            existing_ids = {t.get("id","") for t in todos_p3}
            new_items = []
            for i in sorted(target_indices):
                if 0 <= i < len(todos):
                    t = todos[i]
                    if t.get("id","") not in existing_ids:
                        new_t = copy.deepcopy(t)
                        new_t["status"] = "active"
                        new_t["forwarded"] = True
                        new_items.append(new_t)
                        existing_ids.add(new_t.get("id",""))
            todos_p3.extend(new_items)
            # store-todos에서 전달된 항목 완전 제거 (단방향 소통)
            todos = [t for i, t in enumerate(todos) if i not in target_indices]
            # 전달 성공 시 자동으로 Page3으로 이동
            return todos, todos_p3, f"{len(new_items)}개 항목이 전체 TODO에 추가되었습니다.", True, 3

        # 삭제는 체크 필수
        if not checked:
            return no_update, no_update, "항목을 선택하세요.", True, no_update

        checked_ints = [int(c) for c in checked]
        if triggered == "btn-todo-delete":
            for i in checked_ints:
                todos[i]["status"] = "deleted"
                if todos[i].get("notion_page_id"):
                    todos[i]["pending_sync"] = True
            msg = f"{len(checked_ints)}개 삭제됨"
        else:
            return no_update, no_update, "", False, no_update

        return todos, no_update, msg, True, no_update
    except Exception as e:
        return no_update, no_update, f"오류: {traceback.format_exc()}", True, no_update


# ── Page 2: Trash actions ─────────────────────────────────────────────────────
@app.callback(
    Output("store-todos", "data", allow_duplicate=True),
    Output("p2-toast", "children", allow_duplicate=True),
    Output("p2-toast", "is_open", allow_duplicate=True),
    Input("btn-todo-restore", "n_clicks"),
    Input("btn-todo-perm-delete", "n_clicks"),
    State({"type": "todo-cb-trash-p2", "index": ALL}, "value"),
    State("store-todos", "data"),
    State("store-notion-key", "data"),
    State("store-notion-db", "data"),
    prevent_initial_call=True,
)
def trash_actions_p2(n_restore, n_perm, cb_values, todos, notion_key, notion_db):
    if not todos:
        return no_update, "항목이 없습니다.", True

    cb_states = ctx.states_list[0]
    checked_ints = [s["id"]["index"] for s, v in zip(cb_states, cb_values or []) if v]

    if not checked_ints:
        return no_update, "항목을 선택하세요.", True

    triggered = ctx.triggered_id
    todos = [dict(t) for t in todos]

    if triggered == "btn-todo-restore":
        for i in checked_ints:
            todos[i]["status"] = "active"
            todos[i]["forwarded"] = False
        return todos, f"{len(checked_ints)}개 복구됨", True
    elif triggered == "btn-todo-perm-delete":
        if notion_key and notion_db:
            page_ids = [todos[i]["notion_page_id"] for i in checked_ints if todos[i].get("notion_page_id")]
            if page_ids:
                threading.Thread(
                    target=_archive_notion_bg,
                    args=(notion_key, notion_db, page_ids),
                    daemon=True
                ).start()
        checked_set = set(checked_ints)
        todos = [t for i, t in enumerate(todos) if i not in checked_set]
        return todos, f"{len(checked_ints)}개 영구삭제됨", True

    return no_update, "", False


# ── Analyze modal: confirm close ──────────────────────────────────────────────
@app.callback(
    Output("analyze-modal", "is_open", allow_duplicate=True),
    Output("analyze-interval", "disabled", allow_duplicate=True),
    Output("analyze-progress-bar", "value", allow_duplicate=True),
    Output("analyze-progress-text", "children", allow_duplicate=True),
    Input("btn-analyze-confirm", "n_clicks"),
    prevent_initial_call=True,
)
def close_analyze_modal(n):
    return False, True, 0, ""


# ── AI 답장: 버튼 클릭 → 모달 열기 + 백그라운드 생성 시작 ─────────────────────
@app.callback(
    Output("ai-reply-modal", "is_open"),
    Output("ai-reply-modal-title", "children"),
    Output("ai-reply-loading-wrap", "style"),
    Output("ai-reply-result-wrap", "style"),
    Output("ai-reply-text", "value"),
    Output("btn-ai-reply-open-outlook", "style"),
    Output("store-ai-reply-entry-id", "data"),
    Output("ai-reply-interval", "disabled"),
    Input({"type": "btn-ai-reply-p3", "index": ALL}, "n_clicks"),
    State("store-todos-p3", "data"),
    State("p3-priority-filter", "value"),
    State("p3-sort-combo", "value"),
    State("store-api-key", "data"),
    State("store-user-profile", "data"),
    prevent_initial_call=True,
)
def start_ai_reply(n_clicks_list, todos_p3, priority_filter, sort_combo, api_key, user_profile):
    global _ai_reply
    triggered = ctx.triggered_id
    if not triggered or not any(n for n in (n_clicks_list or []) if n):
        return no_update, no_update, no_update, no_update, no_update, no_update, no_update, no_update

    idx = triggered["index"]
    filtered = _get_filtered_sorted(todos_p3 or [], priority_filter, sort_combo, ["active"])
    if idx >= len(filtered):
        return no_update, no_update, no_update, no_update, no_update, no_update, no_update, no_update

    _, todo = filtered[idx]   # (원본_idx, todo_dict) 튜플
    entry_id = todo.get("entry_id", "")
    email_subject = todo.get("email_subject", "")
    summary = todo.get("summary", "")
    todo_text = todo.get("text", "")

    _ai_reply["status"] = "running"
    _ai_reply["text"] = ""
    _ai_reply["entry_id"] = entry_id

    user_context = build_user_context(user_profile) if user_profile else ""

    def _run():
        global _ai_reply
        try:
            from ai_processor import AIProcessor
            processor = AIProcessor(override_api_key=api_key)
            result = processor.generate_reply(
                email_subject=email_subject,
                summary=summary,
                todo_text=todo_text,
                user_profile=user_context,
            )
            _ai_reply["text"] = result
            _ai_reply["status"] = "done"
        except Exception as e:
            _ai_reply["text"] = str(e)
            _ai_reply["status"] = "error"

    threading.Thread(target=_run, daemon=True).start()

    return True, "AI 답장 생성 중…", {}, {"display": "none"}, "", {"display": "none"}, entry_id, False


# ── AI 답장: 폴링 → 결과 표시 ─────────────────────────────────────────────────
@app.callback(
    Output("ai-reply-modal-title", "children", allow_duplicate=True),
    Output("ai-reply-loading-wrap", "style", allow_duplicate=True),
    Output("ai-reply-result-wrap", "style", allow_duplicate=True),
    Output("ai-reply-text", "value", allow_duplicate=True),
    Output("btn-ai-reply-open-outlook", "style", allow_duplicate=True),
    Output("ai-reply-interval", "disabled", allow_duplicate=True),
    Output("store-ai-reply-done", "data"),
    Input("ai-reply-interval", "n_intervals"),
    State("store-ai-reply-done", "data"),
    prevent_initial_call=True,
)
def poll_ai_reply(n, done_count):
    global _ai_reply
    status = _ai_reply["status"]
    if status == "running":
        return no_update, no_update, no_update, no_update, no_update, no_update, no_update
    if status == "done":
        _ai_reply["status"] = "idle"
        return (
            "AI 답장 초안",
            {"display": "none"},
            {},
            _ai_reply["text"],
            {},           # 버튼 보이기
            True,         # interval 중지
            (done_count or 0) + 1,
        )
    if status == "error":
        _ai_reply["status"] = "idle"
        return (
            "오류 발생",
            {"display": "none"},
            {},
            f"답장 생성 중 오류가 발생했습니다:\n{_ai_reply['text']}",
            {"display": "none"},
            True,
            (done_count or 0) + 1,
        )
    return no_update, no_update, no_update, no_update, no_update, no_update, no_update


# ── AI 답장: Outlook에서 열기 ─────────────────────────────────────────────────
@app.callback(
    Output("p3-toast", "children", allow_duplicate=True),
    Output("p3-toast", "is_open", allow_duplicate=True),
    Output("p3-toast", "color", allow_duplicate=True),
    Input("btn-ai-reply-open-outlook", "n_clicks"),
    State("ai-reply-text", "value"),
    State("store-ai-reply-entry-id", "data"),
    prevent_initial_call=True,
)
def open_ai_reply_in_outlook(n, reply_body, entry_id):
    if not n or not entry_id:
        return no_update, no_update, no_update
    try:
        from outlook_manager import OutlookManager
        OutlookManager().open_reply_with_body(entry_id, reply_body or "")
        return "Outlook 답장 창이 열렸습니다.", True, "success"
    except Exception as e:
        return str(e), True, "danger"


# ── AI 답장: 모달 닫기 ────────────────────────────────────────────────────────
@app.callback(
    Output("ai-reply-modal", "is_open", allow_duplicate=True),
    Output("ai-reply-interval", "disabled", allow_duplicate=True),
    Input("btn-ai-reply-close", "n_clicks"),
    prevent_initial_call=True,
)
def close_ai_reply_modal(n):
    return False, True


# ── Navigation buttons ────────────────────────────────────────────────────────
@app.callback(
    Output("store-page", "data", allow_duplicate=True),
    Input("btn-back-page1", "n_clicks"),
    prevent_initial_call=True,
)
def back_to_page1(n):
    return 1


@app.callback(
    Output("store-page", "data", allow_duplicate=True),
    Input("btn-goto-page3", "n_clicks"),
    prevent_initial_call=True,
)
def goto_page3(n):
    return 3


@app.callback(
    Output("store-page", "data", allow_duplicate=True),
    Input("btn-back-page2", "n_clicks"),
    prevent_initial_call=True,
)
def back_to_page2(n):
    return 2


# ── Page 3: Render todo lists ─────────────────────────────────────────────────
def _priority_badge(priority):
    colors = {"높음": "danger", "보통": "warning", "낮음": "primary"}
    return dbc.Badge(priority, color=colors.get(priority, "secondary"),
                     className="me-1", style={"fontSize": "var(--fs-badge)"})


def _render_todo_item_p3(todo, idx, list_type="active"):
    text = todo.get("text", "")
    priority = todo.get("priority", "보통")
    due_date = todo.get("due_date", "")
    subject = todo.get("email_subject", "")
    summary = todo.get("summary", "")
    status = todo.get("status", "active")
    entry_id = todo.get("entry_id", "")

    text_style = {}
    if status == "completed":
        text_style = {"textDecoration": "line-through", "color": "#888"}
    elif status == "deleted":
        text_style = {"color": "#bbb"}

    cb_id = {"type": f"todo-cb-p3-{list_type}", "index": idx}

    return html.Div([
        dbc.Row([
            dbc.Col(dbc.Checkbox(id=cb_id, value=False), width="auto"),
            dbc.Col([
                html.Div([
                    _priority_badge(priority),
                    html.Span(text[:70] + ("…" if len(text) > 70 else ""),
                               className="small", style=text_style),
                ]),
                html.Div([
                    html.Span(
                        f"출처: {subject[:30]}" + (f" | 마감: {due_date}" if due_date else ""),
                        style={"fontSize": "var(--fs-xs)", "color": "#888"},
                    ),
                    dbc.Button(
                        "원본메일",
                        id={"type": "btn-open-email-p3", "index": idx},
                        size="sm",
                        color="link",
                        className="p-0 ms-2",
                        style={"fontSize": "var(--fs-xs)", "verticalAlign": "baseline"},
                        disabled=not entry_id,
                    ),
                    dbc.Button(
                        "AI 답장",
                        id={"type": "btn-ai-reply-p3", "index": idx},
                        size="sm",
                        color="link",
                        className="p-0 ms-2",
                        style={"fontSize": "var(--fs-xs)", "verticalAlign": "baseline",
                               "color": "#1565C0"},
                        disabled=not entry_id,
                    ),
                ], className="d-flex align-items-center"),
            ]),
            dbc.Col(
                dbc.Button(
                    html.I(className="bi bi-chevron-down"),
                    id={"type": "todo-toggle-p3", "index": idx, "st": status},
                    size="sm", color="link", className="p-0 text-muted",
                    style={"fontSize": "var(--fs-meta)", "lineHeight": "1"},
                ),
                width="auto",
            ),
        ], className="g-0 align-items-center"),
        dbc.Collapse(
            html.Div([
                html.Div("내용 요약", className="text-muted mb-1",
                         style={"fontSize": "var(--fs-xs)", "fontWeight": "600"}),
                html.Div(summary or "요약 없음", className="small",
                         style={"whiteSpace": "pre-wrap", "color": "#555"}),
            ], className="pt-1 pb-2 ps-4"),
            id={"type": "todo-collapse-p3", "index": idx, "st": status},
            is_open=False,
        ),
    ], className="todo-item px-2 pt-1 border-bottom")


def _get_filtered_sorted(todos, filter_val, sort_val, statuses):
    items = [(i, t) for i, t in enumerate(todos) if t.get("status") in statuses]
    if filter_val and filter_val != "all":
        items = [(i, t) for i, t in items if t.get("priority") == filter_val]
    priority_order = {"높음": 0, "보통": 1, "낮음": 2}
    if sort_val == "priority":
        items.sort(key=lambda x: priority_order.get(x[1].get("priority", "보통"), 1))
    elif sort_val == "due_date":
        items.sort(key=lambda x: x[1].get("due_date", "") or "9999")
    return items


@app.callback(
    Output("p3-active-list", "children"),
    Output("p3-completed-list", "children"),
    Output("p3-trash-list", "children"),
    Input("store-todos-p3", "data"),
    Input("p3-priority-filter", "value"),
    Input("p3-sort-combo", "value"),
    Input("store-page", "data"),
)
def render_p3_lists(todos, filter_val, sort_val, page):
    if not todos:
        empty = html.Span("항목 없음", className="text-muted small p-3 d-block")
        return empty, empty, empty

    active_items = _get_filtered_sorted(todos, filter_val, sort_val, ["active"])
    completed_items = _get_filtered_sorted(todos, filter_val, sort_val, ["completed"])
    trash_items = _get_filtered_sorted(todos, filter_val, sort_val, ["deleted"])

    def render(items, list_type):
        if not items:
            return html.Span("항목 없음", className="text-muted small p-3 d-block")
        return [_render_todo_item_p3(t, i, list_type) for i, (_, t) in enumerate(items)]

    return render(active_items, "active"), render(completed_items, "completed"), render(trash_items, "trash")


# ── Page 3: Select all ────────────────────────────────────────────────────────
@app.callback(
    Output({"type": "todo-cb-p3-active", "index": ALL}, "value"),
    Input("p3-active-select-all", "value"),
    prevent_initial_call=True,
)
def select_all_p3_active(checked):
    return [bool(checked)] * len(ctx.outputs_list)


@app.callback(
    Output({"type": "todo-cb-p3-completed", "index": ALL}, "value"),
    Input("p3-completed-select-all", "value"),
    prevent_initial_call=True,
)
def select_all_p3_completed(checked):
    return [bool(checked)] * len(ctx.outputs_list)


@app.callback(
    Output({"type": "todo-cb-p3-trash", "index": ALL}, "value"),
    Input("p3-trash-select-all", "value"),
    prevent_initial_call=True,
)
def select_all_p3_trash(checked):
    return [bool(checked)] * len(ctx.outputs_list)


# ── Page 3: Checkbox states ───────────────────────────────────────────────────
@app.callback(
    Output("store-todo-checked-p3-active", "data"),
    Input({"type": "todo-cb-p3-active", "index": ALL}, "value"),
    State("store-todos-p3", "data"),
    State("p3-priority-filter", "value"),
    State("p3-sort-combo", "value"),
    prevent_initial_call=True,
)
def update_p3_active_checked(values, todos, filter_val, sort_val):
    items = _get_filtered_sorted(todos or [], filter_val, sort_val, ["active"])
    return [items[i][0] for i, v in enumerate(values) if v and i < len(items)]


@app.callback(
    Output("store-todo-checked-p3-completed", "data"),
    Input({"type": "todo-cb-p3-completed", "index": ALL}, "value"),
    State("store-todos-p3", "data"),
    State("p3-priority-filter", "value"),
    State("p3-sort-combo", "value"),
    prevent_initial_call=True,
)
def update_p3_completed_checked(values, todos, filter_val, sort_val):
    items = _get_filtered_sorted(todos or [], filter_val, sort_val, ["completed"])
    return [items[i][0] for i, v in enumerate(values) if v and i < len(items)]


@app.callback(
    Output("store-todo-checked-p3-trash", "data"),
    Input({"type": "todo-cb-p3-trash", "index": ALL}, "value"),
    State("store-todos-p3", "data"),
    State("p3-priority-filter", "value"),
    State("p3-sort-combo", "value"),
    prevent_initial_call=True,
)
def update_p3_trash_checked(values, todos, filter_val, sort_val):
    items = _get_filtered_sorted(todos or [], filter_val, sort_val, ["deleted"])
    return [items[i][0] for i, v in enumerate(values) if v and i < len(items)]


# ── Page 3: Notion 버튼 활성화/비활성화 ──────────────────────────────────────
@app.callback(
    Output("p3-btn-notion-sync", "disabled"),
    Output("p3-btn-notion-sync", "title"),
    Input("store-notion-enabled", "data"),
)
def toggle_notion_btn(notion_enabled):
    if notion_enabled:
        return False, ""
    return True, "Page 1에서 Notion 연동을 활성화하세요."


# ── Page 3: Notion 수동 동기화 (poll과 동일한 로직) ──────────────────────────
@app.callback(
    Output("p3-toast", "children", allow_duplicate=True),
    Output("p3-toast", "is_open", allow_duplicate=True),
    Output("store-todos-p3", "data", allow_duplicate=True),
    Output("store-notion-db-id", "data", allow_duplicate=True),
    Output("p3-notion-sync-status", "children", allow_duplicate=True),
    Output("store-notion-archive-queue", "data", allow_duplicate=True),
    Input("p3-btn-notion-sync", "n_clicks"),
    State("store-todos-p3", "data"),
    State("store-notion-key", "data"),
    State("store-notion-db", "data"),
    State("store-notion-db-id", "data"),
    State("store-perm-deleted-ids", "data"),
    State("store-notion-archive-queue", "data"),
    prevent_initial_call=True,
)
def sync_to_notion(n, todos, notion_key, notion_db, db_id, perm_deleted_ids, archive_queue):
    if not notion_key:
        return "Notion API 키가 없습니다. Page 1에서 입력 후 다시 시작해주세요.", True, no_update, no_update, no_update, no_update
    if not notion_db:
        return "Notion 페이지 URL이 없습니다. Page 1에서 입력 후 다시 시작해주세요.", True, no_update, no_update, no_update, no_update

    perm_deleted_set = set(perm_deleted_ids or [])
    todos = [t for t in (todos or []) if t.get("id") not in perm_deleted_set]
    archive_queue = list(archive_queue or [])

    if not todos and not archive_queue:
        return "동기화할 항목이 없습니다.", True, no_update, no_update, no_update, no_update

    try:
        from notion_sync import NotionSync
        syncer = NotionSync(api_key=notion_key, parent_page_id=notion_db)
        if db_id:
            syncer._db_id = db_id

        # ── 0. 아카이브 큐 처리 ───────────────────────────────────────────
        if archive_queue:
            remaining = []
            for page_id in archive_queue:
                try:
                    syncer.archive_page(page_id)
                except Exception:
                    remaining.append(page_id)
            archive_queue = remaining

        updated_todos = [dict(t) for t in todos]
        now = datetime.now().strftime("%H:%M")
        pushed_ids = set()

        # ── 1. 앱→Notion: pending_sync 항목 상태 push ────────────────────
        for todo in updated_todos:
            if todo.get("pending_sync") and todo.get("notion_page_id"):
                try:
                    syncer.update_page_status(todo["notion_page_id"], todo["status"])
                    pushed_ids.add(todo["id"])
                except Exception as e:
                    logger.warning(f"Notion 상태 push 실패: {e}")
                todo.pop("pending_sync", None)

        # ── 2. 새 항목 Notion에 생성 ─────────────────────────────────────
        new_items = [t for t in updated_todos if not t.get("notion_page_id")]
        new_count = 0
        if new_items:
            _, _, id_map, resolved_db_id = syncer.sync_all_todos(new_items)
            db_id = resolved_db_id
            new_count = len(id_map)
            for todo in updated_todos:
                if todo.get("id") in id_map:
                    todo["notion_page_id"] = id_map[todo["id"]]

        # ── 3. Notion→앱: 상태 pull ──────────────────────────────────────
        notion_id_to_todo_id = {
            t["notion_page_id"]: t["id"]
            for t in updated_todos
            if t.get("notion_page_id") and t["id"] not in pushed_ids
        }
        pull_count = 0
        if notion_id_to_todo_id:
            if not db_id:
                db_id = syncer.get_or_create_db()
            syncer._db_id = db_id
            changes = syncer.fetch_status_changes(db_id, notion_id_to_todo_id)
            for todo in updated_todos:
                if todo.get("id") in changes:
                    todo["status"] = changes[todo["id"]]
            pull_count = len(changes)

        msg = f"동기화 완료 · 신규 {new_count}건 · 상태변경 {pull_count}건 ({now})"
        return msg, True, updated_todos, db_id, f"Notion 연동 중 · 마지막 동기화: {now}", archive_queue

    except Exception as e:
        return f"Notion 오류: {e}", True, no_update, no_update, no_update, no_update


# ── Page 3: Notion polling ────────────────────────────────────────────────────
@app.callback(
    Output("store-todos-p3", "data", allow_duplicate=True),
    Output("store-notion-db-id", "data", allow_duplicate=True),
    Output("p3-notion-sync-status", "children", allow_duplicate=True),
    Output("store-notion-archive-queue", "data", allow_duplicate=True),
    Input("notion-poll-interval", "n_intervals"),
    State("store-todos-p3", "data"),
    State("store-notion-key", "data"),
    State("store-notion-db", "data"),
    State("store-notion-db-id", "data"),
    State("store-perm-deleted-ids", "data"),
    State("store-notion-archive-queue", "data"),
    prevent_initial_call=True,
)
def poll_notion(n_intervals, todos, notion_key, notion_db, db_id, perm_deleted_ids, archive_queue):
    if not notion_key or not notion_db:
        return no_update, no_update, no_update, no_update

    perm_deleted_set = set(perm_deleted_ids or [])
    todos = [t for t in (todos or []) if t.get("id") not in perm_deleted_set]
    forwarded = list(todos)
    archive_queue = list(archive_queue or [])

    try:
        from notion_sync import NotionSync
        syncer = NotionSync(api_key=notion_key, parent_page_id=notion_db)
        if db_id:
            syncer._db_id = db_id

        # ── 0. 아카이브 큐 처리: 영구 삭제된 Notion 페이지 재시도 ──────────
        if archive_queue:
            remaining = []
            for page_id in archive_queue:
                try:
                    syncer.archive_page(page_id)
                except Exception:
                    remaining.append(page_id)
            archive_queue = remaining

        if not forwarded:
            return no_update, no_update, no_update, archive_queue

        updated_todos = [dict(t) for t in todos]
        now = datetime.now().strftime("%H:%M")
        pushed_ids = set()

        # ── 1. 앱→Notion: pending_sync 항목 상태 먼저 push ────────────
        for todo in updated_todos:
            if todo.get("pending_sync") and todo.get("notion_page_id"):
                try:
                    syncer.update_page_status(todo["notion_page_id"], todo["status"])
                    pushed_ids.add(todo["id"])
                except Exception as e:
                    logger.warning(f"Notion 상태 push 실패: {e}")
                todo.pop("pending_sync", None)

        # ── 2. 새 forwarded 항목 Notion에 생성 ────────────────────────
        new_items = [t for t in updated_todos if not t.get("notion_page_id")]
        if new_items:
            _, _, id_map, resolved_db_id = syncer.sync_all_todos(new_items)
            db_id = resolved_db_id
            for todo in updated_todos:
                if todo.get("id") in id_map:
                    todo["notion_page_id"] = id_map[todo["id"]]

        # ── 3. Notion→앱: push한 항목 제외하고 상태 pull ──────────────
        notion_id_to_todo_id = {
            t["notion_page_id"]: t["id"]
            for t in updated_todos
            if t.get("notion_page_id") and t["id"] not in pushed_ids
        }
        if notion_id_to_todo_id:
            if not db_id:
                db_id = syncer.get_or_create_db()
            syncer._db_id = db_id
            changes = syncer.fetch_status_changes(db_id, notion_id_to_todo_id)
            for todo in updated_todos:
                if todo.get("id") in changes:
                    todo["status"] = changes[todo["id"]]

        return updated_todos, db_id, f"Notion 연동 중 · 마지막 확인: {now}", archive_queue
    except Exception as e:
        logger.error(f"Notion 폴링 오류: {e}")
        return no_update, no_update, no_update, no_update


# ── Page 3: Todo actions ──────────────────────────────────────────────────────
@app.callback(
    Output("store-todos-p3", "data", allow_duplicate=True),
    Output("p3-toast", "children", allow_duplicate=True),
    Output("p3-toast", "is_open", allow_duplicate=True),
    Output("store-perm-deleted-ids", "data", allow_duplicate=True),
    Output("store-notion-archive-queue", "data", allow_duplicate=True),
    Input("p3-btn-complete", "n_clicks"),
    Input("p3-btn-delete", "n_clicks"),
    Input("p3-btn-uncomplete", "n_clicks"),
    Input("p3-btn-del-completed", "n_clicks"),
    Input("p3-btn-restore", "n_clicks"),
    Input("p3-btn-perm-delete", "n_clicks"),
    State("store-todo-checked-p3-active", "data"),
    State("store-todo-checked-p3-completed", "data"),
    State("store-todo-checked-p3-trash", "data"),
    State("store-todos-p3", "data"),
    State("store-notion-key", "data"),
    State("store-notion-db", "data"),
    State("store-perm-deleted-ids", "data"),
    State("store-notion-archive-queue", "data"),
    prevent_initial_call=True,
)
def todo_actions_p3(n_complete, n_delete, n_uncomplete, n_del_comp, n_restore, n_perm,
                    checked_active, checked_completed, checked_trash, todos,
                    notion_key, notion_db, perm_deleted_ids, archive_queue):
    triggered = ctx.triggered_id
    if not todos:
        return no_update, "", False, no_update, no_update

    todos = [dict(t) for t in todos]
    perm_deleted_ids = list(perm_deleted_ids or [])

    def _mark_pending(idxs):
        for i in idxs:
            if todos[i].get("notion_page_id"):
                todos[i]["pending_sync"] = True

    if triggered == "p3-btn-complete" and checked_active:
        for i in checked_active:
            todos[i]["status"] = "completed"
        _mark_pending(checked_active)
        msg = f"{len(checked_active)}개 완료 처리됨"
    elif triggered == "p3-btn-delete" and checked_active:
        for i in checked_active:
            todos[i]["status"] = "deleted"
        _mark_pending(checked_active)
        msg = f"{len(checked_active)}개 삭제됨"
    elif triggered == "p3-btn-uncomplete" and checked_completed:
        for i in checked_completed:
            todos[i]["status"] = "active"
        _mark_pending(checked_completed)
        msg = f"{len(checked_completed)}개 미완료로 변경됨"
    elif triggered == "p3-btn-del-completed" and checked_completed:
        for i in checked_completed:
            todos[i]["status"] = "deleted"
        _mark_pending(checked_completed)
        msg = f"{len(checked_completed)}개 삭제됨"
    elif triggered == "p3-btn-restore" and checked_trash:
        for i in checked_trash:
            todos[i]["status"] = "active"
        _mark_pending(checked_trash)
        msg = f"{len(checked_trash)}개 복구됨"
    elif triggered == "p3-btn-perm-delete" and checked_trash:
        ids_to_delete = {todos[i]["id"] for i in checked_trash}
        page_ids = [todos[i]["notion_page_id"] for i in checked_trash if todos[i].get("notion_page_id")]
        # 아카이브 큐에 추가 (실패 시 다음 poll에서 재시도)
        updated_queue = list(set(list(archive_queue or []) + page_ids))
        todos = [t for t in todos if t["id"] not in ids_to_delete]
        perm_deleted_ids = list(set(perm_deleted_ids) | ids_to_delete)
        msg = f"{len(ids_to_delete)}개 영구 삭제됨"
        return todos, msg, True, perm_deleted_ids, updated_queue
    else:
        return no_update, "항목을 선택하세요.", True, no_update, no_update

    return todos, msg, True, no_update, no_update



# ── Cloud Sync: Todos ─────────────────────────────────────────────────────────
@app.callback(
    Output("p3-toast", "style"), # dummy output
    Input("store-todos", "data"),
    State("store-account", "data"),
    State("store-uid", "data"),
    State("store-auth-token", "data"),
    prevent_initial_call=True,
)
def sync_todos_to_cloud(todos, account, uid, id_token):
    if account and uid and id_token and todos is not None:
        try:
            fb_client.save_data(uid, id_token, "todos", account.replace(".", "_"), {"todos": todos})
        except Exception as e:
            print(f"Todo 클라우드 동기화 에러: {e}")
    return no_update


@app.callback(
    Output("store-notion-enabled", "data", allow_duplicate=True), # dummy output
    Input("store-todos-p3", "data"),
    State("store-account", "data"),
    State("store-uid", "data"),
    State("store-auth-token", "data"),
    prevent_initial_call=True,
)
def sync_todos_p3_to_cloud(todos, account, uid, id_token):
    if account and uid and id_token and todos is not None:
        try:
            fb_client.save_data(uid, id_token, "todos-p3", account.replace(".", "_"), {"todos": todos})
        except Exception as e:
            print(f"Todo P3 클라우드 동기화 에러: {e}")
    return no_update


# ── Cloud Sync: Seen-by-date ──────────────────────────────────────────────────
@app.callback(
    Output("store-notion-enabled", "data", allow_duplicate=True),  # dummy output
    Input("store-seen-by-date", "data"),
    State("store-account", "data"),
    State("store-uid", "data"),
    State("store-auth-token", "data"),
    prevent_initial_call=True,
)
def sync_seen_by_date_to_cloud(seen_by_date, account, uid, id_token):
    if account and uid and id_token and seen_by_date:
        def _save():
            try:
                fb_client.save_data(uid, id_token, "seen-by-date",
                                    account.replace(".", "_"), {"data": seen_by_date})
            except Exception as e:
                print(f"seen-by-date 클라우드 동기화 에러: {e}")
        threading.Thread(target=_save, daemon=True).start()
    return no_update

# ── Page 3: Edit modal ────────────────────────────────────────────────────────
@app.callback(
    Output("edit-modal", "is_open"),
    Output("edit-target-id", "data"),
    Output("edit-todo-text", "value"),
    Output("edit-todo-priority", "value"),
    Output("edit-due-date-cb", "value"),
    Output("edit-due-date", "date"),
    Input("p3-btn-edit", "n_clicks"),
    State("store-todo-checked-p3-active", "data"),
    State("store-todos-p3", "data"),
    prevent_initial_call=True,
)
def open_edit_modal(n, checked, todos):
    if not checked or not todos:
        return False, no_update, no_update, no_update, no_update, no_update
    idx = checked[0]
    todo = todos[idx]
    has_due = bool(todo.get("due_date"))
    return (
        True,
        todo["id"],
        todo.get("text", ""),
        todo.get("priority", "보통"),
        has_due,
        todo.get("due_date") or str(date.today()),
    )


@app.callback(
    Output("edit-due-date", "style"),
    Input("edit-due-date-cb", "value"),
)
def toggle_due_date(checked):
    return {"display": "block"} if checked else {"display": "none"}


@app.callback(
    Output("store-todos-p3", "data", allow_duplicate=True),
    Output("edit-modal", "is_open", allow_duplicate=True),
    Input("btn-edit-save", "n_clicks"),
    State("edit-target-id", "data"),
    State("edit-todo-text", "value"),
    State("edit-todo-priority", "value"),
    State("edit-due-date-cb", "value"),
    State("edit-due-date", "date"),
    State("store-todos-p3", "data"),
    prevent_initial_call=True,
)
def save_edit(n, target_id, text, priority, has_due, due_date, todos):
    if not todos or not target_id:
        return no_update, False
    todos = [dict(t) for t in todos]
    for t in todos:
        if t["id"] == target_id:
            t["text"] = text or t["text"]
            t["priority"] = priority or "보통"
            t["due_date"] = due_date if has_due else ""
            break
    return todos, False


@app.callback(
    Output("edit-modal", "is_open", allow_duplicate=True),
    Input("btn-edit-cancel", "n_clicks"),
    prevent_initial_call=True,
)
def close_edit_modal(n):
    return False


# ── Entry point ───────────────────────────────────────────────────────────────
def _create_tray_icon():
    from PIL import Image
    import pystray

    if getattr(sys, 'frozen', False):
        icon_path = os.path.join(sys._MEIPASS, 'assets', 'icon.png')
    else:
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets', 'icon.png')

    image = Image.open(icon_path)

    def on_open(icon, item):
        webbrowser.open("http://localhost:8050")

    def on_quit(icon, item):
        icon.stop()
        os._exit(0)

    menu = pystray.Menu(
        pystray.MenuItem("열기 (브라우저)", on_open),
        pystray.MenuItem("종료", on_quit),
    )
    return pystray.Icon("Email AI Summarizer", image, "Email AI Summarizer", menu)


if __name__ == "__main__":
    # Dash 서버를 백그라운드 스레드에서 실행
    server_thread = threading.Thread(
        target=lambda: app.run(debug=False, port=8050),
        daemon=True,
    )
    server_thread.start()

    # 1.5초 후 브라우저 자동 오픈
    threading.Timer(1.5, lambda: webbrowser.open("http://localhost:8050")).start()

    # 시스템 트레이 아이콘을 메인 스레드에서 실행 (Windows 필수)
    tray = _create_tray_icon()
    tray.run()
