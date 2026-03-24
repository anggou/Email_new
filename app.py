import dash
from dash import dcc, html, Input, Output, State, ALL, ctx, no_update, callback
import dash_bootstrap_components as dbc
import diskcache
import json
import os
import uuid
from datetime import date, datetime

import logging
from firebase_client import FirebaseClient
from dotenv import load_dotenv

logger = logging.getLogger(__name__)
load_dotenv()

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

# ── Setup ────────────────────────────────────────────────────────────────────
os.makedirs("cache", exist_ok=True)
cache = diskcache.Cache("./cache")
background_callback_manager = dash.DiskcacheManager(cache)

app = dash.Dash(
    __name__,
    external_stylesheets=[dbc.themes.BOOTSTRAP, dbc.icons.BOOTSTRAP],
    background_callback_manager=background_callback_manager,
    suppress_callback_exceptions=True,
    title="Email AI Summarizer",
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
PROFILE_FILE = "user_profile.json"

def load_all_profiles():
    if os.path.exists(PROFILE_FILE):
        try:
            with open(PROFILE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_all_profiles(profiles):
    with open(PROFILE_FILE, "w", encoding="utf-8") as f:
        json.dump(profiles, f, ensure_ascii=False, indent=2)

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

                        dbc.Label([
                            "Notion API 키",
                            html.A(" (키 발급/관리)", href="https://www.notion.so/my-integrations",
                                   target="_blank", className="ms-1 small text-primary"),
                        ], className="fw-semibold"),
                        dbc.InputGroup([
                            dbc.Input(id="notion-key-input", type="password", placeholder="secret_xxxxxx"),
                            dbc.Button("표시", id="btn-toggle-notion-key", color="secondary", outline=True, size="sm"),
                        ], className="mb-1"),
                        dbc.Input(id="notion-db-input", placeholder="Notion 페이지 URL (DB가 생성될 위치)",
                                  size="sm", className="mb-3"),

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
                                             className="mb-2", style={"fontSize": "0.85rem"}),
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
                         style={"fontSize": "0.72rem", "opacity": "0.5"}),
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
                    dbc.Col(dbc.Button("전체 TODO 관리 →", id="btn-goto-page3",
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
                        dbc.CardHeader([
                            dbc.Row([
                                dbc.Col(html.Span("메일 목록", className="fw-semibold small"), width="auto"),
                                dbc.Col(
                                    dbc.Checkbox(id="email-select-all-cb", label="전체선택",
                                                  label_class_name="small ms-1"),
                                    className="ms-auto", width="auto"),
                            ], align="center"),
                        ]),
                        dbc.CardBody([
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
                        dbc.CardHeader([
                            dbc.Row([
                                dbc.Col(html.Span("To-Do", className="fw-semibold small"), width="auto"),
                                dbc.Col([
                                    dbc.Button("전체선택", id="btn-todo-select-all", size="sm",
                                               color="secondary", outline=True, className="me-1"),
                                    dbc.Button("완료", id="btn-todo-complete", size="sm",
                                               color="success", outline=True, className="me-1"),
                                    dbc.Button("삭제", id="btn-todo-delete", size="sm",
                                               color="danger", outline=True, className="me-1"),
                                    dbc.Button("전체 TODO로", id="btn-todo-forward", size="sm",
                                               outline=True,
                                               style={"borderColor": "#7B1FA2", "color": "#7B1FA2"}),
                                ], className="ms-auto", width="auto"),
                            ], align="center"),
                        ]),
                        dbc.CardBody([
                            dbc.Tabs([
                                dbc.Tab(
                                    html.Div(id="todo-list-active",
                                             className="todo-list-container"),
                                    label="할 일", tab_id="tab-active",
                                ),
                                dbc.Tab(
                                    html.Div(id="todo-list-trash",
                                             className="todo-list-container"),
                                    label="휴지통", tab_id="tab-trash",
                                ),
                            ], id="todo-tabs", active_tab="tab-active"),
                        ], className="p-0 pt-1"),
                    ], className="card-clean"),
                ], md=7),
            ]),
        ], fluid=True),

        # ── Analyze modal ────────────────────────────────────────────────────
        dbc.Modal([
            dbc.ModalHeader(dbc.ModalTitle("AI 분석 중…")),
            dbc.ModalBody([
                dbc.Progress(id="analyze-progress-bar", value=0, striped=True, animated=True,
                             className="mb-3"),
                html.P(id="analyze-progress-text", className="text-center text-muted small"),
                dbc.Button("취소", id="btn-analyze-cancel", color="danger", outline=True,
                           size="sm", className="d-block mx-auto"),
            ]),
        ], id="analyze-modal", is_open=False, backdrop="static", centered=True),

        html.Div("[ Page 2 — 메일 조회 / AI 요약 / 이메일 TODO ]",
                 style={"textAlign": "center", "fontSize": "0.72rem",
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
                            style={"minWidth": 100, "fontSize": "0.85rem"},
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
                            style={"minWidth": 110, "fontSize": "0.85rem"},
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
                                    dbc.Col(dbc.Checkbox(id="p3-active-select-all", label="전체선택",
                                                          label_class_name="small ms-1"), width="auto"),
                                    dbc.Col([
                                        dbc.Button("완료 처리", id="p3-btn-complete", size="sm",
                                                   color="success", outline=True, className="me-1"),
                                        dbc.Button("삭제", id="p3-btn-delete", size="sm",
                                                   color="danger", outline=True, className="me-1"),
                                        dbc.Button("편집", id="p3-btn-edit", size="sm",
                                                   color="primary", outline=True, className="me-1"),
                                        dbc.Button("Notion 동기화", id="p3-btn-notion-sync", size="sm",
                                                   style={"backgroundColor": "#000", "borderColor": "#000",
                                                          "color": "white"}, className="me-1"),
                                        html.Span(id="p3-notion-sync-status",
                                                  className="text-muted small align-self-center",
                                                  style={"fontSize": "0.75rem"}),
                                    ], className="ms-auto", width="auto"),
                                ], align="center", className="mb-2 g-1"),
                                html.Div(id="p3-active-list", className="todo-list-container"),
                            ]),
                            label="할 일", tab_id="p3-tab-active",
                        ),
                        dbc.Tab(
                            html.Div([
                                dbc.Row([
                                    dbc.Col(dbc.Checkbox(id="p3-completed-select-all", label="전체선택",
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
                                    dbc.Col(dbc.Checkbox(id="p3-trash-select-all", label="전체선택",
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
                              style={"background": "#f5f5f5", "fontSize": "0.85rem"},
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

        html.Div("[ Page 3 — 전체 TODO 관리 / Notion 동기화 ]",
                 style={"textAlign": "center", "fontSize": "0.72rem",
                        "opacity": "0.5", "padding": "6px 0", "color": "#aaa"}),

        dbc.Toast(
            id="p3-toast",
            header="알림",
            is_open=False,
            dismissable=True,
            duration=3000,
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
    dcc.Store(id="store-todo-checked-p3-active", data=[]),
    dcc.Store(id="store-todo-checked-p3-completed", data=[]),
    dcc.Store(id="store-todo-checked-p3-trash", data=[]),
    dcc.Store(id="store-analyze-done", data=0),
    dcc.Store(id="store-profile-lists", data={"projects":[],"superiors":[],"peers":[],"subordinates":[],"clients":[]}),
    dcc.Store(id="store-notion-key", data=""),
    dcc.Store(id="store-notion-db", data=""),
    dcc.Store(id="store-notion-db-id", data=""),
    dcc.Interval(id="notion-poll-interval", interval=60*1000, n_intervals=0, disabled=False),

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
    Input("store-page", "data"),
)
def toggle_pages(page):
    show = {"display": "block"}
    hide = {"display": "none"}
    return (
        show if page == 0 else hide,
        show if page == 1 else hide,
        show if page == 2 else hide,
        show if page == 3 else hide,
    )

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
    Input("account-dropdown", "value"),
    State("store-uid", "data"),
    State("store-auth-token", "data"),
    prevent_initial_call=True,
)
def load_profile(account, uid, id_token):
    p = {}
    keys = {}
    todos_data = no_update
    if uid and id_token and account:
        safe_acc = account.replace(".", "_")
        try:
            cloud_p = fb_client.get_data(uid, id_token, "profiles", safe_acc)
            if cloud_p: p = cloud_p
            cloud_keys = fb_client.get_data(uid, id_token, "keys", safe_acc)
            if cloud_keys: keys = cloud_keys
            cloud_todos = fb_client.get_data(uid, id_token, "todos", safe_acc)
            if cloud_todos and "todos" in cloud_todos: todos_data = cloud_todos["todos"]
        except Exception as e:
            print(f"Firebase 로드 실패: {e}")

    if not p:
        profiles = load_all_profiles()
        p = profiles.get(account, {})
        
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
        todos_data
    )


# ── Page 1: Save profile ─────────────────────────────────────────────────────
@app.callback(
    Output("profile-save-status", "children"),
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
        return "계정을 먼저 선택하세요."
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
    
    profiles = load_all_profiles()
    profiles[account] = profile_data
    save_all_profiles(profiles)

    if uid and id_token:
        fb_save_profile(uid, id_token, profile_data)

    return "저장됨 ✓"


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
        return html.Span("항목 없음", className="text-muted", style={"fontSize": "0.75rem"})
    return [
        html.Span([
            dbc.Badge(item, color="secondary", className="fw-normal me-1",
                      style={"fontSize": "0.78rem"}),
            html.Span("×", id={"type": "pf-del", "field": field, "index": i},
                      n_clicks=0,
                      style={"cursor": "pointer", "color": "#999", "fontSize": "0.85rem",
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
    Output("page1-error", "children"),
    Input("btn-next-page2", "n_clicks"),
    State("account-dropdown", "value"),
    State("api-key-input", "value"),
    State("notion-key-input", "value"),
    State("notion-db-input", "value"),
    State("store-uid", "data"),
    State("store-auth-token", "data"),
    prevent_initial_call=True,
)
def go_to_page2(n, account, api_key, notion_key, notion_db, uid, id_token):
    if not account:
        return no_update, no_update, no_update, no_update, no_update, no_update, "계정을 선택해주세요."
    if not api_key or api_key.strip() == "":
        return no_update, no_update, no_update, no_update, no_update, no_update, "Gemini API 키를 입력해주세요."
        
    if uid and id_token:
        fb_save_keys(uid, id_token, api_key.strip(),
                     (notion_key or "").strip(), (notion_db or "").strip())

    profiles = load_all_profiles()
    profile = profiles.get(account, {})
    return 2, account, api_key.strip(), profile, (notion_key or "").strip(), (notion_db or "").strip(), ""


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
    Input("btn-refresh", "n_clicks"),
    State("store-account", "data"),
    State("p2-date-picker", "date"),
    State("store-uid", "data"),
    State("store-auth-token", "data"),
    prevent_initial_call=True,
)
def fetch_emails(n, account, date_str, uid, id_token):
    if not account:
        return [], "계정을 선택하세요."
    try:
        from outlook_manager import OutlookManager
        mgr = OutlookManager()
        target_date = datetime.strptime(date_str, "%Y-%m-%d").date() if date_str else None
        
        # 1. 클라우드에서 이전 내역 불러오기
        cloud_emails = []
        safe_acc = account.replace(".", "_")
        doc_id = f"{safe_acc}_{date_str}"
        if uid and id_token:
            try:
                res = fb_client.get_data(uid, id_token, "emails", doc_id)
                if res and "emails" in res:
                    cloud_emails = res["emails"]
            except Exception as e:
                print(f"이메일 동기화 로드 실패: {e}")
                
        # 2. 아웃룩에서 가져오기
        outlook_emails = mgr.get_emails_by_date(account_email=account, target_date=target_date, limit=100)
        
        # 3. 중복 병합 로직 (간단히 제목+보낸사람 기준으로 확인)
        existing_keys = {f"{e.get('subject')}_{e.get('sender')}" for e in cloud_emails}
        for e in outlook_emails:
            k = f"{e.get('subject')}_{e.get('sender')}"
            if k not in existing_keys:
                cloud_emails.append(e)
                existing_keys.add(k)
                
        # 4. 다시 클라우드에 백업
        if uid and id_token:
            try:
                fb_client.save_data(uid, id_token, "emails", doc_id, {"emails": cloud_emails})
            except Exception as e:
                print(f"이메일 동기화 저장 실패: {e}")
                
        return cloud_emails, f"클라우드 연동 완료: 총 {len(cloud_emails)}개 메일"
    except Exception as e:
        return [], f"오류: {e}"


# ── Page 2: Render email list ─────────────────────────────────────────────────
@app.callback(
    Output("email-list-container", "children"),
    Input("store-emails", "data"),
    Input("store-analyze-done", "data"),
)
def render_email_list(emails, _):
    if not emails:
        return html.Span("이메일을 조회하세요.", className="text-muted small p-2 d-block")

    items = []
    for i, email in enumerate(emails):
        subject = email.get("subject", "제목 없음")
        sender = email.get("sender", "알 수 없음")
        items.append(
            html.Div([
                dbc.Row([
                    dbc.Col(
                        dbc.Checkbox(
                            id={"type": "email-checkbox", "index": i},
                            value=False,
                        ),
                        width="auto",
                        className="pe-1",
                    ),
                    dbc.Col(
                        html.Div(
                            [
                                html.Span(sender, className="fw-semibold small d-block"),
                                html.Span(subject[:55] + ("…" if len(subject) > 55 else ""),
                                           className="text-muted small"),
                            ],
                            id={"type": "email-row", "index": i},
                            n_clicks=0,
                            style={"cursor": "pointer"},
                        ),
                    ),
                ], className="g-0 align-items-center"),
            ], className="email-item px-2 py-1")
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


# ── Page 2: Select all emails ─────────────────────────────────────────────────
@app.callback(
    Output({"type": "email-checkbox", "index": ALL}, "value"),
    Input("email-select-all-cb", "value"),
    State("store-emails", "data"),
    prevent_initial_call=True,
)
def select_all_emails(checked, emails):
    count = len(emails) if emails else 0
    return [checked] * count


# ── Page 2: AI Analysis (background) ─────────────────────────────────────────
@app.callback(
    output=[
        Output("store-todos", "data"),
        Output("store-analyze-done", "data"),
        Output("p2-toast", "children"),
        Output("p2-toast", "is_open"),
    ],
    inputs=Input("btn-analyze", "n_clicks"),
    state=[
        State("store-email-checked", "data"),
        State("store-emails", "data"),
        State("store-user-profile", "data"),
        State("store-api-key", "data"),
        State("store-todos", "data"),
        State("store-analyze-done", "data"),
    ],
    background=True,
    running=[
        (Output("analyze-modal", "is_open"), True, False),
        (Output("btn-analyze", "disabled"), True, False),
    ],
    progress=[
        Output("analyze-progress-text", "children"),
        Output("analyze-progress-bar", "value"),
    ],
    cancel=Input("btn-analyze-cancel", "n_clicks"),
    prevent_initial_call=True,
)
def analyze_emails(set_progress, n_clicks, checked_indices, emails, user_profile, api_key, todos, done_count):
    if not checked_indices or not emails:
        return todos or [], done_count, "분석할 메일을 선택하세요.", True

    set_progress(("준비 중…", 0))
    results = list(todos or [])

    try:
        from ai_processor import AIProcessor
        processor = AIProcessor(override_api_key=api_key)
    except Exception as e:
        return results, done_count, f"API 오류: {e}", True

    user_context = build_user_context(user_profile) if user_profile else ""
    total = len(checked_indices)

    errors = []
    for i, idx in enumerate(checked_indices):
        if idx >= len(emails):
            continue
        email = emails[idx]
        progress_pct = int((i / total) * 100)
        set_progress((f"{i + 1}/{total} 분석 중… {email.get('subject', '')[:30]}", progress_pct))

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
            with open("error_log.txt", "a", encoding="utf-8") as f:
                f.write(err_detail + "\n")
            continue

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
                "status": "active",
                "forwarded": False,
                "due_date": "",
                "priority": priority,
                "mail_type": mail_type,
            })

    set_progress((f"완료! {len(checked_indices)}개 분석됨", 100))
    if errors:
        all_errors = "\n".join(errors)
        err_msg = f"오류 {len(errors)}건 발생:\n{all_errors}"
        return results, (done_count or 0) + 1, err_msg, True
    return results, (done_count or 0) + 1, f"분석 완료: {total}개 이메일", True


# ── Page 2: Render Todo lists ─────────────────────────────────────────────────
def _render_todo_item_p2(todo, idx):
    status = todo.get("status", "active")
    forwarded = todo.get("forwarded", False)
    text = todo.get("text", "")
    subject = todo.get("email_subject", "")

    if status == "deleted":
        text_style = {"textDecoration": "line-through", "color": "#bbb"}
    elif status == "completed":
        text_style = {"textDecoration": "line-through", "color": "#1976D2"}
    elif forwarded:
        text_style = {"color": "#7B1FA2"}
    else:
        text_style = {"color": "#212121"}

    return html.Div([
        dbc.Row([
            dbc.Col(
                dbc.Checkbox(id={"type": "todo-cb-p2", "index": idx}, value=False),
                width="auto",
            ),
            dbc.Col([
                html.Span(text[:60] + ("…" if len(text) > 60 else ""),
                           className="small d-block", style=text_style),
                html.Span(f"출처: {subject[:30]}", className="text-muted",
                           style={"fontSize": "0.75rem"}),
            ]),
        ], className="g-0 align-items-start"),
    ], className="todo-item px-2 py-1 border-bottom")


@app.callback(
    Output("todo-list-active", "children"),
    Input("store-todos", "data"),
    Input("store-analyze-done", "data"),
)
def render_active_todos(todos, _):
    if not todos:
        return html.Span("Todo 없음", className="text-muted small p-3 d-block")
    items = [
        _render_todo_item_p2(t, i)
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
        _render_todo_item_p2(t, i)
        for i, t in enumerate(todos)
        if t.get("status") == "deleted"
    ]
    return items if items else html.Span("휴지통 비어있음", className="text-muted small p-3 d-block")


# ── Page 2: Todo checkbox state ───────────────────────────────────────────────
@app.callback(
    Output("store-todo-checked-p2", "data"),
    Input({"type": "todo-cb-p2", "index": ALL}, "value"),
    State("store-todos", "data"),
    prevent_initial_call=True,
)
def update_todo_checked_p2(values, todos):
    if not todos:
        return []
    active_indices = [i for i, t in enumerate(todos) if t.get("status") in ("active", "completed")]
    checked = []
    val_idx = 0
    for i, t in enumerate(todos):
        if t.get("status") in ("active", "completed"):
            if val_idx < len(values) and values[val_idx]:
                checked.append(i)
            val_idx += 1
    return checked


# ── Page 2: Todo select all ───────────────────────────────────────────────────
@app.callback(
    Output({"type": "todo-cb-p2", "index": ALL}, "value"),
    Input("btn-todo-select-all", "n_clicks"),
    State("store-todos", "data"),
    prevent_initial_call=True,
)
def select_all_todos_p2(n, todos):
    if not todos:
        return []
    outputs = ctx.outputs_list
    if not outputs:
        return []
    return [
        True if todos[o["id"]["index"]].get("status") in ("active", "completed") else False
        for o in outputs
    ]


# ── Page 2: Todo actions ──────────────────────────────────────────────────────
@app.callback(
    Output("store-todos", "data", allow_duplicate=True),
    Output("p2-toast", "children", allow_duplicate=True),
    Output("p2-toast", "is_open", allow_duplicate=True),
    Input("btn-todo-complete", "n_clicks"),
    Input("btn-todo-delete", "n_clicks"),
    Input("btn-todo-forward", "n_clicks"),
    State("store-todo-checked-p2", "data"),
    State("store-todos", "data"),
    prevent_initial_call=True,
)
def todo_actions_p2(n_complete, n_delete, n_forward, checked, todos):
    if not todos:
        return no_update, "항목이 없습니다.", True

    triggered = ctx.triggered_id
    todos = [dict(t) for t in todos]

    # 전달 버튼은 체크 없으면 전체 active 항목 전달
    if triggered == "btn-todo-forward":
        if checked:
            targets = checked
        else:
            targets = [i for i, t in enumerate(todos) if t.get("status") == "active"]
        if not targets:
            return no_update, "전달할 항목이 없습니다.", True
        for i in targets:
            todos[i]["forwarded"] = True
        return todos, f"{len(targets)}개 전체 TODO로 전달됨", True

    # 완료/삭제는 체크 필수
    if not checked:
        return no_update, "항목을 선택하세요.", True

    if triggered == "btn-todo-complete":
        for i in checked:
            if todos[i]["status"] == "active":
                todos[i]["status"] = "completed"
            elif todos[i]["status"] == "completed":
                todos[i]["status"] = "active"
            if todos[i].get("notion_page_id"):
                todos[i]["pending_sync"] = True
        msg = f"{len(checked)}개 완료 처리됨"
    elif triggered == "btn-todo-delete":
        for i in checked:
            todos[i]["status"] = "deleted"
            if todos[i].get("notion_page_id"):
                todos[i]["pending_sync"] = True
        msg = f"{len(checked)}개 삭제됨"
    else:
        return no_update, "", False

    return todos, msg, True


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
                     className="me-1", style={"fontSize": "0.7rem"})


def _render_todo_item_p3(todo, idx, list_type="active"):
    text = todo.get("text", "")
    priority = todo.get("priority", "보통")
    due_date = todo.get("due_date", "")
    subject = todo.get("email_subject", "")
    status = todo.get("status", "active")

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
                html.Span(
                    f"출처: {subject[:30]}" + (f" | 마감: {due_date}" if due_date else ""),
                    style={"fontSize": "0.72rem", "color": "#888"},
                ),
            ]),
        ], className="g-0 align-items-start"),
    ], className="todo-item px-2 py-1 border-bottom")


def _get_filtered_sorted(todos, filter_val, sort_val, statuses):
    items = [(i, t) for i, t in enumerate(todos) if t.get("status") in statuses and t.get("forwarded", False)]
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
    Input("store-todos", "data"),
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


# ── Page 3: Checkbox states ───────────────────────────────────────────────────
@app.callback(
    Output("store-todo-checked-p3-active", "data"),
    Input({"type": "todo-cb-p3-active", "index": ALL}, "value"),
    State("store-todos", "data"),
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
    State("store-todos", "data"),
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
    State("store-todos", "data"),
    State("p3-priority-filter", "value"),
    State("p3-sort-combo", "value"),
    prevent_initial_call=True,
)
def update_p3_trash_checked(values, todos, filter_val, sort_val):
    items = _get_filtered_sorted(todos or [], filter_val, sort_val, ["deleted"])
    return [items[i][0] for i, v in enumerate(values) if v and i < len(items)]


# ── Page 3: Notion sync ───────────────────────────────────────────────────────
@app.callback(
    Output("p3-toast", "children", allow_duplicate=True),
    Output("p3-toast", "is_open", allow_duplicate=True),
    Output("store-todos", "data", allow_duplicate=True),
    Output("store-notion-db-id", "data"),
    Output("p3-notion-sync-status", "children"),
    Input("p3-btn-notion-sync", "n_clicks"),
    State("store-todos", "data"),
    State("store-todo-checked-p3-active", "data"),
    State("store-notion-key", "data"),
    State("store-notion-db", "data"),
    prevent_initial_call=True,
)
def sync_to_notion(n, todos, checked, notion_key, notion_db):
    if not notion_key:
        return "Notion API 키가 없습니다. Page 1에서 입력 후 다시 시작해주세요.", True, no_update, no_update, no_update
    if not notion_db:
        return "Notion 페이지 URL이 없습니다. Page 1에서 입력 후 다시 시작해주세요.", True, no_update, no_update, no_update

    todos = todos or []
    if checked:
        target = [todos[i] for i in checked if i < len(todos)]
    else:
        target = [t for t in todos if t.get("status") == "active" and t.get("forwarded")]

    if not target:
        return "동기화할 항목이 없습니다.", True, no_update, no_update, no_update

    try:
        from notion_sync import NotionSync
        syncer = NotionSync(api_key=notion_key, parent_page_id=notion_db)
        success, failed, id_map, db_id = syncer.sync_all_todos(target)

        updated_todos = list(todos)
        for todo in updated_todos:
            if todo.get("id") in id_map:
                todo["notion_page_id"] = id_map[todo["id"]]

        now = datetime.now().strftime("%H:%M")
        if failed:
            return f"Notion 동기화 실패: {failed[0]}", True, no_update, no_update, no_update

        return (f"Notion 동기화 완료: {success}건 저장됨", True,
                updated_todos, db_id, f"Notion 연동 중 · 마지막 동기화: {now}")
    except Exception as e:
        return f"Notion 오류: {e}", True, no_update, no_update, no_update


# ── Page 3: Notion polling ────────────────────────────────────────────────────
@app.callback(
    Output("store-todos", "data", allow_duplicate=True),
    Output("store-notion-db-id", "data", allow_duplicate=True),
    Output("p3-notion-sync-status", "children", allow_duplicate=True),
    Input("notion-poll-interval", "n_intervals"),
    State("store-todos", "data"),
    State("store-notion-key", "data"),
    State("store-notion-db", "data"),
    State("store-notion-db-id", "data"),
    prevent_initial_call=True,
)
def poll_notion(n_intervals, todos, notion_key, notion_db, db_id):
    if not notion_key or not notion_db:
        return no_update, no_update, no_update

    todos = todos or []
    forwarded = [t for t in todos if t.get("forwarded")]
    if not forwarded:
        return no_update, no_update, no_update

    try:
        from notion_sync import NotionSync
        syncer = NotionSync(api_key=notion_key, parent_page_id=notion_db)
        if db_id:
            syncer._db_id = db_id

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
        new_items = [t for t in updated_todos if t.get("forwarded") and not t.get("notion_page_id")]
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

        return updated_todos, db_id, f"Notion 연동 중 · 마지막 확인: {now}"
    except Exception as e:
        logger.error(f"Notion 폴링 오류: {e}")
        return no_update, no_update, no_update


# ── Page 3: Todo actions ──────────────────────────────────────────────────────
@app.callback(
    Output("store-todos", "data", allow_duplicate=True),
    Output("p3-toast", "children"),
    Output("p3-toast", "is_open"),
    Input("p3-btn-complete", "n_clicks"),
    Input("p3-btn-delete", "n_clicks"),
    Input("p3-btn-uncomplete", "n_clicks"),
    Input("p3-btn-del-completed", "n_clicks"),
    Input("p3-btn-restore", "n_clicks"),
    Input("p3-btn-perm-delete", "n_clicks"),
    State("store-todo-checked-p3-active", "data"),
    State("store-todo-checked-p3-completed", "data"),
    State("store-todo-checked-p3-trash", "data"),
    State("store-todos", "data"),
    prevent_initial_call=True,
)
def todo_actions_p3(n_complete, n_delete, n_uncomplete, n_del_comp, n_restore, n_perm,
                    checked_active, checked_completed, checked_trash, todos):
    triggered = ctx.triggered_id
    if not todos:
        return no_update, "", False

    todos = [dict(t) for t in todos]

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
        todos = [t for t in todos if t["id"] not in ids_to_delete]
        msg = f"{len(ids_to_delete)}개 영구 삭제됨"
    else:
        return no_update, "항목을 선택하세요.", True

    return todos, msg, True


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
    State("store-todos", "data"),
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
    Output("store-todos", "data", allow_duplicate=True),
    Output("edit-modal", "is_open", allow_duplicate=True),
    Input("btn-edit-save", "n_clicks"),
    State("edit-target-id", "data"),
    State("edit-todo-text", "value"),
    State("edit-todo-priority", "value"),
    State("edit-due-date-cb", "value"),
    State("edit-due-date", "date"),
    State("store-todos", "data"),
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
if __name__ == "__main__":
    app.run(debug=True, port=8050)
