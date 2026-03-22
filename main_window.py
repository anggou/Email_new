from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QListWidget, QListWidgetItem,
    QTextEdit, QMessageBox, QSplitter, QComboBox,
    QStackedWidget, QDateEdit, QDialog, QTabWidget, QCheckBox,
    QLineEdit, QScrollArea, QInputDialog, QFrame
)
from PySide6.QtCore import Qt, QThread, Signal, QDate
from PySide6.QtGui import QColor, QFont
from outlook_manager import OutlookManager
from ai_processor import AIProcessor
import datetime
import uuid
import winsound
import time

class UserProfileDialog(QDialog):
    """내 정보 및 조직 구조를 입력/저장하는 다이얼로그."""

    def __init__(self, parent, profile, account_email=""):
        super().__init__(parent)
        self.setWindowTitle("내 정보 입력")
        self.setMinimumWidth(560)
        self.setMinimumHeight(660)
        self.account_email = account_email
        # 작업용 복사본
        self.profile = {
            "name":        profile.get("name", ""),
            "email":       account_email or profile.get("email", ""),
            "role":        profile.get("role", ""),
            "projects":    list(profile.get("projects", [])),
            "superiors":   list(profile.get("superiors", [])),
            "peers":       list(profile.get("peers", [])),
            "subordinates": list(profile.get("subordinates", [])),
            "clients":     [dict(c) for c in profile.get("clients", [])],
        }

        main_layout = QVBoxLayout(self)

        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("font-size: 10pt;")
        self._init_tab1()
        self._init_tab2()
        main_layout.addWidget(self.tabs)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        save_btn = QPushButton("저장")
        save_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 8pt 20pt;")
        save_btn.clicked.connect(self._save_and_accept)
        cancel_btn = QPushButton("취소")
        cancel_btn.setStyleSheet("padding: 8pt 20pt;")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(cancel_btn)
        main_layout.addLayout(btn_layout)

    # ── Tab 1: 내 정보 ──────────────────────────────────────
    def _init_tab1(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setSpacing(10)

        def row(label_text, widget):
            h = QHBoxLayout()
            lbl = QLabel(label_text)
            lbl.setFixedWidth(80)
            h.addWidget(lbl)
            h.addWidget(widget)
            return h

        self.name_edit = QLineEdit(self.profile["name"])
        self.name_edit.setPlaceholderText("예: 홍길동")
        self.name_edit.setMinimumHeight(35)
        layout.addLayout(row("이름:", self.name_edit))

        self.email_edit = QLineEdit(self.profile["email"])
        self.email_edit.setMinimumHeight(35)
        self.email_edit.setReadOnly(True)
        self.email_edit.setStyleSheet("background-color: #f0f0f0; color: #555555; font-size: 10pt;")
        layout.addLayout(row("이메일:", self.email_edit))

        self.role_edit = QLineEdit(self.profile["role"])
        self.role_edit.setPlaceholderText("예: 프로젝트 매니저")
        self.role_edit.setMinimumHeight(35)
        layout.addLayout(row("역할:", self.role_edit))

        layout.addSpacing(6)
        layout.addWidget(QLabel("<b>현재 참여 프로젝트</b>"))
        self.projects_list = QListWidget()
        self.projects_list.setMaximumHeight(140)
        self.projects_list.setStyleSheet("font-size: 10pt;")
        for p in self.profile["projects"]:
            self.projects_list.addItem(p)
        layout.addWidget(self.projects_list)

        pb = QHBoxLayout()
        add_p = QPushButton("추가")
        add_p.clicked.connect(self._add_project)
        del_p = QPushButton("삭제")
        del_p.clicked.connect(lambda: self._delete_item(self.projects_list))
        pb.addWidget(add_p); pb.addWidget(del_p); pb.addStretch()
        layout.addLayout(pb)
        layout.addStretch()

        self.tabs.addTab(tab, "내 정보")

    # ── Tab 2: 조직 구조 ────────────────────────────────────
    def _init_tab2(self):
        tab = QWidget()
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)

        content = QWidget()
        layout = QVBoxLayout(content)
        layout.setSpacing(8)

        def add_section(title, list_attr, items):
            layout.addWidget(QLabel(f"<b>{title}</b>"))
            lw = QListWidget()
            lw.setMaximumHeight(90)
            lw.setStyleSheet("font-size: 10pt;")
            for it in items:
                lw.addItem(it)
            setattr(self, list_attr, lw)
            layout.addWidget(lw)
            btns = QHBoxLayout()
            add_btn = QPushButton("추가")
            add_btn.clicked.connect(lambda _, lw=lw: self._add_person(lw))
            del_btn = QPushButton("삭제")
            del_btn.clicked.connect(lambda _, lw=lw: self._delete_item(lw))
            btns.addWidget(add_btn); btns.addWidget(del_btn); btns.addStretch()
            layout.addLayout(btns)

        add_section("상사",      "superiors_list",    self.profile["superiors"])
        add_section("동료",      "peers_list",        self.profile["peers"])
        add_section("부하직원",  "subordinates_list", self.profile["subordinates"])

        layout.addSpacing(4)
        layout.addWidget(QLabel("<b>Client</b> (이름 + 연관 프로젝트)"))
        self.clients_list = QListWidget()
        self.clients_list.setMaximumHeight(110)
        self.clients_list.setStyleSheet("font-size: 10pt;")
        for c in self.profile["clients"]:
            item = QListWidgetItem(self._client_display(c))
            item.setData(Qt.UserRole, c)
            self.clients_list.addItem(item)
        layout.addWidget(self.clients_list)

        cb = QHBoxLayout()
        add_c = QPushButton("추가")
        add_c.clicked.connect(self._add_client)
        del_c = QPushButton("삭제")
        del_c.clicked.connect(lambda: self._delete_item(self.clients_list))
        cb.addWidget(add_c); cb.addWidget(del_c); cb.addStretch()
        layout.addLayout(cb)
        layout.addStretch()

        scroll.setWidget(content)
        tab_layout = QVBoxLayout(tab)
        tab_layout.setContentsMargins(0, 0, 0, 0)
        tab_layout.addWidget(scroll)
        self.tabs.addTab(tab, "조직 구조")

    # ── 헬퍼 ────────────────────────────────────────────────
    def _client_display(self, c):
        projs = c.get("projects", [])
        return f"{c['name']} ({', '.join(projs)})" if projs else c["name"]

    def _add_project(self):
        text, ok = QInputDialog.getText(self, "프로젝트 추가", "프로젝트명:")
        if ok and text.strip():
            self.projects_list.addItem(text.strip())

    def _add_person(self, list_widget):
        text, ok = QInputDialog.getText(self, "추가", "이름:")
        if ok and text.strip():
            list_widget.addItem(text.strip())

    def _delete_item(self, list_widget):
        row = list_widget.currentRow()
        if row >= 0:
            list_widget.takeItem(row)
        else:
            QMessageBox.information(self, "선택 안내", "삭제할 항목을 클릭하여 선택해 주세요.")

    def _add_client(self):
        projects = [self.projects_list.item(i).text() for i in range(self.projects_list.count())]

        dlg = QDialog(self)
        dlg.setWindowTitle("Client 추가")
        dlg.setMinimumWidth(350)
        dlg_layout = QVBoxLayout(dlg)

        name_row = QHBoxLayout()
        name_row.addWidget(QLabel("클라이언트 이름:"))
        name_edit = QLineEdit()
        name_edit.setMinimumHeight(35)
        name_row.addWidget(name_edit)
        dlg_layout.addLayout(name_row)

        dlg_layout.addWidget(QLabel("연관 프로젝트 (복수 선택 가능):"))
        check_boxes = []
        if projects:
            for p in projects:
                cb = QCheckBox(p)
                dlg_layout.addWidget(cb)
                check_boxes.append(cb)
        else:
            dlg_layout.addWidget(QLabel("(내 정보 탭에서 프로젝트를 먼저 추가해 주세요)"))

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        ok_btn = QPushButton("추가")
        ok_btn.setStyleSheet("background-color: #4CAF50; color: white; padding: 6pt 15pt;")
        ok_btn.clicked.connect(dlg.accept)
        cancel_btn = QPushButton("취소")
        cancel_btn.clicked.connect(dlg.reject)
        btn_row.addWidget(ok_btn); btn_row.addWidget(cancel_btn)
        dlg_layout.addLayout(btn_row)

        if dlg.exec() == QDialog.Accepted:
            name = name_edit.text().strip()
            if name:
                selected = [cb.text() for cb in check_boxes if cb.isChecked()]
                client_data = {"name": name, "projects": selected}
                item = QListWidgetItem(self._client_display(client_data))
                item.setData(Qt.UserRole, client_data)
                self.clients_list.addItem(item)

    def _save_and_accept(self):
        self.profile["name"]  = self.name_edit.text().strip()
        self.profile["email"] = self.account_email  # 계정 이메일은 변경 불가
        self.profile["role"]  = self.role_edit.text().strip()
        self.profile["projects"]     = [self.projects_list.item(i).text() for i in range(self.projects_list.count())]
        self.profile["superiors"]    = [self.superiors_list.item(i).text() for i in range(self.superiors_list.count())]
        self.profile["peers"]        = [self.peers_list.item(i).text() for i in range(self.peers_list.count())]
        self.profile["subordinates"] = [self.subordinates_list.item(i).text() for i in range(self.subordinates_list.count())]
        self.profile["clients"]      = [self.clients_list.item(i).data(Qt.UserRole) for i in range(self.clients_list.count())]
        self.accept()

    def get_profile(self):
        return self.profile


class EmailFetchThread(QThread):
    finished = Signal(list)
    error = Signal(str)

    def __init__(self, account_email=None, target_date=None):
        super().__init__()
        self.account_email = account_email
        self.target_date = target_date

    def run(self):
        try:
            manager = OutlookManager()
            emails = manager.get_emails_by_date(self.account_email, self.target_date)
            self.finished.emit(emails)
        except Exception as e:
            self.error.emit(str(e))

class EmailAnalyzeThread(QThread):
    finished = Signal(list)
    error = Signal(str)
    progress = Signal(str)

    def __init__(self, emails_list, api_key, user_context=""):
        super().__init__()
        self.emails_list = emails_list
        self.api_key = api_key
        self.user_context = user_context
        self._is_running = True

    def stop(self):
        self._is_running = False

    def run(self):
        try:
            processor = AIProcessor(override_api_key=self.api_key)
            results = []
            total = len(self.emails_list)
            for idx, email_data in enumerate(self.emails_list):
                if not self._is_running:
                    self.error.emit("취소됨")
                    return

                self.progress.emit(f"AI 분석 진행 중... ({idx+1} / {total})")
                result = processor.analyze_email(
                    email_body=email_data['body'],
                    user_context=self.user_context,
                    to_recipients=email_data.get('to_recipients', ''),
                    cc_recipients=email_data.get('cc_recipients', ''),
                )
                result['email_subject'] = email_data['subject']
                results.append(result)

                if idx < total - 1:
                    time.sleep(1.5)

            if self._is_running:
                self.finished.emit(results)
        except Exception as e:
            self.error.emit(str(e))

class TodoDetailDialog(QDialog):
    def __init__(self, parent, todo_data, main_window):
        super().__init__(parent)
        self.todo_data = todo_data
        self.main_window = main_window
        self.setWindowTitle("To-Do 상세 정보")
        self.setMinimumWidth(500)
        self.setMinimumHeight(450)

        layout = QVBoxLayout(self)

        layout.addWidget(QLabel("<b>[출처 메일 제목]</b>"))
        subj_label = QLabel(todo_data['email_subject'])
        subj_label.setWordWrap(True)
        subj_label.setStyleSheet("color: #1976D2; font-size: 10pt;")
        layout.addWidget(subj_label)

        layout.addWidget(QLabel("<b>[지시사항 / 행동할 일]</b>"))
        todo_text_edit = QTextEdit()
        todo_text_edit.setPlainText(todo_data['text'])
        todo_text_edit.setReadOnly(True)
        todo_text_edit.setMaximumHeight(80)
        layout.addWidget(todo_text_edit)

        layout.addWidget(QLabel("<b>[원본 메일 요약]</b>"))
        summary_text = QTextEdit()
        summary_text.setPlainText(todo_data['summary'])
        summary_text.setReadOnly(True)
        layout.addWidget(summary_text)

        btn_layout = QHBoxLayout()

        status = self.todo_data['status']
        if status != 'deleted':
            self.complete_btn = QPushButton("완료 해제" if status == 'completed' else "완료 처리")
            self.complete_btn.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; padding: 10px;")
            self.complete_btn.clicked.connect(self.on_complete_clicked)

            self.delete_btn = QPushButton("삭제하기 (휴지통으로)")
            self.delete_btn.setStyleSheet("background-color: #f44336; color: white; font-weight: bold; padding: 10px;")
            self.delete_btn.clicked.connect(self.on_delete_clicked)

            btn_layout.addWidget(self.complete_btn)
            btn_layout.addWidget(self.delete_btn)
        else:
            self.restore_btn = QPushButton("휴지통에서 복구")
            self.restore_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 10px;")
            self.restore_btn.clicked.connect(self.on_restore_clicked)

            self.perm_delete_btn = QPushButton("영구 삭제")
            self.perm_delete_btn.setStyleSheet("background-color: #f44336; color: white; font-weight: bold; padding: 10px;")
            self.perm_delete_btn.clicked.connect(self.on_perm_delete_clicked)

            btn_layout.addWidget(self.restore_btn)
            btn_layout.addWidget(self.perm_delete_btn)

        layout.addLayout(btn_layout)

    def on_complete_clicked(self):
        new_status = 'completed' if self.todo_data['status'] == 'active' else 'active'
        self.main_window.change_todo_status(self.todo_data['id'], new_status)
        self.accept()

    def on_delete_clicked(self):
        self.main_window.change_todo_status(self.todo_data['id'], 'deleted')
        self.accept()

    def on_restore_clicked(self):
        self.main_window.change_todo_status(self.todo_data['id'], 'active')
        self.accept()

    def on_perm_delete_clicked(self):
        self.main_window.permanently_delete_todo(self.todo_data['id'])
        self.accept()


class TodoEditDialog(QDialog):
    """우선순위와 마감일을 편집하는 다이얼로그."""

    def __init__(self, parent, todo_data):
        super().__init__(parent)
        self.setWindowTitle("날짜 / 우선순위 편집")
        self.setMinimumWidth(420)

        layout = QVBoxLayout(self)

        # 할 일 텍스트 표시
        layout.addWidget(QLabel("<b>[할 일]</b>"))
        text_label = QLabel(todo_data.get('text', ''))
        text_label.setWordWrap(True)
        text_label.setStyleSheet("color: #333333; font-size: 10pt; padding: 6pt; background: #f5f5f5;")
        layout.addWidget(text_label)

        layout.addSpacing(10)

        # 우선순위
        priority_layout = QHBoxLayout()
        priority_layout.addWidget(QLabel("우선순위:"))
        self.priority_combo = QComboBox()
        self.priority_combo.addItems(["높음", "보통", "낮음"])
        self.priority_combo.setCurrentText(todo_data.get('priority', '보통'))
        self.priority_combo.setStyleSheet("font-size: 10pt;")
        self.priority_combo.setMinimumHeight(35)
        priority_layout.addWidget(self.priority_combo)
        priority_layout.addStretch()
        layout.addLayout(priority_layout)

        layout.addSpacing(8)

        # 마감일
        date_layout = QHBoxLayout()
        self.use_due_date_cb = QCheckBox("마감일 지정:")
        due = todo_data.get('due_date', '')
        self.use_due_date_cb.setChecked(bool(due))
        self.use_due_date_cb.stateChanged.connect(self._toggle_due_date)
        date_layout.addWidget(self.use_due_date_cb)

        self.due_date_edit = QDateEdit()
        self.due_date_edit.setCalendarPopup(True)
        self.due_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.due_date_edit.setStyleSheet("font-size: 10pt;")
        self.due_date_edit.setMinimumHeight(35)
        if due:
            self.due_date_edit.setDate(QDate.fromString(due, "yyyy-MM-dd"))
        else:
            self.due_date_edit.setDate(QDate.currentDate())
            self.due_date_edit.setEnabled(False)
        date_layout.addWidget(self.due_date_edit)
        date_layout.addStretch()
        layout.addLayout(date_layout)

        layout.addSpacing(15)

        # 버튼
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        ok_btn = QPushButton("저장")
        ok_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 8pt 20pt;")
        ok_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("취소")
        cancel_btn.setStyleSheet("padding: 8pt 20pt;")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(ok_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)

    def _toggle_due_date(self, state):
        self.due_date_edit.setEnabled(state == Qt.Checked.value)

    def get_result(self):
        priority = self.priority_combo.currentText()
        if self.use_due_date_cb.isChecked():
            due_date = self.due_date_edit.date().toString("yyyy-MM-dd")
        else:
            due_date = ""
        return {"priority": priority, "due_date": due_date}


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Outlook AI Summarizer & To-Do List")
        self.resize(1200, 750)

        self.emails_data = []
        self.all_todos = []
        self.current_account_email = ""
        self.user_profile = {}
        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)

        self.stacked_widget = QStackedWidget()
        main_layout.addWidget(self.stacked_widget)

        self.page1 = QWidget()
        self.init_page1()
        self.stacked_widget.addWidget(self.page1)

        self.page2 = QWidget()
        self.init_page2()
        self.stacked_widget.addWidget(self.page2)

        self.page3 = QWidget()
        self.init_page3()
        self.stacked_widget.addWidget(self.page3)

        self.stacked_widget.setCurrentIndex(0)

    def init_page1(self):
        layout = QVBoxLayout(self.page1)
        layout.setAlignment(Qt.AlignCenter)

        title = QLabel("환영합니다!\n조회할 이메일 계정을 선택해주세요.")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 18pt; font-weight: bold; margin-bottom: 30px;")
        layout.addWidget(title)

        form_container = QWidget()
        form_layout = QVBoxLayout(form_container)

        account_layout = QHBoxLayout()
        label_acc = QLabel("접속할 이메일:")
        label_acc.setStyleSheet("font-size: 12pt; font-weight: bold;")
        account_layout.addWidget(label_acc)

        self.account_combo = QComboBox()
        self.account_combo.setMinimumHeight(45)
        self.account_combo.setMinimumWidth(300)
        self.account_combo.setStyleSheet("""
            QComboBox { font-size: 12pt; padding: 5px; }
            QComboBox QAbstractItemView { background-color: white; selection-background-color: #dddddd; }
        """)

        try:
            manager = OutlookManager()
            accounts = manager.get_accounts()
            if accounts:
                for acc in accounts:
                    self.account_combo.addItem(acc)
            else:
                self.account_combo.addItem("기본 계정")
        except:
            self.account_combo.addItem("기본 계정")

        account_layout.addWidget(self.account_combo)

        self.next_btn = QPushButton("다음 화면으로 이동")
        self.next_btn.setMinimumHeight(45)
        self.next_btn.setMinimumWidth(180)
        self.next_btn.setStyleSheet("font-size: 12pt; font-weight: bold; background-color: #4CAF50; color: white;")
        self.next_btn.clicked.connect(self.go_to_page2)
        account_layout.addWidget(self.next_btn)

        self.profile_btn = QPushButton("내 정보 입력")
        self.profile_btn.setMinimumHeight(45)
        self.profile_btn.setMinimumWidth(150)
        self.profile_btn.setStyleSheet("font-size: 12pt; font-weight: bold; background-color: #1976D2; color: white;")
        self.profile_btn.clicked.connect(self._open_profile_dialog)
        account_layout.addWidget(self.profile_btn)

        form_layout.addLayout(account_layout)
        form_layout.addSpacing(20)

        api_layout = QHBoxLayout()
        label_api = QLabel("Gemini API 키:")
        label_api.setStyleSheet("font-size: 12pt; font-weight: bold;")
        api_layout.addWidget(label_api)

        self.api_key_combo = QComboBox()
        self.api_key_combo.setEditable(True)
        self.api_key_combo.addItem("AIzaSyCalBYf5evm_82ZbkOM5TMbcDuh3D6swCQ")
        self.api_key_combo.setMinimumHeight(45)
        self.api_key_combo.setMinimumWidth(300)
        self.api_key_combo.setStyleSheet("""
            QComboBox { font-size: 10pt; padding: 5px; }
            QComboBox QAbstractItemView { background-color: white; selection-background-color: #dddddd; }
        """)

        api_layout.addWidget(self.api_key_combo)
        api_layout.addStretch()

        form_layout.addLayout(api_layout)

        wrapper_layout = QHBoxLayout()
        wrapper_layout.addStretch()
        wrapper_layout.addWidget(form_container)
        wrapper_layout.addStretch()

        layout.addLayout(wrapper_layout)

    def init_page2(self):
        layout = QVBoxLayout(self.page2)

        top_layout = QHBoxLayout()
        self.back_btn = QPushButton("← 계정 다시 선택하기")
        self.back_btn.setMinimumHeight(40)
        self.back_btn.clicked.connect(self.go_to_page1)

        self.current_account_label = QLabel("접속 정보: ")
        self.current_account_label.setStyleSheet("font-size: 10pt; font-weight: bold; color: #2196F3;")

        self.status_label = QLabel("준비됨")
        self.status_label.setStyleSheet("color: #757575; font-weight: bold;")

        label_date = QLabel("조회 날짜:")
        label_date.setStyleSheet("font-size: 10pt; font-weight: bold;")
        self.date_picker = QDateEdit()
        self.date_picker.setCalendarPopup(True)
        self.date_picker.setDate(QDate.currentDate())
        self.date_picker.setMinimumHeight(40)
        self.date_picker.setStyleSheet("font-size: 10pt;")
        self.date_picker.dateChanged.connect(self.fetch_emails)

        self.refresh_btn = QPushButton("해당 날짜 조회")
        self.refresh_btn.setMinimumHeight(40)
        self.refresh_btn.setStyleSheet("font-weight: bold; background-color: #2196F3; color: white; padding: 0 15px;")
        self.refresh_btn.clicked.connect(self.fetch_emails)

        self.fetch_new_btn = QPushButton("새로 고침 (새 메일)")
        self.fetch_new_btn.setMinimumHeight(40)
        self.fetch_new_btn.setStyleSheet("font-weight: bold; background-color: #FF5722; color: white; padding: 0 15px;")
        self.fetch_new_btn.clicked.connect(self.fetch_new_emails)
        self.fetch_new_btn.setEnabled(False)

        self.goto_page3_btn = QPushButton("전체 TODO 관리 →")
        self.goto_page3_btn.setMinimumHeight(40)
        self.goto_page3_btn.setStyleSheet("font-weight: bold; background-color: #673AB7; color: white; padding: 0 15px;")
        self.goto_page3_btn.clicked.connect(self.go_to_page3)

        top_layout.addWidget(self.back_btn)
        top_layout.addSpacing(15)
        top_layout.addWidget(self.current_account_label)
        top_layout.addStretch()
        top_layout.addWidget(self.status_label)
        top_layout.addSpacing(10)
        top_layout.addWidget(label_date)
        top_layout.addWidget(self.date_picker)
        top_layout.addWidget(self.refresh_btn)
        top_layout.addWidget(self.fetch_new_btn)
        top_layout.addWidget(self.goto_page3_btn)

        layout.addLayout(top_layout)
        layout.addSpacing(10)

        splitter = QSplitter(Qt.Horizontal)
        layout.addWidget(splitter)

        # 좌측 영역: 편지함 분할
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)

        left_splitter = QSplitter(Qt.Vertical)

        # 목록부
        list_container = QWidget()
        list_layout = QVBoxLayout(list_container)
        list_layout.setContentsMargins(0, 0, 0, 0)

        list_title_layout = QHBoxLayout()
        title_label = QLabel("메일 목록")
        title_label.setStyleSheet("font-weight: bold; font-size: 10pt;")
        list_title_layout.addWidget(title_label)

        self.email_select_all_cb = QCheckBox("모든 메일 전체 선택")
        self.email_select_all_cb.setStyleSheet("font-weight: bold; font-size: 10pt;")
        self.email_select_all_cb.stateChanged.connect(self.on_email_select_all)
        list_title_layout.addWidget(self.email_select_all_cb)

        list_layout.addLayout(list_title_layout)

        self.email_list_widget = QListWidget()
        self.email_list_widget.setStyleSheet("font-size: 10pt;")
        self.email_list_widget.itemClicked.connect(self.on_email_selected)
        self.email_list_widget.itemSelectionChanged.connect(self.clear_mail_highlights)
        list_layout.addWidget(self.email_list_widget)
        left_splitter.addWidget(list_container)

        # 내용부 & 버튼
        viewer_container = QWidget()
        viewer_layout = QVBoxLayout(viewer_container)
        viewer_layout.setContentsMargins(0, 0, 0, 0)
        viewer_label = QLabel("원본 메일 본문 (목록 클릭 시 확인 가능)")
        viewer_label.setStyleSheet("font-weight: bold; font-size: 10pt; margin-top: 10px;")
        viewer_layout.addWidget(viewer_label)

        self.email_viewer = QTextEdit()
        self.email_viewer.setReadOnly(True)
        self.email_viewer.setStyleSheet("font-size: 10pt; background-color: #f9f9f9;")
        viewer_layout.addWidget(self.email_viewer)

        self.analyze_btn = QPushButton("메일 AI 일괄 요약하기")
        self.analyze_btn.setMinimumHeight(45)
        self.analyze_btn.setStyleSheet("font-size: 10pt; font-weight: bold; background-color: #FF9800; color: white;")
        self.analyze_btn.clicked.connect(self.analyze_checked_emails)
        self.analyze_btn.setEnabled(False)
        viewer_layout.addWidget(self.analyze_btn)

        left_splitter.addWidget(viewer_container)
        left_splitter.setSizes([350, 250])

        left_layout.addWidget(left_splitter)
        splitter.addWidget(left_widget)

        # 우측 영역: 탭 위젯 (Active vs Trash)
        self.right_tabs = QTabWidget()
        self.right_tabs.setStyleSheet("font-size: 10pt;")

        # 1. 활성/완료 스택 탭
        self.active_tab = QWidget()
        active_layout = QVBoxLayout(self.active_tab)

        active_header_layout = QHBoxLayout()
        self.active_select_all_cb = QCheckBox("모든 항목 전체 선택")
        self.active_select_all_cb.setStyleSheet("font-weight: bold;")
        self.active_select_all_cb.stateChanged.connect(self.on_active_select_all)
        active_header_layout.addWidget(self.active_select_all_cb)
        active_header_layout.addStretch()

        self.complete_selected_btn = QPushButton("체크 항목 완료 / 해제")
        self.complete_selected_btn.setStyleSheet("background-color: #2196F3; color: white;")
        self.complete_selected_btn.clicked.connect(self.complete_selected_active)
        active_header_layout.addWidget(self.complete_selected_btn)

        self.delete_selected_btn = QPushButton("체크 삭제 (휴지통)")
        self.delete_selected_btn.setStyleSheet("background-color: #f44336; color: white;")
        self.delete_selected_btn.clicked.connect(self.delete_selected_active)
        active_header_layout.addWidget(self.delete_selected_btn)

        self.forward_to_p3_btn = QPushButton("전체 TODO로 보내기")
        self.forward_to_p3_btn.setStyleSheet("background-color: #673AB7; color: white; font-weight: bold;")
        self.forward_to_p3_btn.clicked.connect(self.forward_todos_to_page3)
        active_header_layout.addWidget(self.forward_to_p3_btn)

        active_layout.addLayout(active_header_layout)

        self.active_todo_list = QListWidget()
        self.active_todo_list.setSelectionMode(QListWidget.ExtendedSelection)
        self.active_todo_list.itemDoubleClicked.connect(self.on_todo_double_clicked)
        self.active_todo_list.itemSelectionChanged.connect(self.on_todo_selection_changed)
        self.active_todo_list.itemChanged.connect(self.on_todo_selection_changed)
        active_layout.addWidget(self.active_todo_list)

        self.right_tabs.addTab(self.active_tab, "To-Do 리스트")

        # 2. 휴지통 탭
        self.trash_tab = QWidget()
        trash_layout = QVBoxLayout(self.trash_tab)

        trash_header_layout = QHBoxLayout()
        self.trash_select_all_cb = QCheckBox("모든 항목 전체 선택")
        self.trash_select_all_cb.setStyleSheet("font-weight: bold;")
        self.trash_select_all_cb.stateChanged.connect(self.on_trash_select_all)
        trash_header_layout.addWidget(self.trash_select_all_cb)
        trash_header_layout.addStretch()

        self.restore_selected_btn = QPushButton("체크 복구")
        self.restore_selected_btn.setStyleSheet("background-color: #4CAF50; color: white;")
        self.restore_selected_btn.clicked.connect(self.restore_selected_trash)
        trash_header_layout.addWidget(self.restore_selected_btn)

        self.permanent_delete_btn = QPushButton("체크 영구 삭제")
        self.permanent_delete_btn.setStyleSheet("background-color: #9E9E9E; color: white;")
        self.permanent_delete_btn.clicked.connect(self.permanent_delete_selected_trash)
        trash_header_layout.addWidget(self.permanent_delete_btn)
        trash_layout.addLayout(trash_header_layout)

        self.trash_todo_list = QListWidget()
        self.trash_todo_list.setSelectionMode(QListWidget.ExtendedSelection)
        self.trash_todo_list.itemDoubleClicked.connect(self.on_todo_double_clicked)
        self.trash_todo_list.itemSelectionChanged.connect(self.on_todo_selection_changed)
        self.trash_todo_list.itemChanged.connect(self.on_todo_selection_changed)
        trash_layout.addWidget(self.trash_todo_list)

        self.right_tabs.addTab(self.trash_tab, "휴지통")

        splitter.addWidget(self.right_tabs)
        splitter.setSizes([500, 600])

        # 로딩 팝업
        self.loading_dialog = QDialog(self)
        self.loading_dialog.setWindowTitle("시스템 안내 - 분석 중...")
        self.loading_dialog.setModal(True)
        self.loading_dialog.setWindowFlag(Qt.WindowCloseButtonHint, False)
        self.loading_dialog.setFixedSize(350, 150)

        loading_layout = QVBoxLayout(self.loading_dialog)
        self.loading_label = QLabel("로딩 중입니다... 잠시만 기다려주세요.")
        self.loading_label.setAlignment(Qt.AlignCenter)
        self.loading_label.setStyleSheet("font-weight: bold; font-size: 10pt;")
        loading_layout.addWidget(self.loading_label)

        self.loading_cancel_btn = QPushButton("진행 작업 중지 (취소)")
        self.loading_cancel_btn.setMinimumHeight(40)
        self.loading_cancel_btn.setStyleSheet("background-color: #f44336; color: white; font-weight: bold; font-size: 10pt; margin-top: 15px;")
        self.loading_cancel_btn.clicked.connect(self.cancel_analysis)
        loading_layout.addWidget(self.loading_cancel_btn)

    def init_page3(self):
        layout = QVBoxLayout(self.page3)

        # 상단 바
        top_layout = QHBoxLayout()

        self.p3_back_btn = QPushButton("← 메일 조회로 돌아가기")
        self.p3_back_btn.setMinimumHeight(40)
        self.p3_back_btn.clicked.connect(self.go_to_page2_from_page3)

        p3_title = QLabel("전체 TODO 관리")
        p3_title.setStyleSheet("font-size: 14pt; font-weight: bold; color: #333333;")

        filter_label = QLabel("우선순위 필터:")
        filter_label.setStyleSheet("font-size: 10pt;")
        self.p3_priority_filter = QComboBox()
        self.p3_priority_filter.addItems(["전체 우선순위", "높음", "보통", "낮음"])
        self.p3_priority_filter.setStyleSheet("font-size: 10pt;")
        self.p3_priority_filter.setMinimumHeight(35)
        self.p3_priority_filter.currentIndexChanged.connect(self.on_p3_filter_changed)

        sort_label = QLabel("정렬:")
        sort_label.setStyleSheet("font-size: 10pt;")
        self.p3_sort_combo = QComboBox()
        self.p3_sort_combo.addItems(["등록 순", "마감일 오름차순", "마감일 내림차순", "우선순위 높음 순", "우선순위 낮음 순"])
        self.p3_sort_combo.setStyleSheet("font-size: 10pt;")
        self.p3_sort_combo.setMinimumHeight(35)
        self.p3_sort_combo.currentIndexChanged.connect(self.on_p3_filter_changed)

        top_layout.addWidget(self.p3_back_btn)
        top_layout.addSpacing(15)
        top_layout.addWidget(p3_title)
        top_layout.addStretch()
        top_layout.addWidget(filter_label)
        top_layout.addWidget(self.p3_priority_filter)
        top_layout.addSpacing(10)
        top_layout.addWidget(sort_label)
        top_layout.addWidget(self.p3_sort_combo)
        layout.addLayout(top_layout)
        layout.addSpacing(8)

        # 탭 위젯
        self.p3_tabs = QTabWidget()
        self.p3_tabs.setStyleSheet("font-size: 10pt;")

        # ─── 탭 1: 할 일 ───
        p3_active_widget = QWidget()
        p3_active_layout = QVBoxLayout(p3_active_widget)

        p3_active_ctrl = QHBoxLayout()
        self.p3_active_select_all_cb = QCheckBox("전체 선택")
        self.p3_active_select_all_cb.setStyleSheet("font-weight: bold;")
        self.p3_active_select_all_cb.stateChanged.connect(self.on_p3_active_select_all)
        p3_active_ctrl.addWidget(self.p3_active_select_all_cb)
        p3_active_ctrl.addStretch()

        self.p3_complete_btn = QPushButton("체크 항목 완료 처리")
        self.p3_complete_btn.setStyleSheet("background-color: #2196F3; color: white; padding: 5pt 10pt;")
        self.p3_complete_btn.clicked.connect(self.p3_complete_selected)
        p3_active_ctrl.addWidget(self.p3_complete_btn)

        self.p3_delete_active_btn = QPushButton("체크 항목 삭제 (휴지통)")
        self.p3_delete_active_btn.setStyleSheet("background-color: #f44336; color: white; padding: 5pt 10pt;")
        self.p3_delete_active_btn.clicked.connect(self.p3_delete_selected_active)
        p3_active_ctrl.addWidget(self.p3_delete_active_btn)

        self.p3_edit_btn = QPushButton("날짜 / 우선순위 편집")
        self.p3_edit_btn.setStyleSheet("background-color: #FF9800; color: white; padding: 5pt 10pt;")
        self.p3_edit_btn.clicked.connect(self.p3_edit_checked)
        p3_active_ctrl.addWidget(self.p3_edit_btn)

        p3_active_layout.addLayout(p3_active_ctrl)

        self.p3_active_list = QListWidget()
        self.p3_active_list.setSelectionMode(QListWidget.ExtendedSelection)
        self.p3_active_list.setStyleSheet("font-size: 10pt;")
        self.p3_active_list.itemDoubleClicked.connect(self.p3_on_item_double_clicked)
        self.p3_active_list.itemChanged.connect(lambda: None)
        p3_active_layout.addWidget(self.p3_active_list)

        self.p3_tabs.addTab(p3_active_widget, "할 일")

        # ─── 탭 2: 완료됨 ───
        p3_completed_widget = QWidget()
        p3_completed_layout = QVBoxLayout(p3_completed_widget)

        p3_completed_ctrl = QHBoxLayout()
        self.p3_completed_select_all_cb = QCheckBox("전체 선택")
        self.p3_completed_select_all_cb.setStyleSheet("font-weight: bold;")
        self.p3_completed_select_all_cb.stateChanged.connect(self.on_p3_completed_select_all)
        p3_completed_ctrl.addWidget(self.p3_completed_select_all_cb)
        p3_completed_ctrl.addStretch()

        self.p3_uncomplete_btn = QPushButton("체크 항목 미완료로 되돌리기")
        self.p3_uncomplete_btn.setStyleSheet("background-color: #4CAF50; color: white; padding: 5pt 10pt;")
        self.p3_uncomplete_btn.clicked.connect(self.p3_uncomplete_selected)
        p3_completed_ctrl.addWidget(self.p3_uncomplete_btn)

        self.p3_delete_completed_btn = QPushButton("체크 항목 삭제 (휴지통)")
        self.p3_delete_completed_btn.setStyleSheet("background-color: #f44336; color: white; padding: 5pt 10pt;")
        self.p3_delete_completed_btn.clicked.connect(self.p3_delete_selected_completed)
        p3_completed_ctrl.addWidget(self.p3_delete_completed_btn)

        p3_completed_layout.addLayout(p3_completed_ctrl)

        self.p3_completed_list = QListWidget()
        self.p3_completed_list.setSelectionMode(QListWidget.ExtendedSelection)
        self.p3_completed_list.setStyleSheet("font-size: 10pt;")
        self.p3_completed_list.itemDoubleClicked.connect(self.p3_on_item_double_clicked)
        p3_completed_layout.addWidget(self.p3_completed_list)

        self.p3_tabs.addTab(p3_completed_widget, "완료됨")

        # ─── 탭 3: 휴지통 ───
        p3_trash_widget = QWidget()
        p3_trash_layout = QVBoxLayout(p3_trash_widget)

        p3_trash_ctrl = QHBoxLayout()
        self.p3_trash_select_all_cb = QCheckBox("전체 선택")
        self.p3_trash_select_all_cb.setStyleSheet("font-weight: bold;")
        self.p3_trash_select_all_cb.stateChanged.connect(self.on_p3_trash_select_all)
        p3_trash_ctrl.addWidget(self.p3_trash_select_all_cb)
        p3_trash_ctrl.addStretch()

        self.p3_restore_btn = QPushButton("체크 복구")
        self.p3_restore_btn.setStyleSheet("background-color: #4CAF50; color: white; padding: 5pt 10pt;")
        self.p3_restore_btn.clicked.connect(self.p3_restore_selected)
        p3_trash_ctrl.addWidget(self.p3_restore_btn)

        self.p3_perm_delete_btn = QPushButton("체크 영구 삭제")
        self.p3_perm_delete_btn.setStyleSheet("background-color: #9E9E9E; color: white; padding: 5pt 10pt;")
        self.p3_perm_delete_btn.clicked.connect(self.p3_perm_delete_selected)
        p3_trash_ctrl.addWidget(self.p3_perm_delete_btn)

        p3_trash_layout.addLayout(p3_trash_ctrl)

        self.p3_trash_list = QListWidget()
        self.p3_trash_list.setSelectionMode(QListWidget.ExtendedSelection)
        self.p3_trash_list.setStyleSheet("font-size: 10pt;")
        p3_trash_layout.addWidget(self.p3_trash_list)

        self.p3_tabs.addTab(p3_trash_widget, "휴지통")

        layout.addWidget(self.p3_tabs)

    # ──────────────────────────────────────────
    # 사용자 프로필 관련 메서드
    # ──────────────────────────────────────────

    def _default_profile(self, account_email=""):
        return {"name": "", "email": account_email, "role": "", "projects": [],
                "superiors": [], "peers": [], "subordinates": [], "clients": []}

    def _load_user_profile(self, account_email=""):
        import json, os
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "user_profile.json")
        if not os.path.exists(path):
            return self._default_profile(account_email)
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            # 구버전 단일 프로필 형식(name 키가 최상위) → 현재 계정으로 마이그레이션
            if "name" in data:
                migrated = {"name": data.get("name", ""), "email": account_email,
                            "role": data.get("role", ""), "projects": data.get("projects", []),
                            "superiors": data.get("superiors", []), "peers": data.get("peers", []),
                            "subordinates": data.get("subordinates", []), "clients": data.get("clients", [])}
                return migrated
            return data.get(account_email, self._default_profile(account_email))
        except Exception:
            return self._default_profile(account_email)

    def _save_user_profile(self):
        import json, os
        if not self.current_account_email:
            return
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "user_profile.json")
        try:
            all_profiles = {}
            if os.path.exists(path):
                try:
                    with open(path, "r", encoding="utf-8") as f:
                        data = json.load(f)
                    # 구버전 형식이면 버리고 새로 시작
                    if "name" not in data:
                        all_profiles = data
                except Exception:
                    pass
            all_profiles[self.current_account_email] = self.user_profile
            with open(path, "w", encoding="utf-8") as f:
                json.dump(all_profiles, f, ensure_ascii=False, indent=2)
        except Exception as e:
            QMessageBox.critical(self, "저장 오류", f"내 정보 저장 중 오류가 발생했습니다.\n{e}")

    def _open_profile_dialog(self):
        if not self.current_account_email:
            QMessageBox.warning(self, "계정 미선택", "먼저 이메일 계정을 선택한 후 내 정보를 입력해 주세요.")
            return
        dlg = UserProfileDialog(self, self.user_profile, self.current_account_email)
        if dlg.exec() == QDialog.Accepted:
            self.user_profile = dlg.get_profile()
            self._save_user_profile()
            QMessageBox.information(self, "저장 완료",
                f"[{self.current_account_email}] 계정의 내 정보가 저장되었습니다.\nAI 분석 시 자동으로 적용됩니다.")

    def _build_user_context(self):
        """저장된 프로필로 AI 프롬프트 컨텍스트 문자열을 생성한다."""
        p = self.user_profile
        # 아무 정보도 없으면 빈 문자열 반환 → 기존 프롬프트 그대로 사용
        if not any([p.get("name"), p.get("projects"), p.get("superiors"), p.get("clients")]):
            return ""

        def join(lst): return ", ".join(lst) if lst else "없음"

        client_strs = []
        for c in p.get("clients", []):
            cprojs = c.get("projects", [])
            client_strs.append(f"{c['name']} ({', '.join(cprojs)})" if cprojs else c["name"])

        lines = [
            "[내 정보]",
            f"- 이름: {p.get('name', '')}",
            f"- 이메일: {p.get('email', '')}",
            f"- 역할: {p.get('role', '')}",
            f"- 현재 참여 프로젝트: {join(p.get('projects', []))}",
            "",
            "[조직 구조]",
            f"- 상사: {join(p.get('superiors', []))}",
            f"- 동료: {join(p.get('peers', []))}",
            f"- 부하직원: {join(p.get('subordinates', []))}",
            f"- Client: {join(client_strs)}",
            "",
            "[우선순위 기준]",
            "1. 상사 요청 → 최우선",
            "2. 고객/외부 요청 → 높음",
            "3. 동료 요청 → 중간",
            "4. 부하직원 요청 → 내가 결정/피드백 해줘야 함",
            "",
            "[메일 분류 기준]",
            "1. 내가 수신자 or 직접 지칭 → ACTION TODO 생성",
            "2. CC + 내 참여 프로젝트 관련 → AWARENESS TODO 생성",
            "3. CC + 무관 → 무시 (todos 빈 배열 반환)",
        ]
        return "\n".join(lines)

    # ──────────────────────────────────────────
    # Page 2 기존 메서드들
    # ──────────────────────────────────────────

    def on_email_select_all(self, state):
        for i in range(self.email_list_widget.count()):
            item = self.email_list_widget.item(i)
            item.setCheckState(Qt.Checked if state == Qt.Checked.value else Qt.Unchecked)

    def on_active_select_all(self, state):
        for i in range(self.active_todo_list.count()):
            item = self.active_todo_list.item(i)
            item.setCheckState(Qt.Checked if state == Qt.Checked.value else Qt.Unchecked)

    def on_trash_select_all(self, state):
        for i in range(self.trash_todo_list.count()):
            item = self.trash_todo_list.item(i)
            item.setCheckState(Qt.Checked if state == Qt.Checked.value else Qt.Unchecked)

    def change_todo_status(self, todo_id, new_status):
        for t in self.all_todos:
            if t['id'] == todo_id:
                t['status'] = new_status
        self.refresh_todo_lists()

    def permanently_delete_todo(self, todo_id):
        self.all_todos = [t for t in self.all_todos if t['id'] != todo_id]
        self.refresh_todo_lists()

    def refresh_todo_lists(self):
        self.active_todo_list.clear()
        self.trash_todo_list.clear()

        self.active_select_all_cb.setCheckState(Qt.Unchecked)
        self.trash_select_all_cb.setCheckState(Qt.Unchecked)

        for t in self.all_todos:
            display_text = f"• {t['text']}"
            # 전달된 항목 표시
            if t.get('forwarded'):
                display_text = f"[전달됨] {t['text']}"

            item = QListWidgetItem(display_text)
            item.setData(Qt.UserRole, t)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Unchecked)

            if t['status'] == 'deleted':
                item.setForeground(QColor("gray"))
                self.trash_todo_list.addItem(item)
            else:
                if t['status'] == 'completed':
                    item.setForeground(QColor("#2196F3"))
                    font = item.font()
                    if font.pointSize() <= 0: font.setPointSize(10)
                    font.setStrikeOut(True)
                    item.setFont(font)
                elif t.get('forwarded'):
                    item.setForeground(QColor("#7B1FA2"))
                else:
                    item.setForeground(QColor("black"))
                self.active_todo_list.addItem(item)

    def complete_selected_active(self):
        items_processed = 0
        for i in range(self.active_todo_list.count()):
            item = self.active_todo_list.item(i)
            if item.checkState() == Qt.Checked:
                data = item.data(Qt.UserRole)
                new_status = 'completed' if data['status'] == 'active' else 'active'
                for t in self.all_todos:
                    if t['id'] == data['id']:
                        t['status'] = new_status
                items_processed += 1

        if items_processed > 0:
            self.refresh_todo_lists()
        else:
            QMessageBox.information(self, "선택 안내", "완료 처리할 항목의 체크박스를 선택해 주세요.")

    def delete_selected_active(self):
        items_processed = 0
        for i in range(self.active_todo_list.count()):
            item = self.active_todo_list.item(i)
            if item.checkState() == Qt.Checked:
                data = item.data(Qt.UserRole)
                for t in self.all_todos:
                    if t['id'] == data['id']:
                        t['status'] = 'deleted'
                items_processed += 1

        if items_processed > 0:
            self.refresh_todo_lists()
        else:
            QMessageBox.information(self, "선택 안내", "삭제할 항목의 체크박스를 선택해 주세요.")

    def restore_selected_trash(self):
        items_processed = 0
        for i in range(self.trash_todo_list.count()):
            item = self.trash_todo_list.item(i)
            if item.checkState() == Qt.Checked:
                data = item.data(Qt.UserRole)
                for t in self.all_todos:
                    if t['id'] == data['id']:
                        t['status'] = 'active'
                items_processed += 1

        if items_processed > 0:
            self.refresh_todo_lists()
        else:
            QMessageBox.information(self, "선택 안내", "복구할 항목의 체크박스를 선택해 주세요.")

    def permanent_delete_selected_trash(self):
        ids_to_remove = set()
        for i in range(self.trash_todo_list.count()):
            item = self.trash_todo_list.item(i)
            if item.checkState() == Qt.Checked:
                data = item.data(Qt.UserRole)
                ids_to_remove.add(data['id'])

        if ids_to_remove:
            self.all_todos = [t for t in self.all_todos if t['id'] not in ids_to_remove]
            self.refresh_todo_lists()
        else:
            QMessageBox.information(self, "선택 안내", "완전히 영구 삭제할 항목의 체크박스를 선택해 주세요.")

    def on_todo_double_clicked(self, item):
        data = item.data(Qt.UserRole)
        dialog = TodoDetailDialog(self, data, self)
        dialog.exec()

    def on_todo_selection_changed(self, item_arg=None):
        for i in range(self.email_list_widget.count()):
            item = self.email_list_widget.item(i)
            font = item.font()
            if font.pointSize() <= 0: font.setPointSize(10)
            font.setBold(False)
            item.setFont(font)
            if item.data(Qt.UserRole + 1):
                item.setForeground(QColor("red"))
            else:
                item.setForeground(QColor("black"))

        target_subjects = set()

        for t_item in self.active_todo_list.selectedItems():
            target_subjects.add(t_item.data(Qt.UserRole).get('email_subject', ''))
        for i in range(self.active_todo_list.count()):
            t_item = self.active_todo_list.item(i)
            if t_item.checkState() == Qt.Checked:
                target_subjects.add(t_item.data(Qt.UserRole).get('email_subject', ''))

        for t_item in self.trash_todo_list.selectedItems():
            target_subjects.add(t_item.data(Qt.UserRole).get('email_subject', ''))
        for i in range(self.trash_todo_list.count()):
            t_item = self.trash_todo_list.item(i)
            if t_item.checkState() == Qt.Checked:
                target_subjects.add(t_item.data(Qt.UserRole).get('email_subject', ''))

        if not target_subjects:
            return

        for i in range(self.email_list_widget.count()):
            item = self.email_list_widget.item(i)
            idx = item.data(Qt.UserRole)
            if idx < len(self.emails_data):
                email_data = self.emails_data[idx]
                if email_data['subject'] in target_subjects:
                    item.setForeground(QColor("#2196F3"))
                    font = item.font()
                    if font.pointSize() <= 0: font.setPointSize(10)
                    font.setBold(True)
                    item.setFont(font)

    def clear_mail_highlights(self):
        self.active_todo_list.blockSignals(True)
        self.trash_todo_list.blockSignals(True)
        self.active_todo_list.clearSelection()
        self.trash_todo_list.clearSelection()
        self.active_todo_list.blockSignals(False)
        self.trash_todo_list.blockSignals(False)
        for i in range(self.email_list_widget.count()):
            item = self.email_list_widget.item(i)
            font = item.font()
            if font.pointSize() <= 0: font.setPointSize(10)
            font.setBold(False)
            item.setFont(font)
            if item.data(Qt.UserRole + 1):
                item.setForeground(QColor("red"))
            else:
                item.setForeground(QColor("black"))

    def go_to_page1(self):
        self.stacked_widget.setCurrentIndex(0)

    def go_to_page2(self):
        selected_account = self.account_combo.currentText()
        self.current_account_email = selected_account
        self.user_profile = self._load_user_profile(selected_account)
        self.current_account_label.setText(f"접속 정보: {selected_account}")
        self.stacked_widget.setCurrentIndex(1)
        self.fetch_emails()

    def fetch_emails(self):
        self.is_refreshing = False
        self.refresh_btn.setEnabled(False)
        self.fetch_new_btn.setEnabled(False)
        self.back_btn.setEnabled(False)
        self.date_picker.setEnabled(False)
        self.analyze_btn.setEnabled(False)
        self.status_label.setText("Outlook에서 메일을 가져오는 중...")
        self.email_list_widget.clear()
        self.email_viewer.clear()
        self.email_select_all_cb.setCheckState(Qt.Unchecked)

        selected_account = self.account_combo.currentText()
        q_date = self.date_picker.date()
        target_dt = datetime.date(q_date.year(), q_date.month(), q_date.day())

        self.fetch_thread = EmailFetchThread(account_email=selected_account, target_date=target_dt)
        self.fetch_thread.finished.connect(self.on_fetch_success)
        self.fetch_thread.error.connect(self.on_fetch_error)
        self.fetch_thread.start()

    def fetch_new_emails(self):
        self.is_refreshing = True
        self.refresh_btn.setEnabled(False)
        self.fetch_new_btn.setEnabled(False)
        self.back_btn.setEnabled(False)
        self.date_picker.setEnabled(False)
        self.analyze_btn.setEnabled(False)
        self.status_label.setText("Outlook에서 추가된 새 메일을 확인하는 중...")

        selected_account = self.account_combo.currentText()
        q_date = self.date_picker.date()
        target_dt = datetime.date(q_date.year(), q_date.month(), q_date.day())

        self.fetch_thread = EmailFetchThread(account_email=selected_account, target_date=target_dt)
        self.fetch_thread.finished.connect(self.on_fetch_success)
        self.fetch_thread.error.connect(self.on_fetch_error)
        self.fetch_thread.start()

    def on_fetch_success(self, emails):
        if getattr(self, 'is_refreshing', False):
            existing_signatures = {(e.get('sender'), e.get('subject')) for e in self.emails_data}
            new_count = 0
            for email in emails:
                sig = (email.get('sender'), email.get('subject'))
                if sig not in existing_signatures:
                    idx = len(self.emails_data)
                    self.emails_data.append(email)

                    display_text = f"[새 메일] [{email['sender']}] {email['subject']}"
                    item = QListWidgetItem(display_text)
                    item.setData(Qt.UserRole, idx)
                    item.setData(Qt.UserRole + 1, True)
                    item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                    item.setCheckState(Qt.Unchecked)
                    item.setForeground(QColor("red"))
                    self.email_list_widget.addItem(item)
                    new_count += 1
            self.status_label.setText(f"새로고침 완료: {new_count}개의 새 메일이 추가되었습니다.")
        else:
            self.emails_data = emails
            for i, email in enumerate(emails):
                display_text = f"[{email['sender']}] {email['subject']}"
                item = QListWidgetItem(display_text)
                item.setData(Qt.UserRole, i)
                item.setData(Qt.UserRole + 1, False)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Unchecked)
                self.email_list_widget.addItem(item)
            self.status_label.setText(f"요청하신 날짜의 총 {len(emails)}개의 메일을 가져왔습니다.")

        self.refresh_btn.setEnabled(True)
        self.fetch_new_btn.setEnabled(True)
        self.back_btn.setEnabled(True)
        self.date_picker.setEnabled(True)
        self.analyze_btn.setEnabled(self.email_list_widget.count() > 0)

    def on_fetch_error(self, error_msg):
        self.status_label.setText("이메일 가져오기 실패")
        self.refresh_btn.setEnabled(True)
        self.fetch_new_btn.setEnabled(len(self.emails_data) > 0)
        self.back_btn.setEnabled(True)
        self.date_picker.setEnabled(True)
        QMessageBox.critical(self, "Outlook 연동 오류", f"이메일 접근 중 문제가 발생했습니다.\n\n상세: {error_msg}")

    def on_email_selected(self, item):
        idx = item.data(Qt.UserRole)
        is_new = item.data(Qt.UserRole + 1)

        if is_new:
            item.setData(Qt.UserRole + 1, False)
            item.setForeground(QColor("black"))
            text = item.text()
            if text.startswith("[새 메일] "):
                item.setText(text.replace("[새 메일] ", "", 1))

        email_data = self.emails_data[idx]
        body_text = email_data.get('body', "")
        if not body_text.strip():
            body_text = "본문 내용이 없습니다."
        self.email_viewer.setPlainText(body_text)

    def analyze_checked_emails(self):
        emails_to_analyze = []
        for i in range(self.email_list_widget.count()):
            item = self.email_list_widget.item(i)
            if item.checkState() == Qt.Checked:
                idx = item.data(Qt.UserRole)
                emails_to_analyze.append(self.emails_data[idx])

        if not emails_to_analyze:
            QMessageBox.warning(self, "메일 선택 필요", "일괄 분석을 진행할 항목 왼쪽의 박스를 체크(☑)해 주세요.")
            return

        current_api_key = self.api_key_combo.currentText().strip()

        winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)

        self.analyze_thread = EmailAnalyzeThread(emails_to_analyze, current_api_key, self._build_user_context())
        self.analyze_thread.progress.connect(self.update_loading_text)
        self.analyze_thread.finished.connect(self.on_analyze_success)
        self.analyze_thread.error.connect(self.on_analyze_error)
        self.analyze_thread.start()

        self.loading_label.setText(f"총 {len(emails_to_analyze)}건의 선택된 메일을 순차적으로 AI 분석중입니다...\n끝날 때까지 화면을 조작할 수 없습니다.")
        self.loading_cancel_btn.setEnabled(True)
        self.loading_cancel_btn.setText("진행 작업 중지 (취소)")

        self.loading_dialog.exec()

    def update_loading_text(self, text):
        self.loading_label.setText(f"{text}\n중간에 작업 중지를 원하시면 아래 취소 버튼을 누르세요.")

    def cancel_analysis(self):
        if hasattr(self, 'analyze_thread') and self.analyze_thread.isRunning():
            self.analyze_thread.stop()
            self.loading_label.setText("취소 명령 접수됨... 현재 진행중인 데이터 통신을 끊고 종료합니다.\n(안전하게 초기 상태로 돌아옵니다.)")
            self.loading_cancel_btn.setText("종료하는 중...")
            self.loading_cancel_btn.setEnabled(False)

    def on_analyze_success(self, results):
        self.loading_dialog.accept()
        self.status_label.setText(f"총 {len(results)}개의 선택하신 메일 분석이 일괄 완료되었습니다.")

        for res in results:
            summary    = res.get("summary", "요약 정보가 없습니다.")
            raw_todos  = res.get("todos", [])
            ai_priority = res.get("priority", "보통")  # AI가 판단한 우선순위
            email_title = res.get("email_subject", "알 수 없는 메일")

            # todos가 dict 리스트(새 포맷)이면 (text, type) 추출, 문자열이면 호환 처리
            normalized = []
            for t in raw_todos:
                if isinstance(t, dict):
                    normalized.append((t.get("text", ""), t.get("type", "ACTION")))
                else:
                    normalized.append((str(t), "ACTION"))

            if not normalized:
                self.all_todos.append({
                    "id": str(uuid.uuid4()),
                    "text": "할 일 지시사항 없음 (팝업에서 요약을 확인하세요)",
                    "summary": summary,
                    "email_subject": email_title,
                    "status": "active",
                    "forwarded": False,
                    "due_date": "",
                    "priority": ai_priority,
                })
            else:
                for text, todo_type in normalized:
                    prefix = f"[{todo_type}] " if todo_type in ("ACTION", "AWARENESS") else ""
                    self.all_todos.append({
                        "id": str(uuid.uuid4()),
                        "text": prefix + text,
                        "summary": summary,
                        "email_subject": email_title,
                        "status": "active",
                        "forwarded": False,
                        "due_date": "",
                        "priority": ai_priority,
                    })

        self.refresh_todo_lists()
        self.email_select_all_cb.setCheckState(Qt.Unchecked)

    def on_analyze_error(self, error_msg):
        self.loading_dialog.accept()

        if error_msg == "취소됨":
            self.status_label.setText("사용자에 의해 AI 분석이 강제 중단되었습니다 (이전 상태로 복구됨)")
            return

        self.status_label.setText("AI 분석 실패")
        QMessageBox.critical(self, "AI 연동 오류", f"API 호출 중 에러가 발생하여 처리가 중단되었습니다.\n\n상세: {error_msg}")

    # ──────────────────────────────────────────
    # Page 3 네비게이션
    # ──────────────────────────────────────────

    def go_to_page3(self):
        self.stacked_widget.setCurrentIndex(2)
        self.refresh_page3()

    def go_to_page2_from_page3(self):
        self.stacked_widget.setCurrentIndex(1)
        self.refresh_todo_lists()

    def forward_todos_to_page3(self):
        """Page2 active_todo_list에서 체크된 항목을 forwarded=True로 표시하고 Page3로 이동."""
        forwarded_count = 0
        for i in range(self.active_todo_list.count()):
            item = self.active_todo_list.item(i)
            if item.checkState() == Qt.Checked:
                data = item.data(Qt.UserRole)
                for t in self.all_todos:
                    if t['id'] == data['id'] and not t.get('forwarded'):
                        t['forwarded'] = True
                        forwarded_count += 1

        if forwarded_count == 0:
            QMessageBox.information(self, "선택 안내", "전달할 항목의 체크박스를 선택해 주세요.\n이미 전달된 항목은 제외됩니다.")
            return

        self.go_to_page3()

    # ──────────────────────────────────────────
    # Page 3 데이터 / 렌더링
    # ──────────────────────────────────────────

    PRIORITY_ORDER = {"높음": 0, "보통": 1, "낮음": 2}
    PRIORITY_COLOR = {
        "높음": QColor("#D32F2F"),
        "보통": QColor("#E65100"),
        "낮음": QColor("#1565C0"),
    }

    def _get_priority_color(self, priority):
        return self.PRIORITY_COLOR.get(priority, QColor("#333333"))

    def _build_item_text(self, t):
        priority = t.get('priority', '보통')
        due = t.get('due_date', '')
        due_str = due if due else "없음"
        subject = t.get('email_subject', '')
        return f"[{priority}]  {t['text']}  |  출처: {subject}  |  마감: {due_str}"

    def _get_filtered_sorted(self, status):
        result = [t for t in self.all_todos if t['status'] == status and t.get('forwarded', False)]

        priority_filter = self.p3_priority_filter.currentText()
        if priority_filter != "전체 우선순위":
            result = [t for t in result if t.get('priority', '보통') == priority_filter]

        sort_key = self.p3_sort_combo.currentText()
        if sort_key == "마감일 오름차순":
            result.sort(key=lambda t: t.get('due_date', '') or '9999-99-99')
        elif sort_key == "마감일 내림차순":
            result.sort(key=lambda t: t.get('due_date', '') or '', reverse=True)
        elif sort_key == "우선순위 높음 순":
            result.sort(key=lambda t: self.PRIORITY_ORDER.get(t.get('priority', '보통'), 1))
        elif sort_key == "우선순위 낮음 순":
            result.sort(key=lambda t: self.PRIORITY_ORDER.get(t.get('priority', '보통'), 1), reverse=True)

        return result

    def refresh_page3(self):
        # 전체선택 체크박스 초기화
        self.p3_active_select_all_cb.blockSignals(True)
        self.p3_completed_select_all_cb.blockSignals(True)
        self.p3_trash_select_all_cb.blockSignals(True)
        self.p3_active_select_all_cb.setCheckState(Qt.Unchecked)
        self.p3_completed_select_all_cb.setCheckState(Qt.Unchecked)
        self.p3_trash_select_all_cb.setCheckState(Qt.Unchecked)
        self.p3_active_select_all_cb.blockSignals(False)
        self.p3_completed_select_all_cb.blockSignals(False)
        self.p3_trash_select_all_cb.blockSignals(False)

        # 할 일 탭
        self.p3_active_list.blockSignals(True)
        self.p3_active_list.clear()
        for t in self._get_filtered_sorted('active'):
            item = QListWidgetItem(self._build_item_text(t))
            item.setData(Qt.UserRole, t)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Unchecked)
            item.setForeground(self._get_priority_color(t.get('priority', '보통')))
            self.p3_active_list.addItem(item)
        self.p3_active_list.blockSignals(False)

        # 완료됨 탭
        self.p3_completed_list.blockSignals(True)
        self.p3_completed_list.clear()
        for t in self._get_filtered_sorted('completed'):
            item = QListWidgetItem(self._build_item_text(t))
            item.setData(Qt.UserRole, t)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Unchecked)
            font = item.font()
            if font.pointSize() <= 0: font.setPointSize(10)
            font.setStrikeOut(True)
            item.setFont(font)
            item.setForeground(QColor("#757575"))
            self.p3_completed_list.addItem(item)
        self.p3_completed_list.blockSignals(False)

        # 휴지통 탭
        self.p3_trash_list.blockSignals(True)
        self.p3_trash_list.clear()
        for t in self._get_filtered_sorted('deleted'):
            item = QListWidgetItem(self._build_item_text(t))
            item.setData(Qt.UserRole, t)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Unchecked)
            item.setForeground(QColor("gray"))
            self.p3_trash_list.addItem(item)
        self.p3_trash_list.blockSignals(False)

    def on_p3_filter_changed(self):
        self.refresh_page3()

    # ──────────────────────────────────────────
    # Page 3 전체선택
    # ──────────────────────────────────────────

    def on_p3_active_select_all(self, state):
        for i in range(self.p3_active_list.count()):
            self.p3_active_list.item(i).setCheckState(
                Qt.Checked if state == Qt.Checked.value else Qt.Unchecked
            )

    def on_p3_completed_select_all(self, state):
        for i in range(self.p3_completed_list.count()):
            self.p3_completed_list.item(i).setCheckState(
                Qt.Checked if state == Qt.Checked.value else Qt.Unchecked
            )

    def on_p3_trash_select_all(self, state):
        for i in range(self.p3_trash_list.count()):
            self.p3_trash_list.item(i).setCheckState(
                Qt.Checked if state == Qt.Checked.value else Qt.Unchecked
            )

    # ──────────────────────────────────────────
    # Page 3 액션
    # ──────────────────────────────────────────

    def _p3_get_checked_ids(self, list_widget):
        ids = []
        for i in range(list_widget.count()):
            item = list_widget.item(i)
            if item.checkState() == Qt.Checked:
                ids.append(item.data(Qt.UserRole)['id'])
        return ids

    def _p3_change_status(self, ids, new_status):
        for t in self.all_todos:
            if t['id'] in ids:
                t['status'] = new_status
        self.refresh_page3()

    def p3_complete_selected(self):
        ids = self._p3_get_checked_ids(self.p3_active_list)
        if not ids:
            QMessageBox.information(self, "선택 안내", "완료 처리할 항목을 체크해 주세요.")
            return
        self._p3_change_status(set(ids), 'completed')

    def p3_uncomplete_selected(self):
        ids = self._p3_get_checked_ids(self.p3_completed_list)
        if not ids:
            QMessageBox.information(self, "선택 안내", "미완료로 되돌릴 항목을 체크해 주세요.")
            return
        self._p3_change_status(set(ids), 'active')

    def p3_delete_selected_active(self):
        ids = self._p3_get_checked_ids(self.p3_active_list)
        if not ids:
            QMessageBox.information(self, "선택 안내", "삭제할 항목을 체크해 주세요.")
            return
        self._p3_change_status(set(ids), 'deleted')

    def p3_delete_selected_completed(self):
        ids = self._p3_get_checked_ids(self.p3_completed_list)
        if not ids:
            QMessageBox.information(self, "선택 안내", "삭제할 항목을 체크해 주세요.")
            return
        self._p3_change_status(set(ids), 'deleted')

    def p3_restore_selected(self):
        ids = self._p3_get_checked_ids(self.p3_trash_list)
        if not ids:
            QMessageBox.information(self, "선택 안내", "복구할 항목을 체크해 주세요.")
            return
        self._p3_change_status(set(ids), 'active')

    def p3_perm_delete_selected(self):
        ids = set(self._p3_get_checked_ids(self.p3_trash_list))
        if not ids:
            QMessageBox.information(self, "선택 안내", "영구 삭제할 항목을 체크해 주세요.")
            return
        self.all_todos = [t for t in self.all_todos if t['id'] not in ids]
        self.refresh_page3()

    def p3_edit_checked(self):
        """할 일 탭에서 체크된 첫 번째 항목을 편집."""
        for i in range(self.p3_active_list.count()):
            item = self.p3_active_list.item(i)
            if item.checkState() == Qt.Checked:
                self.p3_on_item_double_clicked(item)
                return
        QMessageBox.information(self, "선택 안내", "편집할 항목을 체크해 주세요.")

    def p3_on_item_double_clicked(self, item):
        t = item.data(Qt.UserRole)
        dialog = TodoEditDialog(self, t)
        if dialog.exec() == QDialog.Accepted:
            result = dialog.get_result()
            for todo in self.all_todos:
                if todo['id'] == t['id']:
                    todo['priority'] = result['priority']
                    todo['due_date'] = result['due_date']
                    break
            self.refresh_page3()
