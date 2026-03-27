import sys
import os
import json
import requests
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path

VERSION_URL = "https://raw.githubusercontent.com/anggou/Email_new/master/version.json"
APP_EXE_NAME = "EmailSummarizer.exe"


def get_install_dir():
    """런처 EXE 기준 설치 폴더"""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).parent


def get_local_version(install_dir: Path) -> str:
    vf = install_dir / "app_version.json"
    if vf.exists():
        try:
            return json.loads(vf.read_text(encoding='utf-8')).get("version", "0.0.0")
        except Exception:
            pass
    return "0.0.0"


def save_local_version(install_dir: Path, version: str):
    vf = install_dir / "app_version.json"
    vf.write_text(json.dumps({"version": version}, ensure_ascii=False), encoding='utf-8')


class LauncherApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Email AI Summarizer")
        self.root.geometry("420x180")
        self.root.resizable(False, False)
        self.root.configure(bg="#f0f2f5")

        # 화면 중앙 배치
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() - 420) // 2
        y = (self.root.winfo_screenheight() - 180) // 2
        self.root.geometry(f"+{x}+{y}")

        # UI
        tk.Label(self.root, text="Email AI Summarizer",
                 font=("Segoe UI", 14, "bold"),
                 bg="#f0f2f5", fg="#1a1a2e").pack(pady=(28, 4))

        self.status_var = tk.StringVar(value="시작 중...")
        tk.Label(self.root, textvariable=self.status_var,
                 font=("Segoe UI", 9), bg="#f0f2f5", fg="#666").pack(pady=(0, 12))

        self.progress = ttk.Progressbar(self.root, mode='indeterminate', length=340)
        self.progress.pack(pady=4)
        self.progress.start(10)

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        threading.Thread(target=self._run, daemon=True).start()
        self.root.mainloop()

    def _on_close(self):
        self.root.destroy()
        sys.exit(0)

    def _set_status(self, msg: str):
        self.root.after(0, lambda: self.status_var.set(msg))

    def _set_progress(self, value: float):
        def _update():
            self.progress.stop()
            self.progress.configure(mode='determinate')
            self.progress['value'] = value
        self.root.after(0, _update)

    def _show_error(self, msg: str):
        self.root.after(0, lambda: messagebox.showerror("오류", msg, parent=self.root))

    def _run(self):
        install_dir = get_install_dir()
        app_path = install_dir / APP_EXE_NAME

        # 원격 버전 확인
        remote_version = None
        download_url = None
        try:
            self._set_status("업데이트 확인 중...")
            res = requests.get(VERSION_URL, timeout=8)
            data = res.json()
            remote_version = data.get("version")
            download_url = data.get("download_url")
        except Exception:
            pass  # 오프라인 또는 서버 오류 → 기존 앱 실행 시도

        local_version = get_local_version(install_dir)
        need_update = (
            remote_version
            and download_url
            and (local_version != remote_version or not app_path.exists())
        )

        if need_update:
            success = self._download(download_url, app_path, install_dir, remote_version)
            if not success:
                if not app_path.exists():
                    self._set_status("다운로드 실패: 앱 파일이 없습니다.")
                    self._show_error(
                        f"앱 파일을 다운로드하지 못했습니다.\n"
                        f"인터넷 연결을 확인하고 다시 시도하세요."
                    )
                    return
                # 다운로드 실패지만 기존 앱 있으면 그냥 실행
                self._set_status("업데이트 실패 — 기존 버전으로 실행합니다.")

        if not app_path.exists():
            self._set_status("앱 파일을 찾을 수 없습니다.")
            self._show_error(
                f"{APP_EXE_NAME}을 찾을 수 없습니다.\n인터넷 연결을 확인하세요."
            )
            return

        self._set_status("실행 중...")
        subprocess.Popen([str(app_path)])
        self.root.after(600, self.root.destroy)

    def _download(self, url: str, app_path: Path, install_dir: Path, version: str) -> bool:
        tmp_path = install_dir / f"{APP_EXE_NAME}.tmp"
        try:
            self._set_status("업데이트 다운로드 중...")
            res = requests.get(url, stream=True, timeout=180)
            res.raise_for_status()
            total = int(res.headers.get('content-length', 0))
            downloaded = 0

            with open(tmp_path, 'wb') as f:
                for chunk in res.iter_content(chunk_size=65536):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        if total:
                            pct = downloaded / total * 100
                            self._set_progress(pct)
                            self._set_status(f"다운로드 중... {pct:.0f}%  "
                                             f"({downloaded // 1024 // 1024}MB / "
                                             f"{total // 1024 // 1024}MB)")

            # 기존 파일 교체
            if app_path.exists():
                app_path.unlink()
            tmp_path.rename(app_path)
            save_local_version(install_dir, version)
            self._set_status(f"업데이트 완료! (v{version})")
            return True

        except Exception as e:
            if tmp_path.exists():
                try:
                    tmp_path.unlink()
                except Exception:
                    pass
            self._set_status(f"다운로드 실패: {e}")
            return False


if __name__ == "__main__":
    LauncherApp()
