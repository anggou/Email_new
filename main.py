import sys
import os
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QIcon
from main_window import MainWindow

def main():
    """
    애플리케이션 진입점 (Entry Point)
    """
    # PySide6 애플리케이션 객체 생성
    app = QApplication(sys.argv)
    
    # 아이콘 설정
    icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.png")
    app_icon = QIcon(icon_path)
    app.setWindowIcon(app_icon)
    
    # Windows 작업표시줄 아이콘 정상 표시를 위한 설정
    if os.name == 'nt':
        import ctypes
        myappid = 'com.emailsummarizer.app.1.0'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    
    # 메인 윈도우 인스턴스 생성 및 표시
    window = MainWindow()
    window.setWindowIcon(app_icon)
    window.show()
    
    # 애플리케이션 이벤트 루프 실행 (GUI가 종료될 때까지 대기)
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
