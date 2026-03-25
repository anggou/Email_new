# 🏗️ 로컬 환경 기반 단일 실행 파일(.exe) 배포 계획서

## 1. 배경 및 목적 (Background & Purpose)
현재 `Email_AI` 앱을 클라우드로 마이그레이션(Microsoft Graph API + Azure OAuth 2.0)하는 과정에서 **기업의 IT 보안 정책(Admin Consent 요구)** 벽에 부딪혔습니다.
이에 따라 거대한 클라우드 인프라를 거치지 않고, 가장 확실하고 폐쇄적인 방식인 **"로컬 데스크톱 Outlook 연동 + 단일 프로그램(.exe) 배포"** 아키텍처로 방향을 완전히 전환(Pivot)합니다.

## 2. 🎯 달성하고자 하는 핵심 목표
1. **Azure 종속성 탈피:** Microsoft 서버가 아닌 로컬 PC에 설치된 'Outlook 프로그램'의 권한을 합법적으로 빌려옵니다. (관리자 승인 무력화)
2. **복잡한 설치 과정 제로:** 동료 직원들은 파이썬이나 기타 복잡한 세팅을 할 필요 없이 USB로 전달받은 `.exe` 파일만 클릭하면 즉시 사용할 수 있어야 합니다.
3. **완전한 로컬 웹 인터페이스:** 사용자 경험(UX)은 예전처럼 세련된 브라우저(로컬웹 `localhost:8050`) 형태로 그대로 유지합니다.

---

## 3. 🛠️ 기술 스택 변경점 (Architecture Changes)

| 구분 | (이전) 클라우드 마이그레이션 방향 | **(현재) 단일 실행 프로그램 방향** |
|:---|:---|:---|
| **이메일 접근 구동부** | Graph API (`requests`) | **`win32com.client` (로컬 Outlook C-API 훅)** |
| **로그인/인증 흐름** | Azure OAuth 2.0 라우트 팝업 | **완전 삭제 (로그인 불필요, 로컬 PC 사용자 자동 인식)** |
| **환경/서버** | Docker 컨테이너, Gunicorn (Linux) | **`PyInstaller` 기반 `.exe` 내장 서버 (Windows)** |
| **브라우저 접근** | `https://도메인.com` 접속 | 더블 클릭 시 **`http://localhost:8050` 자동 오픈** |

---

## 4. 📝 3단계 작업 마일스톤 (Action Items)

### [Phase 1] 롤백 및 핵심 구동부 이식
- [ ] `app.py` 구조 간소화
  - Microsoft OAuth 콜백 라우트(`/auth/login`, `/auth/callback`) 및 세션 로직 제거
  - 불필요해진 `AZURE_CLIENT_ID` 등 환경변수 불러오기 로직 삭제
- [ ] `outlook_manager.py`를 `win32com.client.Dispatch("Outlook.Application")` 기반으로 전면 원복하여 로컬 MAPI와 통신하도록 수정

### [Phase 2] 사용자 경험(UX) 개선 및 브라우저 자동화
- [ ] `app.py` 초기화 시 파이썬 `webbrowser` 모듈을 추가하여, 서버가 구동됨과 동시에 크롬이나 기본 브라우저가 자동으로 `localhost:8050` 탭을 열어주도록 구현

### [Phase 3] 컴파일 및 패키징 
- [ ] PyInstaller 설치 (`pip install pyinstaller`)
- [ ] Dash 프레임워크의 숨겨진 에셋(assets, css)과 기타 의존성 파일들이 `.exe`에 안전하게 압축되도록 빌드 명령어(또는 `.spec` 파일) 구성
- [ ] 빌드 테스트 후 최적화된 최종 배포용 `Email_Summarizer.exe` 추출

---

## 5. ⚠️ 필수 선행/제약 조건 (Prerequisites)
이 배포 방식으로 생성된 파일은 **Windows OS**를 사용하며, 컴퓨터 상주 메모리에 **Microsoft Outlook 데스크톱 앱**이 구동 중인 PC에서만 완벽하게 동작합니다. (Mac OS 환경 및 웹 버전 한정 사용자는 지원하지 않습니다.)
