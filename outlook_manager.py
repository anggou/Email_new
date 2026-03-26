import win32com.client
import logging
import pythoncom
import datetime

class OutlookManager:
    def __init__(self):
        self.logger = logging.getLogger(__name__)

    def get_accounts(self):
        """Outlook에 등록된 모든 계정의 이메일 주소 목록을 반환합니다."""
        pythoncom.CoInitialize()
        accounts_list = []
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            for i in range(1, namespace.Accounts.Count + 1):
                acc = namespace.Accounts.Item(i)
                email = getattr(acc, "SmtpAddress", None) or getattr(acc, "DisplayName", f"계정 {i}")
                if email not in accounts_list:
                    accounts_list.append(email)
        except Exception as e:
            self.logger.error(f"계정 조회 중 오류: {e}")
        finally:
            pythoncom.CoUninitialize()
        return accounts_list

    def get_emails_by_date(self, account_email=None, target_date=None, limit=50):
        """
        특정 이메일 계정에서 지정된 날짜의 이메일만 가져옵니다(읽음/읽지않음 상관없이).
        """
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            
            inbox = None
            if account_email and account_email != "기본 계정":
                for i in range(1, namespace.Accounts.Count + 1):
                    acc = namespace.Accounts.Item(i)
                    email = getattr(acc, "SmtpAddress", "") or getattr(acc, "DisplayName", "")
                    if email == account_email:
                        # DeliveryStore → 계정별 실제 받은편지함 (Exchange/로컬 모두 지원)
                        try:
                            store = acc.DeliveryStore
                            inbox = store.GetDefaultFolder(6)
                            self.logger.info(f"DeliveryStore inbox 접근 성공: {email}")
                        except Exception as inner_e:
                            self.logger.warning(f"DeliveryStore 접근 실패, 폴더 직접 탐색 시도: {inner_e}")
                            # 폴더 직접 탐색: Stores에서 해당 계정 store 찾기
                            try:
                                for j in range(1, namespace.Stores.Count + 1):
                                    store = namespace.Stores.Item(j)
                                    store_display = getattr(store, "DisplayName", "")
                                    if account_email in store_display or store_display in account_email:
                                        inbox = store.GetDefaultFolder(6)
                                        self.logger.info(f"Stores 탐색으로 inbox 접근 성공: {store_display}")
                                        break
                            except Exception as e2:
                                self.logger.warning(f"Stores 탐색 실패: {e2}")
                        break

            if not inbox:
                self.logger.warning("계정별 inbox 접근 실패, 기본 받은편지함 사용")
                inbox = namespace.GetDefaultFolder(6)
                
            messages = inbox.Items

            # Restrict 필터로 날짜 범위 직접 쿼리 (Exchange Online에서도 안정적)
            if target_date:
                start_dt = datetime.datetime.combine(target_date, datetime.time.min)
                end_dt   = datetime.datetime.combine(
                    target_date + datetime.timedelta(days=1), datetime.time.min
                )
                filter_str = (
                    f"[ReceivedTime] >= '{start_dt.strftime('%m/%d/%Y %I:%M %p')}' AND "
                    f"[ReceivedTime] < '{end_dt.strftime('%m/%d/%Y %I:%M %p')}'"
                )
                messages = messages.Restrict(filter_str)

            fetched_emails = []
            for message in messages:
                try:
                    email_data = {
                        "subject":       getattr(message, "Subject", "제목 없음"),
                        "sender":        getattr(message, "SenderName", "알 수 없는 발신자"),
                        "sender_email":  getattr(message, "SenderEmailAddress", ""),
                        "to_recipients": getattr(message, "To", ""),
                        "cc_recipients": getattr(message, "CC", ""),
                        "body":          getattr(message, "Body", ""),
                        "entry_id":      getattr(message, "EntryID", "")
                    }
                    fetched_emails.append(email_data)
                    if len(fetched_emails) >= limit:
                        break
                except Exception:
                    pass

            return fetched_emails

        except Exception as e:
            self.logger.error(f"Outlook 연동 중 오류 발생: {e}")
            raise Exception(f"Outlook 메일박스 연동 중 오류 발생: {e}")
        finally:
            pythoncom.CoUninitialize()

    def open_email_by_entry_id(self, entry_id: str):
        """EntryID로 Outlook에서 해당 메일을 열어 표시합니다."""
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.Session.GetItemFromID(entry_id)
            mail.Display()
        except Exception as e:
            self.logger.error(f"메일 열기 실패: {e}")
            raise Exception(f"메일을 열 수 없습니다: {e}")
        finally:
            pythoncom.CoUninitialize()
