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
                        try:
                            inbox = acc.DeliveryStore.GetDefaultFolder(6)
                        except Exception as inner_e:
                            self.logger.warning(f"DeliveryStore 접근 실패: {inner_e}")
                        break
            
            if not inbox:
                inbox = namespace.GetDefaultFolder(6)
                
            messages = inbox.Items
            messages.Sort("[ReceivedTime]", True)

            fetched_emails = []
            
            for message in messages:
                try:
                    msg_date = getattr(message, "ReceivedTime", None)
                    if not msg_date:
                        continue
                        
                    # Outlook ReceivedTime은 pywintypes.datetime이므로 파이썬 date객체와 비교
                    msg_date_only = msg_date.date()
                    
                    if target_date and msg_date_only > target_date:
                        # 타겟 날짜보다 최신 메일은 건너뜀 (내림차순 정렬이므로)
                        continue
                    elif target_date and msg_date_only < target_date:
                        # 그룹 지어있는 날짜 이하로 떨어지면 탐색 종료
                        break
                        
                    # 타겟 날짜와 정확히 일치하는 경우(또는 target_date가 없는경우 전체)
                    if not target_date or msg_date_only == target_date:
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
                            
                except Exception as ex:
                    pass
                        
            return fetched_emails
            
        except Exception as e:
            self.logger.error(f"Outlook 연동 중 오류 발생: {e}")
            raise Exception(f"Outlook 메일박스 연동 중 오류 발생: {e}")
        finally:
            pythoncom.CoUninitialize()
