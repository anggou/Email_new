import os
import json
import logging
import re
from google import genai
from dotenv import load_dotenv

load_dotenv()

class AIProcessor:
    def __init__(self, override_api_key=None):
        self.logger = logging.getLogger(__name__)
        self.api_key = override_api_key or os.getenv("GEMINI_API_KEY")
        
        if not self.api_key or self.api_key == "your_gemini_api_key_here":
            raise ValueError("GEMINI_API_KEY가 등록되어 있지 않습니다. UI에서 API 키를 제대로 입력해주세요.")
            
        # 새로운 최신 SDK 사용
        self.client = genai.Client(api_key=self.api_key)

    def analyze_email(self, email_body, user_context="", to_recipients="", cc_recipients=""):
        if user_context:
            recipient_info = f"""
[이메일 수신 정보]
- 직접 수신자 (TO): {to_recipients if to_recipients else "정보 없음"}
- 참조 수신자 (CC): {cc_recipients if cc_recipients else "정보 없음"}
"""
            prompt = f"""
{user_context}
{recipient_info}
위의 내 정보와 조직 구조, 수신 정보, 우선순위/분류 기준을 바탕으로 아래 이메일을 분석해줘.

[메일 분류 규칙 — 반드시 수신 정보를 기준으로 판단]
- 내 이름 또는 내 이메일이 TO(직접 수신자)에 포함되어 있으면 → "ACTION"
- 내 이름 또는 내 이메일이 CC(참조)에만 있고, 내 참여 프로젝트와 관련된 내용이면 → "AWARENESS"
- 내 이름 또는 내 이메일이 CC(참조)에만 있고, 내 참여 프로젝트와 무관하면 → "IGNORE"
- TO/CC 정보가 불명확한 경우, 메일 본문에서 내가 직접 지시/요청받은 경우 → "ACTION", 단순 공유/참고인 경우 → "AWARENESS"

[우선순위 규칙]
- 발신자가 상사이면 "높음", 고객/외부인이면 "높음", 동료이면 "보통", 부하직원이면 "낮음", 불명확하면 "보통"
- mail_type이 "IGNORE"이면 priority는 "낮음", todos는 반드시 빈 배열

결과는 반드시 아래 JSON 형식으로만 반환해 (설명 없이 JSON만):
{{
  "summary": "세 줄 이내 요약 (내가 CC인 경우 그 사실도 명시)",
  "mail_type": "ACTION",
  "priority": "높음",
  "todos": [
    {{"text": "할 일 텍스트", "type": "ACTION"}}
  ]
}}

이메일 본문:
{email_body}
"""
        else:
            prompt = f"""
이메일 내용을 세 줄 이내로 요약하고, 수신자가 해야 할 명확한 행동(Action Item)을 To-Do 리스트 형태로 추출해 줘.
결과는 반드시 파싱하기 쉬운 JSON 형식(키: 'summary', 'todos')으로 반환해.
todos는 문자열의 리스트 형태로 제공해줘.

이메일 본문:
{email_body}
"""
        response_text = ""
        try:
            try:
                response = self.client.models.generate_content(
                    model='gemini-3.1-flash',
                    contents=prompt
                )
            except Exception as e:
                self.logger.warning(f"3.1-flash 모델 호출 실패, 2.5-flash로 전환합니다. 사유: {e}")
                response = self.client.models.generate_content(
                    model='gemini-2.5-flash',
                    contents=prompt
                )
                
            response_text = response.text
            
            json_match = re.search(r'```(?:json)?\s*(.*?)\s*```', response_text, re.DOTALL)
            if json_match:
                json_str = json_match.group(1)
            else:
                json_str = response_text
                
            result = json.loads(json_str)
            
            if 'summary' not in result:
                result['summary'] = "요약을 생성할 수 없습니다."
            if 'todos' not in result:
                result['todos'] = []
                
            return result
            
        except json.JSONDecodeError as e:
            self.logger.error(f"JSON 파싱 오류: {e}\n원본 응답: {response_text}")
            raise Exception("AI 응답을 JSON 형식으로 파싱할 수 없습니다. 응답 형태가 올바르지 않습니다.")
        except Exception as e:
            # 상세한 에러 메시지를 노출하여 무슨 문제인지 파악하기 쉽게 함
            self.logger.error(f"Gemini API 호출 중 오류 발생: {e}")
            raise Exception(f"{e}")
