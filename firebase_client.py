import requests
import json
import os

_AUTH_ERRORS = {
    "INVALID_EMAIL":               "이메일 형식이 올바르지 않습니다.",
    "EMAIL_NOT_FOUND":             "등록되지 않은 이메일입니다.",
    "INVALID_PASSWORD":            "비밀번호가 올바르지 않습니다.",
    "WRONG_PASSWORD":              "비밀번호가 올바르지 않습니다.",
    "INVALID_LOGIN_CREDENTIALS":   "이메일 또는 비밀번호가 올바르지 않습니다.",
    "USER_DISABLED":               "비활성화된 계정입니다. 관리자에게 문의하세요.",
    "TOO_MANY_ATTEMPTS_TRY_LATER": "로그인 시도가 너무 많습니다. 잠시 후 다시 시도하세요.",
    "EMAIL_EXISTS":                "이미 사용 중인 이메일입니다.",
    "WEAK_PASSWORD":               "비밀번호가 너무 약합니다. (6자 이상 입력하세요)",
    "OPERATION_NOT_ALLOWED":       "이 로그인 방식은 현재 비활성화되어 있습니다.",
}

def _translate_auth_error(raw_msg: str) -> str:
    for code, korean in _AUTH_ERRORS.items():
        if code in raw_msg:
            return korean
    return "인증 오류가 발생했습니다."

class FirebaseClient:
    """
    Firebase Client SDK를 파이썬에서 흉내 내기 위한 REST API 클라이언트입니다.
    배포용 프로그램 보안을 위해 구글 서비스 계정 키(.json) 없이 '웹 API 키'와 유저의 '로그인 토큰'만으로 통신합니다.
    """
    def __init__(self, api_key: str, project_id: str):
        self.api_key = api_key
        self.project_id = project_id
        
        # Firebase Auth Endpoints
        self.auth_login_url = f"https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key={self.api_key}"
        self.auth_signup_url = f"https://identitytoolkit.googleapis.com/v1/accounts:signUp?key={self.api_key}"
        
        # Firestore Base Endpoint
        self.firestore_base_url = f"https://firestore.googleapis.com/v1/projects/{self.project_id}/databases/(default)/documents"

    def sign_up(self, email, password):
        """이메일과 비밀번호로 회원가입합니다."""
        payload = {"email": email, "password": password, "returnSecureToken": True}
        response = requests.post(self.auth_signup_url, json=payload)
        
        if response.status_code != 200:
            raw = response.json().get('error', {}).get('message', '')
            raise Exception(_translate_auth_error(raw))
            
        return response.json() # idToken, email, refreshToken, expiresIn, localId(UID) 반환

    def sign_in(self, email, password):
        """이메일과 비밀번호로 로그인합니다."""
        payload = {"email": email, "password": password, "returnSecureToken": True}
        response = requests.post(self.auth_login_url, json=payload)
        
        if response.status_code != 200:
            raw = response.json().get('error', {}).get('message', '')
            raise Exception(_translate_auth_error(raw))

        return response.json()

    def _get_headers(self, id_token):
        return {"Authorization": f"Bearer {id_token}"}

    def save_data(self, uid: str, id_token: str, collection_name: str, document_id: str, data_dict: dict):
        """
        딕셔너리 데이터를 JSON 문자열로 변환하여 Firestore에 저장합니다.
        (Firestore의 복잡한 REST API 타입 시스템을 피하기 위해 통째로 직렬화하여 'data_json' 필드에 저장)
        """
        url = f"{self.firestore_base_url}/users/{uid}/{collection_name}/{document_id}"
        
        # JSON 직렬화
        json_str = json.dumps(data_dict, ensure_ascii=False)
        
        firestore_payload = {
            "fields": {
                "data_json": {"stringValue": json_str}
            }
        }
        
        response = requests.patch(url, headers=self._get_headers(id_token), json=firestore_payload)
        if response.status_code != 200:
            raise Exception(f"데이터 저장 실패: {response.text}")
        return True

    def get_data(self, uid: str, id_token: str, collection_name: str, document_id: str):
        """Firestore에서 'data_json' 필드를 불러와 파이썬 딕셔너리로 반환합니다."""
        url = f"{self.firestore_base_url}/users/{uid}/{collection_name}/{document_id}"
        
        response = requests.get(url, headers=self._get_headers(id_token))
        if response.status_code == 404:
            return None # 데이터가 아직 없음
        elif response.status_code != 200:
            raise Exception(f"데이터 불러오기 실패: {response.text}")
            
        doc_data = response.json()
        try:
            json_str = doc_data['fields']['data_json']['stringValue']
            return json.loads(json_str)
        except (KeyError, json.JSONDecodeError):
            return None
