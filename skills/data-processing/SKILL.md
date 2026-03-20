---
name: ai-mindmap-generation
description: Streamlit + 멀티 AI(Gemini/GPT/Claude) + Mermaid.js 마인드맵 생성 스킬
---

## AI 마인드맵 생성 스킬

이 스킬은 문서(PDF/TXT/이미지)나 텍스트를 입력받아 AI로 요약하고 Mermaid 마인드맵을 자동 생성하는 패턴을 제공합니다.

### 사용 시점
- PDF, TXT, 이미지 파일에서 텍스트를 추출할 때
- AI API(Gemini, GPT, Claude)를 호출하여 요약/마인드맵을 생성할 때
- Mermaid.js로 다이어그램을 렌더링할 때
- 다국어(한/영) UI를 구현할 때

### 아키텍처 패턴

```
입력 (PDF/TXT/이미지/텍스트)
    ↓
텍스트 추출 (PyPDF2 / pytesseract / 직접 입력)
    ↓
AI 디스패처 (provider에 따라 분기)
    ├─ Gemini: google-genai → client.models.generate_content()
    ├─ GPT:    openai → client.chat.completions.create()
    └─ Claude: anthropic → client.messages.create()
    ↓
응답 파싱 (정규식으로 요약/Mermaid 코드 분리)
    ↓
Mermaid.js CDN 렌더링 (HTML iframe → st.components.html)
```

### 핵심 코드 패턴

#### 1. AI 디스패처 패턴
```python
def call_ai_text(provider: str, api_key: str, content: str, lang: str) -> str:
    if provider == "gemini":
        return call_gemini_text(api_key, content, lang)
    if provider == "gpt":
        return call_gpt_text(api_key, content, lang)
    if provider == "claude":
        return call_claude_text(api_key, content, lang)
```

#### 2. 다국어 번역 테이블 패턴
```python
TRANSLATIONS: dict[str, dict] = {
    "한국어": {"main_title": "🧠 AI 마인드맵 생성기", ...},
    "English": {"main_title": "🧠 AI Mindmap Generator", ...},
}
T = TRANSLATIONS[selected_lang]
st.markdown(T["main_title"])
```

#### 3. Mermaid 렌더링 패턴
```python
def build_mermaid_html(mermaid_code: str, height: int = 560) -> str:
    """mermaid.js CDN을 사용해 HTML 생성 → st.components.v1.html()로 삽입"""
    safe = mermaid_code.replace("\\", "\\\\").replace("`", "\\`").replace("$", "\\$")
    return f"""<script type="module">
      import mermaid from 'https://cdn.jsdelivr.net/npm/mermaid@11/dist/mermaid.esm.min.mjs';
      ...
    </script>"""
```

#### 4. AI 응답 파싱 패턴
```python
def parse_ai_response(text: str) -> tuple[str, str]:
    # 요약: "## 핵심 요약" 또는 "## Key Summary" 이후 ~ "## Mermaid" 이전
    summary = re.search(r"## (?:핵심 요약|Key Summary)\s*(.*?)(?=## )", text, re.DOTALL)
    # Mermaid: ```mindmap ... ``` 블록 추출
    mermaid = re.search(r"```mindmap\s*(.*?)```", text, re.DOTALL)
    return summary, mermaid
```

#### 5. 이미지 Vision 폴백 패턴
```python
# OCR 시도 → 실패 시 AI Vision API로 자동 전환
ocr_text = extract_text_via_ocr(file_bytes)
if ocr_text and len(ocr_text) > 30:
    raw_text = ocr_text      # OCR 성공 → 텍스트 API 사용
else:
    vision_bytes = file_bytes  # OCR 실패 → Vision API로 폴백
    use_vision = True
```

### 참고
- `scripts/` 폴더에 재사용 가능한 헬퍼 스크립트가 있습니다.
- `examples/` 폴더에 사용 예제가 있습니다.
- 현재 프로젝트의 전체 구현은 `app.py`에 통합되어 있습니다.
