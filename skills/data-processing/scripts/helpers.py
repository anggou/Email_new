"""
AI 마인드맵 생성 헬퍼 유틸리티
- 문서 텍스트 추출 (PDF, TXT, 이미지 OCR)
- AI API 호출 (Gemini, GPT, Claude)
- Mermaid 코드 파싱
"""
import io
import re
from typing import Any


def extract_text_from_pdf(file_bytes: bytes) -> str:
    """PyPDF2로 PDF 전 페이지에서 텍스트를 추출한다."""
    import PyPDF2
    reader = PyPDF2.PdfReader(io.BytesIO(file_bytes))
    pages = []
    for i, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""
        if text.strip():
            pages.append(f"[Page {i}]\n{text.strip()}")
    return "\n\n".join(pages)


def extract_text_from_txt(file_bytes: bytes) -> str:
    """여러 인코딩을 시도하여 TXT 파일을 읽는다."""
    for enc in ("utf-8", "euc-kr", "cp949", "latin-1"):
        try:
            return file_bytes.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return file_bytes.decode("utf-8", errors="replace")


def extract_text_via_ocr(file_bytes: bytes) -> str:
    """pytesseract + Pillow로 이미지 OCR. 실패 시 빈 문자열 반환."""
    try:
        import pytesseract
        from PIL import Image
        img = Image.open(io.BytesIO(file_bytes))
        try:
            return pytesseract.image_to_string(img, lang="kor+eng").strip()
        except pytesseract.TesseractError:
            return pytesseract.image_to_string(img, lang="eng").strip()
    except Exception:
        return ""


def parse_ai_response(text: str) -> tuple[str, str]:
    """AI 응답에서 (요약, Mermaid 코드)를 분리한다."""
    summary = ""
    m = re.search(
        r"## (?:핵심 요약|Key Summary)\s*(.*?)(?=## (?:Mermaid|마인드맵)|\Z)",
        text, re.DOTALL,
    )
    if m:
        summary = m.group(1).strip()

    mermaid_code = ""
    m = re.search(r"```mindmap\s*(.*?)```", text, re.DOTALL)
    if m:
        mermaid_code = m.group(1).strip()
        if not mermaid_code.startswith("mindmap"):
            mermaid_code = "mindmap\n" + mermaid_code

    if not mermaid_code:
        m = re.search(r"(mindmap\s*\n(?:[ \t]+.+\n?)+)", text)
        if m:
            mermaid_code = m.group(1).strip()

    return summary, mermaid_code
