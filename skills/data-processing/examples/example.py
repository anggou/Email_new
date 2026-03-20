"""
AI 마인드맵 생성 스킬 사용 예제
- Gemini API로 텍스트를 요약하고 Mermaid 마인드맵 코드를 생성하는 예제
"""
import os
from dotenv import load_dotenv

load_dotenv()


def main():
    from google import genai

    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        print("❌ GEMINI_API_KEY를 .env 파일에 설정해 주세요.")
        return

    # ── 1. 분석할 텍스트 준비
    sample_text = """
    인공지능(AI)은 기계가 인간의 지능을 모방하도록 하는 기술입니다.
    주요 분야로 머신러닝, 딥러닝, 자연어 처리, 컴퓨터 비전이 있으며,
    현재 의료, 금융, 교육, 자율주행 등 다양한 산업에서 활용되고 있습니다.
    """

    # ── 2. Gemini API 호출
    client = genai.Client(api_key=api_key)
    prompt = f"""다음 텍스트를 분석하여 아래 형식으로 응답하세요.

## 핵심 요약
- 요점 1
- 요점 2

## Mermaid 마인드맵
```mindmap
mindmap
  root((주제))
    대주제1
      소주제1
    대주제2
```

[분석할 내용]
{sample_text}
"""
    response = client.models.generate_content(
        model="gemini-3.1-flash-lite-preview",
        contents=prompt,
    )
    print("=== AI 응답 ===")
    print(response.text)

    # ── 3. 응답 파싱
    from scripts.helpers import parse_ai_response
    summary, mermaid_code = parse_ai_response(response.text)

    print("\n=== 요약 ===")
    print(summary)
    print("\n=== Mermaid 코드 ===")
    print(mermaid_code)


if __name__ == "__main__":
    main()
