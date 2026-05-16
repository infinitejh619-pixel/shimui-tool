import streamlit as st
from groq import Groq
import pdfplumber
from pptx import Presentation
import io
import json
import os
from datetime import datetime

st.set_page_config(page_title="심의의견 분석 도구", page_icon="📋", layout="wide")

def get_api_key():
    try:
        return st.secrets["GROQ_API_KEY"]
    except Exception:
        return None

api_key = get_api_key()
if not api_key:
    with st.sidebar:
        api_key = st.text_input("Groq API Key", type="password")
if not api_key:
    st.info("👈 왼쪽 사이드바에 Groq API 키를 입력하면 시작할 수 있어요.")
    st.stop()

client = Groq(api_key=api_key)

HISTORY_FILE = "shimui_history.json"

def load_history():
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []

def save_to_history(email_text, result):
    history = load_history()
    item = {
        "id": datetime.now().strftime("%Y%m%d_%H%M%S"),
        "timestamp": datetime.now().strftime("%Y.%m.%d %H:%M"),
        "preview": email_text[:40].replace("\n", " ") + ("..." if len(email_text) > 40 else ""),
        "email_text": email_text,
        "result": result,
    }
    history.insert(0, item)
    history = history[:30]
    try:
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(history, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    return item

def ask_groq(prompt):
    response = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role": "user", "content": prompt}],
    )
    return response.choices[0].message.content

def extract_file_text(uploaded_file):
    file_text = ""
    data = uploaded_file.read()
    if uploaded_file.name.lower().endswith(".pdf"):
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for i, page in enumerate(pdf.pages, 1):
                text = page.extract_text()
                if text:
                    file_text += f"[{i}페이지]\n{text}\n\n"
    elif uploaded_file.name.lower().endswith(".pptx"):
        prs = Presentation(io.BytesIO(data))
        for i, slide in enumerate(prs.slides, 1):
            slide_text = ""
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text.strip():
                    slide_text += shape.text.strip() + "\n"
            if slide_text:
                file_text += f"[{i}페이지]\n{slide_text}\n"
    return file_text

history = load_history()

with st.sidebar:
    st.header("📁 최근 작업 히스토리")
    if not history:
        st.caption("아직 작업 내역이 없어요.\n1:1 매칭 분석을 하면 자동 저장돼요.")
    else:
        for item in history:
            label = f"🗓 {item['timestamp']}\n{item['preview']}"
            if st.button(label, key=f"hist_{item['id']}", use_container_width=True):
                st.session_state["selected_history"] = item
                st.rerun()

st.title("📋 보험 광고 심의의견 분석 도구")

tab1, tab2 = st.tabs(["🔍 심의의견 ↔ 소재 1:1 매칭", "✅ 수정 반영 검수"])

with tab1:
    st.subheader("심의의견 ↔ 소재 1:1 매칭")
    st.caption("심의 메일과 소재 파일을 함께 넣으면, 각 의견이 소재의 어느 부분인지 찾아 정확한 수정 방향을 알려드려요.")

    col1, col2 = st.columns(2)
    with col1:
        email_text = st.text_area("심의 메일 붙여넣기", height=320,
            placeholder="광고주로부터 받은 심의 메일을 붙여넣으세요...")
    with col2:
        uploaded_file = st.file_uploader("소재 파일 업로드 (PDF 또는 PPT)", type=["pdf", "pptx"])
        if uploaded_file:
            st.info(f"업로드: {uploaded_file.name}")

    if st.button("1:1 매칭 분석하기", type="primary", key="btn_match"):
        if not email_text.strip():
            st.error("심의 메일을 입력해주세요.")
        elif not uploaded_file:
            st.error("소재 파일을 업로드해주세요.")
        else:
            with st.spinner("심의의견과 소재를 매칭하는 중..."):
                try:
                    file_text = extract_file_text(uploaded_file)
                    if not file_text.strip():
                        st.warning("⚠️ 파일에서 텍스트를 추출하지 못했습니다. 이미지 기반 파일일 수 있어요.")
                    else:
                        prompt = f"""당신은 보험 광고 심의 전문가입니다.
심의 메일의 각 의견과 소재의 해당 부분을 1:1로 매칭해주세요.

작업 방법:
1. 심의 메일에서 각 심의의견을 찾아 원문 그대로 인용합니다.
2. 아래 위치 표기 파싱 규칙을 따라 해당 페이지와 배너를 정확히 파악합니다.
3. 해당하는 모든 페이지/배너의 문구를 소재에서 찾습니다.
4. 심의의견에 따른 정확한 수정 방향을 제시합니다.

【위치 표기 파싱 규칙 — 반드시 준수】
- "2,7p" → 2페이지와 7페이지 (쉼표 = '그리고', 절대 첫 번째만 추출하지 말 것)
- "2-7p" → 2페이지부터 7페이지까지
- "[2,7p_1,6배너]" → 2페이지와 7페이지의 1번 배너와 6번 배너
- "_1배너", "_3배너" 등 배너 번호는 반드시 함께 기재
- 쉼표로 구분된 페이지가 여러 개면 해당하는 모든 페이지 내용을 각각 찾아서 보여줄 것
- 원문에 표기된 위치 정보를 절대 생략하거나 축약하지 말 것

출력 형식 (심의의견마다 반복):

---
**[심의의견 N]**
- 📋 의견 원문: (심의 메일 원문 그대로 인용)
- 📄 해당 위치: N페이지 / N번 배너 (원문 위치 표기도 괄호 안에 병기, 예: [2,7p_1,6배너])
- ❌ 현재 문구: (소재에서 찾은 실제 문구 — 해당 위치가 여러 곳이면 각각 표시)
- ✅ 수정 문구: (바꿔야 할 문구)
- 💡 수정 이유: (한 줄 요약)
---

주의: 의견 원문을 절대 요약하거나 변형하지 마세요. 소재에서 찾지 못한 경우 "소재에서 해당 문구를 찾을 수 없음"이라고 명시하세요.

===== 심의 메일 =====
{email_text}

===== 소재 내용 (페이지별) =====
{file_text}"""

                        result = ask_groq(prompt)
                        save_to_history(email_text, result)

                        st.success("매칭 완료! 왼쪽 히스토리에 저장되었어요.")
                        st.markdown("---")
                        st.markdown(result)
                        st.markdown("---")
                        st.download_button("📥 결과 저장", data=result,
                            file_name="수정제안.txt", mime="text/plain")

                except Exception as e:
                    st.error(f"오류: {str(e)}")

with tab2:
    st.subheader("수정 반영 검수")
    st.caption("수정된 소재를 업로드하면, 심의의견이 빠짐없이 반영됐는지 확인해드려요.")

    selected = st.session_state.get("selected_history") or (history[0] if history else None)

    if not selected:
        st.info("👈 왼쪽 히스토리에서 작업을 선택하거나, 먼저 1:1 매칭 분석을 진행해주세요.")
    else:
        st.info(f"📋 선택된 작업: **{selected['timestamp']}** — {selected['preview']}")

        with st.expander("원본 심의의견 분석 결과 보기"):
            st.markdown(selected["result"])

        st.markdown("---")
        revised_file = st.file_uploader("수정된 소재 업로드 (PDF 또는 PPT)",
            type=["pdf", "pptx"], key="revised_file")

        if revised_file:
            st.info(f"업로드: {revised_file.name}")

            if st.button("검수 시작하기", type="primary", key="btn_check"):
                with st.spinner("수정 반영 여부를 검수하는 중..."):
                    try:
                        revised_text = extract_file_text(revised_file)
                        if not revised_text.strip():
                            st.warning("⚠️ 파일에서 텍스트를 추출하지 못했습니다.")
                        else:
                            prompt = f"""당신은 보험 광고 심의 검수 전문가입니다.
원본 심의의견 분석 결과와 수정된 소재를 비교하여, 각 심의의견이 반영됐는지 검수해주세요.

검수 기준:
- ✅ 반영: 심의의견대로 수정됨
- ❌ 미반영: 수정되지 않음
- ⚠️ 부분반영: 일부만 수정됨

【위치 표기 파싱 규칙 — 반드시 준수】
- "2,7p" → 2페이지와 7페이지 모두 확인 (첫 번째만 보지 말 것)
- "[2,7p_1,6배너]" → 2페이지와 7페이지의 1번·6번 배너 모두 확인

출력 형식 (각 심의의견마다):

---
**[심의의견 N]**
- 📋 원래 의견: (원문 인용)
- 판정: ✅ 반영 / ❌ 미반영 / ⚠️ 부분반영
- 확인 내용: (수정 소재에서 어떻게 바뀌었는지 또는 바뀌지 않았는지 구체적으로)
---

**[ 최종 검수 요약 ]**
- 총 N개 의견 중 ✅ 반영 N개 / ❌ 미반영 N개 / ⚠️ 부분반영 N개
- 미반영/부분반영 항목은 재수정 필요

===== 원본 심의의견 분석 결과 =====
{selected['result']}

===== 수정된 소재 내용 =====
{revised_text}"""

                            result = ask_groq(prompt)
                            st.success("검수 완료!")
                            st.markdown("---")
                            st.markdown(result)
                            st.markdown("---")
                            st.download_button("📥 검수 결과 저장", data=result,
                                file_name="검수결과.txt", mime="text/plain")

                    except Exception as e:
                        st.error(f"오류: {str(e)}")
