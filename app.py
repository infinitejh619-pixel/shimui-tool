import streamlit as st
from groq import Groq
import pdfplumber
from pptx import Presentation
import io

st.set_page_config(page_title="심의의견 분석 도구", page_icon="📋", layout="wide")
st.title("📋 보험 광고 심의의견 분석 도구")

def get_api_key():
    try:
        return st.secrets["GROQ_API_KEY"]
    except Exception:
        return None

api_key = get_api_key()

if not api_key:
    with st.sidebar:
        st.header("⚙️ 설정")
        api_key = st.text_input("Groq API Key", type="password")
        if api_key:
            st.success("API 키 입력 완료!")

if not api_key:
    st.info("👈 왼쪽 사이드바에 Groq API 키를 입력하면 시작할 수 있어요.")
    st.stop()

client = Groq(api_key=api_key)

def ask_groq(prompt):
    response = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role": "user", "content": prompt}]
    )
    return response.choices[0].message.content

def extract_file_text(uploaded_file):
    file_text = ""
    if uploaded_file.name.lower().endswith(".pdf"):
        with pdfplumber.open(io.BytesIO(uploaded_file.read())) as pdf:
            for i, page in enumerate(pdf.pages, 1):
                text = page.extract_text()
                if text:
                    file_text += f"[{i}페이지]\n{text}\n\n"
    elif uploaded_file.name.lower().endswith(".pptx"):
        prs = Presentation(io.BytesIO(uploaded_file.read()))
        for i, slide in enumerate(prs.slides, 1):
            slide_text = ""
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text.strip():
                    slide_text += shape.text.strip() + "\n"
            if slide_text:
                file_text += f"[{i}페이지]\n{slide_text}\n"
    return file_text

tab1, tab2 = st.tabs(["📧 심의의견 추출", "🔍 심의의견 ↔ 소재 1:1 매칭"])

# ── 탭 1: 메일에서 심의의견 추출 ──────────────────────────────────────────────
with tab1:
    st.subheader("메일에서 심의의견 추출")
    st.caption("심의 결과 메일을 붙여넣으면, 심의의견만 원문 그대로 정리해드려요.")

    email_text = st.text_area("메일 전문 붙여넣기", height=350,
        placeholder="광고주로부터 받은 메일 내용을 전체 복사해서 여기에 붙여넣으세요...")

    if st.button("심의의견 추출하기", type="primary", key="btn_extract"):
        if not email_text.strip():
            st.error("메일 내용을 입력해주세요.")
        else:
            with st.spinner("추출 중..."):
                prompt = f"""보험 광고 심의 메일에서 심의의견만 추출해주세요.

규칙:
- 수정 요청, 지적 사항, 보완 요청만 추출합니다.
- 인사말, 안내 문구 등은 제외합니다.
- 반드시 원문 표현을 그대로 유지합니다. 요약하거나 바꾸지 마세요.
- 몇 페이지, 어떤 문구인지 언급된 내용을 절대 생략하지 마세요.
- 번호 목록으로 정리합니다.

메일 내용:
{email_text}"""
                try:
                    result = ask_groq(prompt)
                    st.success("추출 완료!")
                    st.markdown("---")
                    st.markdown("### 심의의견 목록")
                    st.markdown(result)
                    st.markdown("---")
                    st.download_button("📥 저장", data=result, file_name="심의의견.txt", mime="text/plain")
                except Exception as e:
                    st.error(f"오류: {str(e)}")

# ── 탭 2: 심의의견 ↔ 소재 1:1 매칭 ────────────────────────────────────────────
with tab2:
    st.subheader("심의의견 ↔ 소재 1:1 매칭 수정 제안")
    st.caption("메일과 소재 파일을 함께 넣으면, 각 심의의견이 소재의 어느 부분인지 찾아서 정확한 수정 방향을 알려드려요.")

    col1, col2 = st.columns(2)
    with col1:
        email_for_match = st.text_area("심의 메일 붙여넣기", height=300,
            placeholder="광고주로부터 받은 심의 메일을 붙여넣으세요...")
    with col2:
        uploaded_file = st.file_uploader("소재 파일 업로드 (PDF 또는 PPT)", type=["pdf", "pptx"])
        if uploaded_file:
            st.info(f"업로드: {uploaded_file.name}")

    if st.button("1:1 매칭 분석하기", type="primary", key="btn_match"):
        if not email_for_match.strip():
            st.error("심의 메일을 입력해주세요.")
        elif not uploaded_file:
            st.error("소재 파일을 업로드해주세요.")
        else:
            with st.spinner("심의의견과 소재를 매칭하는 중..."):
                try:
                    file_text = extract_file_text(uploaded_file)

                    if not file_text.strip():
                        st.warning("⚠️ 파일에서 텍스트를 추출하지 못했습니다. 이미지로만 구성된 파일일 수 있어요.")
                    else:
                        prompt = f"""당신은 보험 광고 심의 전문가입니다.
아래 두 가지 정보를 바탕으로, 각 심의의견이 소재의 어느 부분에 해당하는지 1:1로 매칭해주세요.

작업 방법:
1. 심의 메일에서 각 심의의견을 찾습니다 (원문 그대로 인용).
2. 해당 의견이 가리키는 페이지를 소재에서 찾습니다.
3. 그 페이지에서 문제가 되는 실제 문구를 찾습니다.
4. 심의의견에 따른 정확한 수정 방향을 제시합니다.

출력 형식 (심의의견마다 반복):

---
**[심의의견 N]**
- 📋 의견 원문: (심의 메일 원문 그대로 인용)
- 📄 해당 위치: N페이지
- ❌ 현재 문구: (소재에서 찾은 실제 문구)
- ✅ 수정 문구: (바꿔야 할 문구)
- 💡 수정 이유: (한 줄 요약)
---

주의: 의견 원문을 절대 요약하거나 바꾸지 마세요. 소재에서 찾지 못한 경우 "소재에서 해당 문구를 찾을 수 없음"이라고 명시하세요.

===== 심의 메일 =====
{email_for_match}

===== 소재 내용 (페이지별) =====
{file_text}"""

                        result = ask_groq(prompt)
                        st.success("매칭 완료!")
                        st.markdown("---")
                        st.markdown(result)
                        st.markdown("---")
                        st.download_button("📥 결과 저장", data=result, file_name="수정제안.txt", mime="text/plain")

                except Exception as e:
                    st.error(f"오류: {str(e)}")
