import streamlit as st
from google import genai as google_genai
import pdfplumber
from pptx import Presentation
import io

st.set_page_config(
    page_title="심의의견 분석 도구",
    page_icon="📋",
    layout="wide"
)

st.title("📋 보험 광고 심의의견 분석 도구")
st.caption("메일에서 심의의견을 추출하고, 소재 수정 방향을 제안받아요.")

# API 키 설정 (Streamlit Cloud secrets 우선, 없으면 사이드바 입력)
def get_api_key():
    try:
        return st.secrets["GEMINI_API_KEY"]
    except Exception:
        return None

api_key = get_api_key()

if not api_key:
    with st.sidebar:
        st.header("⚙️ 설정")
        api_key = st.text_input("Gemini API Key", type="password", help="AIza...로 시작하는 키를 입력하세요")
        if api_key:
            st.success("API 키 입력 완료!")

if not api_key:
    st.info("👈 왼쪽 사이드바에 Gemini API 키를 입력하면 시작할 수 있어요.")
    st.stop()

client = google_genai.Client(api_key=api_key)

tab1, tab2 = st.tabs(["📧 심의의견 추출", "📄 소재 수정 제안"])

# ── 탭 1: 메일에서 심의의견 추출 ──────────────────────────────────────────────
with tab1:
    st.subheader("메일에서 심의의견 추출")
    st.caption("광고주에게 받은 심의 결과 메일을 붙여넣으면, 심의의견만 깔끔하게 정리해드려요.")

    email_text = st.text_area(
        "메일 전문 붙여넣기",
        height=350,
        placeholder="광고주로부터 받은 메일 내용을 전체 복사해서 여기에 붙여넣으세요..."
    )

    if st.button("심의의견 추출하기", type="primary", key="btn_extract"):
        if not email_text.strip():
            st.error("메일 내용을 입력해주세요.")
        else:
            with st.spinner("심의의견을 추출하는 중..."):
                prompt = f"""당신은 보험 광고 담당 퍼포먼스 마케터의 어시스턴트입니다.
아래 메일 전문에서 광고 심의 의견만 빠짐없이 추출해주세요.

규칙:
- 수정 요청 사항, 지적 사항, 보완 요청 사항만 추출합니다.
- 인사말, 감사 인사, 일반 안내 문구 등은 제외합니다.
- 번호 목록으로 정리합니다.
- 원문 표현을 최대한 유지합니다.

메일 내용:
{email_text}"""

                try:
                    response = client.models.generate_content(model="gemini-2.0-flash", contents=prompt)
                    st.success("추출 완료!")
                    st.markdown("---")
                    st.markdown("### 심의의견 목록")
                    st.markdown(response.text)
                    st.markdown("---")
                    st.download_button(
                        label="📥 텍스트 파일로 저장",
                        data=response.text,
                        file_name="심의의견.txt",
                        mime="text/plain"
                    )
                except Exception as e:
                    st.error(f"오류가 발생했습니다: {str(e)}")

# ── 탭 2: 소재 수정 제안 ──────────────────────────────────────────────────────
with tab2:
    st.subheader("소재 수정 제안")
    st.caption("심의서 파일을 업로드하면, 어떤 부분을 어떻게 수정해야 하는지 제안해드려요.")

    uploaded_file = st.file_uploader(
        "심의서 파일 업로드 (PDF 또는 PPT)",
        type=["pdf", "pptx"]
    )

    context = st.text_area(
        "광고 소재 추가 정보 (선택사항)",
        height=80,
        placeholder="예: 암보험 인스타그램 배너 소재입니다. 주요 메시지는 '빠른 심사, 빠른 보장'입니다."
    )

    if uploaded_file:
        st.info(f"업로드된 파일: {uploaded_file.name}")

        if st.button("수정 제안 받기", type="primary", key="btn_suggest"):
            with st.spinner("파일을 분석하는 중... 잠시만 기다려주세요."):
                try:
                    file_text = ""

                    if uploaded_file.name.lower().endswith(".pdf"):
                        with pdfplumber.open(io.BytesIO(uploaded_file.read())) as pdf:
                            for page in pdf.pages:
                                text = page.extract_text()
                                if text:
                                    file_text += text + "\n\n"

                    elif uploaded_file.name.lower().endswith(".pptx"):
                        prs = Presentation(io.BytesIO(uploaded_file.read()))
                        for i, slide in enumerate(prs.slides, 1):
                            slide_text = ""
                            for shape in slide.shapes:
                                if hasattr(shape, "text") and shape.text.strip():
                                    slide_text += shape.text.strip() + "\n"
                            if slide_text:
                                file_text += f"[슬라이드 {i}]\n{slide_text}\n"

                    if not file_text.strip():
                        st.warning("⚠️ 파일에서 텍스트를 추출하지 못했습니다. 이미지로만 구성된 파일이거나 스캔본일 수 있어요.")
                    else:
                        context_line = f"\n광고 소재 추가 정보: {context.strip()}" if context.strip() else ""

                        prompt = f"""당신은 보험 광고 심의 전문가입니다.
아래 보험 광고 심의서 내용을 분석하고, 구체적인 수정 방안을 제안해주세요.{context_line}

다음 형식으로 답변해주세요:

## 1. 심의 지적 사항 요약
(어떤 점들이 문제인지 간략히 정리)

## 2. 문구 수정 제안
각 지적 항목에 대해:
- **현재 문구**: (원래 문구)
- **수정 제안**: (바꿔야 할 문구)
- **이유**: (왜 바꿔야 하는지)

## 3. 디자인/구성 수정 제안
(레이아웃, 이미지, 강조 표시 등 비주얼 관련 수정 사항)

## 4. 우선 수정 순위
(가장 중요한 것부터 순서대로)

심의서 내용:
{file_text}"""

                        response = client.models.generate_content(model="gemini-2.0-flash", contents=prompt)
                        st.success("분석 완료!")
                        st.markdown("---")
                        st.markdown(response.text)
                        st.markdown("---")
                        st.download_button(
                            label="📥 수정 제안 저장",
                            data=response.text,
                            file_name="수정제안.txt",
                            mime="text/plain"
                        )

                except Exception as e:
                    st.error(f"오류가 발생했습니다: {str(e)}")
