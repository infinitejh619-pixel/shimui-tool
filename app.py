import streamlit as st
from groq import Groq
import pdfplumber
from pptx import Presentation
import io, json, os, base64
from datetime import datetime, timezone, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed

st.set_page_config(page_title="심의의견 분석 도구", page_icon="📋", layout="wide")

KST = timezone(timedelta(hours=9))

def now_kst():
    return datetime.now(KST).strftime("%Y.%m.%d %H:%M")

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
MEDIA_FILE   = "media_settings.json"

def load_history():
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []

def save_to_history(email_text, result, job_name="", filename=""):
    history = load_history()
    label = job_name.strip() if job_name.strip() else f"작업 {now_kst()}"
    item = {
        "id":        datetime.now(KST).strftime("%Y%m%d_%H%M%S"),
        "timestamp": now_kst(),
        "job_name":  label,
        "filename":  filename,
        "preview":   email_text[:40].replace("\n", " ") + ("..." if len(email_text) > 40 else ""),
        "email_text": email_text,
        "result":    result,
    }
    history.insert(0, item)
    history = history[:30]
    try:
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(history, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

DEFAULT_MEDIA = [
    {"media": "네이버 GFA",  "placement": "커뮤니케이션애즈", "category": "광고문구 1", "notes": "윗배너",                          "limit": 40, "include_spaces": True},
    {"media": "당근",        "placement": "배너",             "category": "브랜드이름", "notes": "브랜드명+보종명까지 같이 기입 가능", "limit": 20, "include_spaces": True},
    {"media": "당근",        "placement": "배너",             "category": "광고 제목",  "notes": "광고 문구",                        "limit": 30, "include_spaces": True},
    {"media": "토스",        "placement": "배너",             "category": "주요 문구",  "notes": "광고 문구 (윗줄)",                 "limit": 18, "include_spaces": True},
    {"media": "토스",        "placement": "배너",             "category": "보조문구",   "notes": "광고 문구 (아랫줄에 작게 들어감)", "limit": 18, "include_spaces": True},
]

def load_media():
    if os.path.exists(MEDIA_FILE):
        try:
            with open(MEDIA_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return DEFAULT_MEDIA[:]
    return DEFAULT_MEDIA[:]

def save_media(settings):
    try:
        with open(MEDIA_FILE, "w", encoding="utf-8") as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def ask_groq(prompt):
    response = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role": "user", "content": prompt}],
    )
    return response.choices[0].message.content

def ocr_image(image_bytes, slide_num, image_num):
    try:
        img_type = "image/png" if image_bytes[:4] == b'\x89PNG' else "image/jpeg"
        img_b64  = base64.b64encode(image_bytes).decode("utf-8")
        response = client.chat.completions.create(
            model="meta-llama/llama-4-scout-17b-16e-instruct",
            temperature=0,
            messages=[{"role": "user", "content": [
                {"type": "text",
                 "text": (
                    f"이 이미지는 보험 광고 소재({slide_num}p, {image_num}번 이미지)입니다.\n"
                    "이미지 안에 보이는 텍스트를 '보이는 그대로' 정확히 옮겨 쓰세요.\n"
                    "절대 하지 말아야 할 것:\n"
                    "- 문장을 완성하거나 추측하지 마세요\n"
                    "- 없는 글자를 추가하지 마세요\n"
                    "- 비슷해 보이는 다른 글자로 바꾸지 마세요\n"
                    "- 문맥에 맞게 고치지 마세요\n"
                    "보이는 글자만, 있는 그대로, 순서대로 출력하세요."
                 )},
                {"type": "image_url", "image_url": {"url": f"data:{img_type};base64,{img_b64}"}}
            ]}]
        )
        return response.choices[0].message.content.strip()
    except Exception:
        return ""

def extract_file_text(uploaded_file, use_ocr=False):
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
        ocr_results = {}

        if use_ocr:
            image_tasks = []
            for i, slide in enumerate(prs.slides, 1):
                img_num = 0
                for shape in slide.shapes:
                    if shape.shape_type == 13:
                        img_num += 1
                        image_tasks.append((i, img_num, shape.image.blob))
            if image_tasks:
                with ThreadPoolExecutor(max_workers=6) as executor:
                    futures = {
                        executor.submit(ocr_image, blob, sn, inn): (sn, inn)
                        for sn, inn, blob in image_tasks
                    }
                    for future in as_completed(futures):
                        sn, inn = futures[future]
                        ocr_results[(sn, inn)] = future.result()

        for i, slide in enumerate(prs.slides, 1):
            slide_text = ""
            img_num = 0
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text.strip():
                    slide_text += shape.text.strip() + "\n"
                if shape.shape_type == 13:
                    img_num += 1
                    if use_ocr:
                        ocr = ocr_results.get((i, img_num), "")
                        slide_text += (f"[{img_num}번 이미지 내 텍스트]\n{ocr}\n" if ocr
                                       else f"[{img_num}번 이미지 내 텍스트: 추출 실패]\n")
                    else:
                        slide_text += f"[{img_num}번 이미지: OCR 비활성화로 읽을 수 없음]\n"
            if slide_text:
                file_text += f"[{i}페이지]\n{slide_text}\n"

    return file_text

# ── 사이드바 히스토리 ──────────────────────────────────────────────────────────
history = load_history()
with st.sidebar:
    st.header("📁 최근 작업 히스토리")
    if not history:
        st.caption("아직 작업 내역이 없어요.")
    else:
        for i, item in enumerate(history, 1):
            job_name = item.get("job_name", "(이름 없음)")
            filename = item.get("filename", "없음")
            timestamp = item.get("timestamp", "")

            with st.expander(f"{i}. {job_name}"):
                st.caption(f"🕐 작업 시간: {timestamp}")
                st.caption(f"📎 파일명: {filename}")
                if st.button("이 작업 불러오기", key=f"load_{item['id']}", use_container_width=True):
                    st.session_state["selected_history"] = item
                    st.rerun()

st.title("📋 보험 광고 심의의견 분석 도구")
tab1, tab2, tab3 = st.tabs(["🔍 심의의견 ↔ 소재 1:1 매칭", "✅ 수정 반영 검수", "✏️ 매체 글자수 체크"])

# ═══════════════════════════════════════════════════════════════════════════════
# 탭 1 — 1:1 매칭
# ═══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.subheader("심의의견 ↔ 소재 1:1 매칭")

    job_name_input = st.text_input("작업명을 입력하세요",
        placeholder="예: 삼성화재 암보험 5월 GFA 심의 1차")

    col1, col2 = st.columns(2)
    with col1:
        email_text = st.text_area("심의 메일 붙여넣기", height=300,
            placeholder="광고주로부터 받은 심의 메일을 붙여넣으세요...")
    with col2:
        uploaded_file = st.file_uploader("소재 파일 업로드 (PDF 또는 PPT)", type=["pdf", "pptx"])
        if uploaded_file:
            st.info(f"업로드: {uploaded_file.name}")
        use_ocr = st.toggle("🖼️ 이미지 내 텍스트도 읽기 (OCR)",
            value=False, help="켜면 이미지 안 텍스트도 추출하지만 시간이 더 걸려요.")

    if st.button("1:1 매칭 분석하기", type="primary", key="btn_match"):
        if not email_text.strip():
            st.error("심의 메일을 입력해주세요.")
        elif not uploaded_file:
            st.error("소재 파일을 업로드해주세요.")
        else:
            msg = "이미지 OCR 포함 분석 중..." if use_ocr else "소재 분석 중..."
            with st.spinner(msg):
                try:
                    file_text = extract_file_text(uploaded_file, use_ocr=use_ocr)
                    if not file_text.strip():
                        st.warning("⚠️ 파일에서 텍스트를 추출하지 못했습니다.")
                    else:
                        if use_ocr:
                            ocr_instruction = "소재 내용에 이미지 OCR 결과가 포함되어 있습니다. 이미지 내 텍스트도 함께 참고하세요."
                        else:
                            ocr_instruction = (
                                "⚠️ OCR 비활성화 상태입니다. '[N번 이미지: OCR 비활성화로 읽을 수 없음]'으로 표시된 부분은 이미지이며 텍스트를 읽을 수 없습니다.\n"
                                "심의의견이 해당 이미지의 문구를 지적하는 경우, 반드시 '❌ 현재 문구: 🔒 이미지 인식 불가 (OCR 기능을 켜고 다시 시도하세요)'라고 표시하세요.\n"
                                "절대 이미지 내용을 추측하거나 지어내지 마세요."
                            )

                        prompt = f"""당신은 보험 광고 심의 전문가입니다.
심의 메일의 각 의견과 소재의 해당 부분을 1:1로 매칭해주세요.

{ocr_instruction}

작업 방법:
1. 심의 메일에서 각 심의의견을 원문 그대로 인용합니다.
2. 위치 표기 파싱 규칙을 따라 페이지 번호와 배너 번호를 각각 파악합니다.
3. 소재에서 해당 문구를 찾습니다.
4. 정확한 수정 방향을 제시합니다.

【위치 표기 파싱 규칙 — 반드시 준수】
- "2,7p" → 페이지 번호: 2페이지와 7페이지 (쉼표 = '그리고', 첫 번째만 추출 절대 금지)
- "2-7p" → 페이지 번호: 2페이지부터 7페이지
- "[2,7p_1,6배너]" → 페이지 번호: 2페이지와 7페이지 / 배너 번호: 1번과 6번
- 배너 번호 언급 없으면 '없음'
- 위치 정보 절대 생략·축약 금지

출력 형식 (심의의견마다 반복):

---
**[심의의견 N]**
- 📋 의견 원문: (심의 메일 원문 그대로 인용)
- 📄 페이지 번호: N페이지 (원문 표기 병기)
- 🏷️ 배너 번호: N번 배너 (없으면 '없음')
- ❌ 현재 문구: (소재에서 찾은 실제 문구. 이미지이고 OCR이 꺼져있으면 '🔒 이미지 인식 불가 (OCR을 켜고 다시 시도하세요)')
- ✅ 수정 문구: (바꿔야 할 문구. 현재 문구를 읽지 못한 경우 '이미지 확인 후 수정 필요')
- 💡 수정 이유: (한 줄 요약)
---

===== 심의 메일 =====
{email_text}

===== 소재 내용 (페이지별) =====
{file_text}"""

                        result = ask_groq(prompt)
                        save_to_history(
                            email_text, result,
                            job_name=job_name_input,
                            filename=uploaded_file.name if uploaded_file else ""
                        )
                        st.success("매칭 완료! 히스토리에 저장되었어요.")
                        st.markdown("---")
                        st.markdown(result)
                        st.download_button("📥 결과 저장", data=result,
                            file_name="수정제안.txt", mime="text/plain")
                except Exception as e:
                    st.error(f"오류: {str(e)}")

# ═══════════════════════════════════════════════════════════════════════════════
# 탭 2 — 검수
# ═══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.subheader("수정 반영 검수")
    selected = st.session_state.get("selected_history") or (history[0] if history else None)

    if not selected:
        st.info("👈 왼쪽 히스토리에서 작업을 선택하거나, 먼저 1:1 매칭 분석을 진행해주세요.")
    else:
        st.info(f"📋 선택된 작업: **{selected.get('job_name', '')}** — {selected.get('timestamp', '')}")
        with st.expander("원본 심의의견 분석 결과 보기"):
            st.markdown(selected["result"])
        st.markdown("---")
        revised_file = st.file_uploader("수정된 소재 업로드", type=["pdf", "pptx"], key="revised_file")
        use_ocr2 = st.toggle("🖼️ 이미지 내 텍스트도 읽기 (OCR)", value=False, key="ocr2")

        if revised_file:
            if st.button("검수 시작하기", type="primary", key="btn_check"):
                with st.spinner("검수 중..."):
                    try:
                        revised_text = extract_file_text(revised_file, use_ocr=use_ocr2)
                        if not revised_text.strip():
                            st.warning("⚠️ 파일에서 텍스트를 추출하지 못했습니다.")
                        else:
                                                       prompt = f"""당신은 보험 광고 심의 검수 전문가입니다.
원본 심의의견 분석 결과와 수정된 소재를 비교하여, 각 심의의견이 반영됐는지 검수해주세요.

작업 방법:
1. 원본 분석 결과에서 의견 원문, 페이지 번호, 배너 번호를 그대로 가져옵니다.
2. 수정된 소재에서 해당 페이지/배너 위치의 실제 문구를 직접 발췌합니다.
3. 의견 원문 기준으로 지적 사항이 해소됐는지 판정합니다.

판정 기준:
- ✅ 반영: 심의의견대로 수정됨
- ❌ 미반영: 수정되지 않음
- ⚠️ 부분반영: 일부만 수정됨

출력 형식 (각 심의의견마다):

---
**[심의의견 N]**
- 📋 의견 원문: (원본 분석 결과에서 그대로 인용)
- 📄 페이지 번호: (원본 분석 결과에서 그대로)
- 🏷️ 배너 번호: (원본 분석 결과에서 그대로)
- 🔄 수정 문구: (수정된 소재 파일의 해당 페이지/배너에서 실제 발췌한 문구)
- ✅ 확인 내용: (의견 원문 기준으로 지적 사항이 해소됐는지 구체적으로 설명)
- 판정: ✅ 반영 / ❌ 미반영 / ⚠️ 부분반영
---

**[ 최종 검수 요약 ]**
총 N개 의견 중 ✅ 반영 N개 / ❌ 미반영 N개 / ⚠️ 부분반영 N개
미반영/부분반영 항목은 재수정 필요

===== 원본 심의의견 분석 결과 =====
{selected['result']}

===== 수정된 소재 내용 =====
{revised_text}"""
                            result = ask_groq(prompt)
                            st.success("검수 완료!")
                            st.markdown("---")
                            st.markdown(result)
                            st.download_button("📥 검수 결과 저장", data=result,
                                file_name="검수결과.txt", mime="text/plain")
                    except Exception as e:
                        st.error(f"오류: {str(e)}")

# ═══════════════════════════════════════════════════════════════════════════════
# 탭 3 — 매체 글자수 체크
# ═══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.subheader("매체 글자수 체크")
    media_settings = load_media()

    if "selected_media_idx" not in st.session_state:
        st.session_state["selected_media_idx"] = 0
    if "editing_idx" not in st.session_state:
        st.session_state["editing_idx"] = None

    col_left, col_right = st.columns([1, 1.6])

    with col_left:
        st.markdown("#### 등록된 설정 목록")
        st.caption("항목을 클릭하면 오른쪽에서 글자수 체크를 할 수 있어요.")

        if not media_settings:
            st.info("아래에서 매체/지면을 먼저 등록해주세요.")
        else:
            for idx, s in enumerate(media_settings):
                is_selected = st.session_state["selected_media_idx"] == idx
                is_editing  = st.session_state["editing_idx"] == idx

                if is_editing:
                    with st.container(border=True):
                        st.markdown("**✏️ 편집 중**")
                        e_media     = st.text_input("매체명", value=s["media"],            key=f"e_media_{idx}")
                        e_placement = st.text_input("지면",   value=s["placement"],        key=f"e_place_{idx}")
                        e_category  = st.text_input("구분",   value=s.get("category",""),  key=f"e_cat_{idx}")
                        e_notes     = st.text_input("비고",   value=s.get("notes",""),     key=f"e_notes_{idx}")
                        e_limit     = st.number_input("글자수", value=s["limit"], min_value=1, key=f"e_limit_{idx}")
                        e_spaces    = st.checkbox("공백 포함", value=s["include_spaces"],  key=f"e_spaces_{idx}")
                        c1, c2 = st.columns(2)
                        with c1:
                            if st.button("💾 저장", key=f"save_{idx}", use_container_width=True):
                                media_settings[idx] = {
                                    "media": e_media, "placement": e_placement,
                                    "category": e_category, "notes": e_notes,
                                    "limit": int(e_limit), "include_spaces": e_spaces,
                                }
                                save_media(media_settings)
                                st.session_state["editing_idx"] = None
                                st.rerun()
                        with c2:
                            if st.button("취소", key=f"cancel_{idx}", use_container_width=True):
                                st.session_state["editing_idx"] = None
                                st.rerun()
                else:
                    notes_text  = f" · {s.get('notes','')}" if s.get('notes') else ""
                    space_label = "공백포함" if s["include_spaces"] else "공백제외"
                    prefix      = "✅ " if is_selected else "　 "
                    label       = f"{prefix}{s['media']} — {s.get('category','')}{notes_text}  [{s['limit']}자/{space_label}]"

                    c_btn, c_edit, c_del = st.columns([5, 1, 1])
                    with c_btn:
                        if st.button(label, key=f"sel_{idx}", use_container_width=True):
                            st.session_state["selected_media_idx"] = idx
                            st.rerun()
                    with c_edit:
                        if st.button("✏️", key=f"edit_{idx}", help="편집"):
                            st.session_state["editing_idx"] = idx
                            st.session_state["selected_media_idx"] = idx
                            st.rerun()
                    with c_del:
                        if st.button("🗑", key=f"del_{idx}", help="삭제"):
                            media_settings.pop(idx)
                            save_media(media_settings)
                            st.session_state["selected_media_idx"] = max(0, min(
                                st.session_state["selected_media_idx"], len(media_settings) - 1))
                            st.rerun()

        st.markdown("---")
        with st.expander("➕ 신규 매체/지면 등록"):
            new_media     = st.text_input("매체명",      placeholder="예: 카카오",     key="nm")
            new_placement = st.text_input("지면",        placeholder="예: 배너",       key="np")
            new_category  = st.text_input("구분",        placeholder="예: 광고문구 1", key="nc")
            new_notes     = st.text_input("비고",        placeholder="예: 윗배너",     key="nn")
            new_limit     = st.number_input("글자수 제한", min_value=1, value=20,      key="nl")
            new_spaces    = st.checkbox("공백 포함 글자수", value=True,                key="ns")

            if st.button("등록", type="primary", use_container_width=True):
                if new_media and new_placement and new_category:
                    media_settings.append({
                        "media": new_media, "placement": new_placement,
                        "category": new_category, "notes": new_notes,
                        "limit": int(new_limit), "include_spaces": new_spaces,
                    })
                    save_media(media_settings)
                    st.session_state["selected_media_idx"] = len(media_settings) - 1
                    st.success("등록 완료!")
                    st.rerun()
                else:
                    st.error("매체명, 지면, 구분은 필수입니다.")

    with col_right:
        st.markdown("#### 글자수 체크")
        if not media_settings:
            st.info("왼쪽에서 매체/지면을 먼저 등록해주세요.")
        else:
            idx     = min(st.session_state.get("selected_media_idx", 0), len(media_settings) - 1)
            setting = media_settings[idx]
            notes   = f" · {setting.get('notes','')}" if setting.get('notes') else ""
            st.info(f"**{setting['media']}** — {setting['placement']} — {setting.get('category','')}{notes}  \n"
                    f"제한: **{setting['limit']}자** ({'공백 포함' if setting['include_spaces'] else '공백 제외'})")

            check_text = st.text_area("문구 입력", height=220,
                placeholder="글자수를 체크할 문구를 붙여넣으세요...")

            if check_text:
                count = len(check_text) if setting["include_spaces"] else len(check_text.replace(" ", ""))
                limit = setting["limit"]
                space_note = "공백 포함" if setting["include_spaces"] else "공백 제외"
                st.markdown("---")
                if count <= limit:
                    st.success(f"✅ 통과!  **{count}자** / {limit}자 ({space_note})")
                    st.progress(count / limit)
                else:
                    over = count - limit
                    st.error(f"❌ 초과!  **{count}자** / {limit}자 — {over}자 초과 ({space_note})")
                    st.progress(min(count / limit, 1.0))
                st.caption(f"공백 포함: {len(check_text)}자  ·  공백 제외: {len(check_text.replace(' ',''))}자")
