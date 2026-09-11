import io, json, os, time
from pathlib import Path

import streamlit as st
from pypdf import PdfReader
from pptx import Presentation
from google import genai
from google.genai import types

st.set_page_config(page_title="AI 강의 노트 & 퀴즈 생성기", page_icon="📚", layout="wide")
st.title("📚 AI 학습 도우미: 상세 요약 & 영문 퀴즈 생성기")
st.caption("강의자료(PPT/PDF)를 폴더별로 정리하고, 페이지별 상세 한국어 요약과 영문 퀴즈(한국어 해설 포함)를 생성합니다.")

# ---------- 세션 상태 초기화 (폴더 & 파일 관리) ----------
if "folders" not in st.session_state:
    st.session_state["folders"] = {"기본 강의자료": {}}

if "current_folder" not in st.session_state:
    st.session_state["current_folder"] = "기본 강의자료"

if "analysis_results" not in st.session_state:
    st.session_state["analysis_results"] = {}

# ---------- extraction ----------
def extract_pdf(data):
    reader = PdfReader(io.BytesIO(data))
    pages = []
    for i, page in enumerate(reader.pages, 1):
        pages.append({"page": i, "text": page.extract_text() or ""})
    return pages

def pptx_to_images(data):
    prs = Presentation(io.BytesIO(data))
    slides = []
    for i, slide in enumerate(prs.slides, 1):
        texts = []
        images = []
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                texts.append(shape.text.strip())
            if shape.shape_type == 13:  # picture
                try:
                    blob = shape.image.blob
                    mime = "image/png" if shape.image.ext in ("png",) else "image/jpeg"
                    images.append({
                        "mime": mime,
                        "raw_bytes": blob
                    })
                except Exception:
                    pass
        slides.append({"page": i, "text": "\n".join(texts), "images": images})
    return slides

def extract(data, filename):
    ext = Path(filename).suffix.lower()
    if ext == ".pdf":
        pages = extract_pdf(data)
        for p in pages:
            p["images"] = []
        return pages
    if ext == ".pptx":
        return pptx_to_images(data)
    raise ValueError("PDF 또는 PPTX만 지원합니다.")

# ---------- Gemini AI 호출 로직 ----------
def ai_analyze(slides, settings):
    api_key = None
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
    elif os.getenv("GEMINI_API_KEY"):
        api_key = os.getenv("GEMINI_API_KEY")
    elif settings.get("api_key_input"):
        api_key = settings["api_key_input"]

    if not api_key:
        return None, "Gemini API 키가 없습니다. Streamlit Cloud Secrets에 'GEMINI_API_KEY'를 추가해주세요."

    client = genai.Client(api_key=api_key)

    system_instruction = f"""
You are an elite university exam tutor and curriculum specialist.
Your task is to analyze the provided lecture materials in depth and generate both comprehensive study notes and rigorous quiz questions.

[CRITICAL INSTRUCTION FOR SUMMARY - KOREAN]
1. 모든 슬라이드/페이지에 대해 내용을 생략하지 말고 누락 없이 매우 상세하게 분석하라.
2. 'page_summaries' 항목에 각 페이지 번호(page_num)별로 다루는 핵심 주제(topic), 주요 개념 및 세부 내용(details)을 상세한 한국어로 작성하라.
3. 전공/학술 용어, 공식, 약어, 핵심 키워드는 원문 영어 용어를 반드시 병기하거나 적절히 활용하라 (예: 활동 전위(Action Potential), 역치(Threshold)).
4. 슬라이드에 다이어그램, 도표, 이미지가 포함되어 있다면 그 시각자료가 설명하는 내용(visual_analysis)도 상세히 서술하라.

[CRITICAL INSTRUCTION FOR QUIZ QUESTIONS]
1. Questions, choices, and answers MUST BE WRITTEN IN ENGLISH.
2. However, the 'explanation' field MUST BE WRITTEN IN THOROUGH AND CLEAR KOREAN (한국어 해설 필수).
   - 해설에는 정답의 이유, 오답 선지가 틀린 이유, 관련 전공 개념을 친절하고 상세하게 한국어로 설명하라.
3. Number of questions: {settings['n']}
4. Question types: {settings['types']}
5. Difficulty: {settings['difficulty']}
6. Style: {settings['styles']}

Output MUST strictly follow this JSON schema:
{{
  "overall_summary": "전체 강의 자료를 아우르는 핵심 요약 (한국어)",
  "page_summaries": [
    {{
      "page_num": 1,
      "topic": "슬라이드 핵심 주제 (한국어/영어)",
      "details": ["세부 학습 내용 1 (상세 설명)", "세부 학습 내용 2", "핵심 수식 또는 정의 등"],
      "visual_analysis": "그림/도표가 있는 경우 분석 내용 (없으면 빈 문자열)"
    }}
  ],
  "key_terminology": [
    {{"term_en": "English Term", "term_kr": "한국어 번역", "definition": "상세한 학술적 설명 (한국어 + 영어 키워드)"}}
  ],
  "questions": [
    {{
      "id": 1,
      "type": "Multiple Choice / True-False / Short Answer",
      "source_page": 1,
      "question": "Question text in English",
      "choices": ["A. ...", "B. ...", "C. ...", "D. ..."],
      "answer": "Correct answer in English",
      "explanation": "정답 및 오답에 대한 상세한 한국어 해설 (필요시 영어 전공 용어 병기)"
    }}
  ]
}}
"""

    contents = []
    for s in slides:
        txt = s.get("text", "")
        contents.append(f"\n\n--- SLIDE/PAGE {s['page']} ---\n{txt[:10000]}")
        for img in s.get("images", []):
            contents.append(
                types.Part.from_bytes(
                    data=img["raw_bytes"],
                    mime_type=img["mime"]
                )
            )

    models_to_try = [settings["model"], "gemini-3.6-flash", "gemini-3.8-flash"]
    models_to_try = list(dict.fromkeys(models_to_try))

    last_error = ""
    for target_model in models_to_try:
        for attempt in range(2):
            try:
                response = client.models.generate_content(
                    model=target_model,
                    contents=contents,
                    config=types.GenerateContentConfig(
                        system_instruction=system_instruction,
                        response_mime_type="application/json",
                        temperature=0.2
                    )
                )
                return json.loads(response.text), None
            except Exception as e:
                last_error = str(e)
                if "503" in last_error:
                    time.sleep(3)
                    continue
                elif "404" in last_error:
                    break
                else:
                    return None, f"오류 발생: {last_error}"

    return None, f"API 서버 과부하로 처리에 실패했습니다. 잠시 후 다시 시도해 주세요. ({last_error})"

# ---------- 사이드바: 폴더 및 파일 관리 & 퀴즈 옵션 ----------
with st.sidebar:
    st.header("📁 자료실 폴더 관리")
    
    new_folder_name = st.text_input("새 폴더 이름", placeholder="예: 생물학 1학기, 컴퓨터구조")
    if st.button("➕ 폴더 생성") and new_folder_name.strip():
        folder_clean = new_folder_name.strip()
        if folder_clean not in st.session_state["folders"]:
            st.session_state["folders"][folder_clean] = {}
            st.session_state["current_folder"] = folder_clean
            st.success(f"'{folder_clean}' 폴더가 생성되었습니다.")
            st.rerun()

    folder_list = list(st.session_state["folders"].keys())
    selected_folder = st.selectbox("현재 작업 폴더", folder_list, index=folder_list.index(st.session_state["current_folder"]))
    st.session_state["current_folder"] = selected_folder

    st.markdown("---")
    st.header("⚙️ 퀴즈 설정")
    n = st.slider("문제 수 (Questions)", 1, 20, 5)
    types_selected = st.multiselect("문제 유형", ["Multiple Choice (객관식)", "True/False (OX)", "Short Answer (단답형)"], ["Multiple Choice (객관식)"])
    difficulty = st.select_slider("난이도", ["Basic", "Intermediate", "Advanced", "Exam Level"], value="Intermediate")
    styles = st.multiselect(
        "출제 스타일",
        ["Key Concepts", "Comparison & Contrast", "Diagram/Graph Interpretation", "Application & Scenarios"],
        ["Key Concepts", "Diagram/Graph Interpretation"]
    )
    model = st.selectbox("Gemini 모델", ["gemini-3.6-flash", "gemini-3.8-flash"], index=0)

    if not ("GEMINI_API_KEY" in st.secrets or os.getenv("GEMINI_API_KEY")):
        api_key_input = st.text_input("Gemini API Key", type="password")
    else:
        api_key_input = None

# ---------- 메인 화면: 파일 업로드 및 보관함 ----------
st.subheader(f"📂 현재 폴더: `{st.session_state['current_folder']}`")

uploaded_files = st.file_uploader(
    "📎 강의 자료 추가 (PPTX, PDF 다중 업로드 가능)", 
    type=["pdf", "pptx"], 
    accept_multiple_files=True
)

if uploaded_files:
    added_count = 0
    for up in uploaded_files:
        if up.name not in st.session_state["folders"][selected_folder]:
            st.session_state["folders"][selected_folder][up.name] = {
                "bytes": up.getvalue(),
                "name": up.name,
                "ext": Path(up.name).suffix.lower()
            }
            added_count += 1
    if added_count > 0:
        st.success(f"{added_count}개의 파일이 '{selected_folder}' 폴더에 저장되었습니다!")

current_files = st.session_state["folders"][selected_folder]
if not current_files:
    st.info("이 폴더에 저장된 파일이 없습니다. 위에서 PPTX 또는 PDF 파일을 업로드해 주세요.")
    st.stop()

st.markdown("### 📄 보관된 자료 목록")
file_names = list(current_files.keys())

col_sel, col_del = st.columns([3, 1])
with col_sel:
    target_filename = st.selectbox("분석할 파일 선택", file_names)
with col_del:
    st.write("")
    if st.button("🗑️ 선택 파일 삭제"):
        del st.session_state["folders"][selected_folder][target_filename]
        cache_key = f"{selected_folder}/{target_filename}"
        if cache_key in st.session_state["analysis_results"]:
            del st.session_state["analysis_results"][cache_key]
        st.rerun()

target_file_data = current_files[target_filename]["bytes"]
try:
    slides = extract(target_file_data, target_filename)
except Exception as e:
    st.error(f"파일을 읽는 도중 오류가 발생했습니다: {e}")
    st.stop()

cache_key = f"{selected_folder}/{target_filename}"

# ---------- 분석 탭 UI ----------
tab1, tab2, tab3 = st.tabs(["📑 슬라이드 원본", "📝 상세 한국어 요약노트", "🎯 English Quiz (한국어 해설)"])

with tab1:
    st.info(f"총 {len(slides)}개의 슬라이드/페이지가 감지되었습니다.")
    for s in slides:
        with st.expander(f"Slide/Page {s['page']}"):
            st.write(s.get("text","") or "(텍스트 없음)")
            if s.get("images"):
                st.caption(f"📸 포함된 이미지/도표: {len(s['images'])}개")

with tab2:
    col_btn, col_status = st.columns([1, 3])
    with col_btn:
        run_analysis = st.button("🚀 상세 요약 & 퀴즈 생성", type="primary")
    
    if run_analysis:
        settings = {
            "n": n,
            "types": ", ".join(types_selected),
            "difficulty": difficulty,
            "styles": ", ".join(styles),
            "model": model,
            "api_key_input": api_key_input
        }
        with st.spinner("AI가 페이지별 세부 내용과 시각자료를 빠짐없이 분석 중입니다..."):
            result, err = ai_analyze(slides, settings)
        if err:
            st.error(err)
        else:
            st.session_state["analysis_results"][cache_key] = result
            st.success("분석이 완료되었습니다!")

    if cache_key in st.session_state["analysis_results"]:
        res = st.session_state["analysis_results"][cache_key]
        
        # 1. 전체 핵심 요약
        st.markdown("### 📌 전반적인 강의 개요")
        st.info(res.get("overall_summary", "요약 내용이 없습니다."))
        
        # 2. 전문 핵심 용어
        if res.get("key_terminology"):
            st.markdown("### 📖 핵심 학술/전문 용어 (Terminology)")
            cols = st.columns(2)
            for idx, term in enumerate(res.get("key_terminology", [])):
                with cols[idx % 2]:
                    st.markdown(f"**🔹 {term.get('term_en','')}** ({term.get('term_kr','')})")
                    st.write(term.get("definition",""))
        
        # 3. 페이지별 누락 없는 상세 요약
        st.markdown("---")
        st.markdown("### 📑 슬라이드/페이지별 상세 정리")
        for page_data in res.get("page_summaries", []):
            p_num = page_data.get("page_num", "?")
            p_topic = page_data.get("topic", "상세 내용")
            
            with st.expander(f"📍 [Page {p_num}] {p_topic}", expanded=True):
                st.markdown("**세부 핵심 내용:**")
                for d in page_data.get("details", []):
                    st.markdown(f"- {d}")
                
                vis = page_data.get("visual_analysis", "")
                if vis and vis.strip():
                    st.markdown(f"🖼️ **시각 자료(도표/그림) 분석:** {vis}")
    else:
        st.info("상단의 **'🚀 상세 요약 & 퀴즈 생성'** 버튼을 눌러 분석을 시작하세요.")

with tab3:
    if cache_key not in st.session_state["analysis_results"]:
        st.info("먼저 '상세 한국어 요약노트' 탭에서 분석을 진행해 주세요.")
    else:
        res = st.session_state["analysis_results"][cache_key]
        questions = res.get("questions", [])
        
        if not questions:
            st.warning("생성된 문제가 없습니다.")
        else:
            st.markdown(f"### 📝 Practice Exam ({len(questions)} Questions)")
            st.caption("Questions & Choices: English | Explanations: Korean")
            
            for idx, q in enumerate(questions, 1):
                st.markdown(f"#### Q{idx}. {q.get('question','')}")
                st.caption(f"Source: Page/Slide {q.get('source_page', '?')} | Type: {q.get('type', '')}")
                
                choices = q.get("choices", [])
                if choices:
                    for ch in choices:
                        st.write(ch)
                
                with st.expander(f"정답 및 한국어 해설 확인 (Q{idx})"):
                    st.markdown(f"**Answer:** `{q.get('answer', '')}`")
                    st.markdown(f"**해설:** {q.get('explanation', '')}")
                st.write("")