import io, json, os, time, base64
from pathlib import Path

import streamlit as st
from pypdf import PdfReader
from pptx import Presentation
from google import genai
from google.genai import types
from supabase import create_client, Client

st.set_page_config(page_title="AI 학습 도우미 (클라우드 저장)", page_icon="📚", layout="wide")
st.title("📚 AI 학습 도우미: 클라우드 영구 보관 & 퀴즈 생성기")

# ---------- 1. Supabase 클라이언트 초기화 ----------
supabase_url = st.secrets.get("SUPABASE_URL") or os.getenv("SUPABASE_URL")
supabase_key = st.secrets.get("SUPABASE_KEY") or os.getenv("SUPABASE_KEY")

if not supabase_url or not supabase_key:
    st.error("⚠️ Supabase 연동 정보(SUPABASE_URL, SUPABASE_KEY)가 Secrets에 설정되지 않았습니다.")
    st.stop()

supabase: Client = create_client(supabase_url, supabase_key)

# ---------- 2. 파일 추출 함수 ----------
def extract_pdf(data):
    reader = PdfReader(io.BytesIO(data))
    pages = []
    for i, page in enumerate(reader.pages, 1):
        pages.append({"page": i, "text": page.extract_text() or "", "images": []})
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
                        "b64": base64.b64encode(blob).decode("utf-8")
                    })
                except Exception:
                    pass
        slides.append({"page": i, "text": "\n".join(texts), "images": images})
    return slides

def extract(data, filename):
    ext = Path(filename).suffix.lower()
    if ext == ".pdf":
        return extract_pdf(data)
    if ext == ".pptx":
        return pptx_to_images(data)
    raise ValueError("PDF 또는 PPTX만 지원합니다.")

# ---------- 3. DB 작업 헬퍼 함수 ----------
def fetch_user_data(username):
    response = supabase.table("user_documents").select("*").eq("username", username).execute()
    return response.data or []

def save_document(username, folder_name, filename, slides):
    existing = supabase.table("user_documents").select("id")\
        .eq("username", username)\
        .eq("folder_name", folder_name)\
        .eq("filename", filename).execute()
    
    if existing.data:
        supabase.table("user_documents").update({"slides_data": slides, "analysis_result": None})\
            .eq("id", existing.data[0]["id"]).execute()
    else:
        supabase.table("user_documents").insert({
            "username": username,
            "folder_name": folder_name,
            "filename": filename,
            "slides_data": slides,
            "analysis_result": None
        }).execute()

def update_analysis_result(username, folder_name, filename, result):
    supabase.table("user_documents").update({"analysis_result": result})\
        .eq("username", username)\
        .eq("folder_name", folder_name)\
        .eq("filename", filename).execute()

def delete_document(doc_id):
    supabase.table("user_documents").delete().eq("id", doc_id).execute()

# ---------- 4. Gemini AI 분석 로직 ----------
def ai_analyze(slides, settings):
    api_key = st.secrets.get("GEMINI_API_KEY") or os.getenv("GEMINI_API_KEY") or settings.get("api_key_input")
    if not api_key:
        return None, "Gemini API 키가 설정되지 않았습니다."

    client = genai.Client(api_key=api_key)

    system_instruction = f"""
You are an elite university exam tutor.
Analyze the provided slides thoroughly and generate detailed Korean lecture notes along with rigorous English exam questions.

[SUMMARY RULES - KOREAN]
1. 모든 슬라이드/페이지에 대해 내용을 생략하지 말고 누락 없이 매우 상세하게 분석하라.
2. 'page_summaries' 항목에 각 페이지 번호(page_num)별 핵심 주제(topic), 주요 개념 및 세부 내용(details)을 상세한 한국어로 작성하라.
3. 전공/학술 용어, 공식, 약어, 핵심 키워드는 원문 영어 용어를 반드시 병기하라 (예: 활동 전위(Action Potential)).
4. 슬라이드에 다이어그램, 도표, 이미지가 포함되어 있다면 시각자료 내용(visual_analysis)도 상세히 서술하라.

[QUIZ RULES - QUESTIONS IN ENGLISH, EXPLANATION IN KOREAN]
1. Questions, choices, and answers MUST BE WRITTEN IN ENGLISH.
2. The 'explanation' field MUST BE WRITTEN IN DETAILED KOREAN (한국어 해설 필수).
3. Number of questions: {settings['n']}
4. Question types: {settings['types']}
5. Difficulty: {settings['difficulty']}

Output MUST strictly follow this JSON schema:
{{
  "overall_summary": "전체 요약 (한국어)",
  "page_summaries": [
    {{
      "page_num": 1,
      "topic": "슬라이드 주제",
      "details": ["세부 설명 1", "세부 설명 2"],
      "visual_analysis": "도표/그림 분석 (없으면 빈칸)"
    }}
  ],
  "key_terminology": [
    {{"term_en": "Term", "term_kr": "한국어 용어", "definition": "설명"}}
  ],
  "questions": [
    {{
      "id": 1,
      "type": "Type",
      "source_page": 1,
      "question": "Question in English",
      "choices": ["A. ...", "B. ...", "C. ...", "D. ..."],
      "answer": "Answer in English",
      "explanation": "상세한 한국어 해설"
    }}
  ]
}}
"""

    contents = []
    for s in slides:
        txt = s.get("text", "")
        contents.append(f"\n\n--- SLIDE/PAGE {s['page']} ---\n{txt[:10000]}")
        for img in s.get("images", []):
            try:
                raw_bytes = base64.b64decode(img["b64"])
                contents.append(types.Part.from_bytes(data=raw_bytes, mime_type=img["mime"]))
            except Exception:
                pass

    models_to_try = [settings["model"], "gemini-3.6-flash", "gemini-3.8-flash"]
    models_to_try = list(dict.fromkeys(models_to_try))

    last_error = ""
    for target_model in models_to_try:
        for _ in range(2):
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

    return None, f"서버 과부하로 실패했습니다. 잠시 후 다시 시도해주세요. ({last_error})"

# ---------- 5. 사이드바 UI (사용자 식별 & 설정) ----------
with st.sidebar:
    st.header("👤 계정 접속")
    username_input = st.text_input("사용자 이름(ID 또는 핀번호)", placeholder="예: user1234, 철수").strip()
    
    if not username_input:
        st.warning("계정 이름을 입력해야 개인 저장소가 활성화됩니다.")
        st.stop()
        
    st.success(f"접속 계정: **{username_input}**")

    # DB에서 현재 사용자의 데이터 로드
    user_records = fetch_user_data(username_input)
    
    st.markdown("---")
    st.header("📁 폴더 관리")
    
    # 등록된 폴더 목록 추출
    existing_folders = sorted(list(set([r["folder_name"] for r in user_records] + ["기본 강의자료"])))
    
    new_folder = st.text_input("새 폴더 생성", placeholder="폴더 이름")
    if st.button("➕ 폴더 추가") and new_folder.strip():
        if new_folder.strip() not in existing_folders:
            existing_folders.append(new_folder.strip())
            st.rerun()

    current_folder = st.selectbox("현재 작업 폴더", existing_folders)

    st.markdown("---")
    st.header("⚙️ 퀴즈 옵션")
    n = st.slider("문제 수", 1, 20, 5)
    types_selected = st.multiselect("문제 유형", ["Multiple Choice", "True/False", "Short Answer"], ["Multiple Choice"])
    difficulty = st.select_slider("난이도", ["Basic", "Intermediate", "Advanced", "Exam Level"], value="Intermediate")
    model = st.selectbox("Gemini 모델", ["gemini-3.6-flash", "gemini-3.8-flash"], index=0)

# ---------- 6. 메인: 파일 업로드 및 DB 동기화 ----------
st.subheader(f"📂 폴더: `{current_folder}` (계정: `{username_input}`)")

uploaded_files = st.file_uploader("📎 새 강의 자료 업로드 (PPTX, PDF)", type=["pdf", "pptx"], accept_multiple_files=True)
if uploaded_files:
    with st.spinner("파일을 파싱하여 클라우드에 영구 보관 중..."):
        for up in uploaded_files:
            try:
                extracted_slides = extract(up.getvalue(), up.name)
                save_document(username_input, current_folder, up.name, extracted_slides)
            except Exception as e:
                st.error(f"{up.name} 처리 중 오류: {e}")
        st.success("클라우드 저장소에 성공적으로 동기화되었습니다!")
        time.sleep(1)
        st.rerun()

# 현재 폴더에 속한 문서만 필터링
folder_docs = [r for r in user_records if r["folder_name"] == current_folder]

if not folder_docs:
    st.info("이 폴더에 저장된 파일이 없습니다. 자료를 업로드해 두면 새로고침해도 영구 유지됩니다.")
    st.stop()

# 파일 선택 및 삭제 UI
st.markdown("### 📄 보관된 자료 목록")
doc_map = {r["filename"]: r for r in folder_docs}
col_sel, col_del = st.columns([3, 1])

with col_sel:
    selected_filename = st.selectbox("분석할 파일 선택", list(doc_map.keys()))
with col_del:
    st.write("")
    if st.button("🗑️ 선택 파일 클라우드에서 삭제"):
        delete_document(doc_map[selected_filename]["id"])
        st.rerun()

current_doc = doc_map[selected_filename]
slides = current_doc["slides_data"]
analysis_result = current_doc.get("analysis_result")

# ---------- 7. 탭 UI ----------
tab1, tab2, tab3 = st.tabs(["📑 슬라이드 원본", "📝 상세 한국어 요약노트", "🎯 English Quiz (한국어 해설)"])

with tab1:
    st.info(f"총 {len(slides)}개의 슬라이드/페이지가 보관되어 있습니다.")
    for s in slides:
        with st.expander(f"Slide/Page {s['page']}"):
            st.write(s.get("text") or "(텍스트 없음)")
            if s.get("images"):
                st.caption(f"📸 포함된 이미지: {len(s['images'])}개")

with tab2:
    col_btn, _ = st.columns([1, 3])
    with col_btn:
        run_btn = st.button("🚀 상세 요약 & 퀴즈 생성", type="primary")

    if run_btn:
        settings = {
            "n": n,
            "types": ", ".join(types_selected),
            "difficulty": difficulty,
            "model": model
        }
        with st.spinner("AI가 분석 중입니다... 완료 시 클라우드에 자동 영구 저장됩니다."):
            result, err = ai_analyze(slides, settings)
        if err:
            st.error(err)
        else:
            update_analysis_result(username_input, current_folder, selected_filename, result)
            st.success("분석 완료 및 영구 저장 완료!")
            analysis_result = result
            st.rerun()

    if analysis_result:
        st.markdown("### 📌 전반적인 강의 개요")
        st.info(analysis_result.get("overall_summary", "요약 내용이 없습니다."))

        if analysis_result.get("key_terminology"):
            st.markdown("### 📖 핵심 학술/전문 용어 (Terminology)")
            cols = st.columns(2)
            for idx, term in enumerate(analysis_result.get("key_terminology", [])):
                with cols[idx % 2]:
                    st.markdown(f"**🔹 {term.get('term_en','')}** ({term.get('term_kr','')})")
                    st.write(term.get("definition",""))

        st.markdown("---")
        st.markdown("### 📑 슬라이드/페이지별 상세 정리")
        for page_data in analysis_result.get("page_summaries", []):
            with st.expander(f"📍 [Page {page_data.get('page_num')}] {page_data.get('topic')}", expanded=True):
                st.markdown("**세부 내용:**")
                for d in page_data.get("details", []):
                    st.markdown(f"- {d}")
                if page_data.get("visual_analysis"):
                    st.markdown(f"🖼️ **시각자료 분석:** {page_data.get('visual_analysis')}")
    else:
        st.info("상단의 **'🚀 상세 요약 & 퀴즈 생성'** 버튼을 누르면 AI 분석이 실행되며 결과가 계정에 영구 저장됩니다.")

with tab3:
    if not analysis_result:
        st.info("먼저 요약 노트를 생성해 주세요.")
    else:
        questions = analysis_result.get("questions", [])
        st.markdown(f"### 📝 Practice Exam ({len(questions)} Questions)")
        for idx, q in enumerate(questions, 1):
            st.markdown(f"#### Q{idx}. {q.get('question')}")
            st.caption(f"Source: Page {q.get('source_page')} | Type: {q.get('type')}")
            for ch in q.get("choices", []):
                st.write(ch)
            with st.expander(f"정답 및 한국어 해설 확인 (Q{idx})"):
                st.markdown(f"**Answer:** `{q.get('answer')}`")
                st.markdown(f"**해설:** {q.get('explanation')}")
            st.write("")
