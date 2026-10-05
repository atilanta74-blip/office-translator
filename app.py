import io
import os
import streamlit as st
import google.generativeai as genai
import openpyxl
from docx import Document
from pptx import Presentation

# Oldal beállításai
st.set_page_config(
    page_title="Universal Office Translator Pro",
    page_icon="🌐",
    layout="centered"
)

# Egyedi modern felület stílus
st.markdown("""
    <style>
    .main-title { font-size: 2.2rem; font-weight: 700; margin-bottom: 0.2rem; }
    .sub-title { color: #94a3b8; font-size: 1rem; margin-bottom: 1.5rem; }
    .footer-bar {
        margin-top: 3rem;
        padding-top: 1rem;
        border-top: 1px solid #334155;
        display: flex;
        justify-content: space-between;
        align-items: center;
        font-size: 0.9rem;
    }
    .powered-by { color: #38bdf8; font-weight: 700; font-size: 1rem; }
    </style>
""", unsafe_allow_html=True)

# Támogatott nyelvek listája
LANGUAGES = {
    "Magyar (Hungarian)": "Hungarian",
    "Angol (English)": "English",
    "Német (German)": "German",
    "Olasz (Italian)": "Italian",
    "Francia (French)": "French",
    "Spanyol (Spanish)": "Spanish",
    "Lengyel (Polish)": "Polish",
    "Cseh (Czech)": "Czech",
    "Szlovák (Slovak)": "Slovak",
    "Román (Romanian)": "Romanian",
    "Japán (Japanese)": "Japanese",
    "Török (Turkish)": "Turkish",
    "Vietnámi (Vietnamese)": "Vietnamese"
}

# Gemini API kulcs beolvasása a Streamlit Secrets tárolóból
gemini_key = st.secrets.get("GEMINI_API_KEY")
if not gemini_key:
    st.error("⚠️ Hiányzik a GEMINI_API_KEY! Kérlek, add meg a Streamlit felületén a 'Secrets' menüben.")
    st.stop()

# Gemini kliens konfigurálása
genai.configure(api_key=gemini_key)
model = genai.GenerativeModel(
    model_name="gemini-1.5-flash",
    system_instruction="You are a professional industrial, technical, and business document translator. Translate the text accurately into the target language. Keep any machine names, codes, formulas, line breaks, and punctuation intact. Return ONLY the translated text, without commentary or extra markdown tags."
)

# Gemini fordító segédfüggvény gyorsítótárazással (Cache)
def translate_text(text, target_lang, cache):
    text_str = str(text).strip()
    if not text_str or text_str.isdigit() or len(text_str) <= 1:
        return text
    if text_str in cache:
        return cache[text_str]
    try:
        prompt = f"Target language: {target_lang}\nText to translate:\n{text_str}"
        response = model.generate_content(prompt)
        translated = response.text.strip()
        cache[text_str] = translated
        return translated
    except Exception:
        return text

# Word (.docx) feldolgozás
def process_docx(file_bytes, target_lang, cache, progress_bar):
    doc = Document(io.BytesIO(file_bytes))
    total_p = len(doc.paragraphs)
    for idx, p in enumerate(doc.paragraphs):
        for run in p.runs:
            if run.text.strip():
                run.text = translate_text(run.text, target_lang, cache)
        if total_p > 0 and idx % 5 == 0:
            progress_bar.progress(int(10 + (idx / total_p) * 40))

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        if run.text.strip():
                            run.text = translate_text(run.text, target_lang, cache)

    progress_bar.progress(90)
    out_stream = io.BytesIO()
    doc.save(out_stream)
    return out_stream.getvalue()

# Excel (.xlsx) feldolgozás
def process_xlsx(file_bytes, target_lang, cache, progress_bar):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    sheets = wb.worksheets
    total_s = len(sheets)

    for idx, ws in enumerate(sheets):
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type == 's' and cell.value:
                    val = str(cell.value)
                    if not val.startswith("="):
                        cell.value = translate_text(val, target_lang, cache)
        if total_s > 0:
            progress_bar.progress(int(10 + ((idx + 1) / total_s) * 80))

    out_stream = io.BytesIO()
    wb.save(out_stream)
    return out_stream.getvalue()

# PowerPoint (.pptx) feldolgozás
def process_pptx(file_bytes, target_lang, cache, progress_bar):
    prs = Presentation(io.BytesIO(file_bytes))
    total_slides = len(prs.slides)

    for idx, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    for run in p.runs:
                        if run.text.strip():
                            run.text = translate_text(run.text, target_lang, cache)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for p in cell.text_frame.paragraphs:
                            for run in p.runs:
                                if run.text.strip():
                                    run.text = translate_text(run.text, target_lang, cache)
        if total_slides > 0:
            progress_bar.progress(int(10 + ((idx + 1) / total_slides) * 80))

    out_stream = io.BytesIO()
    prs.save(out_stream)
    return out_stream.getvalue()

# --- Felület (UI) ---
st.markdown('<div class="main-title">🌐 Office Document Translator Pro</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Gemini AI által vezérelt, formázásmegőrző Office fájlfordító</div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "1. Húzd ide vagy válaszd ki a fájlt",
    type=["docx", "xlsx", "pptx"],
    help="Word (.docx), Excel (.xlsx) és PowerPoint (.pptx) fájlokat tölthetsz fel."
)

col1, col2 = st.columns([2, 1])
with col1:
    target_lang_name = st.selectbox("2. Válassz célnyelvet:", list(LANGUAGES.keys()), index=0)
with col2:
    st.write("")
    st.write("")
    translate_button = st.button("🚀 Fordítás indítása", use_container_width=True, type="primary")

if translate_button:
    if uploaded_file is None:
        st.warning("Kérlek, válassz ki egy fájlt a fordítás megkezdéséhez!")
    else:
        target_lang = LANGUAGES[target_lang_name]
        ext = os.path.splitext(uploaded_file.name)[1].lower()
        file_bytes = uploaded_file.read()

        status_text = st.empty()
        progress_bar = st.progress(5)
        status_text.info(f"Gemini AI fordítás folyamatban ({target_lang_name})...")

        try:
            cache = {}
            translated_bytes = None

            if ext == ".docx":
                translated_bytes = process_docx(file_bytes, target_lang, cache, progress_bar)
                mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            elif ext == ".xlsx":
                translated_bytes = process_xlsx(file_bytes, target_lang, cache, progress_bar)
                mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            elif ext == ".pptx":
                translated_bytes = process_pptx(file_bytes, target_lang, cache, progress_bar)
                mime_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

            progress_bar.progress(100)
            status_text.success("✅ A dokumentum sikeresen lefordítva!")

            base_name, _ = os.path.splitext(uploaded_file.name)
            output_filename = f"{base_name}_forditott_{target_lang[:2].lower()}{ext}"

            st.download_button(
                label=f"📥 Lefordított fájl letöltése ({output_filename})",
                data=translated_bytes,
                file_name=output_filename,
                mime=mime_type,
                type="secondary",
                use_container_width=True
            )

        except Exception as e:
            progress_bar.empty()
            status_text.error(f"Hiba történt a fordítás során: {e}")

# Lábléc
st.markdown("""
    <div class="footer-bar">
        <span style="color: #64748b;">Office Document Translator Pro</span>
        <span class="powered-by">⚡ Powered by Nagy Attila</span>
    </div>
""", unsafe_allow_html=True)
