import io
import os
import streamlit as st
from deep_translator import GoogleTranslator
import openpyxl
from docx import Document
from pptx import Presentation

# Oldal konfiguráció
st.set_page_config(
    page_title="Universal Office Translator",
    page_icon="🌐",
    layout="centered"
)

# Egyedi modern sötét stílus és fejléc/lábléc formázás
st.markdown("""
    <style>
    .main-title {
        font-size: 2.2rem;
        font-weight: 700;
        margin-bottom: 0.2rem;
    }
    .sub-title {
        color: #94a3b8;
        font-size: 1rem;
        margin-bottom: 1.5rem;
    }
    .footer-bar {
        margin-top: 3rem;
        padding-top: 1rem;
        border-top: 1px solid #334155;
        display: flex;
        justify-content: space-between;
        align-items: center;
        font-size: 0.9rem;
    }
    .powered-by {
        color: #38bdf8;
        font-weight: 700;
        font-size: 1rem;
    }
    </style>
""", unsafe_allow_html=True)

# Támogatott nyelvek listája
LANGUAGES = {
    "Magyar (Hungarian)": "hu",
    "Angol (English)": "en",
    "Német (German)": "de",
    "Olasz (Italian)": "it",
    "Francia (French)": "fr",
    "Spanyol (Spanish)": "es",
    "Lengyel (Polish)": "pl",
    "Cseh (Czech)": "cs",
    "Szlovák (Slovak)": "sk",
    "Román (Romanian)": "ro",
    "Kínai (Chinese - Simpl.)": "zh-CN",
    "Japán (Japanese)": "ja",
    "Török (Turkish)": "tr",
    "Vietnámi (Vietnamese)": "vi"
}

# Szöveg fordító segédfüggvény gyorsítótárral
def translate_text(text, translator, cache):
    text_str = str(text).strip()
    if not text_str or text_str.isdigit():
        return text
    if text_str in cache:
        return cache[text_str]
    try:
        translated = translator.translate(text_str)
        cache[text_str] = translated
        return translated
    except Exception:
        return text

# Word (.docx) fordítás memóriában
def process_docx(file_bytes, translator, cache, progress_bar):
    doc = Document(io.BytesIO(file_bytes))
    total_p = len(doc.paragraphs)
    
    for idx, p in enumerate(doc.paragraphs):
        for run in p.runs:
            if run.text.strip():
                run.text = translate_text(run.text, translator, cache)
        if total_p > 0 and idx % 5 == 0:
            progress_bar.progress(int(10 + (idx / total_p) * 40))

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        if run.text.strip():
                            run.text = translate_text(run.text, translator, cache)

    progress_bar.progress(90)
    out_stream = io.BytesIO()
    doc.save(out_stream)
    return out_stream.getvalue()

# Excel (.xlsx) fordítás memóriában
def process_xlsx(file_bytes, translator, cache, progress_bar):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    sheets = wb.worksheets
    total_s = len(sheets)

    for idx, ws in enumerate(sheets):
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type == 's' and cell.value:
                    val = str(cell.value)
                    if not val.startswith("="):
                        cell.value = translate_text(val, translator, cache)
        if total_s > 0:
            progress_bar.progress(int(10 + ((idx + 1) / total_s) * 80))

    out_stream = io.BytesIO()
    wb.save(out_stream)
    return out_stream.getvalue()

# PowerPoint (.pptx) fordítás memóriában
def process_pptx(file_bytes, translator, cache, progress_bar):
    prs = Presentation(io.BytesIO(file_bytes))
    total_slides = len(prs.slides)

    for idx, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    for run in p.runs:
                        if run.text.strip():
                            run.text = translate_text(run.text, translator, cache)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for p in cell.text_frame.paragraphs:
                            for run in p.runs:
                                if run.text.strip():
                                    run.text = translate_text(run.text, translator, cache)
        if total_slides > 0:
            progress_bar.progress(int(10 + ((idx + 1) / total_slides) * 80))

    out_stream = io.BytesIO()
    prs.save(out_stream)
    return out_stream.getvalue()

# --- Felhasználói Felület (UI) ---
st.markdown('<div class="main-title">🌐 Office Document Translator</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Formázásmegőrző Word (.docx), Excel (.xlsx) és PPT (.pptx) fordító</div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "1. Húzd ide vagy válaszd ki a fájlt",
    type=["docx", "xlsx", "pptx"],
    help="Csak az Office formátumokat támogatja a struktúra megtartása érdekében."
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
        st.warning("Kérlek, tölts fel egy fájlt a fordítás megkezdéséhez!")
    else:
        target_code = LANGUAGES[target_lang_name]
        ext = os.path.splitext(uploaded_file.name)[1].lower()
        file_bytes = uploaded_file.read()

        status_text = st.empty()
        progress_bar = st.progress(5)
        status_text.info("Fordítás folyamatban, kérlek várj...")

        try:
            translator = GoogleTranslator(source='auto', target=target_code)
            cache = {}
            translated_bytes = None

            if ext == ".docx":
                translated_bytes = process_docx(file_bytes, translator, cache, progress_bar)
                mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            elif ext == ".xlsx":
                translated_bytes = process_xlsx(file_bytes, translator, cache, progress_bar)
                mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            elif ext == ".pptx":
                translated_bytes = process_pptx(file_bytes, translator, cache, progress_bar)
                mime_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

            progress_bar.progress(100)
            status_text.success("✅ A dokumentum sikeresen lefordítva!")

            base_name, _ = os.path.splitext(uploaded_file.name)
            output_filename = f"{base_name}_forditott_{target_code}{ext}"

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

# Lábléc kiírás
st.markdown("""
    <div class="footer-bar">
        <span style="color: #64748b;">Office Document Translator Pro</span>
        <span class="powered-by">⚡ Powered by Nagy Attila</span>
    </div>
""", unsafe_allow_html=True)