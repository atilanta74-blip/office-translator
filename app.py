import io
import os
import streamlit as st
import google.generativeai as genai
from deep_translator import GoogleTranslator
import openpyxl
from docx import Document
from pptx import Presentation

# Oldal konfiguráció
st.set_page_config(
    page_title="Universal Office Translator Pro",
    page_icon="🌐",
    layout="centered"
)

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

# Támogatott nyelvek és kódjaik
LANG_MAP = {
    "Magyar (Hungarian)": {"name": "Hungarian", "code": "hu"},
    "Angol (English)": {"name": "English", "code": "en"},
    "Német (German)": {"name": "German", "code": "de"},
    "Olasz (Italian)": {"name": "Italian", "code": "it"},
    "Francia (French)": {"name": "French", "code": "fr"},
    "Spanyol (Spanish)": {"name": "Spanish", "code": "es"},
    "Lengyel (Polish)": {"name": "Polish", "code": "pl"},
    "Cseh (Czech)": {"name": "Czech", "code": "cs"},
    "Szlovák (Slovak)": {"name": "Slovak", "code": "sk"},
    "Román (Romanian)": {"name": "Romanian", "code": "ro"},
    "Japán (Japanese)": {"name": "Japanese", "code": "ja"},
    "Török (Turkish)": {"name": "Turkish", "code": "tr"},
    "Vietnámi (Vietnamese)": {"name": "Vietnamese", "code": "vi"}
}

# Gemini API beállítása
gemini_key = st.secrets.get("GEMINI_API_KEY")
gemini_model = None

if gemini_key:
    try:
        genai.configure(api_key=gemini_key)
        # Megkeressük a fiókhoz elérhető működő Gemini modellt
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                if 'gemini' in m.name:
                    gemini_model = genai.GenerativeModel(model_name=m.name)
                    break
    except Exception:
        gemini_model = None

# Megbízható fordító függvény kettős védelemmel
def translate_single_text(text, target_lang_name, target_lang_code, cache):
    text_clean = str(text).strip()
    if not text_clean or text_clean.isdigit() or len(text_clean) <= 1:
        return text
    
    # Ha az Excel elejére aposztrófot tett
    if text_clean.startswith("'"):
        text_clean = text_clean[1:].strip()

    if text_clean in cache:
        return cache[text_clean]

    # 1. Próbálkozás: Gemini AI
    if gemini_model:
        try:
            prompt = (
                f"Translate this industrial/business text accurately into {target_lang_name}. "
                f"Keep acronyms (OEE, KPI, TIR, IPS, UPS, PDCA, BS, TBR, PSR, FI) intact. "
                f"Return ONLY the translated text:\n{text_clean}"
            )
            res = gemini_model.generate_content(prompt)
            if res.text and len(res.text.strip()) > 0:
                translated = res.text.strip()
                cache[text_clean] = translated
                return translated
        except Exception:
            pass

    # 2. Garantált tartalék: Deep Translator (Google Translate)
    try:
        fallback_trans = GoogleTranslator(source='auto', target=target_lang_code).translate(text_clean)
        cache[text_clean] = fallback_trans
        return fallback_trans
    except Exception:
        return text

# Word (.docx) feldolgozás
def process_docx(file_bytes, target_lang_name, target_lang_code, progress_bar):
    doc = Document(io.BytesIO(file_bytes))
    cache = {}
    
    total = len(doc.paragraphs)
    for idx, p in enumerate(doc.paragraphs):
        for run in p.runs:
            if run.text.strip():
                run.text = translate_single_text(run.text, target_lang_name, target_lang_code, cache)
        if total > 0 and idx % 5 == 0:
            progress_bar.progress(int(10 + (idx / total) * 40))

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        if run.text.strip():
                            run.text = translate_single_text(run.text, target_lang_name, target_lang_code, cache)

    progress_bar.progress(95)
    out_stream = io.BytesIO()
    doc.save(out_stream)
    progress_bar.progress(100)
    return out_stream.getvalue()

# Excel (.xlsx) feldolgozás (összevont cellák, stringek, számok kezelése)
def process_xlsx(file_bytes, target_lang_name, target_lang_code, progress_bar):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    cache = {}
    sheets = wb.worksheets
    total_s = len(sheets)

    for s_idx, ws in enumerate(sheets):
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    val_str = str(cell.value).strip()
                    # Képletek kihagyása
                    if val_str.startswith("="):
                        continue
                    # Számok kihagyása
                    if val_str.replace(".", "").replace(",", "").replace("-", "").isdigit():
                        continue
                    
                    if len(val_str) > 1:
                        cell.value = translate_single_text(val_str, target_lang_name, target_lang_code, cache)

        progress_bar.progress(int(10 + ((s_idx + 1) / total_s) * 85))

    out_stream = io.BytesIO()
    wb.save(out_stream)
    progress_bar.progress(100)
    return out_stream.getvalue()

# PowerPoint (.pptx) feldolgozás
def process_pptx(file_bytes, target_lang_name, target_lang_code, progress_bar):
    prs = Presentation(io.BytesIO(file_bytes))
    cache = {}
    total_slides = len(prs.slides)

    for idx, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    for run in p.runs:
                        if run.text.strip():
                            run.text = translate_single_text(run.text, target_lang_name, target_lang_code, cache)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for p in cell.text_frame.paragraphs:
                            for run in p.runs:
                                if run.text.strip():
                                    run.text = translate_single_text(run.text, target_lang_name, target_lang_code, cache)
        if total_slides > 0:
            progress_bar.progress(int(10 + ((idx + 1) / total_slides) * 85))

    out_stream = io.BytesIO()
    prs.save(out_stream)
    progress_bar.progress(100)
    return out_stream.getvalue()

# --- Felület (UI) ---
st.markdown('<div class="main-title">🌐 Office Document Translator Pro</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Intelligens, formázásmegőrző Office fájlfordító</div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "1. Húzd ide vagy válaszd ki a fájlt",
    type=["docx", "xlsx", "pptx"],
    help="Word (.docx), Excel (.xlsx) és PowerPoint (.pptx) fájlokat tölthetsz fel."
)

col1, col2 = st.columns([2, 1])
with col1:
    target_lang_name = st.selectbox("2. Válassz célnyelvet:", list(LANG_MAP.keys()), index=0)
with col2:
    st.write("")
    st.write("")
    translate_button = st.button("🚀 Fordítás indítása", use_container_width=True, type="primary")

if translate_button:
    if uploaded_file is None:
        st.warning("Kérlek, válassz ki egy fájlt a fordítás megkezdéséhez!")
    else:
        lang_info = LANG_MAP[target_lang_name]
        ext = os.path.splitext(uploaded_file.name)[1].lower()
        file_bytes = uploaded_file.read()

        status_text = st.empty()
        progress_bar = st.progress(5)
        status_text.info(f"Fordítás folyamatban ({target_lang_name})...")

        try:
            translated_bytes = None
            if ext == ".docx":
                translated_bytes = process_docx(file_bytes, lang_info["name"], lang_info["code"], progress_bar)
                mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            elif ext == ".xlsx":
                translated_bytes = process_xlsx(file_bytes, lang_info["name"], lang_info["code"], progress_bar)
                mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            elif ext == ".pptx":
                translated_bytes = process_pptx(file_bytes, lang_info["name"], lang_info["code"], progress_bar)
                mime_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

            status_text.success("✅ A dokumentum sikeresen lefordítva!")
            base_name, _ = os.path.splitext(uploaded_file.name)
            output_filename = f"{base_name}_forditott_{lang_info['code']}{ext}"

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
            status_text.error(f"Hiba történt a feldolgozás során: {e}")

st.markdown("""
    <div class="footer-bar">
        <span style="color: #64748b;">Office Document Translator Pro</span>
        <span class="powered-by">⚡ Powered by Nagy Attila</span>
    </div>
""", unsafe_allow_html=True)
