import io
import os
import streamlit as st
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

# Támogatott nyelvek
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
    "Japán (Japanese)": "ja",
    "Török (Turkish)": "tr",
    "Vietnámi (Vietnamese)": "vi"
}

def translate_text(text, translator, cache):
    """Biztonságos szövegfordító gyorsítótárral."""
    if text is None:
        return text
    
    text_str = str(text).strip()
    # Képletek, üres mezők, tiszta számok és 1 karakteres elemek kihagyása
    if not text_str or text_str.startswith("=") or len(text_str) <= 1:
        return text
    if text_str.replace(".", "").replace(",", "").replace("-", "").replace("%", "").replace("/", "").isdigit():
        return text

    # Ha a cella aposztróffal kezdődik
    has_apostrophe = text_str.startswith("'")
    if has_apostrophe:
        text_str = text_str[1:].strip()

    if text_str in cache:
        res = cache[text_str]
        return f"'{res}" if has_apostrophe else res

    try:
        translated = translator.translate(text_str)
        if not translated:
            translated = text_str
        cache[text_str] = translated
        return f"'{translated}" if has_apostrophe else translated
    except Exception:
        return text

# Excel feldolgozás minden cellára kiterjesztve
def process_xlsx(file_bytes, target_lang_code, progress_bar, status_box):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    translator = GoogleTranslator(source='auto', target=target_lang_code)
    cache = {}
    total_translated = 0
    total_sheets = len(wb.worksheets)

    for s_idx, ws in enumerate(wb.worksheets):
        status_box.info(f"Munkalap feldolgozása: {ws.title} ({s_idx + 1}/{total_sheets})...")

        # Közvetlen cellabejárás határok nélkül
        for row in ws.rows:
            for cell in row:
                if cell.value is not None:
                    orig_val = cell.value
                    
                    # Ha RichText típusú objektum, kinyerjük a szövegét
                    if hasattr(orig_val, 'text'):
                        orig_text = str(orig_val.text)
                    else:
                        orig_text = str(orig_val)

                    new_text = translate_text(orig_text, translator, cache)
                    
                    if new_text != orig_text:
                        cell.value = new_text
                        total_translated += 1

        progress_bar.progress(int(10 + ((s_idx + 1) / total_sheets) * 85))

    out_stream = io.BytesIO()
    wb.save(out_stream)
    progress_bar.progress(100)
    return out_stream.getvalue(), total_translated

# Word feldolgozás
def process_docx(file_bytes, target_lang_code, progress_bar, status_box):
    doc = Document(io.BytesIO(file_bytes))
    translator = GoogleTranslator(source='auto', target=target_lang_code)
    cache = {}
    total_translated = 0

    for p in doc.paragraphs:
        for run in p.runs:
            if run.text.strip():
                new_t = translate_text(run.text, translator, cache)
                if new_t != run.text:
                    run.text = new_t
                    total_translated += 1

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        if run.text.strip():
                            new_t = translate_text(run.text, translator, cache)
                            if new_t != run.text:
                                run.text = new_t
                                total_translated += 1

    out_stream = io.BytesIO()
    doc.save(out_stream)
    progress_bar.progress(100)
    return out_stream.getvalue(), total_translated

# PowerPoint feldolgozás
def process_pptx(file_bytes, target_lang_code, progress_bar, status_box):
    prs = Presentation(io.BytesIO(file_bytes))
    translator = GoogleTranslator(source='auto', target=target_lang_code)
    cache = {}
    total_translated = 0

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    for run in p.runs:
                        if run.text.strip():
                            new_t = translate_text(run.text, translator, cache)
                            if new_t != run.text:
                                run.text = new_t
                                total_translated += 1
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for p in cell.text_frame.paragraphs:
                            for run in p.runs:
                                if run.text.strip():
                                    new_t = translate_text(run.text, translator, cache)
                                    if new_t != run.text:
                                        run.text = new_t
                                        total_translated += 1

    out_stream = io.BytesIO()
    prs.save(out_stream)
    progress_bar.progress(100)
    return out_stream.getvalue(), total_translated

# --- Felhasználói felület ---
st.markdown('<div class="main-title">🌐 Office Document Translator Pro</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Intelligens, formázásmegőrző Office fájlfordító</div>', unsafe_allow_html=True)

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
        lang_code = LANGUAGES[target_lang_name]
        ext = os.path.splitext(uploaded_file.name)[1].lower()
        file_bytes = uploaded_file.read()

        status_box = st.empty()
        progress_bar = st.progress(5)
        status_box.info(f"Feldolgozás és fordítás folyamatban ({target_lang_name})...")

        try:
            translated_bytes = None
            count = 0
            if ext == ".docx":
                translated_bytes, count = process_docx(file_bytes, lang_code, progress_bar, status_box)
                mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            elif ext == ".xlsx":
                translated_bytes, count = process_xlsx(file_bytes, lang_code, progress_bar, status_box)
                mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            elif ext == ".pptx":
                translated_bytes, count = process_pptx(file_bytes, lang_code, progress_bar, status_box)
                mime_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

            status_box.success(f"✅ Kész! Összesen {count} db szöveges elem sikeresen lefordítva.")
            base_name, _ = os.path.splitext(uploaded_file.name)
            output_filename = f"{base_name}_forditott_{lang_code}{ext}"

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
            status_box.error(f"Hiba történt a feldolgozás során: {e}")

st.markdown("""
    <div class="footer-bar">
        <span style="color: #64748b;">Office Document Translator Pro</span>
        <span class="powered-by">⚡ Powered by Nagy Attila</span>
    </div>
""", unsafe_allow_html=True)
