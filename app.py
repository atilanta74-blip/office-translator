import io
import os
import streamlit as st
import google.generativeai as genai
import openpyxl
from docx import Document
from pptx import Presentation

# Oldal konfiguráció
st.set_page_config(
    page_title="Universal Office Translator Pro",
    page_icon="🌐",
    layout="centered"
)

# Egyedi felület stílus
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

# API kulcs ellenőrzése
gemini_key = st.secrets.get("GEMINI_API_KEY")
if not gemini_key:
    st.error("⚠️ Hiányzik a GEMINI_API_KEY a Secrets beállításokból!")
    st.stop()

genai.configure(api_key=gemini_key)

# Automatikus kompatibilis modell kiválasztása
available_models = ["gemini-1.5-flash-latest", "gemini-1.5-flash", "gemini-pro"]
model = None

for m_name in available_models:
    try:
        test_model = genai.GenerativeModel(model_name=m_name)
        # Gyors teszt kérés
        test_model.generate_content("ping")
        model = test_model
        break
    except Exception:
        continue

if model is None:
    # Végső tartalékként a garantált alaptípus
    model = genai.GenerativeModel(model_name="gemini-pro")

SEP = " ||| "

def batch_translate(texts_to_translate, target_lang):
    """Kötegelt fordítás szeparátoros technikával."""
    if not texts_to_translate:
        return {}

    unique_texts = list(set([t for t in texts_to_translate if t and len(t.strip()) > 1]))
    results = {}
    batch_size = 20

    for i in range(0, len(unique_texts), batch_size):
        chunk = unique_texts[i:i + batch_size]
        combined_text = SEP.join(chunk)

        prompt = (
            f"You are a professional industrial, technical, TPM, and business translator.\n"
            f"Translate each of the following text segments into {target_lang}.\n"
            f"RULES:\n"
            f"1. The input segments are separated by '{SEP}'.\n"
            f"2. You MUST return the translations separated by EXACTLY the same separator: '{SEP}'.\n"
            f"3. Return EXACTLY {len(chunk)} translated segments in the same order.\n"
            f"4. Keep standard manufacturing acronyms unchanged (e.g., OEE, KPI, TIR, IPS, UPS, PDCA, DDS, WPA, BS, TBR, PSR, FI, CBN).\n"
            f"5. Return ONLY the translated segments with the separators. No explanations or extra commentary.\n\n"
            f"Texts to translate:\n{combined_text}"
        )

        try:
            response = model.generate_content(prompt)
            translated_parts = response.text.strip().split(SEP)

            if len(translated_parts) == len(chunk):
                for orig, trans in zip(chunk, translated_parts):
                    results[orig] = trans.strip()
            else:
                for item in chunk:
                    try:
                        p_single = f"Translate to {target_lang}. Return ONLY translated text:\n{item}"
                        r_single = model.generate_content(p_single)
                        results[item] = r_single.text.strip()
                    except Exception:
                        results[item] = item
        except Exception:
            for item in chunk:
                try:
                    p_single = f"Translate to {target_lang}. Return ONLY translated text:\n{item}"
                    r_single = model.generate_content(p_single)
                    results[item] = r_single.text.strip()
                except Exception:
                    results[item] = item

    return results

# Word feldolgozás
def process_docx(file_bytes, target_lang, progress_bar):
    doc = Document(io.BytesIO(file_bytes))
    all_runs = []

    for p in doc.paragraphs:
        for run in p.runs:
            txt = run.text.strip()
            if txt and not txt.isdigit() and len(txt) > 1:
                all_runs.append((run, run.text))

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        txt = run.text.strip()
                        if txt and not txt.isdigit() and len(txt) > 1:
                            all_runs.append((run, run.text))

    progress_bar.progress(30)
    texts_to_send = [txt for _, txt in all_runs]
    trans_map = batch_translate(texts_to_send, target_lang)
    progress_bar.progress(85)

    for run_obj, orig_txt in all_runs:
        if orig_txt in trans_map:
            run_obj.text = trans_map[orig_txt]

    out_stream = io.BytesIO()
    doc.save(out_stream)
    progress_bar.progress(100)
    return out_stream.getvalue()

# Excel feldolgozás
def process_xlsx(file_bytes, target_lang, progress_bar):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    target_cells = []

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    val_str = str(cell.value).strip()
                    if not val_str.startswith("=") and not val_str.replace(".", "").replace(",", "").isdigit() and len(val_str) > 1:
                        target_cells.append((cell, val_str))

    progress_bar.progress(30)
    texts_to_send = [txt for _, txt in target_cells]
    trans_map = batch_translate(texts_to_send, target_lang)
    progress_bar.progress(85)

    for cell_obj, orig_txt in target_cells:
        if orig_txt in trans_map:
            cell_obj.value = trans_map[orig_txt]

    out_stream = io.BytesIO()
    wb.save(out_stream)
    progress_bar.progress(100)
    return out_stream.getvalue()

# PowerPoint feldolgozás
def process_pptx(file_bytes, target_lang, progress_bar):
    prs = Presentation(io.BytesIO(file_bytes))
    all_runs = []

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    for run in p.runs:
                        txt = run.text.strip()
                        if txt and not txt.isdigit() and len(txt) > 1:
                            all_runs.append((run, run.text))
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for p in cell.text_frame.paragraphs:
                            for run in p.runs:
                                txt = run.text.strip()
                                if txt and not txt.isdigit() and len(txt) > 1:
                                    all_runs.append((run, run.text))

    progress_bar.progress(30)
    texts_to_send = [txt for _, txt in all_runs]
    trans_map = batch_translate(texts_to_send, target_lang)
    progress_bar.progress(85)

    for run_obj, orig_txt in all_runs:
        if orig_txt in trans_map:
            run_obj.text = trans_map[orig_txt]

    out_stream = io.BytesIO()
    prs.save(out_stream)
    progress_bar.progress(100)
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
            translated_bytes = None
            if ext == ".docx":
                translated_bytes = process_docx(file_bytes, target_lang, progress_bar)
                mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            elif ext == ".xlsx":
                translated_bytes = process_xlsx(file_bytes, target_lang, progress_bar)
                mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            elif ext == ".pptx":
                translated_bytes = process_pptx(file_bytes, target_lang, progress_bar)
                mime_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

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
            status_text.error(f"Hiba történt a feldolgozás során: {e}")

st.markdown("""
    <div class="footer-bar">
        <span style="color: #64748b;">Office Document Translator Pro</span>
        <span class="powered-by">⚡ Powered by Nagy Attila</span>
    </div>
""", unsafe_allow_html=True)
