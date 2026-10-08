import io
import os
import json
import time
import zipfile
import xml.etree.ElementTree as ET
import streamlit as st
import google.generativeai as genai
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

st.markdown('<div class="main-title">🌐 Office Document Translator Pro <span style="font-size: 1rem; color: #10b981;">v5.6 Balanced</span></div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Intelligens, formázásmegőrző Office fájlfordító</div>', unsafe_allow_html=True)

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

MODEL_NAME = "gemini-3.8-flash"
model = genai.GenerativeModel(
    model_name=MODEL_NAME,
    generation_config={"response_mime_type": "application/json"}
)

def translate_mega_batch(texts, target_lang):
    """Kiegyensúlyozott csomagméret (40 db), hogy elkerüljük az 504 Timeout hibát."""
    if not texts:
        return {}

    unique_texts = list(set(texts))
    results = {}
    
    # 40 szöveg csomagonként: gyors, megbízható és nem fut ki az időből
    batch_size = 40

    for i in range(0, len(unique_texts), batch_size):
        chunk = unique_texts[i:i + batch_size]
        
        prompt = (
            f"You are a professional industrial, TPM, and technical document translator.\n"
            f"Translate each string in the input JSON array into {target_lang}.\n"
            f"Keep abbreviations (e.g. OEE, KPI, TIR, IPS, UPS, PDCA, DDS, WPA, BS, TBR, PSR, FI, CBN) intact.\n"
            f"Return a JSON object where keys are the original strings and values are the translated strings.\n\n"
            f"Input JSON:\n{json.dumps(chunk, ensure_ascii=False)}"
        )

        try:
            # request_options timeout növeléssel
            resp = model.generate_content(
                prompt,
                request_options={"timeout": 120}
            )
            data = json.loads(resp.text.strip())
            if isinstance(data, dict):
                results.update(data)
            elif isinstance(data, list) and len(data) == len(chunk):
                for orig, trans in zip(chunk, data):
                    results[orig] = trans
        except Exception as e:
            st.warning(f"Gemini API figyelmeztetés (csomag {i // batch_size + 1}): {e}")
            for o in chunk:
                results[o] = o
        
        time.sleep(0.3)

    return results

def is_translatable(txt):
    if not txt:
        return False
    t = txt.strip()
    if len(t) <= 1 or t.startswith("="):
        return False
    num_test = t.replace(".", "").replace(",", "").replace("-", "").replace("%", "").replace("/", "").replace(" ", "")
    if num_test.isdigit():
        return False
    return True

# Excel feldolgozás
def process_xlsx_gemini(file_bytes, target_lang, progress_bar, status_box):
    status_box.info("Excel belső szövegtárának átvizsgálása...")
    progress_bar.progress(15)

    in_zip = zipfile.ZipFile(io.BytesIO(file_bytes), 'r')
    out_zip_buffer = io.BytesIO()
    out_zip = zipfile.ZipFile(out_zip_buffer, 'w', compression=zipfile.ZIP_DEFLATED)

    all_elements = []
    parsed_files = {}

    for item in in_zip.infolist():
        content = in_zip.read(item.filename)
        if item.filename.endswith(".xml") and not item.filename.endswith("styles.xml"):
            try:
                tree = ET.fromstring(content)
                parsed_files[item.filename] = tree
                for elem in tree.iter():
                    tag_name = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
                    if tag_name in ["t", "text", "v"]:
                        if elem.text and is_translatable(elem.text):
                            all_elements.append((item.filename, elem, elem.text.strip()))
            except Exception:
                pass

    status_box.info(f"Összesen {len(all_elements)} db szöveg összegyűjtve. Fordítás folyamatban...")
    progress_bar.progress(40)

    unique_to_translate = [txt for _, _, txt in all_elements]
    translation_map = translate_mega_batch(unique_to_translate, target_lang)
    progress_bar.progress(85)

    count = 0
    for filename, elem, orig_txt in all_elements:
        if orig_txt in translation_map and translation_map[orig_txt] != orig_txt:
            elem.text = translation_map[orig_txt]
            count += 1

    for item in in_zip.infolist():
        if item.filename in parsed_files:
            new_bytes = ET.tostring(parsed_files[item.filename], encoding='utf-8', xml_declaration=True)
            out_zip.writestr(item, new_bytes)
        else:
            out_zip.writestr(item, in_zip.read(item.filename))

    in_zip.close()
    out_zip.close()
    progress_bar.progress(100)

    return out_zip_buffer.getvalue(), count

# Word feldolgozás
def process_docx_gemini(file_bytes, target_lang, progress_bar, status_box):
    doc = Document(io.BytesIO(file_bytes))
    all_runs = []

    for p in doc.paragraphs:
        for run in p.runs:
            if is_translatable(run.text):
                all_runs.append(run)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        if run.text.strip():
                            all_runs.append(run)

    status_box.info(f"Word szövegek összegyűjtve ({len(all_runs)} elem). Fordítás...")
    progress_bar.progress(40)

    texts = [r.text.strip() for r in all_runs]
    t_map = translate_mega_batch(texts, target_lang)
    progress_bar.progress(85)

    count = 0
    for r in all_runs:
        clean_t = r.text.strip()
        if clean_t in t_map and t_map[clean_t] != clean_t:
            r.text = t_map[clean_t]
            count += 1

    out_stream = io.BytesIO()
    doc.save(out_stream)
    progress_bar.progress(100)
    return out_stream.getvalue(), count

# PowerPoint feldolgozás
def process_pptx_gemini(file_bytes, target_lang, progress_bar, status_box):
    prs = Presentation(io.BytesIO(file_bytes))
    all_runs = []

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    for run in p.runs:
                        if is_translatable(run.text):
                            all_runs.append(run)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for p in cell.text_frame.paragraphs:
                            for run in p.runs:
                                if is_translatable(run.text):
                                    all_runs.append(run)

    status_box.info(f"PowerPoint diák összegyűjtve ({len(all_runs)} elem). Fordítás...")
    progress_bar.progress(40)

    texts = [r.text.strip() for r in all_runs]
    t_map = translate_mega_batch(texts, target_lang)
    progress_bar.progress(85)

    count = 0
    for r in all_runs:
        clean_t = r.text.strip()
        if clean_t in t_map and t_map[clean_t] != clean_t:
            r.text = t_map[clean_t]
            count += 1

    out_stream = io.BytesIO()
    prs.save(out_stream)
    progress_bar.progress(100)
    return out_stream.getvalue(), count

# Felhasználói felület
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

        status_box = st.empty()
        progress_bar = st.progress(5)

        try:
            translated_bytes = None
            count = 0

            if ext == ".docx":
                translated_bytes, count = process_docx_gemini(file_bytes, target_lang, progress_bar, status_box)
                mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            elif ext == ".xlsx":
                translated_bytes, count = process_xlsx_gemini(file_bytes, target_lang, progress_bar, status_box)
                mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            elif ext == ".pptx":
                translated_bytes, count = process_pptx_gemini(file_bytes, target_lang, progress_bar, status_box)
                mime_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

            status_box.success(f"✅ Kész! Összesen {count} db szöveges elem sikeresen lefordítva.")

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
            status_box.error(f"Hiba történt a feldolgozás során: {e}")

st.markdown("""
    <div class="footer-bar">
        <span style="color: #64748b;">Office Document Translator Pro</span>
        <span class="powered-by">⚡ Powered by Nagy Attila</span>
    </div>
""", unsafe_allow_html=True)
