import io
import os
import re
import time
import zipfile
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

st.markdown('<div class="main-title">🌐 Office Document Translator Pro <span style="font-size: 1rem; color: #10b981;">v7.5 FullText</span></div>', unsafe_allow_html=True)
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

gemini_key = st.secrets.get("GEMINI_API_KEY")
if not gemini_key:
    st.error("⚠️ Hiányzik a GEMINI_API_KEY a Secrets beállításokból!")
    st.stop()

genai.configure(api_key=gemini_key)
MODEL_NAME = "gemini-3.8-flash"
model = genai.GenerativeModel(model_name=MODEL_NAME)

def is_translatable(txt):
    """Minden olyan elemet átengedünk, amiben van legalább egy betű karakter."""
    if not txt:
        return False
    t = txt.strip()
    if len(t) <= 1 or t.startswith("="):
        return False
    # Ha nincs benne egyetlen betű sem (csak szám, pont, vessző, operátorok), kihagyjuk
    if not re.search(r"[a-zA-Z]", t):
        return False
    # Ha csak dátum formátum (pl. 12-Mar-26)
    if re.match(r"^\d{1,2}-[A-Za-z]{3}-\d{2,4}$", t):
        return False
    return True

def translate_batch_fast(texts, target_lang):
    """Gyors kötegelt fordítás belső sortörés-védelemmel."""
    if not texts:
        return {}

    unique_texts = list(set(texts))
    results = {}
    batch_size = 20

    for i in range(0, len(unique_texts), batch_size):
        chunk = unique_texts[i:i + batch_size]
        
        # Belső sortörések lekezelése egy sorba tömörítve
        prepared_chunk = [t.replace("\r\n", " [BR] ").replace("\n", " [BR] ") for t in chunk]
        lines_input = "\n".join([f"[{idx+1}] {t}" for idx, t in enumerate(prepared_chunk)])
        
        prompt = (
            f"You are a professional industrial, TPM, and technical translator.\n"
            f"Translate each numbered line into {target_lang}.\n"
            f"CRITICAL RULES:\n"
            f"1. Preserve line tags exactly like [1], [2], etc.\n"
            f"2. Keep the placeholder '[BR]' intact where it appears (it represents line breaks).\n"
            f"3. Keep technical acronyms unchanged (e.g. OEE, KPI, TIR, IPS, UPS, PDCA, DDS, WPA, BS, TBR, PSR, FI, CBN, TEAT, SWP, LDR).\n"
            f"4. Translate common business and pillar terms (e.g. 'Pillar owners' -> 'Pillér felelősök/tulajdonosok', 'Plant' -> 'Gyár/Üzem', 'Schedule' -> 'Ütemterv', 'Loss' -> 'Veszteség').\n"
            f"5. Return EXACTLY {len(chunk)} numbered lines in identical sequence. No markdown formatting, no comments.\n\n"
            f"{lines_input}"
        )

        success = False
        for attempt in range(2):
            try:
                resp = model.generate_content(prompt, request_options={"timeout": 60})
                raw_lines = resp.text.strip().split("\n")
                parsed_translations = {}
                for line in raw_lines:
                    line = line.strip()
                    if line.startswith("[") and "]" in line:
                        idx_str = line[1:line.find("]")].strip()
                        if idx_str.isdigit():
                            idx_val = int(idx_str) - 1
                            trans_content = line[line.find("]")+1:].strip()
                            if 0 <= idx_val < len(chunk):
                                # Visszaállítjuk a sortöréseket
                                parsed_translations[idx_val] = trans_content.replace("[BR]", "\n")

                if len(parsed_translations) == len(chunk):
                    for idx_val, orig in enumerate(chunk):
                        results[orig] = parsed_translations[idx_val]
                    success = True
                    break
                else:
                    valid_lines = [l for l in raw_lines if l.strip()]
                    if len(valid_lines) == len(chunk):
                        for orig, line in zip(chunk, valid_lines):
                            cleaned = line[line.find("]")+1:].strip() if "]" in line else line.strip()
                            results[orig] = cleaned.replace("[BR]", "\n")
                        success = True
                        break
            except Exception:
                time.sleep(1.0)

        if not success:
            for o in chunk:
                results[o] = o
        
        time.sleep(0.2)

    return results

# Teljeskörű, biztonságos Excel fordítás
def process_xlsx_ultrasafe(file_bytes, target_lang, progress_bar, status_box):
    status_box.info("Excel szövegtár kinyerése...")
    progress_bar.progress(15)

    in_zip = zipfile.ZipFile(io.BytesIO(file_bytes), 'r')
    out_zip_buffer = io.BytesIO()
    out_zip = zipfile.ZipFile(out_zip_buffer, 'w', compression=zipfile.ZIP_DEFLATED)

    tag_pattern = re.compile(r"(<(?:\w+:)?t(?:\s+[^>]*)?>)(.*?)(</(?:\w+:)?t>)", re.DOTALL)

    target_files = []
    texts_to_translate = []

    for item in in_zip.infolist():
        fn = item.filename
        if fn == "xl/sharedStrings.xml" or fn.startswith("xl/worksheets/sheet") or fn.startswith("xl/drawings/drawing"):
            if fn.endswith(".xml"):
                target_files.append(fn)
                content_str = in_zip.read(fn).decode('utf-8', errors='ignore')
                for match in tag_pattern.finditer(content_str):
                    inner_text = match.group(2)
                    if inner_text and is_translatable(inner_text):
                        texts_to_translate.append(inner_text.strip())

    status_box.info(f"Összesen {len(texts_to_translate)} db szöveg megtalálva. Fordítás a Gemini modellel...")
    progress_bar.progress(35)

    t_map = translate_batch_fast(texts_to_translate, target_lang)
    progress_bar.progress(85)

    count = 0
    for item in in_zip.infolist():
        content_bytes = in_zip.read(item.filename)

        if item.filename in target_files:
            content_str = content_bytes.decode('utf-8', errors='ignore')

            def replace_text(match):
                nonlocal count
                prefix = match.group(1)
                text_val = match.group(2)
                suffix = match.group(3)
                stripped = text_val.strip()
                if stripped in t_map and t_map[stripped] != stripped:
                    count += 1
                    leading_space = text_val[:len(text_val) - len(text_val.lstrip())]
                    trailing_space = text_val[len(text_val.rstrip()):]
                    trans_text = (t_map[stripped]
                                  .replace("&", "&amp;")
                                  .replace("<", "&lt;")
                                  .replace(">", "&gt;"))
                    return f"{prefix}{leading_space}{trans_text}{trailing_space}{suffix}"
                return match.group(0)

            new_str = tag_pattern.sub(replace_text, content_str)
            content_bytes = new_str.encode('utf-8')

        out_zip.writestr(item, content_bytes)

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
                        if is_translatable(run.text):
                            all_runs.append(run)

    status_box.info(f"Word szövegek fordítása ({len(all_runs)} elem)...")
    progress_bar.progress(35)

    texts = [r.text.strip() for r in all_runs]
    t_map = translate_batch_fast(texts, target_lang)
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
                        if run.text.strip():
                            all_runs.append(run)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for p in cell.text_frame.paragraphs:
                            for run in p.runs:
                                if run.text.strip():
                                    all_runs.append(run)

    status_box.info(f"PowerPoint diák fordítása ({len(all_runs)} elem)...")
    progress_bar.progress(35)

    texts = [r.text.strip() for r in all_runs]
    t_map = translate_batch_fast(texts, target_lang)
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

# UI
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
                translated_bytes, count = process_xlsx_ultrasafe(file_bytes, target_lang, progress_bar, status_box)
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
