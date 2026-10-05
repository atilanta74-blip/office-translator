import io
import os
import zipfile
import xml.etree.ElementTree as ET
import streamlit as st
from deep_translator import GoogleTranslator
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

st.markdown('<div class="main-title">🌐 Office Document Translator Pro <span style="font-size: 1rem; color: #10b981;">v3.5 Diagnosztika</span></div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Intelligens, formázásmegőrző Office fájlfordító</div>', unsafe_allow_html=True)

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

def clean_and_translate(text, translator, cache):
    if not text:
        return text
    t_clean = text.strip()
    if len(t_clean) <= 1 or t_clean.startswith("="):
        return text
    
    # Csak a tisztán numerikus értékeket ugorjuk át
    num_test = t_clean.replace(".", "").replace(",", "").replace("-", "").replace("%", "").replace("/", "").replace(" ", "")
    if num_test.isdigit():
        return text

    prefix = ""
    if t_clean.startswith("'"):
        prefix = "'"
        t_clean = t_clean[1:].strip()
    elif t_clean.startswith("-->"):
        prefix = "--> "
        t_clean = t_clean[3:].strip()

    if t_clean in cache:
        return prefix + cache[t_clean]

    try:
        translated = translator.translate(t_clean)
        if not translated:
            translated = t_clean
        cache[t_clean] = translated
        return prefix + translated
    except Exception:
        return text

# Excel feldolgozás diagnosztikával és univerzális XML csomópont-kezeléssel
def process_xlsx_direct(file_bytes, target_lang_code, progress_bar, status_box):
    translator = GoogleTranslator(source='auto', target=target_lang_code)
    cache = {}
    total_translated = 0

    status_box.info("Excel belső archívum elemzése...")
    progress_bar.progress(15)

    in_zip = zipfile.ZipFile(io.BytesIO(file_bytes), 'r')
    out_zip_buffer = io.BytesIO()
    out_zip = zipfile.ZipFile(out_zip_buffer, 'w', compression=zipfile.ZIP_DEFLATED)

    files_list = in_zip.infolist()
    total_files = len(files_list)
    xml_debug_info = []

    # Címkék, ahol szöveg lehet (sharedStrings, inline string, rajzok, cellák)
    valid_tags = {"t", "text", "v"}

    for idx, item in enumerate(files_list):
        content = in_zip.read(item.filename)

        if item.filename.endswith(".xml") and not item.filename.endswith("styles.xml"):
            try:
                tree = ET.fromstring(content)
                modified = False
                file_found = 0

                for elem in tree.iter():
                    tag_name = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
                    if tag_name in valid_tags:
                        if elem.text and elem.text.strip():
                            orig_text = elem.text
                            new_text = clean_and_translate(orig_text, translator, cache)
                            if new_text != orig_text:
                                elem.text = new_text
                                total_translated += 1
                                file_found += 1
                                modified = True

                if file_found > 0:
                    xml_debug_info.append(f"{os.path.basename(item.filename)}: {file_found} szöveg")

                if modified:
                    content = ET.tostring(tree, encoding='utf-8', xml_declaration=True)
            except Exception as e:
                pass

        out_zip.writestr(item, content)
        if total_files > 0:
            progress_bar.progress(int(15 + ((idx + 1) / total_files) * 80))

    in_zip.close()
    out_zip.close()
    progress_bar.progress(100)

    debug_str = ", ".join(xml_debug_info) if xml_debug_info else "Nem talált fordítható XML szövegelemet"
    return out_zip_buffer.getvalue(), total_translated, debug_str

# Word feldolgozás
def process_docx(file_bytes, target_lang_code, progress_bar, status_box):
    doc = Document(io.BytesIO(file_bytes))
    translator = GoogleTranslator(source='auto', target=target_lang_code)
    cache = {}
    total_translated = 0

    for p in doc.paragraphs:
        for run in p.runs:
            if run.text.strip():
                new_t = clean_and_translate(run.text, translator, cache)
                if new_t != run.text:
                    run.text = new_t
                    total_translated += 1

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        if run.text.strip():
                            new_t = clean_and_translate(run.text, translator, cache)
                            if new_t != run.text:
                                run.text = new_t
                                total_translated += 1

    out_stream = io.BytesIO()
    doc.save(out_stream)
    progress_bar.progress(100)
    return out_stream.getvalue(), total_translated, "Word bekezdések és táblázatok"

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
                            new_t = clean_and_translate(run.text, translator, cache)
                            if new_t != run.text:
                                run.text = new_t
                                total_translated += 1
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for p in cell.text_frame.paragraphs:
                            for run in p.runs:
                                if run.text.strip():
                                    new_t = clean_and_translate(run.text, translator, cache)
                                    if new_t != run.text:
                                        run.text = new_t
                                        total_translated += 1

    out_stream = io.BytesIO()
    prs.save(out_stream)
    progress_bar.progress(100)
    return out_stream.getvalue(), total_translated, "PowerPoint diák"

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
        lang_code = LANGUAGES[target_lang_name]
        ext = os.path.splitext(uploaded_file.name)[1].lower()
        file_bytes = uploaded_file.read()

        status_box = st.empty()
        progress_bar = st.progress(5)
        status_box.info(f"Feldolgozás és fordítás folyamatban ({target_lang_name})...")

        try:
            translated_bytes = None
            count = 0
            debug_info = ""

            if ext == ".docx":
                translated_bytes, count, debug_info = process_docx(file_bytes, lang_code, progress_bar, status_box)
                mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            elif ext == ".xlsx":
                translated_bytes, count, debug_info = process_xlsx_direct(file_bytes, lang_code, progress_bar, status_box)
                mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            elif ext == ".pptx":
                translated_bytes, count, debug_info = process_pptx(file_bytes, lang_code, progress_bar, status_box)
                mime_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

            status_box.success(f"✅ Kész! Összesen {count} db szöveges elem sikeresen lefordítva.")
            st.info(f"🔍 Részletek: {debug_info}")

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
