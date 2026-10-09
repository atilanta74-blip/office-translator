import io
import os
import re
import json
import time
import zipfile
import streamlit as st
from google import genai
from docx import Document
from pptx import Presentation

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

st.markdown('<div class="main-title">🌐 Office Document Translator Pro <span style="font-size: 1rem; color: #10b981;">v25.0 AutoSelect</span></div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Intelligens, formázásmegőrző Office fájlfordító (Word, Excel, PowerPoint)</div>', unsafe_allow_html=True)

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

gemini_key = gemini_key.strip()
client = genai.Client(api_key=gemini_key)

def has_letters(text):
    if not text:
        return False
    t = str(text).strip()
    if len(t) <= 1 or t.startswith("="):
        return False
    return bool(re.search(r"[a-zA-ZáéíóöőúüűÁÉÍÓÖŐÚÜŰ]", t))

def get_usable_models():
    """Közvetlenül lekérdezi a Google-től azokat a modelleket, amelyek szöveget tudnak generálni és nem audió/tts modellek."""
    usable = []
    try:
        for m in client.models.list():
            name = m.name.replace("models/", "")
            # Kizárjuk a hang, tts, embedding és vision-only modelleket
            if any(bad in name.lower() for bad in ["tts", "embedding", "audio", "imagen"]):
                continue
            if "gemini" in name.lower():
                usable.append(name)
    except Exception:
        pass
    
    # Ha nem sikerült lekérni, az érvényes alapértelmezett lista
    if not usable:
        usable = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-3-flash-preview"]
    return usable

def translate_batch(unique_texts, target_lang, status_box):
    if not unique_texts:
        return {}

    input_data = {str(idx + 1): txt for idx, txt in enumerate(unique_texts)}
    prompt = (
        f"You are a professional industrial, technical, TPM, and business translator.\n"
        f"Translate the values of the following JSON object into {target_lang}.\n"
        f"Strict Rules:\n"
        f"1. Keep technical acronyms intact (OEE, KPI, TIR, IPS, UPS, PDCA, DDS, WPA, BS, TBR, PSR, FI, CBN, TEAT, SWP, LDR, etc.).\n"
        f"2. Translate common manufacturing/TPM terms accurately (e.g. 'Pillar owners' -> 'Pillér felelősök', 'Plant' -> 'Üzem/Gyár', 'Schedule' -> 'Ütemterv', 'Loss' -> 'Veszteség').\n"
        f"3. Return ONLY a valid JSON object matching the exact keys ('1', '2', etc.) and translated string values. Do not wrap in markdown.\n\n"
        f"{json.dumps(input_data, ensure_ascii=False)}"
    )

    results = {}
    models_to_try = get_usable_models()
    last_error = ""

    for model_name in models_to_try:
        status_box.info(f"AI fordítás próbálkozás ({model_name})...")
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=prompt
            )
            raw_text = response.text.strip()
            if raw_text.startswith("```json"):
                raw_text = raw_text[7:]
            if raw_text.startswith("```"):
                raw_text = raw_text[3:]
            if raw_text.endswith("```"):
                raw_text = raw_text[:-3]

            parsed_json = json.loads(raw_text.strip())
            for idx, orig in enumerate(unique_texts):
                k = str(idx + 1)
                if k in parsed_json and parsed_json[k]:
                    results[orig] = str(parsed_json[k]).strip()

            st.success(f"✅ Sikeres fordítás a(z) **{model_name}** modellel!")
            return results
        except Exception as e:
            err_msg = str(e)
            last_error = f"{model_name}: {err_msg}"
            # Ha kvóta vagy 404 hiba van ennél a modellnél, lépünk a következőre a listában
            continue

    st.error(f"❌ Részletes Google API hiba: {last_error}")
    for orig in unique_texts:
        results[orig] = orig

    return results

def process_xlsx(file_bytes, target_lang, progress_bar, status_box):
    status_box.info("Excel XML réteg beolvasása...")
    progress_bar.progress(15)

    in_zip = zipfile.ZipFile(io.BytesIO(file_bytes), 'r')
    out_zip_buffer = io.BytesIO()
    out_zip = zipfile.ZipFile(out_zip_buffer, 'w', compression=zipfile.ZIP_DEFLATED)

    tag_pattern = re.compile(r"(<(?:\w+:)?t(?:\s+[^>]*)?>)(.*?)(</(?:\w+:)?t>)", re.DOTALL)

    target_files = []
    raw_texts = []

    for item in in_zip.infolist():
        fn = item.filename
        if fn == "xl/sharedStrings.xml" or fn.startswith("xl/worksheets/sheet") or fn.startswith("xl/drawings/drawing"):
            if fn.endswith(".xml"):
                target_files.append(fn)
                content_str = in_zip.read(fn).decode('utf-8', errors='ignore')
                for m in tag_pattern.finditer(content_str):
                    val = m.group(2)
                    if val and has_letters(val):
                        raw_texts.append(val.strip())

    unique_texts = list(set(raw_texts))
    status_box.info(f"{len(raw_texts)} db mező ({len(unique_texts)} egyedi szöveg) átadása a fordítónak...")
    progress_bar.progress(35)

    t_map = translate_batch(unique_texts, target_lang, status_box)
    progress_bar.progress(85)

    replaced_count = 0
    for item in in_zip.infolist():
        content_bytes = in_zip.read(item.filename)

        if item.filename in target_files:
            content_str = content_bytes.decode('utf-8', errors='ignore')

            def replace_text(match):
                nonlocal replaced_count
                prefix = match.group(1)
                text_val = match.group(2)
                suffix = match.group(3)

                stripped = text_val.strip()
                if stripped in t_map:
                    trans = t_map[stripped]
                    if trans and trans != stripped:
                        replaced_count += 1
                        leading = text_val[:len(text_val) - len(text_val.lstrip())]
                        trailing = text_val[len(text_val.rstrip()):]
                        safe_trans = (trans
                                      .replace("&", "&amp;")
                                      .replace("<", "&lt;")
                                      .replace(">", "&gt;"))
                        return f"{prefix}{leading}{safe_trans}{trailing}{suffix}"
                return match.group(0)

            new_str = tag_pattern.sub(replace_text, content_str)
            content_bytes = new_str.encode('utf-8')

        out_zip.writestr(item, content_bytes)

    in_zip.close()
    out_zip.close()
    progress_bar.progress(100)

    return out_zip_buffer.getvalue(), replaced_count, len(raw_texts)

def process_docx(file_bytes, target_lang, progress_bar, status_box):
    doc = Document(io.BytesIO(file_bytes))
    paragraphs_to_translate = []

    for p in doc.paragraphs:
        if has_letters(p.text):
            paragraphs_to_translate.append(p)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if has_letters(p.text):
                        paragraphs_to_translate.append(p)

    status_box.info(f"Word bekezdések kinyerve ({len(paragraphs_to_translate)} db)...")
    progress_bar.progress(30)

    unique_texts = list(set([p.text.strip() for p in paragraphs_to_translate]))
    t_map = translate_batch(unique_texts, target_lang, status_box)
    progress_bar.progress(85)

    count = 0
    for p in paragraphs_to_translate:
        clean_t = p.text.strip()
        if clean_t in t_map and t_map[clean_t] != clean_t:
            p.text = t_map[clean_t]
            count += 1

    out_stream = io.BytesIO()
    doc.save(out_stream)
    progress_bar.progress(100)
    return out_stream.getvalue(), count, len(paragraphs_to_translate)

def process_pptx(file_bytes, target_lang, progress_bar, status_box):
    prs = Presentation(io.BytesIO(file_bytes))
    shapes_to_translate = []

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    if has_letters(p.text):
                        shapes_to_translate.append(p)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for p in cell.text_frame.paragraphs:
                            if has_letters(p.text):
                                shapes_to_translate.append(p)

    status_box.info(f"PowerPoint diák kinyerve ({len(shapes_to_translate)} db)...")
    progress_bar.progress(30)

    unique_texts = list(set([p.text.strip() for p in shapes_to_translate]))
    t_map = translate_batch(unique_texts, target_lang, status_box)
    progress_bar.progress(85)

    count = 0
    for p in shapes_to_translate:
        clean_t = p.text.strip()
        if clean_t in t_map and t_map[clean_t] != clean_t:
            p.text = t_map[clean_t]
            count += 1

    out_stream = io.BytesIO()
    prs.save(out_stream)
    progress_bar.progress(100)
    return out_stream.getvalue(), count, len(shapes_to_translate)

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
            total_found = 0

            if ext == ".xlsx":
                translated_bytes, count, total_found = process_xlsx(file_bytes, target_lang, progress_bar, status_box)
                mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            elif ext == ".docx":
                translated_bytes, count, total_found = process_docx(file_bytes, target_lang, progress_bar, status_box)
                mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            elif ext == ".pptx":
                translated_bytes, count, total_found = process_pptx(file_bytes, target_lang, progress_bar, status_box)
                mime_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

            status_box.success(f"✅ Kész! Összesen {count} db elem sikeresen lefordítva ({total_found} talált mezőből).")

            base_name, _ = os.path.splitext(uploaded_file.name)
            output_filename = f"{base_name}_forditott{ext}"

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
