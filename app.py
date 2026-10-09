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

# Oldalbeállítások
st.set_page_config(
    page_title="TranslateOS | Neural Office Translation",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Futurisztikus Cyber / Enterprise CSS stílus
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap');

html, body, [class*="css"] {
    font-family: 'Plus Jakarta Sans', sans-serif;
    color: #f1f5f9;
}

/* Fő háttér gradiens */
.stApp {
    background: radial-gradient(circle at 50% 0%, #172554 0%, #0b0f19 55%, #030712 100%);
    background-attachment: fixed;
}

/* Fejléc stílus */
.hero-container {
    text-align: center;
    padding: 2.5rem 1rem 1.5rem 1rem;
    position: relative;
}

.hero-badge {
    display: inline-flex;
    align-items: center;
    gap: 0.5rem;
    padding: 0.35rem 1rem;
    border-radius: 9999px;
    background: rgba(30, 58, 138, 0.4);
    border: 1px solid rgba(96, 165, 250, 0.4);
    font-size: 0.8rem;
    font-weight: 600;
    color: #93c5fd;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    margin-bottom: 1rem;
    box-shadow: 0 0 20px rgba(59, 130, 246, 0.2);
}

.hero-title {
    font-size: 3rem;
    font-weight: 800;
    background: linear-gradient(135deg, #ffffff 30%, #93c5fd 70%, #38bdf8 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    margin-bottom: 0.5rem;
    letter-spacing: -0.03em;
}

.hero-subtitle {
    font-size: 1.1rem;
    color: #94a3b8;
    max-width: 650px;
    margin: 0 auto 2rem auto;
    line-height: 1.6;
}

/* Glassmorphism Vezérlő Panel */
.glass-panel {
    background: rgba(15, 23, 42, 0.65);
    backdrop-filter: blur(16px);
    -webkit-backdrop-filter: blur(16px);
    border: 1px solid rgba(255, 255, 255, 0.08);
    border-radius: 20px;
    padding: 2rem;
    box-shadow: 0 20px 40px -15px rgba(0, 0, 0, 0.7), 0 0 0 1px rgba(255, 255, 255, 0.05);
    margin-bottom: 2rem;
}

/* Fájl típus jelvények */
.format-grid {
    display: flex;
    justify-content: center;
    gap: 1rem;
    margin-bottom: 2rem;
}

.format-card {
    background: rgba(30, 41, 59, 0.5);
    border: 1px solid rgba(255, 255, 255, 0.06);
    border-radius: 12px;
    padding: 0.6rem 1.2rem;
    font-size: 0.85rem;
    font-weight: 600;
    color: #cbd5e1;
    display: flex;
    align-items: center;
    gap: 0.5rem;
}

.dot-green {
    width: 7px;
    height: 7px;
    background: #10b981;
    border-radius: 50%;
    box-shadow: 0 0 10px #10b981;
}

/* Streamlit gomb futurisztikus felülbírálása */
div.stButton > button {
    width: 100%;
    background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 50%, #0284c7 100%) !important;
    color: #ffffff !important;
    font-weight: 700 !important;
    font-size: 1.05rem !important;
    border: 1px solid rgba(255, 255, 255, 0.2) !important;
    border-radius: 14px !important;
    padding: 0.75rem 1.5rem !important;
    box-shadow: 0 0 25px rgba(37, 99, 235, 0.5) !important;
    transition: all 0.25s ease-in-out !important;
}

div.stButton > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 0 35px rgba(56, 189, 248, 0.7) !important;
    border-color: rgba(255, 255, 255, 0.4) !important;
}

/* Letöltés gomb */
div.stDownloadButton > button {
    background: linear-gradient(135deg, #059669 0%, #047857 100%) !important;
    color: #ffffff !important;
    font-weight: 700 !important;
    border-radius: 14px !important;
    box-shadow: 0 0 25px rgba(16, 185, 129, 0.4) !important;
}

/* Footer sáv */
.footer-hud {
    margin-top: 4rem;
    padding-top: 1.5rem;
    border-top: 1px solid rgba(255, 255, 255, 0.08);
    display: flex;
    justify-content: space-between;
    align-items: center;
    color: #64748b;
    font-size: 0.85rem;
}

.engine-status {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.8rem;
    color: #38bdf8;
    background: rgba(14, 165, 233, 0.1);
    padding: 0.25rem 0.75rem;
    border-radius: 6px;
    border: 1px solid rgba(14, 165, 233, 0.2);
}
</style>
""", unsafe_allow_html=True)

# Hero szekció
st.markdown("""
<div class="hero-container">
    <div class="hero-badge">
        <span class="dot-green"></span> Core Engine v26.0 Active
    </div>
    <div class="hero-title">TranslateOS Neural Pro</div>
    <div class="hero-subtitle">Vállalati szintű intelligens Office fordítórendszer. XML-alapú formázásmegőrzés gépi tanulási kontextuskezeléssel.</div>
    <div class="format-grid">
        <div class="format-card">📑 Microsoft Word (.docx)</div>
        <div class="format-card">📊 Microsoft Excel (.xlsx)</div>
        <div class="format-card">📽️ PowerPoint (.pptx)</div>
    </div>
</div>
""", unsafe_allow_html=True)

LANGUAGES = {
    "🇭🇺 Magyar (Hungarian)": "Hungarian",
    "🇬🇧 Angol (English)": "English",
    "🇩🇪 Német (German)": "German",
    "🇮🇹 Olasz (Italian)": "Italian",
    "🇫🇷 Francia (French)": "French",
    "🇪🇸 Spanyol (Spanish)": "Spanish",
    "🇵🇱 Lengyel (Polish)": "Polish",
    "🇨🇿 Cseh (Czech)": "Czech",
    "🇸🇰 Szlovák (Slovak)": "Slovak",
    "🇷🇴 Román (Romanian)": "Romanian",
    "🇯🇵 Japán (Japanese)": "Japanese",
    "🇹🇷 Török (Turkish)": "Turkish",
    "🇻🇳 Vietnámi (Vietnamese)": "Vietnamese"
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
    usable = []
    try:
        for m in client.models.list():
            name = m.name.replace("models/", "")
            if any(bad in name.lower() for bad in ["tts", "embedding", "audio", "imagen"]):
                continue
            if "gemini" in name.lower():
                usable.append(name)
    except Exception:
        pass
    if not usable:
        usable = ["gemini-flash-lite-latest", "gemini-2.5-flash", "gemini-2.0-flash"]
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
        status_box.info(f"⚡ Neural Pipeline inicializálása: **{model_name}**...")
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

            st.success(f"🚀 Feldolgozás kész: **{model_name}** motorral!")
            return results
        except Exception as e:
            last_error = f"{model_name}: {str(e)}"
            continue

    st.error(f"❌ Rendszerhiba: {last_error}")
    for orig in unique_texts:
        results[orig] = orig

    return results

def process_xlsx(file_bytes, target_lang, progress_bar, status_box):
    status_box.info("Excel XML réteg elemzése és DOM faépítés...")
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
    status_box.info(f"Kinyerve {len(raw_texts)} mező ({len(unique_texts)} egyedi). AI csomag átadása...")
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

# Konténerbe helyezett vezérlőpult
with st.container():
    col_l, col_center, col_r = st.columns([1, 10, 1])
    with col_center:
        uploaded_file = st.file_uploader(
            "📁 Dokumentum feltöltése (.docx, .xlsx, .pptx)",
            type=["docx", "xlsx", "pptx"],
            help="Húzd be a lefordítani kívánt fájlt az ablakba."
        )

        c1, c2 = st.columns([2, 1])
        with c1:
            target_lang_name = st.selectbox(
                "🌐 Célnyelv kijelölése",
                list(LANGUAGES.keys()),
                index=0
            )
        with c2:
            st.write("")
            st.write("")
            translate_button = st.button("⚡ Fordítás Indítása", use_container_width=True)

        if translate_button:
            if uploaded_file is None:
                st.warning("⚠️ Kérlek, csatolj egy fájlt az indításhoz!")
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

                    status_box.success(f"✨ Teljes feldolgozás sikeres! {count} db kifejezés lefordítva ({total_found} összes mezőből).")

                    base_name, _ = os.path.splitext(uploaded_file.name)
                    output_filename = f"{base_name}_forditott{ext}"

                    st.download_button(
                        label=f"📥 Lefordított Dokumentum Letöltése ({output_filename})",
                        data=translated_bytes,
                        file_name=output_filename,
                        mime=mime_type,
                        use_container_width=True
                    )

                except Exception as e:
                    progress_bar.empty()
                    status_box.error(f"Kritikus pipeline hiba: {e}")

# Lábléc
st.markdown("""
<div class="footer-hud">
    <div>TranslateOS Enterprise Architecture</div>
    <div class="engine-status">SYSTEM STATUS: OPTIMAL // POWERED BY NAGY ATTILA</div>
</div>
""", unsafe_allow_html=True)
