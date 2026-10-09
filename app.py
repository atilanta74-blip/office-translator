import io
import os
import re
import json
import time
import urllib.parse
import urllib.request
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

st.markdown('<div class="main-title">🌐 Office Document Translator Pro <span style="font-size: 1rem; color: #10b981;">v18.0 Universal</span></div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Intelligens, formázásmegőrző Office fájlfordító (Word, Excel, PowerPoint)</div>', unsafe_allow_html=True)

LANGUAGES = {
    "Magyar (Hungarian)": {"gemini": "Hungarian", "code": "hu"},
    "Angol (English)": {"gemini": "English", "code": "en"},
    "Német (German)": {"gemini": "German", "code": "de"},
    "Olasz (Italian)": {"gemini": "Italian", "code": "it"},
    "Francia (French)": {"gemini": "French", "code": "fr"},
    "Spanyol (Spanish)": {"gemini": "Spanish", "code": "es"},
    "Lengyel (Polish)": {"gemini": "Polish", "code": "pl"},
    "Cseh (Czech)": {"gemini": "Czech", "code": "cs"},
    "Szlovák (Slovak)": {"gemini": "Slovak", "code": "sk"},
    "Román (Romanian)": {"gemini": "Romanian", "code": "ro"},
    "Japán (Japanese)": {"gemini": "Japanese", "code": "ja"},
    "Török (Turkish)": {"gemini": "Turkish", "code": "tr"},
    "Vietnámi (Vietnamese)": {"gemini": "Vietnamese", "code": "vi"}
}

gemini_key = st.secrets.get("GEMINI_API_KEY")

def has_letters(text):
    if not text:
        return False
    t = text.strip()
    if len(t) <= 1 or t.startswith("="):
        return False
    return bool(re.search(r"[a-zA-ZáéíóöőúüűÁÉÍÓÖŐÚÜŰ]", t))

def mymemory_translate(text, target_code="hu"):
    if not text or not has_letters(text):
        return text
    try:
        query = urllib.parse.quote(text)
        url = f"https://api.mymemory.translated.net/get?q={query}&langpair=en|{target_code}"
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=8) as resp:
            data = json.loads(resp.read().decode('utf-8'))
            if data and "responseData" in data and "translatedText" in data["responseData"]:
                res = data["responseData"]["translatedText"]
                if res and res != text:
                    return res
    except Exception:
        pass
    return text

def translate_texts_all(unique_texts, lang_info, status_box, progress_bar):
    target_gemini = lang_info["gemini"]
    target_code = lang_info["code"]
    results = {}
    use_fallback = False

    # 1. Próbálkozás a Geminivel (1 db összefogott hívással)
    if gemini_key:
        status_box.info("Fordítás kísérlet a Gemini AI modellel...")
        try:
            genai.configure(api_key=gemini_key)
            model = genai.GenerativeModel(model_name="gemini-3.8-flash")
            
            input_data = {str(idx + 1): txt for idx, txt in enumerate(unique_texts)}
            prompt = (
                f"You are a professional industrial, technical, TPM, and business translator.\n"
                f"Translate the values of the JSON object into {target_gemini}.\n"
                f"Rules:\n"
                f"1. Keep technical acronyms intact (OEE, KPI, TIR, IPS, UPS, PDCA, DDS, WPA, BS, TBR, PSR, FI, CBN, TEAT, SWP, LDR, etc.).\n"
                f"2. Translate standard terms accurately.\n"
                f"3. Return ONLY a valid JSON object matching the input keys. No markdown backticks.\n\n"
                f"{json.dumps(input_data, ensure_ascii=False)}"
            )
            
            resp = model.generate_content(prompt, request_options={"timeout": 60})
            raw_text = resp.text.strip()
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
            
            st.success("✅ Sikeres fordítás a Gemini AI segítségével!")
            return results
        except Exception as e:
            err_str = str(e)
            st.warning(f"⚠️ Gemini nem elérhető ({err_str[:60]}...). Automatikus átváltás a tartalék motorra!")
            use_fallback = True

    # 2. Tartalék MyMemory motor
    status_box.info(f"Szövegek fordítása a tartalék motorral ({len(unique_texts)} elem)...")
    total = len(unique_texts)
    for idx, txt in enumerate(unique_texts):
        res = mymemory_translate(txt, target_code=target_code)
        results[txt] = res
        if idx % 5 == 0 or idx == total - 1:
            perc = int(35 + ((idx + 1) / total) * 50)
            progress_bar.progress(perc)
            status_box.info(f"Tartalék motor fordítás: {idx + 1}/{total} kész...")
        time.sleep(0.15)

    return results

# Excel feldolgozás
def process_xlsx(file_bytes, lang_info, progress_bar, status_box):
    status_box.info("Excel belső szövegtár kinyerése...")
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
    status_box.info(f"{len(raw_texts)} db szöveges mező ({len(unique_texts)} egyedi) megtalálva. Fordítás...")
    progress_bar.progress(35)

    t_map = translate_texts_all(unique_texts, lang_info, status_box, progress_bar)
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

# Word (.docx) feldolgozás
def process_docx(file_bytes, lang_info, progress_bar, status_box):
    doc = Document(io.BytesIO(file_bytes))
    all_runs = []

    for p in doc.paragraphs:
        for run in p.runs:
            if has_letters(run.text):
                all_runs.append(run)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        if has_letters(run.text):
                            all_runs.append(run)

    status_box.info(f"Word szövegek átvizsgálása ({len(all_runs)} elem)...")
    progress_bar.progress(30)

    unique_texts = list(set([r.text.strip() for r in all_runs]))
    t_map = translate_texts_all(unique_texts, lang_info, status_box, progress_bar)
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
    return out_stream.getvalue(), count, len(all_runs)

# PowerPoint (.pptx) feldolgozás
def process_pptx(file_bytes, lang_info, progress_bar, status_box):
    prs = Presentation(io.BytesIO(file_bytes))
    all_runs = []

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    for run in p.runs:
                        if has_letters(run.text):
                            all_runs.append(run)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for p in cell.text_frame.paragraphs:
                            for run in p.runs:
                                if has_letters(run.text):
                                    all_runs.append(run)

    status_box.info(f"PowerPoint diák átvizsgálása ({len(all_runs)} elem)...")
    progress_bar.progress(30)

    unique_texts = list(set([r.text.strip() for r in all_runs]))
    t_map = translate_texts_all(unique_texts, lang_info, status_box, progress_bar)
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
    return out_stream.getvalue(), count, len(all_runs)

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
        lang_info = LANGUAGES[target_lang_name]
        ext = os.path.splitext(uploaded_file.name)[1].lower()
        file_bytes = uploaded_file.read()

        status_box = st.empty()
        progress_bar = st.progress(5)

        try:
            translated_bytes = None
            count = 0
            total_found = 0

            if ext == ".xlsx":
                translated_bytes, count, total_found = process_xlsx(file_bytes, lang_info, progress_bar, status_box)
                mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            elif ext == ".docx":
                translated_bytes, count, total_found = process_docx(file_bytes, lang_info, progress_bar, status_box)
                mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            elif ext == ".pptx":
                translated_bytes, count, total_found = process_pptx(file_bytes, lang_info, progress_bar, status_box)
                mime_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

            status_box.success(f"✅ Kész! Összesen {count} db szöveges elem sikeresen lefordítva ({total_found} talált mezőből).")

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
            status_box.error(f"Hiba történt a feldolgozás során: {e}")

st.markdown("""
    <div class="footer-bar">
        <span style="color: #64748b;">Office Document Translator Pro</span>
        <span class="powered-by">⚡ Powered by Nagy Attila</span>
    </div>
""", unsafe_allow_html=True)
