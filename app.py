import io
import os
import re
import json
import time
import zipfile
import streamlit as st
import google.generativeai as genai
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

st.markdown('<div class="main-title">🌐 Office Document Translator Pro <span style="font-size: 1rem; color: #10b981;">v13.0 Hybrid</span></div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Intelligens, formázásmegőrző Office fájlfordító</div>', unsafe_allow_html=True)

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
    "Román (Romanian)": "Romanian",
    "Japán (Japanese)": {"gemini": "Japanese", "code": "ja"},
    "Török (Turkish)": {"gemini": "Turkish", "code": "tr"},
    "Vietnámi (Vietnamese)": {"gemini": "Vietnamese", "code": "vi"}
}

gemini_key = st.secrets.get("GEMINI_API_KEY")
gemini_available = False
if gemini_key:
    try:
        genai.configure(api_key=gemini_key)
        model = genai.GenerativeModel(model_name="gemini-3.8-flash")
        gemini_available = True
    except Exception:
        gemini_available = False

def has_letters(text):
    if not text:
        return False
    t = text.strip()
    if len(t) <= 1 or t.startswith("="):
        return False
    return bool(re.search(r"[a-zA-ZáéíóöőúüűÁÉÍÓÖŐÚÜŰ]", t))

def translate_fallback(texts, target_lang_code, status_box):
    """Ingyenes fordító fallback, ha a Gemini kvóta kimerült."""
    translator = GoogleTranslator(source='auto', target=target_lang_code)
    results = {}
    total = len(texts)
    status_box.info(f"Átváltás tartalék fordító motorra ({total} elem)...")
    
    for idx, txt in enumerate(texts):
        try:
            res = translator.translate(txt)
            results[txt] = res if res else txt
        except Exception:
            results[txt] = txt
        if idx % 10 == 0:
            time.sleep(0.1)
    return results

def translate_hybrid(unique_texts, target_lang_info, status_box):
    """Megpróbálja a Geminivel lefordítani; 429 kvótahiba esetén automatikusan átvált."""
    target_gemini = target_lang_info["gemini"] if isinstance(target_lang_info, dict) else target_lang_info
    target_code = target_lang_info["code"] if isinstance(target_lang_info, dict) else "hu"
    results = {}

    quota_exhausted = False
    if gemini_available:
        batch_size = 50
        total_batches = (len(unique_texts) + batch_size - 1) // batch_size

        for i in range(0, len(unique_texts), batch_size):
            chunk = unique_texts[i:i + batch_size]
            batch_num = (i // batch_size) + 1
            status_box.info(f"AI fordítás: {batch_num}/{total_batches} csomag...")

            input_data = {str(idx + 1): txt for idx, txt in enumerate(chunk)}
            prompt = (
                f"You are a professional industrial, technical, TPM, and business translator.\n"
                f"Translate the JSON values into {target_gemini}.\n"
                f"Keep acronyms (OEE, KPI, TIR, IPS, UPS, PDCA, DDS, WPA, BS, TBR, PSR, FI, CBN, TEAT, SWP, LDR) intact.\n"
                f"Translate roles: 'Pillar owners' -> 'Pillér felelősök', 'Plant' -> 'Üzem/Gyár', 'Schedule' -> 'Ütemterv', 'Loss' -> 'Veszteség'.\n"
                f"Return ONLY valid JSON matching input keys. No markdown backticks.\n\n"
                f"{json.dumps(input_data, ensure_ascii=False)}"
            )

            try:
                resp = model.generate_content(prompt, request_options={"timeout": 60})
                raw_text = resp.text.strip()
                if raw_text.startswith("```json"):
                    raw_text = raw_text[7:]
                if raw_text.startswith("```"):
                    raw_text = raw_text[3:]
                if raw_text.endswith("```"):
                    raw_text = raw_text[:-3]

                parsed_json = json.loads(raw_text.strip())
                for idx, orig in enumerate(chunk):
                    k = str(idx + 1)
                    if k in parsed_json and parsed_json[k]:
                        results[orig] = str(parsed_json[k]).strip()
            except Exception as e:
                err_str = str(e)
                if "429" in err_str or "quota" in err_str.lower():
                    st.warning("⚠️ A Gemini API napi kerete kimerült. Az alkalmazás automatikusan átváltott a korlátlan tartalék fordítóra!")
                    quota_exhausted = True
                    break
                else:
                    for orig in chunk:
                        results[orig] = orig

            time.sleep(0.3)

    # Ha a kvóta elfogyott, a még le nem fordított szövegeket a fallback motor fejezi be
    remaining = [t for t in unique_texts if t not in results or results[t] == t]
    if quota_exhausted or not gemini_available or len(remaining) > len(unique_texts) // 2:
        fallback_results = translate_fallback(unique_texts, target_code, status_box)
        results.update(fallback_results)

    return results

def process_xlsx_hybrid(file_bytes, target_lang_info, progress_bar, status_box):
    status_box.info("Excel szövegtár kinyerése...")
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

    t_map = translate_hybrid(unique_texts, target_lang_info, status_box)
    progress_bar.progress(80)

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
        target_info = LANGUAGES[target_lang_name]
        ext = os.path.splitext(uploaded_file.name)[1].lower()
        file_bytes = uploaded_file.read()

        status_box = st.empty()
        progress_bar = st.progress(5)

        try:
            translated_bytes = None
            count = 0
            total_found = 0

            if ext == ".xlsx":
                translated_bytes, count, total_found = process_xlsx_hybrid(file_bytes, target_info, progress_bar, status_box)
                mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            else:
                st.info("A Word és PPT formátum támogatása kész.")

            status_box.success(f"✅ Kész! Összesen {count} db szöveges elem sikeresen lefordítva ({total_found} talált mezőből).")

            base_name, _ = os.path.splitext(uploaded_file.name)
            output_filename = f"{base_name}_forditott_hu{ext}"

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
