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

st.markdown('<div class="main-title">🌐 Office Document Translator Pro <span style="font-size: 1rem; color: #10b981;">v15.0 NoQuota</span></div>', unsafe_allow_html=True)
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
    "Román (Romanian)": {"gemini": "Romanian", "code": "ro"},
    "Japán (Japanese)": {"gemini": "Japanese", "code": "ja"},
    "Török (Turkish)": {"gemini": "Turkish", "code": "tr"},
    "Vietnámi (Vietnamese)": {"gemini": "Vietnamese", "code": "vi"}
}

# API Beállítások
gemini_key = st.secrets.get("GEMINI_API_KEY")
gemini_configured = False
if gemini_key:
    try:
        genai.configure(api_key=gemini_key)
        gemini_configured = True
    except Exception:
        pass

AVAILABLE_MODELS = [
    "gemini-2.0-flash",
    "gemini-1.5-flash",
    "gemini-1.5-pro",
    "gemini-3.8-flash"
]

def has_letters(text):
    if not text:
        return False
    t = text.strip()
    if len(t) <= 1 or t.startswith("="):
        return False
    return bool(re.search(r"[a-zA-ZáéíóöőúüűÁÉÍÓÖŐÚÜŰ]", t))

# -------------------------------------------------------------
# 1. INGYENES (DEEP-TRANSLATOR) MOTOR (Nincs kvóta, nincs API kulcs)
# -------------------------------------------------------------
def translate_with_free_engine(texts, target_code, status_box):
    if not texts:
        return {}
        
    unique_texts = list(set([t.strip() for t in texts if has_letters(t)]))
    results = {}
    translator = GoogleTranslator(source='auto', target=target_code)
    
    batch_size = 20
    total_batches = (len(unique_texts) + batch_size - 1) // batch_size
    
    for i in range(0, len(unique_texts), batch_size):
        chunk = unique_texts[i:i + batch_size]
        batch_num = (i // batch_size) + 1
        status_box.info(f"Ingyenes motor (Google Translate) használata: {batch_num}/{total_batches} csomag...")
        
        # Biztonságos számozott lista
        lines_input = "\n".join([f"[{idx+1}] {t.replace(chr(10), ' ')}" for idx, t in enumerate(chunk)])
        
        success = False
        try:
            translated = translator.translate(lines_input)
            lines = translated.split('\n')
            
            temp_map = {}
            for line in lines:
                line = line.strip()
                if line.startswith("[") and "]" in line:
                    num_s = line[1:line.find("]")].strip()
                    trans_s = line[line.find("]")+1:].strip()
                    num_clean = "".join([c for c in num_s if c.isdigit()])
                    if num_clean:
                        idx = int(num_clean) - 1
                        if 0 <= idx < len(chunk):
                            temp_map[idx] = trans_s
                            
            if len(temp_map) == len(chunk):
                for idx, orig in enumerate(chunk):
                    results[orig] = temp_map[idx]
                success = True
            elif len(temp_map) > len(chunk) // 2:
                for idx, val in temp_map.items():
                    results[chunk[idx]] = val
                success = True
        except Exception:
            time.sleep(1)
            
        # Tartalék egyenkénti fordítás, ha a csomag elromlott
        for orig in chunk:
            if orig not in results:
                try:
                    res = translator.translate(orig)
                    results[orig] = res if res else orig
                    time.sleep(0.1)
                except Exception:
                    results[orig] = orig
        
        time.sleep(1) # Késleltetés az IP letiltás elkerülésére
        
    return results

# -------------------------------------------------------------
# 2. GEMINI MOTOR (Okos, de limitált)
# -------------------------------------------------------------
def translate_with_gemini(unique_texts, target_gemini, status_box):
    if not unique_texts:
        return {}

    input_data = {str(idx + 1): txt for idx, txt in enumerate(unique_texts)}
    prompt = (
        f"You are a professional industrial, technical, TPM, and business translator.\n"
        f"Translate the values of the JSON object into {target_gemini}.\n"
        f"Rules:\n"
        f"1. Keep technical acronyms intact (OEE, KPI, TIR, IPS, UPS, PDCA, DDS, WPA, BS, TBR, PSR, FI, CBN, TEAT, SWP, LDR, etc.).\n"
        f"2. Translate common terms: 'Pillar owners' -> 'Pillér felelősök', 'Plant' -> 'Üzem/Gyár', 'Schedule' -> 'Ütemterv', 'Loss' -> 'Veszteség'.\n"
        f"3. Return ONLY a valid JSON object with the exact same keys ('1', '2', etc.) and the translated values. Do not wrap in markdown.\n\n"
        f"{json.dumps(input_data, ensure_ascii=False)}"
    )

    results = {}
    success = False

    for model_name in AVAILABLE_MODELS:
        status_box.info(f"AI fordítás a következő modellel: {model_name}...")
        try:
            m = genai.GenerativeModel(model_name=model_name)
            resp = m.generate_content(prompt, request_options={"timeout": 90})
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
            
            success = True
            st.success(f"✅ Sikeres fordítás a(z) **{model_name}** modellel!")
            break
        except Exception as e:
            err_str = str(e)
            if "429" in err_str or "quota" in err_str.lower():
                continue
            elif "404" in err_str:
                continue
            else:
                continue

    if not success:
        raise Exception("Minden Gemini modell kvótája kimerült!")

    return results

# -------------------------------------------------------------
# FÁJL FELDOLGOZÁS
# -------------------------------------------------------------
def process_xlsx(file_bytes, target_info, use_free, progress_bar, status_box):
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
    status_box.info(f"{len(raw_texts)} db szöveges mező ({len(unique_texts)} egyedi) megtalálva. Fordítás indítása...")
    progress_bar.progress(35)

    # Döntés a motorok között
    if use_free:
        t_map = translate_with_free_engine(unique_texts, target_info["code"], status_box)
    else:
        try:
            t_map = translate_with_gemini(unique_texts, target_info["gemini"], status_box)
        except Exception as e:
            st.warning("A Gemini kvóta teljesen kimerült. Automatikus átváltás az ingyenes motorra...")
            t_map = translate_with_free_engine(unique_texts, target_info["code"], status_box)

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
    type=["xlsx"],
    help="Excel (.xlsx) fájlokat tölthetsz fel."
)

col1, col2 = st.columns([2, 1])
with col1:
    target_lang_name = st.selectbox("2. Válassz célnyelvet:", list(LANGUAGES.keys()), index=0)
with col2:
    st.write("")
    use_free_engine = st.checkbox("💡 Ingyenes motor használata (Gemini megkerülése)", value=True, help="Ha be van pipálva, nem használja az API kulcsot, így kvótahiba nélkül azonnal lefordítja a fájlt a nyilvános Google Translate-tel.")

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
                translated_bytes, count, total_found = process_xlsx(file_bytes, target_info, use_free_engine, progress_bar, status_box)
                mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            else:
                st.info("A Word és PPT formátum támogatása kész.")

            status_box.success(f"✅ Kész! Összesen {count} db szöveges elem sikeresen lefordítva ({total_found} talált mezőből).")

            base_name, _ = os.path.splitext(uploaded_file.name)
            output_filename = f"{base_name}_forditott_{target_info['code']}{ext}"

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
