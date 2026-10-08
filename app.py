import io
import os
import re
import json
import time
import urllib.parse
import urllib.request
import zipfile
import streamlit as st

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

st.markdown('<div class="main-title">🌐 Office Document Translator Pro <span style="font-size: 1rem; color: #10b981;">v16.0 DirectEngine</span></div>', unsafe_allow_html=True)
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

def has_letters(text):
    if not text:
        return False
    t = text.strip()
    if len(t) <= 1 or t.startswith("="):
        return False
    return bool(re.search(r"[a-zA-ZáéíóöőúüűÁÉÍÓÖŐÚÜŰ]", t))

def direct_translate(text, target_lang="hu"):
    """Közvetlen Google Translate kérés - nincs API kulcs, nincs letiltás."""
    if not text or not has_letters(text):
        return text
    
    url = "https://translate.googleapis.com/translate_a/single?client=gtx&sl=auto&tl=" + target_lang + "&dt=t&q=" + urllib.parse.quote(text)
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }
    req = urllib.request.Request(url, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=10) as response:
            result = json.loads(response.read().decode('utf-8'))
            translated = "".join([part[0] for part in result[0] if part[0]])
            return translated if translated else text
    except Exception as e:
        return text

def translate_all_direct(unique_texts, target_code, status_box, progress_bar):
    """Kötegelt közvetlen fordítás progress bar visszajelzéssel és hibajelzéssel."""
    results = {}
    total = len(unique_texts)
    success_count = 0
    
    for idx, txt in enumerate(unique_texts):
        trans = direct_translate(txt, target_lang=target_code)
        if trans != txt:
            results[txt] = trans
            success_count += 1
        else:
            results[txt] = txt
        
        # Frissítjük a folyamatjelzőt 10 elemenként
        if idx % 5 == 0 or idx == total - 1:
            perc = int(35 + ((idx + 1) / total) * 45)
            progress_bar.progress(perc)
            status_box.info(f"Szövegek fordítása: {idx + 1}/{total} kész ({success_count} lefordítva)...")
        time.sleep(0.05) # Kis késleltetés a szerver védelmére
        
    return results

def process_xlsx(file_bytes, target_code, progress_bar, status_box):
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
    status_box.info(f"{len(raw_texts)} db szöveges mező ({len(unique_texts)} egyedi) megtalálva. Közvetlen fordítás indítása...")
    progress_bar.progress(35)

    t_map = translate_all_direct(unique_texts, target_code, status_box, progress_bar)
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
    st.write("")
    translate_button = st.button("🚀 Fordítás indítása", use_container_width=True, type="primary")

if translate_button:
    if uploaded_file is None:
        st.warning("Kérlek, válassz ki egy fájlt a fordítás megkezdéséhez!")
    else:
        target_code = LANGUAGES[target_lang_name]
        ext = os.path.splitext(uploaded_file.name)[1].lower()
        file_bytes = uploaded_file.read()

        status_box = st.empty()
        progress_bar = st.progress(5)

        try:
            translated_bytes, count, total_found = process_xlsx(file_bytes, target_code, progress_bar, status_box)
            mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

            status_box.success(f"✅ Kész! Összesen {count} db szöveges elem sikeresen lefordítva ({total_found} talált mezőből).")

            base_name, _ = os.path.splitext(uploaded_file.name)
            output_filename = f"{base_name}_forditott_{target_code}{ext}"

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
