import streamlit as st
import pandas as pd
import requests
from concurrent.futures import ThreadPoolExecutor
from io import BytesIO
import time
import random
from openpyxl.styles import PatternFill

# ========== DAFTAR BAHASA LENGKAP ==========
LANGUAGES = {
    "Abkhaz": "ab", "Acehnese": "ace", "Acholi": "ach", "Afar": "aa", "Afrikaans": "af",
    "Albanian": "sq", "Alur": "alz", "Amharic": "am", "Arabic": "ar", "Armenian": "hy",
    "Assamese": "as", "Avar": "av", "Awadhi": "awa", "Aymara": "ay", "Azerbaijani": "az",
    "Balinese": "ban", "Baluchi": "bal", "Bambara": "bm", "Bashkir": "ba", "Basque": "eu",
    "Batak Karo": "btx", "Batak Simalungun": "bts", "Batak Toba": "bbc", "Belarusian": "be",
    "Bemba": "bem", "Bengali": "bn", "Betawi": "bew", "Bhojpuri": "bho", "Bikol": "bik",
    "Bosnian": "bs", "Breton": "br", "Bulgarian": "bg", "Buryat": "bua",
    "Cantonese": "yue", "Catalan": "ca", "Cebuano": "ceb", "Chamorro": "ch",
    "Chechen": "ce", "Chichewa (Nyanja)": "ny", "Chinese (Simplified)": "zh-CN",
    "Chinese (Traditional)": "zh-TW", "Chuukese": "chk", "Chuvash": "cv",
    "Corsican": "co", "Crimean Tatar": "crh", "Croatian": "hr", "Czech": "cs",
    "Danish": "da", "Dari": "prs", "Dinka": "din", "Divehi": "dv", "Dogri": "doi",
    "Dombe": "dov", "Dutch": "nl", "Dzongkha": "dz",
    "English": "en", "Esperanto": "eo", "Estonian": "et", "Ewe": "ee",
    "Faroese": "fo", "Fijian": "fj", "Filipino (Tagalog)": "tl", "Finnish": "fi",
    "Fon": "fon", "French": "fr", "French (Canada)": "fr-CA", "French (France)": "fr-FR",
    "Frisian": "fy", "Friulian": "fur", "Fulfulde": "ff",
    "Ga": "gaa", "Galician": "gl", "Ganda (Luganda)": "lg", "Georgian": "ka",
    "German": "de", "Greek": "el", "Guarani": "gn", "Gujarati": "gu",
    "Haitian Creole": "ht", "Hakha Chin": "cnh", "Hausa": "ha", "Hawaiian": "haw",
    "Hebrew": "he", "Hiligaynon": "hil", "Hindi": "hi", "Hmong": "hmn",
    "Hungarian": "hu", "Hunsrik": "hrx",
    "Icelandic": "is", "Igbo": "ig", "Iloko": "ilo", "Indonesian": "id",
    "Irish": "ga", "Italian": "it",
    "Jamaican Patois": "jam", "Japanese": "ja", "Javanese": "jv", "Jingpo": "kac",
    "Kalaallisut": "kl", "Kannada": "kn", "Kanuri": "kr", "Kapampangan": "pam",
    "Kazakh": "kk", "Khasi": "kha", "Khmer": "km", "Kiga": "cgg", "Kikongo": "kg",
    "Kinyarwanda": "rw", "Kituba": "ktu", "Kokborok": "trp", "Komi": "kv",
    "Konkani": "gom", "Korean": "ko", "Kurdish (Kurmanji)": "ku", "Kyrgyz": "ky",
    "Lao": "lo", "Latgalian": "ltg", "Latin": "la", "Latvian": "lv",
    "Ligurian": "lij", "Limburgish": "li", "Lithuanian": "lt", "Lombard": "lmo",
    "Luo": "luo", "Luxembourgish": "lb",
    "Macedonian": "mk", "Madurese": "mad", "Makassar": "mak", "Malagasy": "mg",
    "Malay": "ms", "Malay (Jawi)": "ms-Arab", "Malayalam": "ml", "Maltese": "mt",
    "Mam": "mam", "Manx": "gv", "Maori": "mi", "Marathi": "mr", "Marshallese": "mh",
    "Marwadi": "mwr", "Mauritian Creole": "mfe", "Meadow Mari": "mhr",
    "Minang": "min", "Mongolian": "mn", "Myanmar (Burmese)": "my",
    "Nahuatl (Eastern Huasteca)": "nhe", "Ndau": "ndc", "Ndebele (South)": "nr",
    "Nepalbhasa (Newari)": "new", "Nepali": "ne", "NKo": "nqo", "Norwegian": "no",
    "Nuer": "nus",
    "Occitan": "oc", "Odia (Oriya)": "or", "Oromo": "om", "Ossetian": "os",
    "Pangasinan": "pag", "Papiamento": "pap", "Pashto": "ps", "Persian": "fa",
    "Polish": "pl", "Portuguese": "pt", "Portuguese (Portugal)": "pt-PT",
    "Punjabi": "pa", "Punjabi (Shahmukhi)": "pa-Arab",
    "Q'eqchi'": "kek",
    "Romani": "rom", "Romanian": "ro", "Rundi": "rn", "Russian": "ru",
    "Samoan": "sm", "Sami (North)": "se", "Sango": "sg", "Santali": "sat",
    "Scots Gaelic": "gd", "Sepedi": "nso", "Serbian": "sr", "Sesotho": "st",
    "Seychellois Creole": "crs", "Shan": "shn", "Shona": "sn", "Sicilian": "scn",
    "Silesian": "szl", "Sindhi": "sd", "Sinhala": "si", "Slovak": "sk",
    "Slovenian": "sl", "Somali": "so", "Spanish": "es", "Sundanese": "su",
    "Susu": "sus", "Swahili": "sw", "Swati": "ss", "Swedish": "sv",
    "Tahitian": "ty", "Tajik": "tg", "Tamazight": "ber", "Tamazight (Tifinagh)": "ber-Tfng",
    "Tamil": "ta", "Tatar": "tt", "Telugu": "te", "Tetum": "tet", "Thai": "th",
    "Tibetan": "bo", "Tigrinya": "ti", "Tiv": "tiv", "Tok Pisin": "tpi",
    "Tongan": "to", "Tswana": "tn", "Tulu": "tcy", "Tumbuka": "tum", "Turkish": "tr",
    "Turkmen": "tk", "Tuvan": "tyv",
    "Udmurt": "udm", "Ukrainian": "uk", "Urdu": "ur", "Uyghur": "ug", "Uzbek": "uz",
    "Venda": "ve", "Venetian": "vec", "Vietnamese": "vi",
    "Waray": "war", "Welsh": "cy", "Wolof": "wo",
    "Xhosa": "xh",
    "Yakut": "sah", "Yiddish": "yi", "Yoruba": "yo", "Yucatec Maya": "yua",
    "Zapotec": "zap", "Zulu": "zu",
    "--- Lainnya (ketik kode manual) ---": "other"
}

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/119.0"
]

# ========== FUNGSI TRANSLATE PER PROVIDER ==========
def translate_google_gtx(text, target, source='id'):
    headers = {"User-Agent": random.choice(USER_AGENTS)}
    url = "https://translate.googleapis.com/translate_a/single"
    params = {"client": "gtx", "sl": source, "tl": target, "dt": "t", "q": text.strip()}
    try:
        resp = requests.get(url, params=params, headers=headers, timeout=10)
        if resp.status_code == 200:
            parts = [part[0] for part in resp.json()[0] if part[0]]
            return "".join(parts)
        elif resp.status_code == 429:
            return "ERR_LIMIT"
        else:
            return "ERR_UNSUPPORTED"
    except:
        return None

def translate_deep_translator(text, target, source='id'):
    """Fallback via library deep-translator (Google Web)."""
    try:
        from deep_translator import GoogleTranslator
        # deep-translator bisa menerima 'auto' untuk source
        translated = GoogleTranslator(source='auto', target=target).translate(text.strip())
        return translated
    except Exception:
        return None

# ========== FUNGSI UTAMA DENGAN FALLBACK ==========
def translate_with_fallback(text, target, source='id'):
    """Coba provider sesuai urutan di FALLBACK_ORDER (global dari sidebar)."""
    if not text or str(text).strip().lower() in ["nan", "none", ""]:
        return ""

    text = str(text).strip()
    # Jika teks panjang, pecah dulu (fallback hanya untuk potongan)
    if len(text) > 4500:
        chunks = [text[i:i+4000] for i in range(0, len(text), 4000)]
        results = []
        for chunk in chunks:
            res = translate_with_fallback(chunk, target, source)
            if res and "ERR" not in res:
                results.append(res)
            elif res == "ERR_LIMIT":
                return "ERR_LIMIT"
            elif res == "ERR_UNSUPPORTED":
                return "ERR_UNSUPPORTED"
            else:
                return None
        return " ".join(results)

    for provider in FALLBACK_ORDER:
        if provider == "google_gtx":
            result = translate_google_gtx(text, target, source)
        elif provider == "deep_translator":
            result = translate_deep_translator(text, target, source)
        elif provider == "google_cloud":
            # Hanya bisa dipakai jika kunci diisi
            if st.session_state.get("gc_key"):
                result = translate_google_cloud_api(text, target, source, st.session_state.gc_key)
            else:
                continue
        else:
            continue

        if result is None or "ERR" in str(result):
            continue
        return result
    return "ERR_UNSUPPORTED"

# Placeholder cloud function (tidak berubah, hanya jika ada key)
def translate_google_cloud_api(text, target, source, api_key):
    url = "https://translation.googleapis.com/language/translate/v2"
    params = {"q": text, "target": target, "source": source, "format": "text", "key": api_key}
    try:
        resp = requests.get(url, params=params, timeout=10)
        if resp.status_code == 200:
            return resp.json()["data"]["translations"][0]["translatedText"]
        elif resp.status_code == 403:
            return "ERR_LIMIT"
        else:
            return "ERR_UNSUPPORTED"
    except:
        return None

# ========== STREAMLIT UI ==========
st.set_page_config(page_title="Turbo Translator + Fallback", page_icon="⚡", layout="wide")
st.title("⚡ Turbo Excel Translator + Fallback")
st.markdown("... (deskripsi) ...")

# Sidebar
st.sidebar.header("⚙️ Bahasa & Provider")
selected_lang = st.sidebar.selectbox("🌍 Bahasa Tujuan", list(LANGUAGES.keys()),
                                     index=list(LANGUAGES.values()).index("en"))
if LANGUAGES[selected_lang] == "other":
    target_lang = st.sidebar.text_input("Kode bahasa manual", "")
else:
    target_lang = LANGUAGES[selected_lang]

max_workers = st.sidebar.slider("Workers", 1, 15, 5)

st.sidebar.header("🔄 Fallback Order")
fallback_options = st.sidebar.multiselect(
    "Urutan provider yang dicoba:",
    ["google_gtx", "deep_translator", "google_cloud"],
    default=["google_gtx", "deep_translator"],
    help="Akan dicoba berurutan. google_cloud hanya muncul jika Anda isi kunci di bawah."
)
# Simpan di global list agar dipakai di fungsi
FALLBACK_ORDER = fallback_options

# Input kunci API opsional (hanya untuk google_cloud)
if "google_cloud" in FALLBACK_ORDER:
    gc_key = st.sidebar.text_input("🔑 Kunci API Google Cloud (opsional)", type="password")
    if gc_key:
        st.session_state.gc_key = gc_key
else:
    st.session_state.gc_key = None

st.sidebar.info("ℹ️ **deep_translator** adalah fallback gratis yang menggunakan endpoint Google Web.\n"
                "Mungkin bisa menangani bahasa baru seperti N'Ko.\n"
                "Install dengan `pip install deep-translator`.")

# Upload file
uploaded_file = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])

if uploaded_file and target_lang:
    df = pd.read_excel(uploaded_file)
    st.success(f"File '{uploaded_file.name}' dimuat. {len(df)} baris.")
    if st.button("🚀 Mulai Terjemahkan"):
        texts = df.iloc[:, 1].tolist()
        total = len(texts)
        results = [None]*total
        progress_bar = st.progress(0)
        status = st.empty()
        timer = st.empty()
        start = time.time()

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_idx = {
                executor.submit(translate_with_fallback, t, target_lang): i
                for i, t in enumerate(texts)
            }
            for completed, future in enumerate(future_to_idx):
                idx = future_to_idx[future]
                try:
                    results[idx] = future.result()
                except:
                    results[idx] = None
                progress_bar.progress((completed+1)/total)
                elapsed = time.time()-start
                eta = int((elapsed/(completed+1))*(total-completed-1))
                status.write(f"⏳ {completed+1}/{total}")
                timer.markdown(f"⏱️ Sisa: **{eta} detik**")

        df['Hasil Translate'] = results
        st.subheader("Preview")
        st.dataframe(df[['Hasil Translate']].head(5))

        # Download dengan warna merah untuk error
        nama_file = f"{uploaded_file.name.rsplit('.',1)[0]} ({target_lang}).xlsx"
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            ws = writer.sheets['Sheet1']
            red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            for row_num, val in enumerate(results, start=2):
                if val is None or "ERR" in str(val) or val == "":
                    for col in range(1, df.shape[1]+1):
                        ws.cell(row=row_num, column=col).fill = red_fill
        st.download_button("📥 Download", data=output.getvalue(), file_name=nama_file,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
