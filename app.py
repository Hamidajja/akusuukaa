import streamlit as st
import pandas as pd
import requests
from concurrent.futures import ThreadPoolExecutor
from io import BytesIO
import time
import random
from openpyxl.styles import PatternFill

# --- DAFTAR BAHASA LENGKAP (ISO 639-1 yang umum didukung Google Translate) ---
# Daftar ini mencakup lebih dari 190 bahasa yang didukung oleh Google Translate
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
    # Opsi untuk kode manual
    "--- Lainnya (ketik kode manual) ---": "other"
}

# --- KONFIGURASI USER-AGENT ---
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/119.0"
]

# --- FUNGSI INTI TRANSLATE ---
def translate_core(text, target, source='id'):
    """Request ke Google API gtx."""
    if not text or str(text).strip().lower() in ["nan", "none", ""]:
        return ""
    
    headers = {"User-Agent": random.choice(USER_AGENTS)}
    base_url = "https://translate.googleapis.com/translate_a/single"
    
    params = {
        "client": "gtx",
        "sl": source,
        "tl": target,
        "dt": "t",
        "q": str(text).strip()
    }
    
    try:
        response = requests.get(base_url, params=params, headers=headers, timeout=12)
        if response.status_code == 200:
            result_json = response.json()
            translated_parts = [part[0] for part in result_json[0] if part[0]]
            return "".join(translated_parts)
        elif response.status_code == 429:
            return "ERR_LIMIT"
        else:
            return "ERR_UNSUPPORTED"  # kemungkinan kode bahasa tidak didukung
    except Exception:
        pass
    return None

def translate_smart(text, target):
    """Pecah teks panjang >4500 karakter jika perlu."""
    text_str = str(text).strip()
    
    if len(text_str) <= 4500:
        return translate_core(text_str, target)
    
    chunks = [text_str[i:i+4000] for i in range(0, len(text_str), 4000)]
    translated_results = []
    
    for c in chunks:
        res = translate_core(c, target)
        if res and res != "ERR_LIMIT" and res != "ERR_UNSUPPORTED":
            translated_results.append(res)
        elif res == "ERR_LIMIT":
            return "ERR_LIMIT"
        elif res == "ERR_UNSUPPORTED":
            return "ERR_UNSUPPORTED"
        else:
            return None
    return " ".join(translated_results)

# --- STREAMLIT UI ---
st.set_page_config(page_title="Turbo Translator Pro v2", page_icon="⚡", layout="wide")

st.title("⚡ Turbo Excel Translator")
st.markdown("Alat translasi otomatis untuk file Excel buatan fadhil ganteng kece keren hebat slebew. kalo gatau kodenya tanya gugel nulisnya gini 639-1 kode bahasa ..... bahasa mu ketiken. JANGAN LUPA DIKASIH LETI 1 BARIS DIATAS NYA")
st.markdown("PAKAILAH 1 TAB AJA JANGAN MULTI TAB WOYYYY RUSAK HOST E, NDAK TAK HOST NO MANEH WM")

# --- SIDEBAR ---
st.sidebar.header("⚙️ Pengaturan")

# Pilih bahasa dengan dropdown
selected_lang = st.sidebar.selectbox(
    "🌍 Bahasa Tujuan",
    options=list(LANGUAGES.keys()),
    index=list(LANGUAGES.values()).index("en")  # default English
)

# Jika memilih opsi "lainnya", tampilkan input manual
if LANGUAGES[selected_lang] == "other":
    target_lang = st.sidebar.text_input(
        "✍️ Masukkan kode bahasa (contoh: nqo untuk N'Ko, eo untuk Esperanto)",
        value="",
        help="Ketik kode ISO 639-1/639-2 untuk bahasa yang tidak ada di daftar."
    )
else:
    target_lang = LANGUAGES[selected_lang]

max_workers = st.sidebar.slider("Kecepatan (Workers)", 1, 15, 5, help="Disarankan 5-10 agar aman.")

st.sidebar.markdown("---")
st.sidebar.info(
    "📌 **Catatan Penting:**\n"
    "• Jika hasil berwarna merah, artinya gagal (mungkin limit atau kode tidak didukung).\n"
    "• Google Translate sekarang mendukung banyak bahasa, tapi jika kode tidak dikenali, akan muncul error.\n"
    "• Jika ada kode tidak dikenal, tapi yakin sudah didukung, coba refresh halaman atau periksa daftar resmi Google.\n"
    "• Untuk N'Ko, gunakan kode 'nqo' (tanpa tanda kutip)."
)

# --- PROSES UTAMA ---
uploaded_file = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        st.success(f"File dimuat: '{uploaded_file.name}' | Total: {len(df)} baris.")

        if not target_lang:
            st.warning("⚠️ Silakan pilih atau masukkan kode bahasa tujuan terlebih dahulu.")
        elif st.button("🚀 Mulai Terjemahkan"):
            if df.shape[1] < 2:
                st.error("Kolom B (Kolom ke-2) tidak ditemukan!")
            else:
                texts_to_process = df.iloc[:, 1].tolist()
                total_rows = len(texts_to_process)
                results = [None] * total_rows
                
                progress_bar = st.progress(0)
                status_placeholder = st.empty()
                time_placeholder = st.empty()
                
                start_time = time.time()

                with ThreadPoolExecutor(max_workers=max_workers) as executor:
                    future_to_idx = {executor.submit(translate_smart, texts_to_process[i], target_lang): i for i in range(total_rows)}
                    
                    completed = 0
                    for future in future_to_idx:
                        idx = future_to_idx[future]
                        try:
                            results[idx] = future.result()
                        except:
                            results[idx] = None
                        
                        completed += 1
                        elapsed = time.time() - start_time
                        avg_time = elapsed / completed
                        eta = int(avg_time * (total_rows - completed))
                        
                        progress_bar.progress(completed / total_rows)
                        status_placeholder.write(f"⏳ Memproses: {completed}/{total_rows} baris")
                        time_placeholder.markdown(f"⏱️ Sisa waktu: **{eta} detik**")

                df['Hasil Translate'] = results
                
                # Cek jika banyak error unsupported
                unsupported_count = sum(1 for res in results if res == "ERR_UNSUPPORTED")
                if unsupported_count > 0:
                    st.warning(f"⚠️ {unsupported_count} baris gagal karena kemungkinan kode bahasa **'{target_lang}'** tidak didukung Google Translate. Cek kembali kode atau gunakan API lain untuk bahasa langka.")

                # --- PREVIEW ---
                st.subheader("📋 Preview Hasil (5 Baris Pertama)")
                st.dataframe(df[['Hasil Translate']].head(5))

                # --- GENERASI FILE DENGAN WARNA ---
                nama_file_murni = uploaded_file.name.rsplit('.', 1)[0]
                nama_file_baru = f"{nama_file_murni} ({target_lang}).xlsx"

                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Sheet1')
                    
                    workbook = writer.book
                    worksheet = writer.sheets['Sheet1']
                    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                    
                    # Warnai merah untuk hasil gagal/error/kosong
                    for row_num, val in enumerate(results, start=2):
                        if val is None or "ERR" in str(val) or val == "":
                            for col_num in range(1, df.shape[1] + 1):
                                worksheet.cell(row=row_num, column=col_num).fill = red_fill

                st.success(f"✅ Selesai! Nama file: {nama_file_baru}")
                
                st.download_button(
                    label="📥 Download Hasil Terjemahan",
                    data=output.getvalue(),
                    file_name=nama_file_baru,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
