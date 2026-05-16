import streamlit as st
import pandas as pd
import requests
from concurrent.futures import ThreadPoolExecutor
from io import BytesIO
import time
import random
from openpyxl.styles import PatternFill

# --- KONFIGURASI USER-AGENT ---
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/119.0"
]

# --- FUNGSI INTI TRANSLATE ---
def translate_core(text, target, source):
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
    except Exception:
        pass
    return None

def translate_smart(text, target, source):
    """Logika Chunking: Memecah teks >4500 karakter."""
    text_str = str(text).strip()
    
    if len(text_str) <= 4500:
        return translate_core(text_str, target, source)
    
    # Pecah per 4000 karakter agar aman dari limit URL
    chunks = [text_str[i:i+4000] for i in range(0, len(text_str), 4000)]
    translated_results = []
    
    for c in chunks:
        res = translate_core(c, target, source)
        if res and res != "ERR_LIMIT":
            translated_results.append(res)
        else:
            return "ERR_LIMIT" if res == "ERR_LIMIT" else None
            
    return " ".join(translated_results)

# --- ANTARMUKA STREAMLIT ---
st.set_page_config(page_title="Penerjemah Bambara → N'Ko", page_icon="🔤", layout="wide")

st.title("🔤 Penerjemah Excel: Bambara ke N'Ko")
st.markdown("Unggah file Excel Anda, pilih kolom teks, dan terjemahkan otomatis dari Bahasa Bambara ke aksara N'Ko.")

# --- SIDEBAR ---
st.sidebar.header("⚙️ Pengaturan Bahasa")
source_lang = st.sidebar.text_input("Kode Bahasa Sumber", value="bm", 
                                    help="Kode ISO 639-1 untuk bahasa sumber. Default: bm (Bambara)")
target_lang = st.sidebar.text_input("Kode Bahasa Tujuan", value="nqo", 
                                    help="Kode bahasa tujuan. Default: nqo (N'Ko)")

max_workers = st.sidebar.slider("Kecepatan (Workers)", 1, 15, 5, 
                                help="Jumlah permintaan simultan. Disarankan 5–10.")

st.sidebar.markdown("---")
st.sidebar.info(
    "📌 **Catatan:**\n"
    "- Kolom pertama di Excel akan diabaikan, kolom kedua (indeks 1) dianggap sebagai teks sumber.\n"
    "- Jika sel hasil berwarna merah saat diunduh, berarti baris tersebut gagal diterjemahkan (limit/kosong).\n"
    "- Kurangi *Workers* jika sering terkena limit."
)

# --- PROSES UTAMA ---
uploaded_file = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        st.success(f"File berhasil dimuat: **{uploaded_file.name}** | Total baris: {len(df)}")

        if st.button("🚀 Mulai Terjemahkan"):
            if df.shape[1] < 2:
                st.error("File harus memiliki minimal 2 kolom. Kolom ke-2 akan dijadikan sumber teks.")
            else:
                texts_to_process = df.iloc[:, 1].tolist()
                total_rows = len(texts_to_process)
                results = [None] * total_rows
                
                # UI Progress
                progress_bar = st.progress(0)
                status_placeholder = st.empty()
                time_placeholder = st.empty()
                
                start_time = time.time()

                # --- MULTITHREADING ---
                with ThreadPoolExecutor(max_workers=max_workers) as executor:
                    future_to_idx = {
                        executor.submit(translate_smart, texts_to_process[i], target_lang, source_lang): i 
                        for i in range(total_rows)
                    }
                    
                    completed = 0
                    for future in future_to_idx:
                        idx = future_to_idx[future]
                        try:
                            results[idx] = future.result()
                        except:
                            results[idx] = None
                        
                        completed += 1
                        
                        # Estimasi waktu
                        elapsed = time.time() - start_time
                        avg_time = elapsed / completed
                        eta = int(avg_time * (total_rows - completed))
                        
                        progress_bar.progress(completed / total_rows)
                        status_placeholder.write(f"⏳ Memproses: {completed}/{total_rows} baris")
                        time_placeholder.markdown(f"⏱️ Sisa waktu: **{eta} detik**")

                df['Hasil Terjemahan (N\'Ko)'] = results
                
                # --- PREVIEW ---
                st.subheader("📋 Preview (5 Baris Pertama)")
                st.dataframe(df[['Hasil Terjemahan (N\'Ko)']].head(5))

                # --- GENERASI FILE DENGAN WARNA & NAMA DINAMIS ---
                nama_file_murni = uploaded_file.name.rsplit('.', 1)[0]
                nama_file_baru = f"{nama_file_murni} ({source_lang} → {target_lang}).xlsx"

                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Sheet1')
                    
                    workbook = writer.book
                    worksheet = writer.sheets['Sheet1']
                    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                    
                    # Tandai baris yang gagal
                    for row_num, val in enumerate(results, start=2):
                        if val is None or val == "ERR_LIMIT" or val == "":
                            for col_num in range(1, df.shape[1] + 1):
                                worksheet.cell(row=row_num, column=col_num).fill = red_fill

                st.success(f"✅ Terjemahan selesai! File siap diunduh: **{nama_file_baru}**")
                
                st.download_button(
                    label="📥 Unduh Hasil Terjemahan",
                    data=output.getvalue(),
                    file_name=nama_file_baru,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
