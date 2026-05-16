import streamlit as st
import pandas as pd
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed
from io import BytesIO
import time
import random
from openpyxl.styles import PatternFill
import urllib.parse

# --- KONFIGURASI USER-AGENT ---
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0"
]

# --- FUNGSI INTI TRANSLATE (DENGAN RETRY) ---
def safe_request(url, params, headers, retries=3):
    """Melakukan request dengan retry jika terkena limit."""
    for attempt in range(retries):
        try:
            # Tambahkan delay acak agar terlihat seperti manusia
            time.sleep(random.uniform(0.5, 1.5)) 
            response = requests.get(url, params=params, headers=headers, timeout=15)
            
            if response.status_code == 200:
                return response.json()
            elif response.status_code == 429:
                wait_time = (attempt + 1) * 5 # Wait 5s, 10s, 15s
                time.sleep(wait_time)
                continue
            else:
                return None
        except Exception:
            time.sleep(2)
            continue
    return "ERR_LIMIT"

def translate_single_step(text, source, target):
    """Translate satu langkah via gtx."""
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
    
    result_json = safe_request(base_url, params, headers)
    
    if result_json == "ERR_LIMIT":
        return "ERR_LIMIT"
    
    if result_json:
        try:
            translated_parts = [part[0] for part in result_json[0] if part[0]]
            return "".join(translated_parts)
        except:
            return None
    return None

def translate_to_dari_strategy(text):
    """
    Strategi Khusus Dari:
    1. ID -> EN (Inggris lebih stabil)
    2. EN -> fa-AF (Dari Afghanistan)
    """
    if not text or str(text).strip().lower() in ["nan", "none", ""]:
        return ""
        
    text_str = str(text).strip()
    
    # Langkah 1: Indonesia ke Inggris
    en_result = translate_single_step(text_str, 'id', 'en')
    if not en_result or en_result == "ERR_LIMIT":
        return en_result if en_result == "ERR_LIMIT" else None
        
    # Langkah 2: Inggris ke Dari (fa-AF)
    dari_result = translate_single_step(en_result, 'en', 'fa-AF')
    return dari_result

# --- ANTARMUKA STREAMLIT ---
st.set_page_config(page_title="Turbo Translator Pro v3 (Dari Edition)", page_icon="⚡", layout="wide")

st.title("⚡ Turbo Excel Translator - Edisi Dari")
st.markdown("Alat translasi otomatis khusus Bahasa Dari (Afghanistan).")
st.info("💡 **Strategi:** Sistem akan menerjemahkan ID → EN → fa-AF secara otomatis untuk hasil terbaik.")

# --- SIDEBAR ---
st.sidebar.header("⚙️ Pengaturan")
# Kita kunci targetnya ke fa-AF untuk kasus ini, tapi bisa diubah jika mau
target_lang_display = "fa-AF (Dari)"
max_workers = st.sidebar.slider("Kecepatan (Workers)", 1, 5, 2, help="Untuk gtx gratis, JANGAN lebih dari 3-5 agar IP tidak dibanned.")

st.sidebar.markdown("---")
st.sidebar.warning("⚠️ **PENTING:**\nGunakan Workers rendah (2-3). Jika muncul 'ERR_LIMIT', istirahat sejenak.")

# --- PROSES UTAMA ---
uploaded_file = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        st.success(f"File dimuat: '{uploaded_file.name}' | Total: {len(df)} baris.")

        if st.button("🚀 Mulai Terjemahkan ke Dari"):
            if df.shape[1] < 2:
                st.error("Kolom B (Kolom ke-2) tidak ditemukan!")
            else:
                texts_to_process = df.iloc[:, 1].tolist()
                total_rows = len(texts_to_process)
                results = [None] * total_rows
                
                # UI Progress
                progress_bar = st.progress(0)
                status_placeholder = st.empty()
                
                start_time = time.time()

                # --- MULTITHREADING ---
                # Gunakan ThreadPoolExecutor
                with ThreadPoolExecutor(max_workers=max_workers) as executor:
                    # Submit semua tugas
                    future_to_idx = {executor.submit(translate_to_dari_strategy, texts_to_process[i]): i for i in range(total_rows)}
                    
                    completed = 0
                    for future in as_completed(future_to_idx):
                        idx = future_to_idx[future]
                        try:
                            results[idx] = future.result()
                        except Exception as e:
                            results[idx] = f"Error: {str(e)}"
                        
                        completed += 1
                        
                        # Update Progress UI
                        progress_bar.progress(completed / total_rows)
                        elapsed = time.time() - start_time
                        avg_time = elapsed / completed if completed > 0 else 1
                        eta = int(avg_time * (total_rows - completed))
                        status_placeholder.write(f"⏳ Memproses: {completed}/{total_rows} baris (Estimasi sisa: {eta} detik)")

                df['Hasil Translate (Dari)'] = results
                
                # --- PREVIEW ---
                st.subheader("📋 Preview Hasil (5 Baris Pertama)")
                st.dataframe(df[['Hasil Translate (Dari)']].head(5))

                # --- GENERASI FILE ---
                nama_file_murni = uploaded_file.name.rsplit('.', 1)[0]
                nama_file_baru = f"{nama_file_murni} (Dari_fa-AF).xlsx"

                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Sheet1')
                    
                    workbook = writer.book
                    worksheet = writer.sheets['Sheet1']
                    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                    
                    # Warnai error
                    for row_num, val in enumerate(results, start=2):
                        if val is None or val == "ERR_LIMIT" or str(val).startswith("Error"):
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
        st.error(f"Terjadi kesalahan sistem: {e}")
