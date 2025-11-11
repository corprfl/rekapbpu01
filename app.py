# app.py - Revisi ke-202507211915-2
# Tambahan:
# ✅ Deskripsi aplikasi
# ✅ Panduan penggunaan
# ✅ Peringatan: hanya untuk PDF Bukti Potong Unifikasi dari Coretax (bukan hasil scan)
# ✅ Disclaimer keamanan & tanggung jawab pengguna
# ✅ Warna & tema tetap seperti versi sebelumnya

import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO
import zipfile

st.set_page_config(page_title="Rekap & Rename Bukti Potong Unifikasi", layout="centered")

# ======== THEME SETUP ========
st.markdown("""
<style>
    .stApp {
        background-color: #0d1117;
        color: white;
    }
    h1, h2, h3, h4, h5, h6, p, label, .markdown-text-container, .stText, .stMarkdown {
        color: white !important;
    }
    .stButton>button, .stDownloadButton>button {
        background-color: #0070C0 !important;
        color: white !important;
        border-radius: 8px;
        padding: 0.5em 1em;
    }
    .stFileUploader {
        background-color: #161b22;
        color: white;
        border-radius: 8px;
        padding: 0.5em;
    }
</style>
""", unsafe_allow_html=True)

# ======== HEADER ========
st.markdown("## 🧾 Rekap & Rename Bukti Potong Unifikasi")
st.markdown("*By: Reza Fahlevi Lubis BKP @zavibis*")

# ======== DESKRIPSI APLIKASI ========
st.markdown("""
Aplikasi ini digunakan untuk **membaca dan menamai ulang (rename)** file PDF **Bukti Potong Unifikasi**
berdasarkan metadata yang terdapat di dalam dokumen, seperti:
- Nomor Bukti Potong  
- Jenis PPh dan Kode Objek Pajak  
- Nama Pemotong dan Pihak Dipotong  
- Masa dan Tahun Pajak  
- Tanggal Pemotongan, Penandatangan, dan informasi lainnya  

⚠️ **Catatan penting:**
Aplikasi ini hanya dapat membaca **PDF hasil unduhan langsung dari sistem Coretax DJP**,  
**bukan hasil scan atau hasil foto.**  
Pastikan file berasal dari portal DJP (biasanya teks masih bisa diseleksi di PDF).
""")

# ======== PANDUAN PENGGUNAAN ========
st.markdown("### 📘 Cara Menggunakan Aplikasi")
st.markdown("""
1. Klik tombol **Browse files** untuk mengunggah satu atau beberapa file **PDF Bukti Potong Unifikasi** dari Coretax.  
2. Aplikasi akan otomatis membaca informasi penting dari setiap file (nama pemotong, masa pajak, kode objek pajak, dll).  
3. Setelah data tampil, pilih kolom mana saja yang ingin digunakan untuk format penamaan file.  
4. Klik tombol **Rename PDF & Download** untuk memproses semua file.  
5. File hasil rename akan digabungkan dalam satu file ZIP yang bisa langsung diunduh.
""")

# ======== DISCLAIMER ========
st.markdown("---")
st.markdown("### ⚠️ Disclaimer")
st.markdown("""
Aplikasi ini **tidak menyimpan, mengunggah, atau mengirimkan data apa pun ke server**.  
Semua proses dijalankan **sepenuhnya di sisi pengguna (local processing)**.

Aplikasi ini **bukan situs resmi Direktorat Jenderal Pajak (DJP)** dan tidak berafiliasi dengan otoritas pajak.  
Pengguna bertanggung jawab penuh atas **penggunaan, isi, dan keamanan dokumen** yang diproses melalui aplikasi ini.
""")

# ======== UTILITAS EKSTRAKSI ========
def extract_safe(text, pattern, group=1, default=""):
    match = re.search(pattern, text)
    return match.group(group).strip() if match else default

def smart_extract_dpp_tarif_pph(text):
    for line in text.splitlines():
        if re.search(r"\b\d{2}-\d{3}-\d{2}\b", line):
            numbers = re.findall(r"\d[\d.]*", line)
            if len(numbers) >= 6:
                try:
                    dpp = int(numbers[3].replace(".", ""))
                    tarif = int(numbers[4])
                    pph = int(numbers[5].replace(".", ""))
                    return dpp, tarif, pph
                except:
                    continue
    return 0, 0, 0

def extract_data_from_pdf(file_like):
    with pdfplumber.open(file_like) as pdf:
        text = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())

    try:
        data = {}
        data["Nomor Bukti Potong"] = extract_safe(text, r"\n(\S{9})\s+\d{2}-\d{4}")
        masa_pajak = extract_safe(text, r"\n\S{9}\s+(\d{2}-\d{4})")
        data["MASA PAJAK"] = masa_pajak
        if "-" in masa_pajak:
            data["MASA"], data["TAHUN"] = masa_pajak.split("-")
        else:
            data["MASA"], data["TAHUN"] = "", ""

        data["SIFAT PEMOTONGAN"] = extract_safe(text, r"(TIDAK FINAL|FINAL)")
        data["STATUS BUKTI"] = extract_safe(text, r"(NORMAL|PEMBETULAN)")

        data["NPWP / NIK PENERIMA PENGHASILAN"] = extract_safe(text, r"A\.1 NPWP / NIK\s*:\s*(\d+)")
        data["NAMA PENERIMA PENGHASILAN"] = extract_safe(text, r"A\.2 NAMA\s*:\s*(.+)")
        data["NOMOR IDENTITAS TEMPAT KEGIATAN USAHA"] = extract_safe(text, r"A\.3 NOMOR IDENTITAS.*?:\s*(\d+)")

        data["JENIS PPH"] = extract_safe(text, r"B\.2 Jenis PPh\s*:\s*(Pasal \d+)")
        data["KODE OBJEK PAJAK"] = extract_safe(text, r"(\d{2}-\d{3}-\d{2})")
        data["OBJEK PAJAK"] = extract_safe(text, r"\d{2}-\d{3}-\d{2}\s+([A-Za-z ]+)")
        dpp, tarif, pph = smart_extract_dpp_tarif_pph(text)
        data["DPP"] = dpp
        data["TARIF %"] = tarif
        data["PAJAK PENGHASILAN"] = pph

        data["JENIS DOKUMEN"] = extract_safe(text, r"Jenis Dokumen\s*:\s*(.+)")
        data["TANGGAL DOKUMEN"] = extract_safe(text, r"Tanggal\s*:\s*(\d{2} .+ \d{4})")
        data["NOMOR DOKUMEN"] = extract_safe(text, r"Nomor Dokumen\s*:\s*(.+)")

        data["NPWP / NIK PEMOTONG"] = extract_safe(text, r"C\.1 NPWP / NIK\s*:\s*(\d+)")
        data["NOMOR IDENTITAS TEMPAT USAHA PEMOTONG"] = extract_safe(text, r"C\.2.*?:\s*(\d+)")
        data["NAMA PEMOTONG"] = extract_safe(text, r"C\.3 NAMA PEMOTONG.*?:\s*(.+)")
        data["TANGGAL PEMOTONGAN"] = extract_safe(text, r"C\.4 TANGGAL\s*:\s*(\d{2} .+ \d{4})")
        data["PENANDATANGAN PEMOTONG"] = extract_safe(text, r"C\.5 NAMA PENANDATANGAN\s*:\s*(.+)")
        return data
    except Exception as e:
        st.warning(f"Gagal ekstrak data: {e}")
        return None

def sanitize_filename(text):
    return re.sub(r'[\\/*?:"<>|]', "_", str(text))

def generate_filename(row, selected_cols):
    parts = [sanitize_filename(str(row[col])) for col in selected_cols]
    return "Bukti Potong " + "_".join(parts) + ".pdf"

# ======== FILE UPLOAD SECTION ========
uploaded_files = st.file_uploader("📎 Upload PDF Bukti Potong (hasil unduhan Coretax)", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    data_rows = []
    for uploaded_file in uploaded_files:
        with st.spinner(f"📄 Membaca {uploaded_file.name}..."):
            pdf_bytes = uploaded_file.read()
            raw_data = extract_data_from_pdf(BytesIO(pdf_bytes))
            if raw_data:
                raw_data["OriginalName"] = uploaded_file.name
                raw_data["FileBytes"] = pdf_bytes
                data_rows.append(raw_data)

    if data_rows:
        df = pd.DataFrame(data_rows).drop(columns=["FileBytes", "OriginalName"])
        st.markdown("### ✅ Data berhasil diekstrak dari file berikut:")
        st.dataframe(df)

        column_options = df.columns.tolist()
        selected_columns = st.multiselect("### ✏️ Pilih Kolom untuk Format Nama File", column_options, default=[], key="formatselector")

        if st.button("🔁 Rename PDF & Download"):
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                for i, row in df.iterrows():
                    filename = generate_filename(row, selected_columns)
                    zipf.writestr(filename, data_rows[i]["FileBytes"])
            zip_buffer.seek(0)
            st.success("✅ Berhasil! Klik tombol di bawah ini untuk mengunduh file ZIP.")
            st.download_button("📦 Download ZIP Bukti Potong", zip_buffer, file_name="bukti_potong_renamed.zip", mime="application/zip")
