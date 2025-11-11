# app.py - Revisi ke-202507211915-4
# ✅ Aplikasi Rekap Bukti Potong Unifikasi (bukan rename)
# ✅ Output ke Excel (.xlsx)
# ✅ Dilengkapi deskripsi, cara pakai, catatan Coretax, dan disclaimer
# ✅ Tema gelap profesional biru

import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Rekap Bukti Potong Unifikasi (Coretax PDF Only)", layout="wide")

# ======== THEME ========
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
        font-weight: 600;
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
st.markdown("## 📊 Rekap Bukti Potong Unifikasi (PDF Coretax Only)")
st.markdown("*By: Reza Fahlevi Lubis BKP @zavibis*")

# ======== DESKRIPSI APLIKASI ========
st.markdown("""
Aplikasi ini berfungsi untuk **membaca, mengekstrak, dan merekap data Bukti Potong Unifikasi**
langsung dari file **PDF hasil unduhan Coretax DJP**, lalu menyimpannya ke format **Excel (.xlsx)**.

Data yang dibaca antara lain:
- Nomor Bukti Potong, Masa & Tahun Pajak  
- Jenis PPh dan Kode Objek Pajak  
- DPP, Tarif (%), dan PPh Dipotong  
- Nama & NPWP Pemotong dan Pihak Dipotong  
- Tanggal Dokumen, Tanggal Pemotongan, dan Penandatangan

⚠️ **Catatan penting:**  
Aplikasi ini hanya bisa membaca **PDF hasil unduhan resmi dari Coretax DJP**,  
**bukan hasil scan atau hasil foto.**  
Pastikan teks PDF masih dapat diseleksi (copy-pasteable).
""")

# ======== CARA PAKAI ========
st.markdown("### 📘 Cara Menggunakan Aplikasi")
st.markdown("""
1. Klik **Browse files** untuk mengunggah satu atau beberapa file **Bukti Potong Unifikasi (PDF)** dari Coretax.  
2. Aplikasi akan menampilkan tabel hasil ekstraksi otomatis dari setiap file.  
3. Periksa hasil rekap di tampilan bawah.  
4. Klik tombol **💾 Download Excel Rekap** untuk mengunduh hasil dalam format `.xlsx`.
""")

# ======== DISCLAIMER ========
st.markdown("---")
st.markdown("### ⚠️ Disclaimer")
st.markdown("""
Aplikasi ini **tidak menyimpan, mengunggah, atau mengirimkan data ke server mana pun.**  
Semua proses dijalankan **sepenuhnya di perangkat Anda (local processing).**

Aplikasi ini **bukan situs resmi Direktorat Jenderal Pajak (DJP)**  
dan tidak berafiliasi dengan otoritas pajak mana pun.  
Pengguna bertanggung jawab penuh atas **penggunaan dan keamanan dokumen** yang diproses.
""")

# ======== FUNGSI PENDUKUNG ========
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
        data["Masa Pajak"] = masa_pajak
        if "-" in masa_pajak:
            data["Masa"], data["Tahun"] = masa_pajak.split("-")
        else:
            data["Masa"], data["Tahun"] = "", ""

        data["Sifat Pemotongan"] = extract_safe(text, r"(TIDAK FINAL|FINAL)")
        data["Status Bukti"] = extract_safe(text, r"(NORMAL|PEMBETULAN)")

        data["NPWP / NIK Pihak Dipotong"] = extract_safe(text, r"A\.1 NPWP / NIK\s*:\s*(\d+)")
        data["Nama Pihak Dipotong"] = extract_safe(text, r"A\.2 NAMA\s*:\s*(.+)")
        data["NITKU Pihak Dipotong"] = extract_safe(text, r"A\.3 NOMOR IDENTITAS.*?:\s*(\d+)")

        data["Jenis PPh"] = extract_safe(text, r"B\.2 Jenis PPh\s*:\s*(Pasal \d+)")
        data["Kode Objek Pajak"] = extract_safe(text, r"(\d{2}-\d{3}-\d{2})")
        data["Objek Pajak"] = extract_safe(text, r"\d{2}-\d{3}-\d{2}\s+([A-Za-z ]+)")
        dpp, tarif, pph = smart_extract_dpp_tarif_pph(text)
        data["DPP (Rp)"] = dpp
        data["Tarif (%)"] = tarif
        data["PPh Dipotong (Rp)"] = pph

        data["Jenis Dokumen"] = extract_safe(text, r"Jenis Dokumen\s*:\s*(.+)")
        data["Tanggal Dokumen"] = extract_safe(text, r"Tanggal\s*:\s*(\d{2} .+ \d{4})")
        data["Nomor Dokumen"] = extract_safe(text, r"Nomor Dokumen\s*:\s*(.+)")

        data["NPWP / NIK Pemotong"] = extract_safe(text, r"C\.1 NPWP / NIK\s*:\s*(\d+)")
        data["NITKU Pemotong"] = extract_safe(text, r"C\.2.*?:\s*(\d+)")
        data["Nama Pemotong"] = extract_safe(text, r"C\.3 NAMA PEMOTONG.*?:\s*(.+)")
        data["Tanggal Pemotongan"] = extract_safe(text, r"C\.4 TANGGAL\s*:\s*(\d{2} .+ \d{4})")
        data["Penandatangan"] = extract_safe(text, r"C\.5 NAMA PENANDATANGAN\s*:\s*(.+)")

        return data
    except Exception as e:
        st.warning(f"Gagal ekstrak data: {e}")
        return None

# ======== UNGGAH FILE ========
uploaded_files = st.file_uploader("📎 Upload PDF Bukti Potong (hasil unduhan Coretax, bukan hasil scan)", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    data_rows = []
    for uploaded_file in uploaded_files:
        with st.spinner(f"📄 Membaca {uploaded_file.name}..."):
            pdf_bytes = uploaded_file.read()
            result = extract_data_from_pdf(BytesIO(pdf_bytes))
            if result:
                result["Nama File Asli"] = uploaded_file.name
                data_rows.append(result)

    if data_rows:
        df = pd.DataFrame(data_rows)
        st.markdown("### ✅ Berikut data yang berhasil diekstrak:")
        st.dataframe(df, use_container_width=True)

        # ======== DOWNLOAD EXCEL ========
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name="Rekap Bukti Potong")
        output.seek(0)

        st.download_button(
            label="💾 Download Excel Rekap",
            data=output,
            file_name="rekap_bukti_potong_unifikasi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
