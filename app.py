# app.py - Revisi ke-202507211915-5 (Final Fix)
# Creator: Reza Fahlevi Lubis BKP @zavibis
# Aplikasi Rekap Bukti Potong Unifikasi (PDF Coretax Only)
# - FIX: Status bukti (NORMAL / PEMBETULAN / DIBATALKAN) sekarang terbaca akurat
# - Tema gelap biru profesional
# - Output Excel

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
st.markdown("## 📊 Rekap Bukti Potong Unifikasi — PDF Coretax")
st.markdown("*By: **Reza Fahlevi Lubis BKP @zavibis***")

# ======== DESKRIPSI ========
st.markdown("""
Aplikasi ini digunakan untuk **mengekstrak data Bukti Potong Unifikasi** (format resmi Coretax DJP)
dan menghasilkan **file Excel (.xlsx)** secara otomatis.

Informasi yang diekstrak meliputi:
- Nomor Bukti Potong  
- Masa & Tahun Pajak  
- Status Bukti: **NORMAL, PEMBETULAN, atau DIBATALKAN**  
- Jenis PPh & Kode Objek Pajak  
- DPP, Tarif, PPh Dipotong  
- Identitas Pemotong & Pihak Dipotong  
- Tanggal Dokumen & Tanggal Pemotongan  
- Penandatangan  

⚠️ Khusus untuk **PDF asli dari Coretax** (bukan scan/foto).
""")

# ======== CARA PAKAI ========
st.markdown("### 📘 Cara Menggunakan")
st.markdown("""
1. Upload 1 atau banyak file Bukti Potong Unifikasi (PDF Coretax).  
2. Sistem akan membaca dan mengekstrak semua data otomatis.  
3. Lihat hasilnya pada tabel di bawah.  
4. Klik **Download Excel Rekap** untuk menyimpan hasilnya.
""")

# ======== DISCLAIMER ========
st.markdown("---")
st.markdown("### ⚠️ Disclaimer")
st.markdown("""
Semua proses dilakukan **sepenuhnya di perangkat Anda** (local).  
Tidak ada data yang dikirim atau disimpan ke server mana pun.  
Ini bukan aplikasi resmi DJP.
""")

# ------------------------
# Helpers
# ------------------------

def extract_safe(text, pattern, group=1, default=""):
    match = re.search(pattern, text, flags=re.DOTALL)
    return match.group(group).strip() if match else default

def extract_status(text):
    """
    Contoh header Coretax:
    NOMOR   MASA    SIFAT    STATUS
    2500ZK0YY 03-2025 TIDAK FINAL DIBATALKAN
    """
    # Ambil baris pertama saja
    first_lines = text.splitlines()[0:4]
    block = " ".join(first_lines)

    # POLA UMUM UNTUK STATUS
    # NORMAL
    # PEMBETULAN
    # PEMBETULAN KE-X
    # DIBATALKAN
    status = re.search(r"(NORMAL|DIBATALKAN|PEMBETULAN(?: KE-?\d+)?)", block, re.IGNORECASE)
    return status.group(1).upper() if status else ""

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

# ------------------------
# PDF PARSER
# ------------------------
def extract_data_from_pdf(file_like):
    with pdfplumber.open(file_like) as pdf:
        text = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())

    try:
        data = {}

        # ===== HEADER =====
        data["Nomor Bukti Potong"] = extract_safe(text, r"\n([A-Z0-9]{9})\s+\d{2}-\d{4}")
        masa_pajak = extract_safe(text, r"[A-Z0-9]{9}\s+(\d{2}-\d{4})")
        data["Masa Pajak"] = masa_pajak
        if "-" in masa_pajak:
            data["Masa"], data["Tahun"] = masa_pajak.split("-")
        else:
            data["Masa"], data["Tahun"] = "", ""

        data["Sifat Pemotongan"] = extract_safe(text, r"(FINAL|TIDAK FINAL)")
        data["Status Bukti"] = extract_status(text)

        # ===== IDENTITAS DIPOTONG =====
        data["NPWP / NIK Pihak Dipotong"] = extract_safe(text, r"A\.1 NPWP / NIK\s*:\s*(\d+)")
        data["Nama Pihak Dipotong"] = extract_safe(text, r"A\.2 NAMA\s*:\s*(.+)")
        data["NITKU Pihak Dipotong"] = extract_safe(text, r"A\.3.*?:\s*(\d+)")

        # ===== PENGHASILAN =====
        data["Jenis PPh"] = extract_safe(text, r"B\.2 Jenis PPh\s*:\s*(Pasal \d+)")
        data["Kode Objek Pajak"] = extract_safe(text, r"(\d{2}-\d{3}-\d{2})")
        data["Objek Pajak"] = extract_safe(text, r"\d{2}-\d{3}-\d{2}\s+(.+)")

        dpp, tarif, pph = smart_extract_dpp_tarif_pph(text)
        data["DPP (Rp)"] = dpp
        data["Tarif (%)"] = tarif
        data["PPh Dipotong (Rp)"] = pph

        # ===== DOKUMEN =====
        data["Jenis Dokumen"] = extract_safe(text, r"Jenis Dokumen\s*:\s*(.+)")
        data["Tanggal Dokumen"] = extract_safe(text, r"Tanggal\s*:\s*(\d{2} .+ \d{4})")
        data["Nomor Dokumen"] = extract_safe(text, r"Nomor Dokumen\s*:\s*(.+)")

        # ===== PEMOTONG =====
        data["NPWP / NIK Pemotong"] = extract_safe(text, r"C\.1 NPWP / NIK\s*:\s*(\d+)")
        data["NITKU Pemotong"] = extract_safe(text, r"C\.2.*?:\s*(\d+)")
        data["Nama Pemotong"] = extract_safe(text, r"C\.3 NAMA PEMOTONG.*?:\s*(.+)")
        data["Tanggal Pemotongan"] = extract_safe(text, r"C\.4 TANGGAL\s*:\s*(\d{2} .+ \d{4})")
        data["Penandatangan"] = extract_safe(text, r"C\.5 NAMA PENANDATANGAN\s*:\s*(.+)")

        return data

    except Exception as e:
        st.warning(f"⚠️ Gagal ekstrak data: {e}")
        return None

# ------------------------
# UI — Upload
# ------------------------
uploaded_files = st.file_uploader(
    "📎 Upload PDF Bukti Potong Unifikasi (PDF Coretax, bukan scan)",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    all_rows = []

    for f in uploaded_files:
        pdf_bytes = f.read()
        result = extract_data_from_pdf(BytesIO(pdf_bytes))
        if result:
            result["Nama File Asli"] = f.name
            all_rows.append(result)

    if all_rows:
        df = pd.DataFrame(all_rows)

        st.markdown("### ✅ Berikut data yang berhasil diekstrak:")
        st.dataframe(df, use_container_width=True)

        # EXPORT
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name="Rekap", index=False)
        output.seek(0)

        st.download_button(
            "💾 Download Excel Rekap",
            output,
            file_name="rekap_bukti_potong_unifikasi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
