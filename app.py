import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

# =====================================
# 🏷️ Konfigurasi Streamlit
# =====================================
st.set_page_config(page_title="Rekap Bukti Potong PPh dari PDF ke Excel", layout="wide")
st.title("📄 Rekap Bukti Potong PPh dari PDF ke Excel")

# =====================================
# 📝 DESKRIPSI SINGKAT
# =====================================
st.markdown("""
Aplikasi ini berfungsi untuk mengekstrak **Bukti Potong PPh Unifikasi (Coretax DJP)** 
langsung dari file PDF resmi (bukan hasil scan).  
Hasil ekstraksi otomatis akan disajikan dalam bentuk tabel dan dapat diunduh menjadi file **Excel (.xlsx)**.
""")

# =====================================
# 📘 PETUNJUK PENGGUNAAN
# =====================================
st.markdown("""
### 📘 Cara Menggunakan
1. Upload satu atau lebih file **PDF Bukti Potong Unifikasi** yang diunduh dari Coretax.
2. Sistem akan membaca isi PDF dan mengekstrak data penting seperti:  
   - Nomor Bukti  
   - Masa Pajak  
   - Status Bukti (**NORMAL / PEMBETULAN / DIBATALKAN**)  
   - Identitas Pihak Dipotong & Pemotong  
   - Jenis PPh, Kode Objek Pajak, Objek Pajak  
   - DPP, Tarif, dan PPh Dipotong  
3. Lihat hasil ekstraksi pada tabel di bawah.
4. Klik **Unduh Excel** untuk menyimpan hasil rekap.
""")

# =====================================
# ⚠️ DISCLAIMER
# =====================================
st.markdown("""
### ⚠️ Disclaimer
- Semua proses dilakukan **sepenuhnya di perangkat Anda (local processing)**.  
- Aplikasi ini **tidak mengirim atau menyimpan dokumen** ke server mana pun.  
- Aplikasi ini **bukan aplikasi resmi DJP** dan tidak memiliki afiliasi dengan otoritas pajak.  
""")

# =====================================
# 🔍 Fungsi bantu regex aman
# =====================================
def extract_safe(text, pattern, group=1, default=""):
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(group).strip() if match else default

# =====================================
# 💡 Ekstraksi DPP, Tarif, dan PPh
# =====================================
def smart_extract_dpp_tarif_pph(text):
    """Ekstrak nilai DPP, tarif (%), dan PPh"""
    for line in text.splitlines():
        if re.search(r"\b\d{2}-\d{3}-\d{2}\b", line):
            numbers = re.findall(r"\d[\d.,]*", line)
            if len(numbers) >= 6:
                try:
                    dpp = float(numbers[3].replace(".", "").replace(",", ""))
                    tarif = float(numbers[4].replace(",", "."))
                    pph = float(numbers[5].replace(".", "").replace(",", ""))
                    return dpp, tarif, pph
                except:
                    continue
    return 0, 0, 0

# =====================================
# 🔎 Ekstraksi OBJEK PAJAK multi-baris
# =====================================
def extract_objek_pajak(text):
    joined = " ".join(text.splitlines())
    match = re.search(r"\b\d{2}-\d{3}-\d{2}\b\s+(.+?)B\.8", joined, re.DOTALL | re.IGNORECASE)
    if not match:
        return ""
    objek = match.group(1)
    objek = re.sub(r"\d[\d.,\s]+(?=[A-Za-z])", "", objek)
    objek = re.sub(r"[\d.,]+", "", objek)
    objek = re.sub(r"Rp", "", objek, flags=re.IGNORECASE)
    objek = re.sub(r"\s+", " ", objek).strip()
    return objek

# =====================================
# 📄 Ekstraksi Data dari PDF
# =====================================
def extract_data_from_pdf(file):
    with pdfplumber.open(file) as pdf:
        text = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())

    try:
        data = {}

        # =========================
        # FIX STATUS BUKTI ✓
        # =========================
        data["STATUS BUKTI"] = extract_safe(
            text,
            r"(NORMAL|DIBATALKAN|PEMBETULAN(?: KE-?\d+)?)",
            group=1,
            default=""
        ).upper()

        # HEADER
        data["NOMOR"] = extract_safe(text, r"\n(\S{9})\s+\d{2}-\d{4}")
        data["MASA PAJAK"] = extract_safe(text, r"\n\S{9}\s+(\d{2}-\d{4})")
        data["SIFAT PEMOTONGAN"] = extract_safe(text, r"(TIDAK FINAL|FINAL)")

        # A. IDENTITAS
        data["NPWP / NIK"] = extract_safe(text, r"A\.1 NPWP / NIK\s*:\s*(\d+)")
        data["NAMA"] = extract_safe(text, r"A\.2 NAMA\s*:\s*(.+)")
        data["NOMOR IDENTITAS TEMPAT USAHA"] = extract_safe(text, r"A\.3.*?:\s*(\d+)")

        # B. PEMOTONGAN
        data["JENIS FASILITAS"] = extract_safe(text, r"B\.1\s*Jenis Fasilitas\s*:\s*(.+)")
        data["JENIS PPH"] = extract_safe(text, r"B\.2\s*Jenis PPh\s*:\s*(Pasal\s*\d+)")
        data["KODE OBJEK"] = extract_safe(text, r"(\d{2}-\d{3}-\d{2})")
        data["OBJEK PAJAK"] = extract_objek_pajak(text)
        data["DPP"], data["TARIF %"], data["PAJAK PENGHASILAN"] = smart_extract_dpp_tarif_pph(text)

        # Dokumen dasar
        data["JENIS DOKUMEN"] = extract_safe(text, r"Jenis Dokumen\s*:\s*(.+)")
        data["TANGGAL DOKUMEN"] = extract_safe(text, r"Tanggal\s*:\s*(\d{1,2} .+ \d{4})")
        data["NOMOR DOKUMEN"] = extract_safe(text, r"Nomor Dokumen\s*:\s*(.+)")

        # B.10 dan B.11
        b10_raw = extract_safe(text, r"B\.10\s*Untuk Instansi Pemerintah.*:\s*(.+)")
        if re.search(r"B\.10\s*Untuk Instansi Pemerintah.*:\s*(B\.11|C\.)", text, re.IGNORECASE):
            b10_raw = "-"
        data["UNTUK INSTANSI PEMERINTAH"] = b10_raw

        sp2d_raw = extract_safe(text, r"B\.11\s*Nomor SP2D\s*:\s*(.+)")
        if re.search(r"B\.11\s*Nomor SP2D\s*:\s*C\. IDENTITAS PEMOTONG", text, re.IGNORECASE):
            sp2d_raw = "-"
        data["NOMOR SP2D"] = sp2d_raw

        # C. IDENTITAS PEMOTONG
        data["NPWP / NIK PEMOTONG"] = extract_safe(text, r"C\.1 NPWP / NIK\s*:\s*(\d+)")
        data["NOMOR IDENTITAS TEMPAT USAHA PEMOTONG"] = extract_safe(text, r"C\.2.*?:\s*(\d+)")
        data["NAMA PEMOTONG"] = extract_safe(text, r"C\.3.*?:\s*(.+)")
        data["TANGGAL PEMOTONGAN"] = extract_safe(text, r"C\.4 TANGGAL\s*:\s*(\d{1,2} .+ \d{4})")
        data["NAMA PENANDATANGAN"] = extract_safe(text, r"C\.5 NAMA PENANDATANGAN\s*:\s*(.+)")

        return data

    except Exception as e:
        st.warning(f"Gagal ekstrak data: {e}")
        return None

# =====================================
# 🚀 Streamlit UI
# =====================================
uploaded_files = st.file_uploader(
    "📤 Upload satu atau lebih file PDF Bukti Potong",
    type="pdf",
    accept_multiple_files=True,
)

if uploaded_files:
    all_data = []
    for file in uploaded_files:
        with st.spinner(f"🔎 Memproses {file.name}..."):
            result = extract_data_from_pdf(file)
            if result:
                result["FILE"] = file.name
                all_data.append(result)

    if all_data:
        df = pd.DataFrame(all_data)
        st.success(f"✅ Berhasil mengekstrak {len(df)} bukti potong")
        st.dataframe(df, use_container_width=True)

        output = BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)
        st.download_button(
            "⬇️ Unduh Excel",
            output,
            file_name="Rekap_Bukti_Potong.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.error("❌ Tidak ada data berhasil diproses.")
else:
    st.info("Silakan upload satu atau beberapa file PDF Bukti Potong untuk mulai ekstraksi.")
