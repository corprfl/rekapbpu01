import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

# =========================
# 🏷️ Konfigurasi Streamlit
# =========================
st.set_page_config(page_title="Rekap Bukti Potong PPh dari PDF ke Excel", layout="wide")
st.title("📄 Rekap Bukti Potong PPh dari PDF ke Excel")

# =========================
# 🔍 Fungsi bantu regex aman
# =========================
def extract_safe(text, pattern, group=1, default=""):
    """Ekstraksi teks dengan regex aman"""
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(group).strip() if match else default

# =========================
# 💡 Ekstraksi DPP, Tarif, dan PPh
# =========================
def smart_extract_dpp_tarif_pph(text):
    """Membaca nilai DPP, tarif %, dan pajak penghasilan"""
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

# =========================
# 🔎 Ekstraksi OBJEK PAJAK panjang
# =========================
def extract_objek_pajak(text):
    """
    Menangkap seluruh kalimat objek pajak multi-baris,
    berhenti sebelum nilai angka (DPP).
    """
    pattern = r"\d{2}-\d{3}-\d{2}\s+([A-Za-z0-9/()., -]+(?:\n[A-Za-z0-9/()., -]+)*)"
    match = re.search(pattern, text)
    if match:
        obj = match.group(1)
        obj = re.sub(r"\s*\n\s*", " ", obj).strip()
        obj = re.split(r"\s\d[\d.,]+", obj)[0].strip()
        return obj
    return ""

# =========================
# 📄 Ekstraksi Data dari PDF
# =========================
def extract_data_from_pdf(file):
    with pdfplumber.open(file) as pdf:
        text = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())

    try:
        data = {}

        # Header
        data["NOMOR"] = extract_safe(text, r"\n(\S{9})\s+\d{2}-\d{4}")
        data["MASA PAJAK"] = extract_safe(text, r"\n\S{9}\s+(\d{2}-\d{4})")
        data["SIFAT PEMOTONGAN"] = extract_safe(text, r"(TIDAK FINAL|FINAL)")
        data["STATUS BUKTI"] = extract_safe(text, r"(NORMAL|PEMBETULAN)")

        # A. IDENTITAS
        data["NPWP / NIK"] = extract_safe(text, r"A\.1 NPWP / NIK\s*:\s*(\d+)")
        data["NAMA"] = extract_safe(text, r"A\.2 NAMA\s*:\s*(.+)")
        data["NOMOR IDENTITAS TEMPAT USAHA"] = extract_safe(text, r"A\.3.*?:\s*(\d+)")

        # B. PEMOTONGAN
        data["JENIS FASILITAS"] = extract_safe(text, r"B\.1\s*Jenis Fasilitas\s*:\s*(.+)")
        data["JENIS PPH"] = extract_safe(text, r"B\.2\s*Jenis PPh\s*:\s*(Pasal\s*\d+)")

        # Ambil kode objek & objek pajak multi-baris
        data["KODE OBJEK"] = extract_safe(text, r"(\d{2}-\d{3}-\d{2})")
        data["OBJEK PAJAK"] = extract_objek_pajak(text)

        data["DPP"], data["TARIF %"], data["PAJAK PENGHASILAN"] = smart_extract_dpp_tarif_pph(text)

        # Dokumen
        data["JENIS DOKUMEN"] = extract_safe(text, r"Jenis Dokumen\s*:\s*(.+)")
        data["TANGGAL DOKUMEN"] = extract_safe(text, r"Tanggal\s*:\s*(\d{1,2} .+ \d{4})")
        data["NOMOR DOKUMEN"] = extract_safe(text, r"Nomor Dokumen\s*:\s*(.+)")

        # =========================
        # Tambahan kolom: B.10 dan B.11
        # =========================
        b10_raw = extract_safe(text, r"B\.10\s*Untuk Instansi Pemerintah.*:\s*(.+)")
        # Jika setelah B.10 langsung B.11 → kosongkan
        if re.search(r"B\.10\s*Untuk Instansi Pemerintah.*:\s*B\.11", text, re.IGNORECASE):
            b10_raw = "-"
        data["UNTUK INSTANSI PEMERINTAH"] = b10_raw

        sp2d_raw = extract_safe(text, r"B\.11\s*Nomor SP2D\s*:\s*(.+)")
        # Jika langsung heading "C."
        if re.search(r"B\.11\s*Nomor SP2D\s*:\s*C\. IDENTITAS PEMOTONG", text, re.IGNORECASE):
            sp2d_raw = "-"
        data["NOMOR SP2D"] = sp2d_raw

        # =========================
        # C. IDENTITAS PEMOTONG
        # =========================
        data["NPWP / NIK PEMOTONG"] = extract_safe(text, r"C\.1 NPWP / NIK\s*:\s*(\d+)")
        data["NOMOR IDENTITAS TEMPAT USAHA PEMOTONG"] = extract_safe(text, r"C\.2.*?:\s*(\d+)")
        data["NAMA PEMOTONG"] = extract_safe(text, r"C\.3.*?:\s*(.+)")
        data["TANGGAL PEMOTONGAN"] = extract_safe(text, r"C\.4 TANGGAL\s*:\s*(\d{1,2} .+ \d{4})")
        data["NAMA PENANDATANGAN"] = extract_safe(text, r"C\.5 NAMA PENANDATANGAN\s*:\s*(.+)")

        return data
    except Exception as e:
        st.warning(f"Gagal ekstrak data: {e}")
        return None

# =========================
# 📥 Upload & Proses PDF
# =========================
uploaded_files = st.file_uploader(
    "📤 Upload satu atau lebih file PDF Bukti Potong",
    type="pdf",
    accept_multiple_files=True
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
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("❌ Tidak ada data berhasil diproses.")
else:
    st.info("Silakan upload satu atau beberapa file PDF Bukti Potong untuk mulai ekstraksi.")
