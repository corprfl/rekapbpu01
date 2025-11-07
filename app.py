import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Rekap Bukti Potong PDF", layout="wide")
st.title("📄 Rekap Bukti Potong PPh dari PDF ke Excel")

def extract_safe(text, pattern, group=1, default=""):
    match = re.search(pattern, text)
    return match.group(group).strip() if match else default

def smart_extract_dpp_tarif_pph(text):
    for line in text.splitlines():
        if re.search(r"\b\d{2}-\d{3}-\d{2}\b", line):
            numbers = re.findall(r"\d[\d.]*", line)
            if len(numbers) >= 6:
                try:
                    # Lewati 3 angka pertama (kode objek: 24, 104, 07)
                    dpp = int(numbers[3].replace(".", ""))
                    tarif = int(numbers[4])
                    pph = int(numbers[5].replace(".", ""))
                    return dpp, tarif, pph
                except:
                    continue
    return 0, 0, 0

def extract_data_from_pdf(file):
    with pdfplumber.open(file) as pdf:
        text = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())

    try:
        data = {}
        data["NOMOR"] = extract_safe(text, r"\n(\S{9})\s+\d{2}-\d{4}")
        data["MASA PAJAK"] = extract_safe(text, r"\n\S{9}\s+(\d{2}-\d{4})")
        data["SIFAT PEMOTONGAN"] = extract_safe(text, r"(TIDAK FINAL|FINAL)")
        data["STATUS BUKTI"] = extract_safe(text, r"(NORMAL|PEMBETULAN)")

        data["NPWP / NIK"] = extract_safe(text, r"A\.1 NPWP / NIK\s*:\s*(\d+)")
        data["NAMA"] = extract_safe(text, r"A\.2 NAMA\s*:\s*(.+)")
        data["NOMOR IDENTITAS TEMPAT USAHA"] = extract_safe(text, r"A\.3 NOMOR IDENTITAS.*?:\s*(\d+)")

        data["JENIS PPH"] = extract_safe(text, r"B\.2 Jenis PPh\s*:\s*(Pasal \d+)")
        data["KODE OBJEK"] = extract_safe(text, r"(\d{2}-\d{3}-\d{2})")
        data["OBJEK PAJAK"] = extract_safe(text, r"\d{2}-\d{3}-\d{2}\s+([A-Za-z ]+)")
        data["DPP"], data["TARIF %"], data["PAJAK PENGHASILAN"] = smart_extract_dpp_tarif_pph(text)

        data["JENIS DOKUMEN"] = extract_safe(text, r"Jenis Dokumen\s*:\s*(.+)")
        data["TANGGAL DOKUMEN"] = extract_safe(text, r"Tanggal\s*:\s*(\d{2} .+ \d{4})")
        data["NOMOR DOKUMEN"] = extract_safe(text, r"Nomor Dokumen\s*:\s*(.+)")

        data["NPWP / NIK PEMOTONG"] = extract_safe(text, r"C\.1 NPWP / NIK\s*:\s*(\d+)")
        data["NOMOR IDENTITAS TEMPAT USAHA PEMOTONG"] = extract_safe(text, r"C\.2.*?:\s*(\d+)")
        data["NAMA PEMOTONG"] = extract_safe(text, r"C\.3.*?:\s*(.+)")
        data["TANGGAL PEMOTONGAN"] = extract_safe(text, r"C\.4 TANGGAL\s*:\s*(\d{2} .+ \d{4})")
        data["NAMA PENANDATANGAN"] = extract_safe(text, r"C\.5 NAMA PENANDATANGAN\s*:\s*(.+)")
        return data
    except Exception as e:
        st.warning(f"Gagal ekstrak data: {e}")
        return None

uploaded_files = st.file_uploader("Upload satu atau lebih file PDF", type="pdf", accept_multiple_files=True)

if uploaded_files:
    all_data = []
    for file in uploaded_files:
        with st.spinner(f"Memproses {file.name}..."):
            result = extract_data_from_pdf(file)
            if result:
                result["FILE"] = file.name
                all_data.append(result)

    if all_data:
        df = pd.DataFrame(all_data)
        st.success(f"Berhasil mengekstrak {len(df)} bukti potong")
        st.dataframe(df)

        output = BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)
        st.download_button("⬇️ Unduh Excel", output, file_name="Rekap_Bukti_Potong.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.error("Tidak ada data berhasil diproses.")
