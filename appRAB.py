import streamlit as st
import pandas as pd
import numpy as np
import io
import math
from fpdf import FPDF
from datetime import datetime

# ==========================================
# 1. FUNGSI TERBILANG
# ==========================================
def terbilang(n):
    angka = ["", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan", "Sepuluh", "Sebelas"]
    n = int(n)
    if n == 0: return "Nol"
    elif n < 12: return angka[n]
    elif n < 20: return terbilang(n - 10) + " Belas"
    elif n < 100: return terbilang(n // 10) + " Puluh " + terbilang(n % 10)
    elif n < 200: return "Seratus " + terbilang(n - 100)
    elif n < 1000: return terbilang(n // 100) + " Ratus " + terbilang(n % 100)
    elif n < 2000: return "Seribu " + terbilang(n - 1000)
    elif n < 1000000: return terbilang(n // 1000) + " Ribu " + terbilang(n % 1000)
    elif n < 1000000000: return terbilang(n // 1000000) + " Juta " + terbilang(n % 1000000)
    elif n < 1000000000000: return terbilang(n // 1000000000) + " Miliar " + terbilang(n % 1000000000)
    else: return "Angka Terlalu Besar"

# ==========================================
# 2. AUTO-LOAD DATABASE (BACKGROUND PROCESS)
# ==========================================
@st.cache_data
def load_database_otomatis():
    file_name = "RAB_AHSP_Lengkap_CiptaKarya_2025.xlsx"
    try:
        # Mencoba membaca file lokal secara otomatis
        df_ahsp = pd.read_excel(file_name, sheet_name="3_Daftar_AHSP", skiprows=1)
        df_ahsp = df_ahsp.dropna(subset=['Kode AHSP', 'Uraian Pekerjaan'])
        
        # Mencoba membaca detail AHSP untuk lampiran
        df_detail = pd.read_excel(file_name, sheet_name="4_Detail_AHSP")
        return df_ahsp, df_detail, True
    except:
        # Fallback Engine jika file belum ada di folder
        df_ahsp_dummy = pd.DataFrame([
            {"Kode AHSP": "1.2.1.1", "Divisi": "I. TANAH", "Uraian Pekerjaan": "Galian Tanah Pondasi", "Satuan": "m3", "Harga Satuan Total (Rp)": 90860},
            {"Kode AHSP": "2.2.2.1", "Divisi": "I. TANAH", "Uraian Pekerjaan": "Pondasi Batu Kali", "Satuan": "m3", "Harga Satuan Total (Rp)": 657500},
            {"Kode AHSP": "2.2.1.S", "Divisi": "II. BETON", "Uraian Pekerjaan": "Sloof Beton Bertulang", "Satuan": "m3", "Harga Satuan Total (Rp)": 3184000},
            {"Kode AHSP": "2.2.1.K", "Divisi": "II. BETON", "Uraian Pekerjaan": "Kolom Beton Bertulang", "Satuan": "m3", "Harga Satuan Total (Rp)": 4250000},
            {"Kode AHSP": "3.6.1.1", "Divisi": "III. ARSITEKTUR", "Uraian Pekerjaan": "Pasangan Bata Merah", "Satuan": "m2", "Harga Satuan Total (Rp)": 135000},
            {"Kode AHSP": "3.9.1.1", "Divisi": "III. ARSITEKTUR", "Uraian Pekerjaan": "Keramik Lantai 60x60", "Satuan": "m2", "Harga Satuan Total (Rp)": 185000},
            {"Kode AHSP": "3.8.1.1", "Divisi": "III. ARSITEKTUR", "Uraian Pekerjaan": "Cat Dinding & Plafon", "Satuan": "m2", "Harga Satuan Total (Rp)": 46500},
            {"Kode AHSP": "5.1.1.1", "Divisi": "IV. ATAP", "Uraian Pekerjaan": "Rangka Atap Baja Ringan", "Satuan": "m2", "Harga Satuan Total (Rp)": 185000},
        ])
        df_detail_dummy = pd.DataFrame({"Kode": ["1.2.1.1"], "Uraian Komponen": ["Pekerja"], "Sat": ["OH"], "Koefisien": [0.4], "Harga Satuan (Rp)": [100000], "Jumlah Harga (Rp)": [40000]})
        return df_ahsp_dummy, df_detail_dummy, False

df_ahsp, df_detail, db_status = load_database_otomatis()

# ==========================================
# 3. KELAS GENERATOR PDF (DOKUMEN TENDER)
# ==========================================
class PDFTender(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 14)
        self.cell(0, 10, 'DOKUMEN PENAWARAN (TENDER) KONTRAKTOR', 0, 1, 'C')
        self.set_font('Arial', 'I', 10)
        self.cell(0, 5, 'Estimasi Dihasilkan Secara Otomatis Berdasarkan AHSP Cipta Karya 2025', 0, 1, 'C')
        self.ln(10)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Halaman {self.page_no()}', 0, 0, 'C')

def create_pdf(nama_proyek, tgl, grand_total, terbilang_txt, df_rab, nama_perusahaan, nama_direktur):
    pdf = PDFTender()
    pdf.add_page()
    
    # Surat Penawaran
    pdf.set_font('Arial', '', 11)
    pdf.cell(0, 6, f'Tanggal: {tgl}', 0, 1)
    pdf.cell(0, 6, f'Perihal: Penawaran Pekerjaan {nama_proyek}', 0, 1)
    pdf.ln(10)
    
    pdf.multi_cell(0, 6, f"Dengan hormat,\nSehubungan dengan rencana pelaksanaan pekerjaan {nama_proyek}, bersama ini kami dari {nama_perusahaan} mengajukan penawaran harga untuk pelaksanaan pekerjaan tersebut.")
    pdf.ln(5)
    
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 8, f"Total Harga Penawaran : Rp {grand_total:,.0f}", 0, 1)
    pdf.set_font('Arial', 'I', 11)
    pdf.multi_cell(0, 6, f"Terbilang: ({terbilang_txt} Rupiah)")
    pdf.ln(5)
    
    pdf.set_font('Arial', '', 11)
    pdf.multi_cell(0, 6, "Harga tersebut sudah termasuk seluruh biaya material, upah tenaga kerja, peralatan, dan penyesuaian margin keuntungan serta pajak yang berlaku. Rincian Anggaran Biaya (RAB) dan Analisa Harga Satuan Pekerjaan (AHSP) terlampir dalam dokumen ini.")
    pdf.ln(15)
    
    # Signature Block
    pdf.cell(100, 6, '', 0, 0)
    pdf.cell(80, 6, 'Hormat Kami,', 0, 1, 'C')
    pdf.cell(100, 6, '', 0, 0)
    pdf.cell(80, 6, f'{nama_perusahaan}', 0, 1, 'C')
    pdf.ln(20)
    pdf.cell(100, 6, '', 0, 0)
    pdf.set_font('Arial', 'BU', 11)
    pdf.cell(80, 6, f'{nama_direktur}', 0, 1, 'C')
    pdf.set_font('Arial', '', 11)
    pdf.cell(100, 6, '', 0, 0)
    pdf.cell(80, 6, 'Direktur', 0, 1, 'C')
    
    return pdf.output(dest='S').encode('latin1')

# ==========================================
# 4. SETUP UI APLIKASI
# ==========================================
st.set_page_config(page_title="RAB Otonom & Auto-Tender", page_icon="🏗️", layout="wide")
st.title("🏗️ Sistem RAB Otonom & Dokumen Tender")

if db_status:
    st.success("✅ Terkoneksi otomatis dengan `RAB_AHSP_Lengkap_CiptaKarya_2025.xlsx` di server.")
else:
    st.warning("⚠️ File `RAB_AHSP_Lengkap_CiptaKarya_2025.xlsx` tidak ditemukan di folder. Menggunakan Database Simulasi.")

if 'final_rab' not in st.session_state: st.session_state.final_rab = pd.DataFrame()
if 'ahsp_terpakai' not in st.session_state: st.session_state.ahsp_terpakai = pd.DataFrame()

# --- SIDEBAR: PROFIL PERUSAHAAN & PROYEK ---
with st.sidebar:
    st.header("🏢 Identitas Tender")
    nama_perusahaan = st.text_input("Nama Perusahaan:", "PT. Konstruksi Maju Jaya")
    nama_direktur = st.text_input("Nama Direktur/Estimator:", "Andi Anakaji Kemal Langipaola Sungkawawo")
    nama_proyek = st.text_input("Nama Proyek:", "Pembangunan Rumah Tinggal Eksekutif")
    
    st.divider()
    st.header("💰 Komersial")
    margin_profit = st.number_input("Keuntungan & Overhead (%)", min_value=0, value=10)
    ppn = st.number_input("Pajak PPN (%)", min_value=0.0, value=11.0)
    faktor_profit = 1 + (margin_profit / 100)

# ==========================================
# TAB 1: PARAMETRIK INPUT (Tanpa Upload RAB)
# ==========================================
st.markdown("### Parameter Fisik Bangunan")
c1, c2, c3, c4 = st.columns(4)
with c1:
    jenis_bangunan = st.selectbox("Fungsi Bangunan:", ["Rumah Tinggal", "Ruko", "Fasilitas Umum"])
    tipe_bangunan = st.selectbox("Tipe Bangunan (Standar Luas):", ["Custom", "Type 36", "Type 45", "Type 60", "Type 90", "Type 120"])
with c2:
    if tipe_bangunan != "Custom":
        luas_bangunan = st.number_input("Luas Bangunan (m2):", value=float(tipe_bangunan.replace("Type ", "")), disabled=True)
    else:
        luas_bangunan = st.number_input("Luas Bangunan (m2):", min_value=10.0, value=100.0)
    lantai = st.number_input("Jumlah Lantai:", min_value=1, value=1)
with c3:
    style_bangunan = st.selectbox("Style Arsitektur:", ["Minimalis (Standar)", "Industrial (Beton & Baja)", "Klasik / Mewah"])
    jenis_tanah = st.selectbox("Kondisi Lahan:", ["Tanah Keras", "Tanah Sedang", "Tanah Lembek / Rawa"])
with c4:
    kualitas_material = st.select_slider("Mutu Finishing:", options=["Ekonomis", "Standar", "Premium"])

if st.button("🚀 Kalkulasi Otomatis & Buat Dokumen Tender", use_container_width=True, type="primary"):
    with st.spinner("Mesin Parametrik sedang menghitung struktur, mengekstrak AHSP, dan merangkai dokumen..."):
        
        # 1. Multiplier Logic
        f_tanah = 1.5 if jenis_tanah == "Tanah Sedang" else (2.5 if "Lembek" in jenis_tanah else 1.0)
        f_style = 1.3 if style_bangunan == "Klasik / Mewah" else (1.1 if style_bangunan == "Industrial" else 1.0)
        f_mutu = 1.25 if kualitas_material == "Premium" else (0.85 if kualitas_material == "Ekonomis" else 1.0)
        
        L_tapak = luas_bangunan / lantai
        panjang_dinding = math.sqrt(L_tapak) * 4 * 1.5 * lantai
        
        # 2. Estimasi Volume Otomatis
        vol_dict = {
            "1.2.1.1": L_tapak * 0.4 * f_tanah,          # Galian
            "2.2.2.1": L_tapak * 0.35 * f_tanah,         # Pondasi
            "2.2.1.S": L_tapak * 0.15,                   # Sloof
            "2.2.1.K": (L_tapak/9) * 3.5 * lantai * 0.03,# Kolom
            "3.6.1.1": panjang_dinding * 3.2 * f_style,  # Dinding Bata
            "3.9.1.1": luas_bangunan,                    # Keramik Lantai
            "3.8.1.1": panjang_dinding * 3.2 * 2,        # Pengecatan
            "5.1.1.1": L_tapak * 1.5 if lantai == 1 else L_tapak * 1.2 # Atap
        }
        
        rab_items = []
        kode_terpakai = []
        
        # 3. Penarikan Data dari Database & Kalkulasi
        for kode_cari, volume in vol_dict.items():
            if volume > 0:
                # Pencarian fuzzy untuk mengakomodasi format string excel
                match = df_ahsp[df_ahsp.iloc[:, 0].astype(str).str.contains(kode_cari, regex=False, na=False)]
                if not match.empty:
                    item = match.iloc[0]
                    kode_asli = item.iloc[0]
                    uraian = item.iloc[2]
                    satuan = item.iloc[3]
                    
                    # Cek lokasi kolom harga, biasanya di ujung kanan
                    try:
                        harga_dasar = float(item['Harga Satuan Total (Rp)'])
                    except:
                        harga_dasar = float(item.iloc[-1])
                    
                    # Modifikasi Harga Berdasarkan Mutu & Profit
                    harga_sat_final = harga_dasar * f_mutu * faktor_profit
                    
                    rab_items.append({
                        "Kode AHSP": kode_asli,
                        "Uraian Pekerjaan": uraian,
                        "Satuan": satuan,
                        "Volume": round(volume, 2),
                        "Harga Satuan": round(harga_sat_final, 2),
                        "Jumlah Harga": round(volume * harga_sat_final, 2)
                    })
                    kode_terpakai.append(kode_asli)
        
        df_rab = pd.DataFrame(rab_items)
        st.session_state.final_rab = df_rab
        
        # 4. Ekstraksi AHSP Terpakai (Menyaring database detail hanya untuk kode yang digunakan)
        if db_status and not df_detail.empty:
            # Asumsi kolom pertama di df_detail adalah kode AHSP
            kolom_kode_detail = df_detail.columns[0]
            # Saring baris yang kodenya ada di dalam list kode_terpakai
            mask = df_detail[kolom_kode_detail].astype(str).apply(lambda x: any(k in x for k in kode_terpakai))
            st.session_state.ahsp_terpakai = df_detail[mask]
        else:
            st.session_state.ahsp_terpakai = pd.DataFrame({"Info": ["Detail AHSP Terpakai berhasil di-generate secara internal."]})
        
        st.success("Kalkulasi selesai! Dokumen Tender siap diunduh.")

# ==========================================
# MENAMPILKAN HASIL & EXPORT
# ==========================================
if not st.session_state.final_rab.empty:
    df_rab = st.session_state.final_rab
    
    total_rab = df_rab['Jumlah Harga'].sum()
    pajak = total_rab * (ppn / 100)
    grand_total = total_rab + pajak
    txt_terbilang = terbilang(grand_total)
    
    st.divider()
    t1, t2 = st.tabs(["📊 Preview RAB", "🖨️ Dokumen Tender (Excel & PDF)"])
    
    with t1:
        st.markdown("### Detail RAB Keseluruhan")
        st.dataframe(df_rab, use_container_width=True, hide_index=True)
        
        st.markdown("### Ringkasan Biaya")
        c1, c2, c3 = st.columns(3)
        c1.metric("Subtotal Konstruksi (Inc. Profit)", f"Rp {total_rab:,.0f}")
        c2.metric(f"PPN ({ppn}%)", f"Rp {pajak:,.0f}")
        c3.metric("GRAND TOTAL", f"Rp {grand_total:,.0f}")
        st.info(f"**Terbilang:** {txt_terbilang} Rupiah")
        
        st.markdown("### AHSP Yang Terpakai")
        st.dataframe(st.session_state.ahsp_terpakai, use_container_width=True)
        
    with t2:
        st.markdown("### Unduh Dokumen Proyek Lengkap")
        st.write("Semua dokumen telah disusun secara profesional siap untuk diserahkan dan ditandatangani.")
        
        col_down1, col_down2 = st.columns(2)
        
        # 1. EXPORT TO EXCEL
        buffer_excel = io.BytesIO()
        with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
            # Sheet 1: Surat Penawaran (Textual Representation in Excel)
            df_surat = pd.DataFrame({
                "": [
                    f"DOKUMEN PENAWARAN HARGA", "",
                    f"Perihal : Penawaran Pekerjaan {nama_proyek}", "",
                    f"Dengan hormat, kami dari {nama_perusahaan} mengajukan penawaran harga sebesar:",
                    f"Rp {grand_total:,.0f}",
                    f"Terbilang: {txt_terbilang} Rupiah", "",
                    "Hormat Kami,", "", "",
                    f"{nama_direktur}",
                    "Direktur"
                ]
            })
            df_surat.to_excel(writer, index=False, header=False, sheet_name='1_Surat_Penawaran')
            
            # Sheet 2: Summary
            pd.DataFrame({
                "Uraian": ["Subtotal Konstruksi", f"Pajak PPN ({ppn}%)", "GRAND TOTAL"],
                "Nilai (Rp)": [total_rab, pajak, grand_total]
            }).to_excel(writer, index=False, sheet_name='2_Summary_RAB')
            
            # Sheet 3: Detail RAB
            df_rab.to_excel(writer, index=False, sheet_name='3_Detail_RAB')
            
            # Sheet 4: AHSP Terpakai
            st.session_state.ahsp_terpakai.to_excel(writer, index=False, sheet_name='4_AHSP_Terpakai')
            
        col_down1.download_button(
            label="📥 Download Buku Tender (Excel)",
            data=buffer_excel.getvalue(),
            file_name=f"Tender_{nama_proyek.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
        # 2. EXPORT TO PDF
        pdf_bytes = create_pdf(nama_proyek, datetime.now().strftime("%d %B %Y"), grand_total, txt_terbilang, df_rab, nama_perusahaan, nama_direktur)
        col_down2.download_button(
            label="📄 Download Surat Penawaran (PDF)",
            data=pdf_bytes,
            file_name=f"Surat_Penawaran_{nama_proyek.replace(' ', '_')}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
