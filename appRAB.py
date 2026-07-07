import streamlit as st
import pandas as pd
import io

# ==========================================
# 1. FUNGSI KONVERSI ANGKA KE HURUF (TERBILANG)
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
# 2. SETUP UI & INISIALISASI STATE
# ==========================================
st.set_page_config(page_title="Mesin RAB Otomatis Cipta Karya", page_icon="🏢", layout="wide")
st.title("🏢 Auto-RAB Estimator (Excel Database Engine)")

if 'keranjang_rab' not in st.session_state:
    st.session_state.keranjang_rab = []

# ==========================================
# 3. SIDEBAR: UPLOAD & KONEKSI DATABASE
# ==========================================
st.sidebar.header("📂 1. Upload Database Excel")
st.sidebar.caption("Unggah file RAB_AHSP_Lengkap_CiptaKarya_2025.xlsx Anda di sini.")
db_file = st.sidebar.file_uploader("Pilih File Excel:", type=['xlsx', 'xls'])

if db_file:
    # Membaca seluruh sheet yang ada di dalam Excel
    xls = pd.ExcelFile(db_file)
    sheet_names = xls.sheet_names
    
    st.sidebar.success("Database berhasil dimuat!")
    st.sidebar.header("⚙️ 2. Pemetaan Data (Mapping)")
    
    # Membiarkan Anda memilih sheet mana yang berfungsi sebagai Master AHSP
    sheet_master = st.sidebar.selectbox("Pilih Sheet Daftar Harga Satuan (AHSP):", sheet_names)
    
    # Load Sheet Master AHSP
    try:
        df_master = pd.read_excel(db_file, sheet_name=sheet_master)
        # Menghapus baris yang kosong
        df_master = df_master.dropna(how='all')
        
        # Mencari kolom secara otomatis berdasarkan teks
        kolom_kode = st.sidebar.selectbox("Pilih Kolom KODE AHSP:", df_master.columns)
        kolom_uraian = st.sidebar.selectbox("Pilih Kolom URAIAN PEKERJAAN:", df_master.columns)
        kolom_satuan = st.sidebar.selectbox("Pilih Kolom SATUAN:", df_master.columns)
        kolom_harga = st.sidebar.selectbox("Pilih Kolom HARGA SATUAN:", df_master.columns)
        
        # Membuat kolom referensi untuk dropdown
        df_master['Label_Pilihan'] = df_master[kolom_kode].astype(str) + " - " + df_master[kolom_uraian].astype(str)
        
    except Exception as e:
        st.error(f"Terjadi kesalahan saat membaca sheet: {e}")
        st.stop()
    
    # Pengaturan Pajak & SMKK
    st.sidebar.header("💰 3. Pengaturan Proyek")
    persen_ppn = st.sidebar.number_input("Pajak PPN (%)", value=11.0, step=1.0)
    persen_smkk = st.sidebar.number_input("Biaya SMKK (%)", value=2.0, step=0.5, help="Persentase dari total konstruksi")

    # ==========================================
    # 4. TAB UTAMA: INPUT RAB
    # ==========================================
    tab1, tab2, tab3 = st.tabs(["📝 1. Susun RAB", "📊 2. Preview Detail & Rangkuman", "📥 3. Cetak Output (Excel)"])
    
    with tab1:
        st.subheader("Cari & Tambahkan Pekerjaan")
        with st.form("form_input"):
            c1, c2, c3 = st.columns([3, 1, 1])
            pilihan = c1.selectbox("Ketik/Pilih Pekerjaan dari Database:", df_master['Label_Pilihan'].tolist())
            volume = c2.number_input("Input Volume:", min_value=0.01, value=1.0, step=0.1)
            btn_tambah = c3.form_submit_button("➕ Tambahkan ke RAB")
            
            if btn_tambah:
                item_terpilih = df_master[df_master['Label_Pilihan'] == pilihan].iloc[0]
                harga_satuan = float(item_terpilih[kolom_harga])
                
                st.session_state.keranjang_rab.append({
                    "Kode AHSP": item_terpilih[kolom_kode],
                    "Uraian Pekerjaan": item_terpilih[kolom_uraian],
                    "Satuan": item_terpilih[kolom_satuan],
                    "Volume": volume,
                    "Harga Satuan": harga_satuan,
                    "Jumlah Harga": volume * harga_satuan
                })
                st.success(f"Berhasil menambahkan: {item_terpilih[kolom_uraian]}")

        # Tabel Live Update
        st.divider()
        if st.session_state.keranjang_rab:
            st.markdown("### Daftar Pekerjaan Sementara")
            df_rab_sementara = pd.DataFrame(st.session_state.keranjang_rab)
            edited_rab = st.data_editor(
                df_rab_sementara, 
                num_rows="dynamic",
                use_container_width=True,
                column_config={
                    "Harga Satuan": st.column_config.NumberColumn(format="Rp %.0f"),
                    "Jumlah Harga": st.column_config.NumberColumn(format="Rp %.0f")
                }
            )
            # Hitung ulang jika ada editan di tabel
            edited_rab['Jumlah Harga'] = edited_rab['Volume'] * edited_rab['Harga Satuan']
            st.session_state.keranjang_rab = edited_rab.to_dict('records')
            
            if st.button("Kosongkan Semua"):
                st.session_state.keranjang_rab = []
                st.rerun()

    # ==========================================
    # 5. TAB: PREVIEW & KALKULASI TOTAL
    # ==========================================
    with tab2:
        if st.session_state.keranjang_rab:
            df_final_rab = pd.DataFrame(st.session_state.keranjang_rab)
            
            total_konstruksi = df_final_rab['Jumlah Harga'].sum()
            biaya_smkk = total_konstruksi * (persen_smkk / 100)
            subtotal = total_konstruksi + biaya_smkk
            pajak_ppn = subtotal * (persen_ppn / 100)
            grand_total = subtotal + pajak_ppn
            
            st.subheader("Rangkuman Biaya Proyek")
            col_met1, col_met2, col_met3 = st.columns(3)
            col_met1.metric("1. Total Konstruksi Dasar", f"Rp {total_konstruksi:,.0f}")
            col_met2.metric(f"2. Biaya SMKK ({persen_smkk}%)", f"Rp {biaya_smkk:,.0f}")
            col_met3.metric(f"3. PPN ({persen_ppn}%)", f"Rp {pajak_ppn:,.0f}")
            
            st.markdown(f"## 💰 GRAND TOTAL: Rp {grand_total:,.0f}")
            st.info(f"**TERBILANG:** *{terbilang(grand_total)} Rupiah*")
        else:
            st.info("RAB masih kosong. Tambahkan data di Tab 1.")

    # ==========================================
    # 6. TAB: EXPORT EXCEL
    # ==========================================
    with tab3:
        if st.session_state.keranjang_rab:
            st.markdown("### Ekspor Laporan Lengkap")
            st.write("Sistem akan membungkus RAB Detail, Rangkuman Biaya, dan Kalkulasi SMKK ke dalam format Excel siap cetak.")
            
            # Membuat File Excel dalam Memory
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                # 1. Sheet Rangkuman
                df_summary = pd.DataFrame({
                    "Keterangan": [
                        "A. TOTAL BIAYA KONSTRUKSI", 
                        f"B. BIAYA SMKK / K3 ({persen_smkk}%)", 
                        "C. SUBTOTAL (A+B)", 
                        f"D. PAJAK PPN ({persen_ppn}%)", 
                        "E. GRAND TOTAL", 
                        "TERBILANG:"
                    ],
                    "Nilai": [
                        total_konstruksi, 
                        biaya_smkk, 
                        subtotal, 
                        pajak_ppn, 
                        grand_total, 
                        f"{terbilang(grand_total)} Rupiah"
                    ]
                })
                df_summary.to_excel(writer, index=False, sheet_name='1_Ringkasan_RAB')
                
                # 2. Sheet RAB Detail
                df_final_rab.to_excel(writer, index=False, sheet_name='2_Detail_RAB')
                
                # 3. Sheet Analisa SMKK (Breakdown Proporsional Otomatis)
                df_smkk = pd.DataFrame({
                    "Uraian Keselamatan Konstruksi (SMKK)": [
                        "1. Penyiapan RKK & Dokumen", 
                        "2. Sosialisasi, Promosi & Pelatihan K3", 
                        "3. Alat Pelindung Kerja & Diri (APD)", 
                        "4. Asuransi BPJS Ketenagakerjaan", 
                        "5. Rambu, Barikade, & Lantas"
                    ],
                    "Alokasi Dana (Rp)": [
                        biaya_smkk * 0.10,
                        biaya_smkk * 0.15,
                        biaya_smkk * 0.40,
                        biaya_smkk * 0.20,
                        biaya_smkk * 0.15
                    ]
                })
                df_smkk.to_excel(writer, index=False, sheet_name='3_Analisa_SMKK')
                
                # Opsi tambahan: Jika di dalam Excel asli ada sheet Upah/Bahan, copy juga ke laporan akhir
                for sheet in ['Upah Bahan', 'Database_Bahan', 'Database_Upah']:
                    if sheet in sheet_names:
                        pd.read_excel(db_file, sheet_name=sheet).to_excel(writer, index=False, sheet_name=sheet[:31])
                
            st.download_button(
                label="⬇️ Download Output Laporan RAB (.xlsx)",
                data=buffer.getvalue(),
                file_name="Laporan_Output_RAB_Otomatis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
else:
    st.info("Menunggu Anda mengunggah File Excel Database...")
