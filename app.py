import streamlit as st
import pandas as pd
import os
import io

# Konfigurasi Halaman Utama
st.set_page_config(
    page_title="Aplikasi Nilai Bahasa Inggris - SDN Pringapus 01",
    page_icon="🏫",
    layout="wide"
)

FILE_EXCEL = 'data_sekolah.xlsx'

# List kolom nilai sesuai urutan permintaan
KOLOM_NILAI = ['Tugas 1', 'Tugas 2', 'Tugas 3', 'Ulangan 1', 'Ulangan 2', 'Ulangan 3', 'PTS', 'PAS']

# Fungsi Inisialisasi Database Excel (Membuat template jika file belum ada)
def init_database():
    if not os.path.exists(FILE_EXCEL):
        with pd.ExcelWriter(FILE_EXCEL, engine='openpyxl') as writer:
            for kls in range(1, 7):
                for smt in [1, 2]:
                    # 1. Sheet Capaian Pembelajaran (CP)
                    sheet_cp = f'CP_Kls_{kls}_Smt_{smt}'
                    df_cp = pd.DataFrame({
                        "No": [1, 2, 3],
                        "Materi / Elemen": [f"Listening & Speaking (Kls {kls} Smt {smt})", f"Reading & Viewing (Kls {kls} Smt {smt})", f"Writing & Presenting (Kls {kls} Smt {smt})"],
                        "Capaian Pembelajaran": [f"Peserta didik mampu memahami materi Bahasa Inggris Semester {smt}."] * 3
                    })
                    df_cp.to_excel(writer, sheet_name=sheet_cp, index=False)
                    
                    # 2. Sheet Data Nilai
                    sheet_nilai = f'Nilai_Kls_{kls}_Smt_{smt}'
                    df_nilai = pd.DataFrame({
                        "No": [1, 2, 3, 4, 5],
                        "NISN": [f"100{kls}01", f"100{kls}02", f"100{kls}03", f"100{kls}04", f"100{kls}05"],
                        "Nama Siswa": [f"Siswa {chr(65+j)}" for j in range(5)],
                        "Tugas 1": [80, 65, 88, 70, 60],
                        "Tugas 2": [85, 70, 82, 75, 62],
                        "Tugas 3": [80, 68, 85, 72, 65],
                        "Ulangan 1": [78, 60, 90, 70, 58],
                        "Ulangan 2": [82, 65, 88, 74, 60],
                        "Ulangan 3": [80, 67, 85, 71, 62],
                        "PTS": [85, 65, 92, 73, 64],
                        "PAS": [88, 60, 95, 75, 60],
                        "Rata-Rata": [0.0]*5
                    })
                    # Hitung rata-rata awal
                    df_nilai['Rata-Rata'] = df_nilai[KOLOM_NILAI].mean(axis=1).round(1)
                    df_nilai.to_excel(writer, sheet_name=sheet_nilai, index=False)

init_database()

# Fungsi Membaca Sheet dari Excel
def load_sheet(sheet_name):
    try:
        return pd.read_excel(FILE_EXCEL, sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()

# Fungsi Menyimpan Data Kembali ke Excel
def save_sheet(df, sheet_name):
    # Hitung ulang rata-rata sebelum disimpan
    for col in KOLOM_NILAI:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    df['Rata-Rata'] = df[KOLOM_NILAI].mean(axis=1).round(1)
    
    with pd.ExcelWriter(FILE_EXCEL, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    return df

# Fungsi konversi DataFrame ke file Excel untuk Fitur Download
def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Data_Nilai')
    return output.getvalue()

# ================= SIDEBAR NAVIGASI =================
st.sidebar.title("🏫 SDN PRINGAPUS 01")
st.sidebar.caption("Sistem Nilai Bahasa Inggris")
st.sidebar.markdown("---")

menu = st.sidebar.radio(
    "Pilih Menu Utama:",
    ["🏠 Dashboard Profil", "📘 Capaian Pembelajaran (CP)", "📝 Input & Data Nilai"]
)

st.sidebar.markdown("---")
st.sidebar.info("📌 **Database:** Menggunakan file `data_sekolah.xlsx`.")

# ================= 1. MENU DASHBOARD =================
if menu == "🏠 Dashboard Profil":
    st.title("🏫 Dashboard Profil Sekolah")
    st.subheader("SDN Pringapus 01 - Mapel Bahasa Inggris")
    st.markdown("---")
    
    col_logo, col_desc = st.columns([1, 3])
    with col_logo:
        st.image("https://cdn-icons-png.flaticon.com/512/167/167707.png", width=140)
    with col_desc:
        st.write("""
        Selamat datang di **Aplikasi Pengolahan Nilai Bahasa Inggris SDN Pringapus 01**.
        
        **Fitur Utama Aplikasi:**
        * ✏️ **Input Nilai Interaktif:** Guru dapat menginput/mengedit nilai langsung dari tabel di aplikasi.
        * 💾 **Tersimpan Otomatis ke Excel:** Data siswa dan nilai tersimpan aman di database Excel.
        * 📥 **Fitur Download Table:** Rekap nilai dapat diunduh kapan saja ke format Excel/CSV.
        * 🔴 **Penanda KKM Otomatis:** Nilai di bawah standar KKM akan otomatis berwarna merah.
        """)

    st.markdown("---")
    st.write("### 🎯 Ketentuan KKM (Kriteria Ketuntasan Minimal)")
    col1, col2 = st.columns(2)
    with col1:
        st.warning("🔴 **Kelas 1 - Kelas 3**\n\nStandar KKM: **70**\n*(Nilai < 70 ditandai warna merah)*")
    with col2:
        st.warning("🔴 **Kelas 4 - Kelas 6**\n\nStandar KKM: **75**\n*(Nilai < 75 ditandai warna merah)*")

# ================= 2. MENU CAPAIAN PEMBELAJARAN =================
elif menu == "📘 Capaian Pembelajaran (CP)":
    st.title("📘 Data Capaian Pembelajaran (CP)")
    st.markdown("---")
    
    col_k, col_s = st.columns(2)
    with col_k:
        pilihan_kelas = st.selectbox("Pilih Kelas:", [f"Kelas {i}" for i in range(1, 7)])
    with col_s:
        pilihan_semester = st.selectbox("Pilih Semester:", ["Semester 1", "Semester 2"])
        
    kelas_num = pilihan_kelas.split(" ")[1]
    smt_num = pilihan_semester.split(" ")[1]
    sheet_cp = f'CP_Kls_{kelas_num}_Smt_{smt_num}'
    
    st.write(f"### Capaian Pembelajaran - {pilihan_kelas} ({pilihan_semester})")
    df_cp = load_sheet(sheet_cp)
    
    if not df_cp.empty:
        # Editor CP
        edited_cp = st.data_editor(df_cp, num_rows="dynamic", use_container_width=True, hide_index=True)
        if st.button("💾 Simpan Perubahan CP"):
            with pd.ExcelWriter(FILE_EXCEL, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                edited_cp.to_excel(writer, sheet_name=sheet_cp, index=False)
            st.success("Data CP berhasil diperbarui ke Excel!")
    else:
        st.warning("Data CP tidak ditemukan.")

# ================= 3. MENU INPUT & DATA NILAI =================
elif menu == "📝 Input & Data Nilai":
    st.title("📝 Input & Pengolahan Nilai Siswa")
    st.markdown("---")
    
    # Filter Kelas & Semester
    c1, c2 = st.columns(2)
    with c1:
        pilihan_kelas = st.selectbox("Pilih Kelas:", [f"Kelas {i}" for i in range(1, 7)])
    with c2:
        pilihan_semester = st.selectbox("Pilih Semester:", ["Semester 1", "Semester 2"])
        
    kelas_num = int(pilihan_kelas.split(" ")[1])
    smt_num = pilihan_semester.split(" ")[1]
    sheet_nilai = f'Nilai_Kls_{kelas_num}_Smt_{smt_num}'
    
    # Hitung batas KKM
    kkm = 70 if kelas_num <= 3 else 75
    
    st.info(f"📌 **{pilihan_kelas} | {pilihan_semester}** — Batas KKM = **{kkm}** (Nilai < {kkm} ditandai warna merah)")
    
    df_nilai = load_sheet(sheet_nilai)
    
    if not df_nilai.empty:
        st.write("### ✏️ Tabel Input Nilai Siswa")
        st.caption("Silakan ketik/edit nilai secara langsung pada tabel di bawah ini, lalu klik tombol **Simpan Data Nilai**.")
        
        # 1. TABEL INPUT NILAI INTERAKTIF
        edited_df = st.data_editor(
            df_nilai,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            disabled=["No", "Rata-Rata"] # Rata-Rata dihitung otomatis
        )
        
        col_btn1, col_btn2 = st.columns([1, 4])
        with col_btn1:
            if st.button("💾 Simpan Data Nilai", type="primary"):
                updated_df = save_sheet(edited_df, sheet_nilai)
                st.success("Data nilai berhasil disimpan ke Excel!")
                st.rerun()
                
        st.markdown("---")
        
        # 2. TABEL REKAPITULASI & PRATINJAU WARNA (HIGHLIGHT KKM)
        st.write("### 🔍 Rekapitulasi Nilai & Penanda KKM")
        
        # Fungsi Penanda Warna Merah (Pandas 2.1+ Menggunakan .map)
        def highlight_kkm(val):
            try:
                val_num = float(val)
                if val_num < kkm:
                    return 'background-color: #ffcccc; color: #cc0000; font-weight: bold;'
            except (ValueError, TypeError):
                pass
            return ''

        # Terapkan styling warna merah ke kolom nilai & rata-rata
        cols_to_style = KOLOM_NILAI + ['Rata-Rata']
        styled_df = edited_df.style.map(highlight_kkm, subset=cols_to_style)
        
        st.dataframe(styled_df, use_container_width=True, hide_index=True)
        
        st.markdown("---")
        
        # 3. MENU DOWNLOAD HASIL NILAI
        st.write("### 📥 Download Tabel Nilai")
        d_col1, d_col2 = st.columns(2)
        
        # Download format Excel
        excel_data = convert_df_to_excel(edited_df)
        with d_col1:
            st.download_button(
                label="📄 Download Rekap Nilai (Format Excel .xlsx)",
                data=excel_data,
                file_name=f"Nilai_Bahasa_Inggris_{pilihan_kelas}_{pilihan_semester}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
        # Download format CSV
        csv_data = edited_df.to_csv(index=False).encode('utf-8')
        with d_col2:
            st.download_button(
                label="📊 Download Rekap Nilai (Format CSV)",
                data=csv_data,
                file_name=f"Nilai_Bahasa_Inggris_{pilihan_kelas}_{pilihan_semester}.csv",
                mime="text/csv",
                use_container_width=True
            )

    else:
        st.warning("Data Nilai tidak ditemukan pada database.")
