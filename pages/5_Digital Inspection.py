import os
import streamlit as st
import pandas as pd
from datetime import date
import openpyxl
from openpyxl.styles import Alignment

# Judul halaman
st.title("Part Digital Inspection")

# Input nama petugas inspeksi
nama_petugas = st.text_input("Masukkan nama petugas inspeksi")

# Mengambil tanggal inspeksi (hari ini)
tanggal_inspeksi = date.today().strftime("%d-%m-%Y")

# Base directory input dan output
base_upload_directory = r"C:\Users\madan\OneDrive - Institut Teknologi Bandung\Desktop\inspection_data\input"
base_output_directory = r'C:\Users\madan\OneDrive - Institut Teknologi Bandung\Desktop\inspection_data\output'

# Select box untuk memilih nomor order
order_number = st.selectbox("Pilih nomor order", os.listdir(base_output_directory))

if order_number and nama_petugas:
    # Membuat path folder order untuk upload dan output
    order_folder_path_upload = os.path.join(base_upload_directory, str(order_number))
    order_folder_path_output = os.path.join(base_output_directory, str(order_number))

    # Menampilkan gambar jika ada di folder upload
    if os.path.exists(order_folder_path_upload):
        image_files = [f for f in os.listdir(order_folder_path_upload) if f.endswith('.jpg')]
        if image_files:
            st.write("Gambar Inspeksi:")
            for image_file in image_files:
                image_path = os.path.join(order_folder_path_upload, image_file)
                st.image(image_path, caption=image_file)
        else:
            st.error(f"Tidak ada gambar dengan format JPG di folder {order_folder_path_upload}.")
    else:
        st.error(f"Folder untuk order number {order_number} tidak ditemukan di direktori upload.")

    # Proses inspeksi (berdasarkan kode sebelumnya)
    if not os.path.exists(order_folder_path_output):
        st.error(f"Folder untuk order number {order_number} tidak ditemukan di direktori output.")
    else:
        # Mencari file Excel di folder order
        excel_files = [f for f in os.listdir(order_folder_path_output) if f.endswith('.xlsx')]
        if not excel_files:
            st.error(f"Tidak ada file Excel di folder {order_folder_path_output}.")
        else:
            # Mengambil file Excel pertama yang ditemukan
            file_path = os.path.join(order_folder_path_output, excel_files[0])

            # Membaca data dari Excel
            df = pd.read_excel(file_path, header=None)

            # Menyiapkan data untuk tabel pada Streamlit
            data = []

            for index, row in df.iterrows():
                if index >= 9:  # Mulai dari baris 10 (index 9 karena 0-based index)
                    dimensi = row[1]  # Kolom B
                    max_value = row[3]  # Kolom D
                    min_value = row[4]  # Kolom E
                    hasil_pengukuran = row[5]  # Kolom F
                    if not pd.isna(dimensi) and not pd.isna(max_value) and not pd.isna(min_value):
                        data.append([
                            f"{dimensi:.1f}" if isinstance(dimensi, (int, float)) else dimensi,
                            f"{max_value:.1f}" if isinstance(max_value, (int, float)) else max_value,
                            f"{min_value:.1f}" if isinstance(min_value, (int, float)) else min_value,
                            hasil_pengukuran  # Biarkan hasil pengukuran tetap original
                        ])

            # Membuat DataFrame untuk ditampilkan di Streamlit
            df_display = pd.DataFrame(data, columns=["Dimensi", "Max", "Min", "Hasil Pengukuran"])

            # Menampilkan tabel di halaman Streamlit menggunakan st.data_editor
            edited_df = st.data_editor(df_display)

            # Memulai tampilan form
            with st.form(key='my_form'):
                # Tambahkan tombol submit untuk submit
                submit_button = st.form_submit_button(label='Submit')

            # Tampilkan hasil jika tombol submit ditekan
            if submit_button:
                # Tambahkan kolom Status ke DataFrame
                edited_df['Status'] = ''

                # Perbarui kolom status dan format hasil pengukuran berdasarkan hasil pengukuran
                for idx, row in edited_df.iterrows():
                    try:
                        hasil_pengukuran = float(row['Hasil Pengukuran'])
                        min_value = float(row['Min'])
                        max_value = float(row['Max'])
                        if min_value <= hasil_pengukuran <= max_value:
                            edited_df.at[idx, 'Status'] = 'LOLOS'
                        else:
                            edited_df.at[idx, 'Status'] = 'TAK LOLOS'
                        edited_df.at[idx, 'Hasil Pengukuran'] = f"{hasil_pengukuran:.1f}"  # Format 1 angka di belakang koma
                    except ValueError:
                        edited_df.at[idx, 'Status'] = 'TAK LOLOS'  # Jika hasil_pengukuran bukan angka

                st.write(f"Tabel Data Pengukuran (Petugas: {nama_petugas}, Tanggal: {tanggal_inspeksi})")

                # Menambahkan CSS kustom untuk memperbesar font
                st.markdown(
                    """
                    <style>
                    .dataframe {font-size: 30px !important;}
                    </style>
                    """,
                    unsafe_allow_html=True
                )

                # Menampilkan tabel dengan font yang diperbesar
                st.table(edited_df)

                # Menulis data kembali ke file Excel
                workbook = openpyxl.load_workbook(file_path)
                sheet = workbook.active

                # Menuliskan 'Hasil Pengukuran' pada tabel di excel input pada mulai sel G10 ke bawah
                for idx, row in edited_df.iterrows():
                    cell_hasil_pengukuran = sheet[f'G{10 + idx}']
                    cell_hasil_pengukuran.value = row['Hasil Pengukuran']
                    cell_hasil_pengukuran.alignment = Alignment(horizontal='center')

                    if row['Status'] == 'LOLOS':
                        cell_status = sheet[f'J{10 + idx}']
                    else:
                        cell_status = sheet[f'K{10 + idx}']
                    
                    cell_status.value = '✔'
                    cell_status.alignment = Alignment(horizontal='center')

                # Menuliskan nama_petugas pada L4 pada workbook excel
                cell_nama_petugas = sheet['L4']
                cell_nama_petugas.value = nama_petugas
                cell_nama_petugas.alignment = Alignment(horizontal='center')

                # Menuliskan tanggal pada L5 pada workbook excel
                cell_tanggal = sheet['L5']
                cell_tanggal.value = tanggal_inspeksi
                cell_tanggal.alignment = Alignment(horizontal='center')

                # Menyimpan file Excel baru
                new_directory = r"C:\Users\madan\OneDrive - Institut Teknologi Bandung\Desktop\inspection_data\dokumen inspeksi"
                new_folder_path = os.path.join(new_directory, str(order_number))
                if not os.path.exists(new_folder_path):
                    os.makedirs(new_folder_path)
                new_file_path = os.path.join(new_folder_path, f"{tanggal_inspeksi}_{os.path.basename(file_path)}")
                workbook.save(new_file_path)

                st.success(f"Data berhasil disimpan pada {new_file_path}")