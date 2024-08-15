import streamlit as st
import os
import shutil

# Direktori penyimpanan file
base_upload_directory = r"C:\Users\madan\OneDrive - Institut Teknologi Bandung\Desktop\inspection_data\input"
croppedview_directory = r"C:\Users\madan\OneDrive - Institut Teknologi Bandung\Desktop\inspection_data\cropped_view"
croppeddim_directory = r"C:\Users\madan\OneDrive - Institut Teknologi Bandung\Desktop\inspection_data\cropped_dimension"
base_output_directory = r'C:\Users\madan\OneDrive - Institut Teknologi Bandung\Desktop\inspection_data\output'

# Membuat direktori dasar jika belum ada
directories = [base_upload_directory, croppedview_directory, croppeddim_directory, base_output_directory]
for directory in directories:
    if not os.path.exists(directory):
        os.makedirs(directory)

st.title("Halaman Upload Gambar Teknik")
st.header("Upload gambar teknik baru")
# Input nomor order
order_number = st.text_input("Masukkan nomor order:")

# Hanya tampilkan uploader jika nomor order sudah diinput
if order_number:
    # Membuat subdirektori berdasarkan nomor order
    upload_directory = os.path.join(base_upload_directory, order_number)
    if not os.path.exists(upload_directory):
        os.makedirs(upload_directory)

    st.header("Upload gambar teknik (pdf) di sini")
    uploaded_file = st.file_uploader("Pilih file PDF", type="pdf")

    if uploaded_file is not None:
        # Menyimpan file yang diunggah ke direktori lokal berdasarkan nomor order
        file_path = os.path.join(upload_directory, uploaded_file.name)
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        st.success(f"File {uploaded_file.name} telah disimpan di {upload_directory}.")
    else:
        st.write("Belum ada file yang diunggah.")
else:
    st.write("Silakan masukkan nomor order terlebih dahulu.")

# Bagian untuk menghapus nomor order yang telah diunggah
st.header("Hapus nomor order yang telah diunggah")
order_dirs = [d for d in os.listdir(base_upload_directory) if os.path.isdir(os.path.join(base_upload_directory, d))]

if order_dirs:
    selected_order = st.selectbox("Pilih nomor order untuk dihapus:", order_dirs)

    if st.button("Hapus nomor order"):
        directories_to_delete = [
            os.path.join(base_upload_directory, selected_order),
            os.path.join(croppedview_directory, selected_order),
            os.path.join(croppeddim_directory, selected_order),
            os.path.join(base_output_directory, selected_order)
        ]
        
        errors = []
        for delete_directory in directories_to_delete:
            if os.path.exists(delete_directory):
                shutil.rmtree(delete_directory)
            else:
                errors.append(delete_directory)

        if errors:
            st.warning(f"Beberapa direktori tidak ditemukan: {', '.join(errors)}")
        else:
            st.success(f"Nomor order {selected_order} telah dihapus beserta semua file yang terkait.")
else:
    st.write("Belum ada nomor order yang diunggah.")
