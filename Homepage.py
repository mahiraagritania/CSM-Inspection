import streamlit as st

st.set_page_config(
    page_title="CSM Inspection"
)

st.title("Aplikasi Digital Inspeksi CV CSM")

st.write("Prosedur penggunaan aplikasi digital inspeksi:")

st.subheader("1. Halaman Upload")
st.write("""
Halaman untuk mengupload pdf gambar teknik yang akan dibuat lembar inspeksi. 
""")

st.subheader("2. Halaman Inspection Sheet Generation")
st.write("""
Halaman untuk melakukan pembuatan lembar inspeksi gambar teknik yang sudah diupload.
""")

st.subheader("3. Halaman Digital Inspection")
st.write("""
Halaman untuk memasukkan hasil pengukuran dimensi dan menghasilkan dokumen lembar inspeksi. 
""")