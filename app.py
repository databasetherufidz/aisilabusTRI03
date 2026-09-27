import streamlit as OpenAIClient
import streamlit as st
from openai import OpenAI

# Konfigurasi halaman Streamlit
st.set_page_title_co_kwargs = {}
st.set_page_config(
    page_title="Generator Silabus & RPP Kitab Kuning",
    page_icon="📚",
    layout="wide",
)

st.title("📚 Generator Silabus & RPP Berbasis AI (ChatGPT)")
st.markdown(
    "Aplikasi otomatis untuk menghasilkan Silabus dan RPP mendetail berdasarkan Kitab dan parameter kelas."
)

# Sidebar untuk konfigurasi API Key dan Input Data
with st.sidebar:
  st.header("⚙️ Pengaturan & Parameter")

  # Input API Key OpenAI
  api_key_input = st.text_input(
      "Masukkan OpenAI API Key:", type="password", help="Dapatkan dari platform.openai.com"
  )

  st.markdown("---")
  st.subheader("📝 Data Kitab & Pembelajaran")

  # Form Input sesuai parameter yang Anda berikan
  nama_kitab = st.text_input("Nama Kitab", value="Al-Qawaidus Sharfiyyah")
  fan_ilmu = st.text_input("Fan Ilmu", value="Sharaf")
  tingkat_kelas = st.text_input("Tingkat Kelas", value="7")
  total_pertemuan = st.number_input("Total Pertemuan", value=134, min_value=1)
  alokasi_waktu = st.text_input("Alokasi Waktu", value="2x45 Menit")
  sumber = st.text_input("Sumber / Bab", value="PDF Kitab (Halaman 1-30)")

  generate_btn = st.button("🚀 Buat Silabus & RPP", type="primary")

# Area Utama Tampilan Hasil
if generate_btn:
  if not api_key_input:
    st.error("⚠️ Mohon masukkan OpenAI API Key terlebih dahulu di sidebar sebelah kiri.")
  else:
    with st.spinner("Sedang merancang Silabus dan RPP melalui ChatGPT..."):
      try:
        # Inisialisasi Client OpenAI
        client = OpenAI(api_key=api_key_input)

        # Prompt sistem & instruksi untuk ChatGPT
        prompt_system = (
            "Bertindaklah sebagai sistem kurikulum dan pengajar ahli bahasa"
            " Arab pondok pesantren. Buatkan Silabus dan Rencana Pelaksanaan"
            " Pembelajaran (RPP) yang sangat detail dan terstruktur berdasarkan"
            " data yang diberikan."
        )

        prompt_user = f"""
        Tolong buatkan Silabus dan RPP dengan detail berikut:
        - Nama Kitab: {nama_kitab}
        - Fan Ilmu: {fan_ilmu}
        - Tingkat Kelas: {tingkat_kelas}
        - Total Pertemuan: {total_pertemuan} Pertemuan
        - Alokasi Waktu: {alokasi_waktu}
        - Sumber: {sumber}

        Tolong susun hasilnya dengan format Markdown yang rapi mencakup:
        1. Silabus Pembelajaran (dibagi secara proporsional dari pertemuan 1 sampai {total_pertemuan} mencakup materi terkait).
        2. Rencana Pelaksanaan Pembelajaran (RPP) makro dan mikro yang memuat Tujuan Pembelajaran, Langkah-langkah Pembelajaran (Pendahuluan, Kegiatan Inti dengan metode bandongan/sorogan/tasrif, Penutup), serta Penilaian Hasil Belajar (Sikap, Pengetahuan, dan Keterampilan).
        """

        # Panggilan API ke ChatGPT (Model gpt-4o atau gpt-3.5-turbo)
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": prompt_system},
                {"role": "user", "content": prompt_user},
            ],
            temperature=0.7,
        )

        hasil_output = response.choices[0].message.content

        # Tampilkan hasil di aplikasi Streamlit
        st.success("✅ Silabus dan RPP berhasil dibuat!")
        st.markdown("---")
        st.markdown(hasil_output)

        # Tombol Download Hasil
        st.download_button(
            label="📥 Download Hasil (Markdown)",
            data=hasil_output,
            file_name=f"Silabus_RPP_{nama_kitab.replace(' ', '_')}.md",
            mime="text/markdown",
        )

      except Exception as e:
        st.error(f"Terjadi kesalahan saat menghubungkan ke API: {e}")
else:
  st.info(
      "👈 Silakan lengkapi parameter di sidebar kiri lalu klik tombol **Buat"
      " Silabus & RPP**."
  )
