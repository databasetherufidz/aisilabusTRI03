import io
import streamlit as st
import pypdf
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from google import genai

# ---------------------------------------------------------
# Konfigurasi Halaman Streamlit
# ---------------------------------------------------------
st.set_page_config(
    page_title="Generator Silabus & RPP Kitab",
    page_icon="📖",
    layout="wide"
)

st.title("📖 Generator Silabus & RPP Otomatis dari PDF Kitab")
st.caption("Unggah file PDF kitab/materi, dan AI akan merancang Silabus serta RPP siap pakai!")

# ---------------------------------------------------------
# Manajemen API Key (Sidebar)
# ---------------------------------------------------------
st.sidebar.header("🔑 Konfigurasi API Key")
api_key = st.sidebar.text_input("Masukkan Gemini API Key:", type="password")

if not api_key:
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
    else:
        st.info("💡 Masukkan Gemini API Key Anda di sidebar untuk melanjutkan.")
        st.stop()

# Inisialisasi Google GenAI Client
client = genai.Client(api_key=api_key)

# ---------------------------------------------------------
# Fungsi Helper: Membaca Teks PDF
# ---------------------------------------------------------
def extract_text_from_pdf(pdf_file) -> str:
    reader = pypdf.PdfReader(pdf_file)
    text = ""
    for page in reader.pages:
        extracted = page.extract_text()
        if extracted:
            text += extracted + "\n"
    return text

# ---------------------------------------------------------
# Fungsi Helper: Membuat Word (.docx)
# ---------------------------------------------------------
def create_docx(content_text: str, title: str) -> io.BytesIO:
    doc = Document()
    
    # Header Utama
    h = doc.add_heading(title, level=1)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Isi Teks
    for paragraph in content_text.split('\n'):
        if paragraph.strip():
            p = doc.add_paragraph(paragraph.strip())
            p.style.font.name = 'Arial'
            p.style.font.size = Pt(11)
            
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ---------------------------------------------------------
# Form Input & Upload File
# ---------------------------------------------------------
col1, col2 = st.columns([1, 1])

with col1:
    nama_kitab = st.text_input("Nama Kitab / Pelajaran:", placeholder="Contoh: Nahwu / Safinatun Najah")
    # PERUBAHAN: Input ketik manual untuk Tingkat / Kelas
    tingkat_kelas = st.text_input("Tingkat / Kelas:", placeholder="Contoh: Ula / Kelas 1 / Kelas 7 MTs")

with col2:
    # PERUBAHAN: st.file_uploader yang benar
    uploaded_pdf = st.file_uploader("Unggah PDF Kitab/Bab (Max 10MB):", type=["pdf"])

# ---------------------------------------------------------
# Proses Generasi
# ---------------------------------------------------------
if st.button("🚀 Buat Silabus & RPP Sekarang", type="primary"):
    if not uploaded_pdf:
        st.error("Harap unggah berkas PDF kitab terlebih dahulu.")
        st.stop()
        
    with st.spinner("Membaca dan menganalisis teks PDF kitab..."):
        pdf_text = extract_text_from_pdf(uploaded_pdf)
        
        if not pdf_text.strip():
            st.error("Teks pada PDF tidak dapat dibaca (kemungkinan PDF berbentuk scan gambar). Harap gunakan PDF berbasis teks.")
            st.stop()
            
        # Batasi panjang teks (30.000 karakter pertama)
        pdf_text_truncated = pdf_text[:30000]

    with st.spinner("Gemini AI sedang menyusun Silabus dan RPP..."):
        prompt = f"""
        Anda adalah seorang pakar kurikulum pendidikan madrasah/pesantren.
        
        Tugas Anda: Buatkan Silabus dan RPP (Rencana Pelaksanaan Pembelajaran) 1 Lembar yang sistematis dan rapi berdasarkan rujukan teks kitab berikut:
        
        --- DETAIL INPUT ---
        - Nama Kitab/Materi: {nama_kitab}
        - Tingkat/Kelas: {tingkat_kelas}
        --- TEKS KITAB ---
        {pdf_text_truncated}
        --------------------
        
        Harap sajikan output dalam format Markdown yang rapi dengan struktur:
        1. SILABUS PEMBELAJARAN
           - Identitas
           - Capaian Pembelajaran / Standar Kompetensi
           - Pemetaan Bab/Materi Pokok & Alokasi Waktu
           - Indikator Pencapaian
           - Metode & Media Pembelajaran
           - Penilaian (Evaluasi)
        
        2. RPP 1 LEMBAR (RENCANA PELAKSANAAN PEMBELAJARAN)
           - Tujuan Pembelajaran
           - Langkah-Langkah Kegiatan Pembelajaran (Pendahuluan, Inti, Penutup)
           - Asesmen/Penilaian Hasil Belajar
        """
        
        try:
            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=prompt
            )
            
            result_text = response.text
            
            st.success("✅ Silabus & RPP Berhasil Dibuat!")
            
            # Tampilkan Hasil di Tab
            tab1, tab2 = st.tabs(["📄 Tampilan Teks", "📥 Download File Docx"])
            
            with tab1:
                st.markdown(result_text)
                
            with tab2:
                docx_buffer = create_docx(result_text, f"Silabus & RPP - {nama_kitab}")
                st.download_button(
                    label="⬇️ Download Dokumen Word (.docx)",
                    data=docx_buffer,
                    file_name=f"Silabus_RPP_{nama_kitab.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
        except Exception as e:
            st.error(f"Terjadi kesalahan saat menghubungi API: {str(e)}")
