import io
import json
import os
import streamlit as st
import pypdf
import pandas as pd
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

st.title("📖 Generator Silabus & RPP Otomatis dari Kitab")
st.caption("Unggah file PDF atau masukkan teks kitab secara manual untuk merancang Perangkat Ajar Lengkap siap pakai!")

# ---------------------------------------------------------
# Inisialisasi Google GenAI Client (Tanpa UI API Key)
# ---------------------------------------------------------
api_key = os.environ.get("GEMINI_API_KEY")
if not api_key and "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]

if not api_key:
    st.error("⚠️ GEMINI_API_KEY tidak ditemukan di Secrets atau Environment Variables.")
    st.stop()

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
# Fungsi Helper: Format Dokumen Word (.docx)
# ---------------------------------------------------------
def create_full_docx(data: dict) -> io.BytesIO:
    doc = Document()
    
    # Header Utama
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_header = p_title.add_run("PERANGKAT AJAR LENGKAP\nSILABUS & RENCANA PELAKSANAAN PEMBELAJARAN (RPP)\n")
    run_header.bold = True
    run_header.font.size = Pt(14)
    run_header.font.name = 'Arial'

    # I. IDENTITAS PERANGKAT AJAR
    h1 = doc.add_heading("I. IDENTITAS PERANGKAT AJAR", level=2)
    h1.style.font.name = 'Arial'
    
    identitas_text = (
        f"Nama Kitab/Pelajaran: {data.get('nama_kitab', '')}\n"
        f"Fan Ilmu: {data.get('fan_ilmu', '')}\n"
        f"Tingkat / Kelas: {data.get('tingkat_kelas', '')}\n"
        f"Total Pertemuan: {data.get('total_pertemuan', '')}"
    )
    p_id = doc.add_paragraph(identitas_text)
    p_id.style.font.name = 'Arial'
    p_id.style.font.size = Pt(11)

    # II. SILABUS PEMBELAJARAN
    h2 = doc.add_heading("II. SILABUS PEMBELAJARAN", level=2)
    h2.style.font.name = 'Arial'
    
    silabus_list = data.get("silabus", [])
    if silabus_list:
        table = doc.add_table(rows=1, cols=8)
        table.style = 'Table Grid'
        
        headers = ["Ptm", "Bab / Fasal", "Halaman", "Alokasi Waktu", "Metode Klasik", "Capaian Indikator", "Indikator Ketercapaian", "Bentuk Evaluasi"]
        hdr_cells = table.rows[0].cells
        for idx, text in enumerate(headers):
            hdr_cells[idx].text = text
            hdr_cells[idx].paragraphs[0].runs[0].font.bold = True
            hdr_cells[idx].paragraphs[0].runs[0].font.size = Pt(9)

        for item in silabus_list:
            row_cells = table.add_row().cells
            vals = [
                str(item.get("ptm", "")),
                str(item.get("bab", "")),
                str(item.get("halaman", "")),
                str(item.get("alokasi_waktu", "")),
                str(item.get("metode", "")),
                str(item.get("capaian", "")),
                str(item.get("indikator", "")),
                str(item.get("evaluasi", ""))
            ]
            for idx, val in enumerate(vals):
                row_cells[idx].text = val
                if row_cells[idx].paragraphs[0].runs:
                    row_cells[idx].paragraphs[0].runs[0].font.size = Pt(9)

    doc.add_page_break()

    # III. RENCANA PELAKSANAAN PEMBELAJARAN (RPP PER PERTEMUAN)
    h3 = doc.add_heading("III. RENCANA PELAKSANAAN PEMBELAJARAN (RPP)", level=2)
    h3.style.font.name = 'Arial'
    
    rpp_list = data.get("rpp_list", [])
    for idx_rpp, rpp in enumerate(rpp_list, 1):
        doc.add_heading(f"RPP Pertemuan Ke-{rpp.get('ptm', idx_rpp)}", level=3)
        
        rpp_id = (
            f"Sekolah       : {rpp.get('sekolah', 'Lembaga Diniyah/Pesantren')}\n"
            f"Mata Pelajaran: {data.get('fan_ilmu', '')} ({data.get('nama_kitab', '')})\n"
            f"Kelas / Tahap : {data.get('tingkat_kelas', '')}\n"
            f"Materi Pokok  : {rpp.get('materi_pokok', '')}\n"
            f"Alokasi Waktu : {rpp.get('alokasi_waktu', '')}"
        )
        doc.add_paragraph(rpp_id).style.font.name = 'Arial'

        # A. Tujuan Pembelajaran
        doc.add_heading("A. Tujuan Pembelajaran", level=4)
        for t in rpp.get("tujuan_pembelajaran", []):
            p = doc.add_paragraph(f"• {t}")
            p.style.font.name = 'Arial'

        # B. Langkah-Langkah Pembelajaran
        doc.add_heading("B. Langkah-Langkah Pembelajaran", level=4)
        
        doc.add_paragraph("Kegiatan Pendahuluan").runs[0].bold = True
        for p_item in rpp.get("langkah_pembelajaran", {}).get("pendahuluan", []):
            doc.add_paragraph(f"• {p_item}")

        doc.add_paragraph("\nKegiatan Inti").runs[0].bold = True
        inti_list = rpp.get("langkah_pembelajaran", {}).get("inti", [])
        if inti_list:
            table_inti = doc.add_table(rows=1, cols=2)
            table_inti.style = 'Table Grid'
            hdr = table_inti.rows[0].cells
            hdr[0].text = "Aspek"
            hdr[1].text = "Kegiatan Pembelajaran"
            hdr[0].paragraphs[0].runs[0].font.bold = True
            hdr[1].paragraphs[0].runs[0].font.bold = True
            
            for item in inti_list:
                r_cells = table_inti.add_row().cells
                r_cells[0].text = item.get("aspek", "")
                r_cells[1].text = item.get("kegiatan", "")

        doc.add_paragraph("\nKegiatan Penutup").runs[0].bold = True
        for p_item in rpp.get("langkah_pembelajaran", {}).get("penutup", []):
            doc.add_paragraph(f"• {p_item}")

        # C. Penilaian
        doc.add_heading("C. Penilaian Hasil Pembelajaran", level=4)
        penilaian = rpp.get("penilaian", {})
        doc.add_paragraph(f"• Penilaian Sikap: {penilaian.get('sikap', '')}")
        doc.add_paragraph(f"• Penilaian Pengetahuan: {penilaian.get('pengetahuan', '')}")
        doc.add_paragraph(f"• Penilaian Keterampilan: {penilaian.get('keterampilan', '')}")

        # Tanda Tangan
        doc.add_paragraph("\n")
        ttd_table = doc.add_table(rows=2, cols=2)
        ttd_cells_0 = ttd_table.rows[0].cells
        ttd_cells_1 = ttd_table.rows[1].cells
        
        ttd_cells_0[0].text = "Mengetahui,\nKepala Sekolah / Mudir"
        ttd_cells_0[1].text = "\nPengampu Mapel"
        ttd_cells_1[0].text = "\n\n( .................................... )"
        ttd_cells_1[1].text = "\n\n( .................................... )"
        
        if idx_rpp < len(rpp_list):
            doc.add_page_break()

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ---------------------------------------------------------
# Form Input
# ---------------------------------------------------------
col1, col2 = st.columns([1, 1])

with col1:
    nama_kitab = st.text_input("Nama Kitab / Pelajaran:", placeholder="Contoh: Matan Al-Ajurrumiyyah")
    fan_ilmu = st.text_input("Fan Ilmu:", placeholder="Contoh: Nahwu / Fiqih")
    tingkat_kelas = st.text_input("Tingkat / Kelas:", placeholder="Contoh: Kelas 7 / Ula 1")
    
    col_ptm, col_waktu = st.columns(2)
    with col_ptm:
        total_pertemuan_input = st.text_input("Total Pertemuan:", placeholder="Contoh: 6 Pertemuan / 12 Pertemuan")
    with col_waktu:
        alokasi_waktu_input = st.text_input("Alokasi Waktu per Pertemuan:", placeholder="Contoh: 2 x 45 Menit")

with col2:
    metode_input = st.radio(
        "Pilih Metode Masukan Teks Kitab:",
        ["📁 Upload File PDF", "✍️ Ketik / Paste Teks Manual"],
        horizontal=True
    )
    
    kitab_text_input = ""
    if metode_input == "📁 Upload File PDF":
        uploaded_pdf = st.file_uploader("Unggah PDF Kitab/Bab (Max 10MB):", type=["pdf"])
    else:
        uploaded_pdf = None
        kitab_text_input = st.text_area(
            "Masukkan / Paste Teks Kitab di Sini:",
            placeholder="Salin dan tempelkan teks materi/kitab di sini...",
            height=200
        )

# ---------------------------------------------------------
# Proses Generasi
# ---------------------------------------------------------
if st.button("🚀 Buat Silabus & RPP Sekarang", type="primary"):
    final_text = ""
    
    if metode_input == "📁 Upload File PDF":
        if not uploaded_pdf:
            st.error("Harap unggah berkas PDF kitab terlebih dahulu.")
            st.stop()
        with st.spinner("Membaca dan menganalisis teks PDF kitab..."):
            extracted = extract_text_from_pdf(uploaded_pdf)
            if not extracted.strip():
                st.error("Teks pada PDF tidak dapat dibaca. Harap gunakan PDF berbasis teks atau pilih opsi 'Ketik / Paste Teks Manual'.")
                st.stop()
            final_text = extracted
    else:
        if not kitab_text_input.strip():
            st.error("Harap masukkan atau salin teks kitab terlebih dahulu.")
            st.stop()
        final_text = kitab_text_input
        
    pdf_text_truncated = final_text[:30000]

    with st.spinner("Gemini AI sedang menyusun Silabus dan RPP untuk setiap pertemuan..."):
        prompt = f"""
        Anda adalah seorang pakar kurikulum madrasah/pesantren.
        
        Tugas Anda: Analisis teks kitab berikut dan buatkan Silabus serta RPP 1 Lembar UNTUK TIAP PERTEMUAN dalam bentuk format JSON murni.
        
        --- DETAIL INPUT ---
        - Nama Kitab: {nama_kitab}
        - Fan Ilmu: {fan_ilmu}
        - Tingkat / Kelas: {tingkat_kelas}
        - Total Pertemuan: {total_pertemuan_input if total_pertemuan_input else 'Sesuaikan dengan cakupan materi'}
        - Alokasi Waktu per Pertemuan: {alokasi_waktu_input if alokasi_waktu_input else '2 x 45 Menit'}
        --- TEKS KITAB ---
        {pdf_text_truncated}
        --------------------
        
        Instruksi Khusus RPP:
        1. Buatkan array `rpp_list` yang berisi RPP 1 lembar untuk SETIAP PERTEMUAN yang ada di silabus.
        2. Setiap elemen di `rpp_list` harus spesifik membahas materi/bab sesuai nomor pertemuannya.
        
        Keluarkan respons HANYA dalam bentuk JSON valid tanpa tanda backtick markdown dengan struktur persis seperti ini:
        {{
            "nama_kitab": "{nama_kitab}",
            "fan_ilmu": "{fan_ilmu}",
            "tingkat_kelas": "{tingkat_kelas}",
            "total_pertemuan": "{total_pertemuan_input if total_pertemuan_input else '6 Pertemuan'}",
            "silabus": [
                {{
                    "ptm": 1,
                    "bab": "Babu Al-Kalam",
                    "halaman": "Hal 1 - 3",
                    "alokasi_waktu": "{alokasi_waktu_input if alokasi_waktu_input else '2 x 45 Menit'}",
                    "metode": "Ceramah Interaktif dan Tanya Jawab",
                    "capaian": "Santri memahami definisi Al-Kalam...",
                    "indikator": "Santri dapat menyebutkan definisi Al-Kalam...",
                    "evaluasi": "Tes Lisan dan Latihan Soal Tertulis"
                }}
            ],
            "rpp_list": [
                {{
                    "ptm": 1,
                    "sekolah": "Lembaga Diniyah & Pesantren",
                    "materi_pokok": "Pengertian Al-Kalam dan Pembagiannya",
                    "alokasi_waktu": "{alokasi_waktu_input if alokasi_waktu_input else '2 x 45 Menit'}",
                    "tujuan_pembelajaran": [
                        "Santri dapat menjelaskan pengertian Al-Kalam...",
                        "Santri dapat menyebutkan 4 syarat pembentuk Al-Kalam..."
                    ],
                    "langkah_pembelajaran": {{
                        "pendahuluan": [
                            "Guru membuka pembelajaran dengan salam dan doa.",
                            "Guru mengabsen kehadiran santri."
                        ],
                        "inti": [
                            {{"aspek": "Kegiatan Literasi", "kegiatan": "Santri membaca bersama matan..."}},
                            {{"aspek": "Critical Thinking", "kegiatan": "Santri menganalisis 4 syarat..."}}
                        ],
                        "penutup": [
                            "Guru bersama santri merangkum poin penting.",
                            "Pembelajaran ditutup dengan doa."
                        ]
                    }},
                    "penilaian": {{
                        "sikap": "Kedisiplinan dan kerapian selama pembelajaran.",
                        "pengetahuan": "Tes lisan hafalan matan.",
                        "keterampilan": "Kemampuan membuat contoh kalimat."
                    }}
                }}
            ]
        }}
        """
        
        try:
            response = client.models.generate_content(
                model="gemini-3.5-flash-lite",
                contents=prompt
            )
            
            raw_text = response.text.strip()
            if raw_text.startswith("```"):
                raw_text = raw_text.split("\n", 1)[1]
                if raw_text.endswith("```"):
                    raw_text = raw_text.rsplit("\n", 1)[0]
            
            data = json.loads(raw_text)
            st.session_state["result_data"] = data
            
        except Exception as e:
            st.error(f"Gagal memproses AI / JSON: {str(e)}")

# ---------------------------------------------------------
# Tampilan Hasil (Sesuai Desain UI)
# ---------------------------------------------------------
if "result_data" in st.session_state:
    data = st.session_state["result_data"]
    
    st.markdown("---")
    st.header(f"📖 Identitas Kurikulum: {data.get('nama_kitab', '')}")
    
    # 3 Metric Cards
    m1, m2, m3 = st.columns(3)
    with m1:
        st.caption("Fan Ilmu")
        st.markdown(f"### {data.get('fan_ilmu', '-')}")
    with m2:
        st.caption("Tingkat / Kelas")
        st.markdown(f"### {data.get('tingkat_kelas', '-')}")
    with m3:
        st.caption("Total Pertemuan")
        st.markdown(f"### {data.get('total_pertemuan', '-')}")
    
    st.write("")
    
    # Tab Tampilan
    tab_silabus, tab_rpp = st.tabs(["📑 Silabus Pembelajaran", "📝 RPP Tiap Pertemuan"])
    
    with tab_silabus:
        silabus_list = data.get("silabus", [])
        if silabus_list:
            df_silabus = pd.DataFrame(silabus_list)
            df_silabus.rename(columns={
                "ptm": "Ptm",
                "bab": "Bab / Fasal",
                "halaman": "Halaman",
                "alokasi_waktu": "Alokasi Waktu",
                "metode": "Metode Klasik",
                "capaian": "Capaian Indikator",
                "indikator": "Indikator Ketercapaian",
                "evaluasi": "Bentuk Evaluasi"
            }, inplace=True)
            st.dataframe(df_silabus, use_container_width=True, hide_index=True)
            
    with tab_rpp:
        rpp_list = data.get("rpp_list", [])
        if rpp_list:
            ptm_options = [f"Pertemuan Ke-{r.get('ptm', idx+1)}" for idx, r in enumerate(rpp_list)]
            selected_ptm = st.selectbox("Pilih Pertemuan untuk Dilihat:", ptm_options)
            
            selected_index = ptm_options.index(selected_ptm)
            rpp = rpp_list[selected_index]
            
            st.markdown(f"### RPP Pertemuan Ke-{rpp.get('ptm', selected_index+1)}")
            st.markdown(f"**Sekolah**: {rpp.get('sekolah', '-')}")
            st.markdown(f"**Materi Pokok**: {rpp.get('materi_pokok', '-')}")
            st.markdown(f"**Alokasi Waktu**: {rpp.get('alokasi_waktu', '-')}")
            
            st.markdown("#### A. Tujuan Pembelajaran")
            for t in rpp.get("tujuan_pembelajaran", []):
                st.markdown(f"* {t}")
                
            st.markdown("#### B. Langkah-Langkah Pembelajaran")
            st.markdown("**Kegiatan Inti:**")
            df_inti = pd.DataFrame(rpp.get("langkah_pembelajaran", {}).get("inti", []))
            if not df_inti.empty:
                df_inti.rename(columns={"aspek": "Aspek", "kegiatan": "Kegiatan Pembelajaran"}, inplace=True)
                st.table(df_inti)
                
            st.markdown("#### C. Penilaian")
            pen = rpp.get("penilaian", {})
            st.write(f"* **Sikap**: {pen.get('sikap')}")
            st.write(f"* **Pengetahuan**: {pen.get('pengetahuan')}")
            st.write(f"* **Keterampilan**: {pen.get('keterampilan')}")

    st.markdown("---")
    st.subheader("📬 Download Perangkat Ajar Lengkap:")
    
    full_docx_buffer = create_full_docx(data)
    
    st.download_button(
        label="📦 Download Paket Lengkap Silabus & RPP Semua Pertemuan (.docx)",
        data=full_docx_buffer,
        file_name=f"Perangkat_Ajar_Lengkap_{data.get('nama_kitab','').replace(' ', '_')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
        type="primary"
    )
