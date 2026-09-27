import io
import json
import os
import re
import Streamlit as st
import pypdf
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
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
# Inisialisasi Google GenAI Client
# ---------------------------------------------------------
api_key = os.environ.get("GEMINI_API_KEY")
if not api_key and "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]

if not api_key:
    st.error("⚠️ GEMINI_API_KEY tidak ditemukan di Secrets atau Environment Variables.")
    st.stop()

client = genai.Client(api_key=api_key)

# ---------------------------------------------------------
# Fungsi Helper: Membaca & Memfilter Halaman PDF
# ---------------------------------------------------------
def parse_page_range(range_str: str, max_pages: int) -> set:
    """Mengubah input teks seperti '1-5, 8, 10-12' menjadi set indeks halaman (0-indexed)."""
    pages = set()
    if not range_str.strip():
        return set(range(max_pages))
        
    parts = range_str.split(',')
    for part in parts:
        part = part.strip()
        if '-' in part:
            sub = part.split('-')
            if len(sub) == 2 and sub[0].isdigit() and sub[1].isdigit():
                start = max(1, int(sub[0]))
                end = min(max_pages, int(sub[1]))
                for p in range(start, end + 1):
                    pages.add(p - 1)
        elif part.isdigit():
            p = int(part)
            if 1 <= p <= max_pages:
                pages.add(p - 1)
                
    return pages if pages else set(range(max_pages))

def extract_text_from_pdf(pdf_file, page_range_str: str = "") -> tuple[str, int, int]:
    reader = pypdf.PdfReader(pdf_file)
    total_pdf_pages = len(reader.pages)
    selected_indices = parse_page_range(page_range_str, total_pdf_pages)
    
    text = ""
    for idx in sorted(list(selected_indices)):
        page = reader.pages[idx]
        extracted = page.extract_text()
        if extracted:
            text += f"\n--- [Halaman {idx + 1}] ---\n" + extracted
            
    return text, len(selected_indices), total_pdf_pages

# ---------------------------------------------------------
# Fungsi Helper: Format Dokumen Word (.docx)
# ---------------------------------------------------------
def create_full_docx(data: dict) -> io.BytesIO:
    doc = Document()
    
    # ---------------- SILABUS SECTION ----------------
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_header = p_title.add_run("PERANGKAT AJAR LENGKAP\nSILABUS & RENCANA PELAKSANAAN PEMBELAJARAN (RPP)\n")
    run_header.bold = True
    run_header.font.size = Pt(14)
    run_header.font.name = 'Arial'

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

    # ---------------- RPP SECTION ----------------
    rpp_list = data.get("rpp_list", [])
    for idx_rpp, rpp in enumerate(rpp_list, 1):
        p_lembaga = doc.add_paragraph()
        p_lembaga.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_inst = p_lembaga.add_run("THE RUFIDZ INDONESIA\n")
        run_inst.bold = True
        run_inst.font.size = Pt(14)
        run_inst.font.name = 'Arial'
        
        run_rpp_title = p_lembaga.add_run("RENCANA PELAKSANAAN PEMBELAJARAN (RPP)")
        run_rpp_title.bold = True
        run_rpp_title.font.size = Pt(12)
        run_rpp_title.font.name = 'Arial'

        rpp_id = (
            f"Sekolah       : {rpp.get('sekolah', 'MDT / Pesantren Rufidz Tahfidz & Diniyah Indonesia')}\n"
            f"Mata Pelajaran: {data.get('fan_ilmu', '')} ({data.get('nama_kitab', '')})\n"
            f"Kelas / Tahap : {data.get('tingkat_kelas', '')}\n"
            f"Materi Pokok  : {rpp.get('materi_pokok', '')}\n"
            f"Alokasi Waktu : {rpp.get('alokasi_waktu', '')}"
        )
        p_ident = doc.add_paragraph(rpp_id)
        p_ident.style.font.name = 'Arial'
        p_ident.style.font.size = Pt(11)

        # A. Tujuan Pembelajaran
        p_tujuan_head = doc.add_paragraph()
        r_tujuan = p_tujuan_head.add_run("A. Tujuan Pembelajaran")
        r_tujuan.bold = True
        r_tujuan.font.size = Pt(11)
        r_tujuan.font.name = 'Arial'
        
        for t in rpp.get("tujuan_pembelajaran", []):
            p = doc.add_paragraph(f"• {t}")
            p.style.font.name = 'Arial'
            p.style.font.size = Pt(11)

        # B. Langkah-Langkah Pembelajaran
        p_langkah_head = doc.add_paragraph()
        r_langkah = p_langkah_head.add_run("B. Langkah-Langkah Pembelajaran")
        r_langkah.bold = True
        r_langkah.font.size = Pt(11)
        r_langkah.font.name = 'Arial'
        
        # 1. Kegiatan Pendahuluan
        p_pend = doc.add_paragraph()
        r_pend = p_pend.add_run("1. Kegiatan Pendahuluan")
        r_pend.bold = True
        r_pend.font.size = Pt(11)
        r_pend.font.name = 'Arial'
        
        for p_item in rpp.get("langkah_pembelajaran", {}).get("pendahuluan", []):
            p = doc.add_paragraph(f"• {p_item}")
            p.style.font.name = 'Arial'
            p.style.font.size = Pt(11)

        # 2. Kegiatan Inti
        p_inti_head = doc.add_paragraph()
        r_inti_head = p_inti_head.add_run("2. Kegiatan Inti")
        r_inti_head.bold = True
        r_inti_head.font.size = Pt(11)
        r_inti_head.font.name = 'Arial'
        
        inti_list = rpp.get("langkah_pembelajaran", {}).get("inti", [])
        if inti_list:
            table_inti = doc.add_table(rows=1, cols=2)
            table_inti.style = 'Table Grid'
            hdr = table_inti.rows[0].cells
            hdr[0].text = "Aspek"
            hdr[1].text = "Kegiatan Pembelajaran"
            hdr[0].paragraphs[0].runs[0].font.bold = True
            hdr[1].paragraphs[0].runs[0].font.bold = True
            hdr[0].paragraphs[0].runs[0].font.size = Pt(10)
            hdr[1].paragraphs[0].runs[0].font.size = Pt(10)
            
            for item in inti_list:
                r_cells = table_inti.add_row().cells
                r_cells[0].text = item.get("aspek", "")
                r_cells[1].text = item.get("kegiatan", "")
                r_cells[0].paragraphs[0].runs[0].font.size = Pt(10)
                r_cells[1].paragraphs[0].runs[0].font.size = Pt(10)

        # 3. Kegiatan Penutup
        p_penut = doc.add_paragraph()
        r_penut = p_penut.add_run("\n3. Kegiatan Penutup")
        r_penut.bold = True
        r_penut.font.size = Pt(11)
        r_penut.font.name = 'Arial'
        
        for p_item in rpp.get("langkah_pembelajaran", {}).get("penutup", []):
            p = doc.add_paragraph(f"• {p_item}")
            p.style.font.name = 'Arial'
            p.style.font.size = Pt(11)

        # C. Penilaian Hasil Pembelajaran
        p_penilaian_head = doc.add_paragraph()
        r_penilaian = p_penilaian_head.add_run("C. Penilaian Hasil Pembelajaran")
        r_penilaian.bold = True
        r_penilaian.font.size = Pt(11)
        r_penilaian.font.name = 'Arial'
        
        penilaian = rpp.get("penilaian", {})
        st_p1 = doc.add_paragraph(f"• Penilaian Sikap: {penilaian.get('sikap', '')}")
        st_p1.style.font.name = 'Arial'
        st_p1.style.font.size = Pt(11)
        
        st_p2 = doc.add_paragraph(f"• Penilaian Pengetahuan: {penilaian.get('pengetahuan', '')}")
        st_p2.style.font.name = 'Arial'
        st_p2.style.font.size = Pt(11)
        
        st_p3 = doc.add_paragraph(f"• Penilaian Keterampilan: {penilaian.get('keterampilan', '')}")
        st_p3.style.font.name = 'Arial'
        st_p3.style.font.size = Pt(11)

        # Tanda Tangan
        doc.add_paragraph("\n")
        ttd_table = doc.add_table(rows=2, cols=2)
        ttd_cells_0 = ttd_table.rows[0].cells
        ttd_cells_1 = ttd_table.rows[1].cells
        
        ttd_cells_0[0].text = "Mengetahui,\nKepala Sekolah"
        ttd_cells_0[1].text = "\nPengampu Mapel"
        ttd_cells_1[0].text = "\n\n( Mudir Sekolah )"
        ttd_cells_1[1].text = "\n\n( Ustadz Pengampu )"
        
        for row in ttd_table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.font.name = 'Arial'
                        run.font.size = Pt(11)
        
        if idx_rpp < len(rpp_list):
            doc.add_page_break()

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ---------------------------------------------------------
# Form Input UI
# ---------------------------------------------------------
col1, col2 = st.columns([1, 1])

with col1:
    nama_kitab = st.text_input("Nama Kitab / Pelajaran:", placeholder="Contoh: Nahwu / Matan Al-Ajurrumiyyah")
    fan_ilmu = st.text_input("Fan Ilmu:", placeholder="Contoh: Ilmu Nahwu")
    tingkat_kelas = st.text_input("Tingkat / Kelas:", placeholder="Contoh: Kelas 7")
    
    col_ptm, col_waktu = st.columns(2)
    with col_ptm:
        total_pertemuan_input = st.text_input("Total Pertemuan:", placeholder="Contoh: 6 Pertemuan")
    with col_waktu:
        alokasi_waktu_input = st.text_input("Alokasi Waktu per Pertemuan:", placeholder="Contoh: 2 x 45 Menit")

with col2:
    metode_input = st.radio(
        "Pilih Metode Masukan Teks Kitab:",
        ["📁 Upload File PDF", "✍️ Ketik / Paste Teks Manual"],
        horizontal=True
    )
    
    kitab_text_input = ""
    page_range_input = ""
    if metode_input == "📁 Upload File PDF":
        uploaded_pdf = st.file_uploader("Unggah PDF Kitab/Bab (Max 10MB):", type=["pdf"])
        page_range_input = st.text_input(
            "Rentang Halaman yang Ingin Dianalisis (Opsional):",
            placeholder="Contoh: 1-10 atau 5, 8, 12-20 (Kosongkan jika semua)",
            help="Tentukan halaman tertentu pada PDF yang ingin dijadikan bahan RPP/Silabus."
        )
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
        with st.spinner("Membaca dan memproses halaman PDF..."):
            extracted, selected_count, total_count = extract_text_from_pdf(uploaded_pdf, page_range_input)
            if not extracted.strip():
                st.error("Teks pada halaman PDF terpilih tidak terbaca/kosong. Pastikan file PDF mengandung teks.")
                st.stop()
            final_text = extracted
            st.success(f"Berhasil mengekstrak {selected_count} dari total {total_count} halaman PDF.")
    else:
        if not kitab_text_input.strip():
            st.error("Harap masukkan atau salin teks kitab terlebih dahulu.")
            st.stop()
        final_text = kitab_text_input
        
    pdf_text_truncated = final_text[:30000]

    with st.spinner("Gemini AI sedang menyusun Silabus dan RPP..."):
        prompt = f"""
        Anda adalah seorang pakar kurikulum madrasah/pesantren.
        
        Tugas Anda: Analisis teks kitab berikut dan buatkan Silabus serta RPP 1 Lembar UNTUK TIAP PERTEMUAN dalam format JSON murni.
        
        --- DETAIL INPUT ---
        - Nama Kitab: {nama_kitab}
        - Fan Ilmu: {fan_ilmu}
        - Tingkat / Kelas: {tingkat_kelas}
        - Total Pertemuan: {total_pertemuan_input if total_pertemuan_input else 'Sesuaikan dengan materi'}
        - Alokasi Waktu per Pertemuan: {alokasi_waktu_input if alokasi_waktu_input else '2 x 45 Menit'}
        --- TEKS KITAB ---
        {pdf_text_truncated}
        --------------------
        
        Gunakan acuan struktur RPP 1 Lembar persis seperti ini:
        - Sekolah: MDT / Pesantren Rufidz Tahfidz & Diniyah Indonesia
        - Pendahuluan: Berisi salam, istighfar, doa bersama, presensi, motivasi, apersepsi, dan penyampaian tujuan.
        - Kegiatan Inti (Tabel 5 C): "Kegiatan Literasi", "Critical Thinking", "Collaboration", "Communication", "Creativity".
        - Penutup: Menyimpulkan poin utama, umpan balik/tugas, doa kafaratul majlis, dan salam.
        - Penilaian: Penilaian Sikap, Pengetahuan, dan Keterampilan.
        
        Keluarkan respons HANYA dalam bentuk JSON valid tanpa tanda backtick markdown dengan struktur persis seperti ini:
        {{
            "nama_kitab": "{nama_kitab}",
            "fan_ilmu": "{fan_ilmu}",
            "tingkat_kelas": "{tingkat_kelas}",
            "total_pertemuan": "{total_pertemuan_input if total_pertemuan_input else '6 Pertemuan'}",
            "silabus": [
                {{
                    "ptm": 1,
                    "bab": "Bab Al-Kalam (Pengertian Kalam dan Pembagian Kata)",
                    "halaman": "Hal 1 - 3",
                    "alokasi_waktu": "{alokasi_waktu_input if alokasi_waktu_input else '2 x 45 Menit'}",
                    "metode": "Ceramah Interaktif dan Diskusi Kelompok",
                    "capaian": "Santri mampu menjelaskan definisi Kalam menurut istilah ilmu Nahwu dengan tepat...",
                    "indikator": "Santri mampu mengidentifikasi pembagian kata beserta tanda-tandanya...",
                    "evaluasi": "Tes Lisan dan Tulis"
                }}
            ],
            "rpp_list": [
                {{
                    "ptm": 1,
                    "sekolah": "MDT / Pesantren Rufidz Tahfidz & Diniyah Indonesia",
                    "materi_pokok": "Bab Al-Kalam (Pengertian Kalam dan Pembagian Kata)",
                    "alokasi_waktu": "{alokasi_waktu_input if alokasi_waktu_input else '2 x 45 Menit'}",
                    "tujuan_pembelajaran": [
                        "Santri mampu menjelaskan definisi Kalam menurut istilah ilmu Nahwu dengan tepat.",
                        "Santri mampu mengidentifikasi pembagian kata (Isim, Fi'il, dan Huruf) beserta tanda-tandanya dengan benar."
                    ],
                    "langkah_pembelajaran": {{
                        "pendahuluan": [
                            "Guru membuka pembelajaran dengan salam, istighfar, dan doa bersama.",
                            "Guru memeriksa kehadiran santri dan memberikan motivasi pentingnya belajar ilmu Nahwu.",
                            "Guru menyampaikan apersepsi serta tujuan pembelajaran untuk Bab Al-Kalam."
                        ],
                        "inti": [
                            {{"aspek": "Kegiatan Literasi", "kegiatan": "Santri membaca dan mencermati matan Bab Al-Kalam dalam Kitab secara bersama-sama."}},
                            {{"aspek": "Critical Thinking", "kegiatan": "Guru memberikan kesempatan kepada santri untuk mendiskusikan perbedaan karakteristik Isim, Fi'il, dan Huruf."}},
                            {{"aspek": "Collaboration", "kegiatan": "Santri dibentuk dalam kelompok kecil untuk mengklasifikasikan lafad-lafad ke dalam Isim, Fi'il, atau Huruf."}},
                            {{"aspek": "Communication", "kegiatan": "Masing-masing kelompok menyampaikan hasil klasifikasinya di depan kelas."}},
                            {{"aspek": "Creativity", "kegiatan": "Santri membuat bagan atau peta konsep sederhana tentang unsur pembentuk Kalam."}}
                        ],
                        "penutup": [
                            "Guru bersama santri menyimpulkan poin-poin utama materi Bab Al-Kalam.",
                            "Guru memberikan umpan balik, apresiasi, serta tugas latihan mandiri di rumah.",
                            "Kegiatan pembelajaran ditutup dengan doa kafaratul majlis dan salam."
                        ]
                    }},
                    "penilaian": {{
                        "sikap": "Observasi kedisiplinan dan keaktifan santri selama kegiatan belajar mengajar.",
                        "pengetahuan": "Tes lisan dan tulis mengenai definisi serta pembagian kalam.",
                        "keterampilan": "Unjuk kerja kelengkapan dan kerapihan bagan peta konsep pembagian kata."
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
# Tampilan Hasil
# ---------------------------------------------------------
if "result_data" in st.session_state:
    data = st.session_state["result_data"]
    
    st.markdown("---")
    st.header(f"📖 Identitas Kurikulum: {data.get('nama_kitab', '')}")
    
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
            
            st.markdown("### THE RUFIDZ INDONESIA")
            st.markdown("#### RENCANA PELAKSANAAN PEMBELAJARAN (RPP)")
            st.markdown(f"**Sekolah**: {rpp.get('sekolah', '-')}")
            st.markdown(f"**Mata Pelajaran**: {data.get('fan_ilmu', '')} ({data.get('nama_kitab', '')})")
            st.markdown(f"**Kelas / Tahap**: {data.get('tingkat_kelas', '')}")
            st.markdown(f"**Materi Pokok**: {rpp.get('materi_pokok', '-')}")
            st.markdown(f"**Alokasi Waktu**: {rpp.get('alokasi_waktu', '-')}")
            
            st.markdown("#### A. Tujuan Pembelajaran")
            for t in rpp.get("tujuan_pembelajaran", []):
                st.markdown(f"• {t}")
                
            st.markdown("#### B. Langkah-Langkah Pembelajaran")
            st.markdown("**1. Kegiatan Pendahuluan**")
            for p_item in rpp.get("langkah_pembelajaran", {}).get("pendahuluan", []):
                st.markdown(f"• {p_item}")
                
            st.markdown("**2. Kegiatan Inti**")
            df_inti = pd.DataFrame(rpp.get("langkah_pembelajaran", {}).get("inti", []))
            if not df_inti.empty:
                df_inti.rename(columns={"aspek": "Aspek", "kegiatan": "Kegiatan Pembelajaran"}, inplace=True)
                st.table(df_inti)
                
            st.markdown("**3. Kegiatan Penutup**")
            for p_item in rpp.get("langkah_pembelajaran", {}).get("penutup", []):
                st.markdown(f"• {p_item}")
                
            st.markdown("#### C. Penilaian Hasil Pembelajaran")
            pen = rpp.get("penilaian", {})
            st.write(f"• **Penilaian Sikap**: {pen.get('sikap')}")
            st.write(f"• **Penilaian Pengetahuan**: {pen.get('pengetahuan')}")
            st.write(f"• **Penilaian Keterampilan**: {pen.get('keterampilan')}")

    st.markdown("---")
    st.subheader("📬 Download Perangkat Ajar Lengkap:")
    
    full_docx_buffer = create_full_docx(data)
    
    st.download_button(
        label="📦 Download Paket Lengkap Silabus & RPP Semua Pertemuan (.docx)",
        data=full_docx_buffer,
        file_name=f"RPP_{data.get('nama_kitab','').replace(' ', '_')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
        type="primary"
    )
