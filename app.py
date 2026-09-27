import io
import json
import os
import re
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
# Fungsi Helper: Membaca Teks PDF Berdasarkan Rentang Halaman
# ---------------------------------------------------------
def parse_page_range(range_str: str, total_pages: int):
    """Mengubah string seperti '1-5' atau '1,3,5-7' menjadi set indeks halaman (0-indexed)."""
    if not range_str.strip():
        return set(range(total_pages))
    
    selected_pages = set()
    parts = range_str.split(',')
    for part in parts:
        part = part.strip()
        if '-' in part:
            sub = part.split('-')
            if len(sub) == 2 and sub[0].isdigit() and sub[1].isdigit():
                start = max(1, int(sub[0]))
                end = min(total_pages, int(sub[1]))
                for p in range(start, end + 1):
                    selected_pages.add(p - 1)
        elif part.isdigit():
            p = int(part)
            if 1 <= p <= total_pages:
                selected_pages.add(p - 1)
    
    return sorted(list(selected_pages))

def extract_text_from_pdf(pdf_file, page_range_str: str = "") -> str:
    reader = pypdf.PdfReader(pdf_file)
    total_pages = len(reader.pages)
    
    page_indices = parse_page_range(page_range_str, total_pages)
    if not page_indices:
        page_indices = list(range(total_pages))
        
    text = ""
    for idx in page_indices:
        extracted = reader.pages[idx].extract_text()
        if extracted:
            text += f"\n--- Halaman {idx + 1} ---\n" + extracted
    return text

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

        # A. Tujuan
        p_tujuan_head = doc.add_paragraph()
        r_tujuan = p_tujuan_head.add_run("A. Tujuan Pembelajaran")
        r_tujuan.bold = True
        r_tujuan.font.size = Pt(11)
        r_tujuan.font.name = 'Arial'
        
        for t in rpp.get("tujuan_pembelajaran", []):
            p = doc.add_paragraph(f"• {t}")
            p.style.font.name = 'Arial'
            p.style.font.size = Pt(11)

        # B. Langkah Pembelajaran
        p_langkah_head = doc.add_paragraph()
        r_langkah = p_langkah_head.add_run("B. Langkah-Langkah Pembelajaran")
        r_langkah.bold = True
        r_langkah.font.size = Pt(11)
        r_langkah.font.name = 'Arial'
        
        # 1. Pendahuluan
        p_pend = doc.add_paragraph()
        r_pend = p_pend.add_run("1. Kegiatan Pendahuluan")
        r_pend.bold = True
        r_pend.font.size = Pt(11)
        r_pend.font.name = 'Arial'
        
        for p_item in rpp.get("langkah_pembelajaran", {}).get("pendahuluan", []):
            p = doc.add_paragraph(f"• {p_item}")
            p.style.font.name = 'Arial'
            p.style.font.size = Pt(11)

        # 2. Inti
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

        # 3. Penutup
        p_penut = doc.add_paragraph()
        r_penut = p_penut.add_run("\n3. Kegiatan Penutup")
        r_penut.bold = True
        r_penut.font.size = Pt(11)
        r_penut.font.name = 'Arial'
        
        for p_item in rpp.get("langkah_pembelajaran", {}).get("penutup", []):
            p = doc.add_paragraph(f"• {p_item}")
            p.style.font.name = 'Arial'
            p.style.font.size = Pt(11)

        # C. Penilaian
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
# Form Input
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
            "📄 Rentang Halaman yang Ingin Dianalisis (Opsional):",
            placeholder="Contoh: 1-5 atau 3,5,7-10 (Kosongkan jika semua halaman)"
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
        with st.spinner("Membaca dan menganalisis halaman PDF yang dipilih..."):
            extracted = extract_text_from_pdf(uploaded_pdf, page_range_input)
            if not extracted.strip():
                st.error("Teks pada PDF tidak terbaca. Pastikan PDF bukan hasil scan/gambar.")
                st.stop()
            final_text = extracted
    else:
        if not kitab_text_input.strip():
            st.error("Harap masukkan teks kitab terlebih dahulu.")
            st.stop()
        final_text = kitab_text_input
        
    pdf_text = final_text 

    with st.spinner("Gemini AI sedang menyusun Silabus dan RPP dari teks yang Anda berikan..."):
        prompt = f"""
        Anda adalah seorang pakar kurikulum madrasah/pesantren.
        
        Tugas Anda: Analisis TEKS KITAB yang diberikan secara mendalam dan buatkan Silabus serta RPP 1 Lembar UNTUK TIAP PERTEMUAN.
        Pastikan materi yang Anda buat BENAR-BENAR bersumber dari TEKS KITAB di bawah ini, BUKAN dari pengetahuan umum.
        
        --- DETAIL INPUT ---
        - Nama Kitab: {nama_kitab}
        - Fan Ilmu: {fan_ilmu}
        - Tingkat / Kelas: {tingkat_kelas}
        - Total Pertemuan: {total_pertemuan_input if total_pertemuan_input else 'Sesuaikan dengan cakupan materi'}
        - Alokasi Waktu per Pertemuan: {alokasi_waktu_input if alokasi_waktu_input else '2 x 45 Menit'}
        --- TEKS KITAB ---
        {pdf_text}
        --------------------
        
        Keluarkan respons HANYA dalam bentuk JSON valid. Gunakan struktur di bawah ini sebagai TEMPLATE. 
        Ganti teks berawalan "[" dan diakhiri "]" dengan hasil analisis Anda yang SEBENARNYA dari teks kitab!
        
        {{
            "nama_kitab": "{nama_kitab}",
            "fan_ilmu": "{fan_ilmu}",
            "tingkat_kelas": "{tingkat_kelas}",
            "total_pertemuan": "{total_pertemuan_input if total_pertemuan_input else '6 Pertemuan'}",
            "silabus": [
                {{
                    "ptm": 1,
                    "bab": "[Isi dengan nama bab/fasal dari teks kitab pada halaman terkait]",
                    "halaman": "[Rentang halaman yang dibahas di ptm ini]",
                    "alokasi_waktu": "{alokasi_waktu_input if alokasi_waktu_input else '2 x 45 Menit'}",
                    "metode": "[Contoh: Ceramah Interaktif]",
                    "capaian": "[Rumuskan capaian berdasarkan materi kitab tersebut]",
                    "indikator": "[Rumuskan indikator keberhasilan]",
                    "evaluasi": "[Tentukan evaluasinya]"
                }}
            ],
            "rpp_list": [
                {{
                    "ptm": 1,
                    "sekolah": "MDT / Pesantren Rufidz Tahfidz & Diniyah Indonesia",
                    "materi_pokok": "[Isi dengan materi spesifik pada pertemuan ini]",
                    "alokasi_waktu": "{alokasi_waktu_input if alokasi_waktu_input else '2 x 45 Menit'}",
                    "tujuan_pembelajaran": [
                        "[Tujuan 1 berdasarkan teks]",
                        "[Tujuan 2 berdasarkan teks]"
                    ],
                    "langkah_pembelajaran": {{
                        "pendahuluan": [
                            "Guru membuka pembelajaran dengan salam, istighfar, dan doa bersama.",
                            "Guru menyampaikan apersepsi serta tujuan pembelajaran materi ini."
                        ],
                        "inti": [
                            {{"aspek": "Kegiatan Literasi", "kegiatan": "[Jelaskan kegiatan santri membaca materi kitab ini]"}},
                            {{"aspek": "Critical Thinking", "kegiatan": "[Jelaskan kegiatan berpikir kritis dari materi ini]"}},
                            {{"aspek": "Collaboration", "kegiatan": "[Jelaskan aktivitas kerja kelompok]"}},
                            {{"aspek": "Communication", "kegiatan": "[Jelaskan aktivitas komunikasi santri]"}},
                            {{"aspek": "Creativity", "kegiatan": "[Jelaskan aktivitas kreatif santri]"}}
                        ],
                        "penutup": [
                            "Guru bersama santri menyimpulkan materi.",
                            "Kegiatan ditutup dengan doa kafaratul majlis."
                        ]
                    }},
                    "penilaian": {{
                        "sikap": "[Metode penilaian sikap]",
                        "pengetahuan": "[Metode penilaian pengetahuan]",
                        "keterampilan": "[Metode penilaian keterampilan]"
                    }}
                }}
            ]
        }}
        """
        
        try:
            response = client.models.generate_content(
                model="gemini-2.5-flash",
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
            
            # Pemetaan nama kolom
            column_mapping = {
                "ptm": "Ptm",
                "bab": "Bab / Fasal",
                "halaman": "Halaman",
                "alokasi_waktu": "Alokasi Waktu",
                "metode": "Metode Klasik",
                "capaian": "Capaian Indikator",
                "indikator": "Indikator Ketercapaian",
                "evaluasi": "Bentuk Evaluasi"
            }
            
            # Hanya ubah kolom yang ada dan belum memiliki nama baru
            df_silabus.rename(columns={k: v for k, v in column_mapping.items() if k in df_silabus.columns and v not in df_silabus.columns}, inplace=True)
            
            # Eliminasi kolom yang terduplikasi secara aman
            df_silabus = df_silabus.loc[:, ~df_silabus.columns.duplicated()]
            
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
                if "aspek" in df_inti.columns and "kegiatan" in df_inti.columns:
                    df_inti.rename(columns={"aspek": "Aspek", "kegiatan": "Kegiatan Pembelajaran"}, inplace=True)
                df_inti = df_inti.loc[:, ~df_inti.columns.duplicated()]
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
