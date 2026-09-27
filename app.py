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
from google.genai import types

# ---------------------------------------------------------
# Konfigurasi Halaman Streamlit
# ---------------------------------------------------------
st.set_page_config(
    page_title="Generator Silabus & RPP Kitab",
    page_icon="📖",
    layout="wide"
)

st.title("📖 Generator Silabus & RPP Otomatis dari Kitab")
st.caption("Unggah file PDF atau masukkan teks kitab secara manual untuk merancang Perangkat Ajar Bertahap anti-limit!")

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
# Fungsi Helper: Format Dokumen Word (.docx) untuk RPP Tunggal/Semua
# ---------------------------------------------------------
def create_single_rpp_docx(data: dict, rpp: dict) -> io.BytesIO:
    doc = Document()
    
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
    
    p_pend = doc.add_paragraph()
    r_pend = p_pend.add_run("1. Kegiatan Pendahuluan")
    r_pend.bold = True
    r_pend.font.size = Pt(11)
    r_pend.font.name = 'Arial'
    
    for p_item in rpp.get("langkah_pembelajaran", {}).get("pendahuluan", []):
        p = doc.add_paragraph(f"• {p_item}")
        p.style.font.name = 'Arial'
        p.style.font.size = Pt(11)

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

    p_penut = doc.add_paragraph()
    r_penut = p_penut.add_run("\n3. Kegiatan Penutup")
    r_penut.bold = True
    r_penut.font.size = Pt(11)
    r_penut.font.name = 'Arial'
    
    for p_item in rpp.get("langkah_pembelajaran", {}).get("penutup", []):
        p = doc.add_paragraph(f"• {p_item}")
        p.style.font.name = 'Arial'
        p.style.font.size = Pt(11)

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

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ---------------------------------------------------------
# Form Input
# ---------------------------------------------------------
col1, col2 = st.columns([1, 1])

with col1:
    nama_kitab = st.text_input("Nama Kitab / Pelajaran:", placeholder="Contoh: Al-Qawaidus Sharfiyyah")
    fan_ilmu = st.text_input("Fan Ilmu:", placeholder="Contoh: Sharaf")
    tingkat_kelas = st.text_input("Tingkat / Kelas:", placeholder="Contoh: Kelas 7")
    
    col_ptm, col_waktu = st.columns(2)
    with col_ptm:
        total_pertemuan_input = st.text_input("Total Pertemuan:", placeholder="Contoh: 134 Pertemuan")
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
            placeholder="Contoh: 1-30 (Kosongkan jika semua halaman)"
        )
    else:
        uploaded_pdf = None
        kitab_text_input = st.text_area(
            "Masukkan / Paste Teks Kitab di Sini:",
            placeholder="Salin dan tempelkan teks materi/kitab di sini...",
            height=200
        )

# Inisialisasi session state rpp cache
if "rpp_cache" not in st.session_state:
    st.session_state["rpp_cache"] = {}

# ---------------------------------------------------------
# Tombol 1: Buat Silabus Saja (Ringan & Cepat)
# ---------------------------------------------------------
if st.button("🚀 1. Buat Silabus Pembelajaran Terlebih Dahulu", type="primary"):
    final_text = ""
    if metode_input == "📁 Upload File PDF":
        if not uploaded_pdf:
            st.error("Harap unggah berkas PDF kitab terlebih dahulu.")
            st.stop()
        with st.spinner("Membaca dan menganalisis halaman PDF..."):
            final_text = extract_text_from_pdf(uploaded_pdf, page_range_input)
            if not final_text.strip():
                st.error("Teks pada PDF tidak terbaca.")
                st.stop()
    else:
        if not kitab_text_input.strip():
            st.error("Harap masukkan teks kitab terlebih dahulu.")
            st.stop()
        final_text = kitab_text_input
        
    st.session_state["pdf_text"] = final_text
    st.session_state["nama_kitab"] = nama_kitab
    st.session_state["fan_ilmu"] = fan_ilmu
    st.session_state["tingkat_kelas"] = tingkat_kelas
    st.session_state["total_pertemuan"] = total_pertemuan_input
    st.session_state["alokasi_waktu"] = alokasi_waktu_input
    st.session_state["rpp_cache"] = {} # Reset cache rpp

    with st.spinner("Gemini AI sedang menyusun Silabus..."):
        prompt = f"""
        Anda adalah seorang pakar kurikulum madrasah/pesantren.
        Buatkan pemetaan Silabus secara lengkap dan proporsional untuk seluruh pertemuan dari teks kitab di bawah ini.
        
        --- DETAIL INPUT ---
        - Nama Kitab: {nama_kitab}
        - Fan Ilmu: {fan_ilmu}
        - Tingkat / Kelas: {tingkat_kelas}
        - Total Pertemuan: {total_pertemuan_input if total_pertemuan_input else 'Sesuaikan'}
        - Alokasi Waktu per Pertemuan: {alokasi_waktu_input if alokasi_waktu_input else '2 x 45 Menit'}
        --- TEKS KITAB ---
        {final_text[:30000]}  # Batasi teks agar aman
        --------------------
        
        Keluarkan respons HANYA dalam bentuk JSON valid dengan struktur berikut tanpa teks lain:
        {{
            "nama_kitab": "{nama_kitab}",
            "fan_ilmu": "{fan_ilmu}",
            "tingkat_kelas": "{tingkat_kelas}",
            "total_pertemuan": "{total_pertemuan_input}",
            "silabus": [
                {{
                    "ptm": 1,
                    "bab": "Nama bab/fasal dari teks",
                    "halaman": "Rentang halaman",
                    "alokasi_waktu": "{alokasi_waktu_input}",
                    "metode": "Metode pembelajaran",
                    "capaian": "Capaian indikator",
                    "indikator": "Indikator ketercapaian",
                    "evaluasi": "Bentuk evaluasi"
                }}
            ]
        }}
        """
        try:
            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=prompt,
                config=types.GenerateContentConfig(max_output_tokens=8192, temperature=0.7)
            )
            raw_text = response.text.strip()
            if raw_text.startswith("```"):
                lines = raw_text.splitlines()
                if lines[0].startswith("```"): lines = lines[1:]
                if lines and lines[-1].startswith("```"): lines = lines[:-1]
                raw_text = "\n".join(lines).strip()
            
            data = json.loads(raw_text)
            st.session_state["result_data"] = data
            st.success("✅ Silabus berhasil dibuat! Silakan lihat tab di bawah untuk menggenerate RPP per pertemuan.")
        except Exception as e:
            st.error(f"Gagal memproses AI / JSON: {str(e)}")

# ---------------------------------------------------------
# Tampilan Hasil & Generator RPP Per Pertemuan
# ---------------------------------------------------------
if "result_data" in st.session_state:
    data = st.session_state["result_data"]
    
    st.markdown("---")
    st.header(f"📖 Kurikulum: {data.get('nama_kitab', '')}")
    
    tab_silabus, tab_rpp = st.tabs(["📑 Silabus Pembelajaran", "📝 Generator RPP Per Pertemuan"])
    
    with tab_silabus:
        silabus_list = data.get("silabus", [])
        if silabus_list:
            df_silabus = pd.DataFrame(silabus_list)
            column_mapping = {
                "ptm": "Ptm", "bab": "Bab / Fasal", "halaman": "Halaman",
                "alokasi_waktu": "Alokasi Waktu", "metode": "Metode Klasik",
                "capaian": "Capaian Indikator", "indikator": "Indikator Ketercapaian", "evaluasi": "Bentuk Evaluasi"
            }
            df_silabus.rename(columns={k: v for k, v in column_mapping.items() if k in df_silabus.columns}, inplace=True)
            df_silabus = df_silabus.loc[:, ~df_silabus.columns.duplicated()]
            st.dataframe(df_silabus, use_container_width=True, hide_index=True)
            
    with tab_rpp:
        silabus_items = data.get("silabus", [])
        if silabus_items:
            ptm_options = [f"Pertemuan Ke-{item.get('ptm')} ({item.get('bab', 'Materi')})" for item in silabus_items]
            selected_ptm_label = st.selectbox("Pilih Pertemuan untuk Generate RPP:", ptm_options)
            
            selected_idx = ptm_options.index(selected_ptm_label)
            selected_item = silabus_items[selected_idx]
            ptm_num = selected_item.get("ptm")
            
            if st.button(f"✨ Buat RPP untuk Pertemuan Ke-{ptm_num}", type="secondary"):
                with st.spinner(f"AI sedang merancang RPP khusus untuk Pertemuan Ke-{ptm_num}..."):
                    rpp_prompt = f"""
                    Anda adalah pakar kurikulum pesantren. Buatkan SATU Rencana Pelaksanaan Pembelajaran (RPP) yang sangat detail untuk Pertemuan ke-{ptm_num} berdasarkan data silabus berikut:
                    - Bab/Materi: {selected_item.get('bab')}
                    - Halaman: {selected_item.get('halaman')}
                    - Capaian: {selected_item.get('capaian')}
                    - Indikator: {selected_item.get('indikator')}
                    - Alokasi Waktu: {selected_item.get('alokasi_waktu')}
                    
                    Gunakan teks sumber kitab jika relevan. Berikan keluaran HANYA dalam format JSON valid berikut tanpa teks lain:
                    {{
                        "ptm": {ptm_num},
                        "sekolah": "MDT / Pesantren Rufidz Tahfidz & Diniyah Indonesia",
                        "materi_pokok": "{selected_item.get('bab')}",
                        "alokasi_waktu": "{selected_item.get('alokasi_waktu')}",
                        "tujuan_pembelajaran": [
                            "Tujuan 1...",
                            "Tujuan 2..."
                        ],
                        "langkah_pembelajaran": {{
                            "pendahuluan": [
                                "Guru membuka pembelajaran dengan salam dan doa.",
                                "Guru menyampaikan apersepsi."
                            ],
                            "inti": [
                                {{"aspek": "Kegiatan Literasi", "kegiatan": "Santri menyimak..."}},
                                {{"aspek": "Critical Thinking", "kegiatan": "Santri menganalisis..."}},
                                {{"aspek": "Collaboration", "kegiatan": "Diskusi kelompok..."}},
                                {{"aspek": "Communication", "kegiatan": "Presentasi hasil..."}},
                                {{"aspek": "Creativity", "kegiatan": "Membuat kesimpulan..."}}
                            ],
                            "penutup": [
                                "Guru bersama santri menyimpulkan materi.",
                                "Doa penutup majelis."
                            ]
                        }},
                        "penilaian": {{
                            "sikap": "Observasi keaktifan dan kedisiplinan santri.",
                            "pengetahuan": "Tes lisan / Tanya jawab wazan/kaidah.",
                            "keterampilan": "Praktik membaca teks dan tashrif."
                        }}
                    }}
                    """
                    try:
                        rpp_response = client.models.generate_content(
                            model="gemini-2.5-flash",
                            contents=rpp_prompt,
                            config=types.GenerateContentConfig(max_output_tokens=4000, temperature=0.7)
                        )
                        raw_rpp = rpp_response.text.strip()
                        if raw_rpp.startswith("```"):
                            lines = raw_rpp.splitlines()
                            if lines[0].startswith("```"): lines = lines[1:]
                            if lines and lines[-1].startswith("```"): lines = lines[:-1]
                            raw_rpp = "\n".join(lines).strip()
                        
                        rpp_data = json.loads(raw_rpp)
                        st.session_state["rpp_cache"][ptm_num] = rpp_data
                        st.success(f"✅ RPP Pertemuan ke-{ptm_num} berhasil dibuat!")
                    except Exception as e:
                        st.error(f"Gagal generate RPP: {str(e)}")
            
            # Tampilkan RPP jika sudah di-generate
            if ptm_num in st.session_state["rpp_cache"]:
                rpp = st.session_state["rpp_cache"][ptm_num]
                
                st.markdown("---")
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
                    st.table(df_inti)
                    
                st.markdown("**3. Kegiatan Penutup**")
                for p_item in rpp.get("langkah_pembelajaran", {}).get("penutup", []):
                    st.markdown(f"• {p_item}")
                    
                st.markdown("#### C. Penilaian Hasil Pembelajaran")
                pen = rpp.get("penilaian", {})
                st.write(f"• **Penilaian Sikap**: {pen.get('sikap')}")
                st.write(f"• **Penilaian Pengetahuan**: {pen.get('pengetahuan')}")
                st.write(f"• **Penilaian Keterampilan**: {pen.get('keterampilan')}")
                
                # Tombol Download RPP Pertemuan Ini
                docx_buffer = create_single_rpp_docx(data, rpp)
                st.download_button(
                    label=f"📥 Download RPP Pertemuan ke-{ptm_num} (.docx)",
                    data=docx_buffer,
                    file_name=f"RPP_Pertemuan_{ptm_num}_{data.get('nama_kitab','').replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
