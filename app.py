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
st.caption("Unggah file PDF, teks manual, atau file Silabus .docx untuk merancang Perangkat Ajar Bertahap!")

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
# Fungsi Helper: Parse Silabus dari File .docx yang Diunggah
# ---------------------------------------------------------
def parse_silabus_from_docx(docx_file) -> dict:
    doc = Document(docx_file)
    silabus_list = []
    
    if doc.tables:
        table = doc.tables[0]
        headers = [cell.text.strip().lower() for cell in table.rows[0].cells]
        
        for row in table.rows[1:]:
            row_data = [cell.text.strip() for cell in row.cells]
            if len(row_data) >= len(headers):
                item = {
                    "ptm": int(row_data[0]) if row_data[0].isdigit() else len(silabus_list) + 1,
                    "bab": row_data[1] if len(row_data) > 1 else "",
                    "halaman": row_data[2] if len(row_data) > 2 else "",
                    "alokasi_waktu": row_data[3] if len(row_data) > 3 else "",
                    "metode": row_data[4] if len(row_data) > 4 else "",
                    "capaian": row_data[5] if len(row_data) > 5 else "",
                    "indikator": row_data[6] if len(row_data) > 6 else "",
                    "evaluasi": row_data[7] if len(row_data) > 7 else ""
                }
                silabus_list.append(item)
                
    if not silabus_list:
        current_item = {}
        for p in doc.paragraphs:
            text = p.text.strip()
            if not text:
                continue
            if "pertemuan" in text.lower() or "ptm" in text.lower():
                if current_item:
                    silabus_list.append(current_item)
                current_item = {
                    "ptm": len(silabus_list) + 1,
                    "bab": text,
                    "halaman": "-",
                    "alokasi_waktu": "2 x 45 Menit",
                    "metode": "Bandongan / Sorogan / Diskusi",
                    "capaian": text,
                    "indikator": "Memahami isi materi",
                    "evaluasi": "Tes lisan"
                }
        if current_item:
            silabus_list.append(current_item)

    if not silabus_list:
        silabus_list = [{
            "ptm": 1,
            "bab": "Materi Pembelajaran dari Dokumen",
            "halaman": "1-end",
            "alokasi_waktu": "2 x 45 Menit",
            "metode": "Klasik",
            "capaian": "Pemahaman awal",
            "indikator": "Keaktifan santri",
            "evaluasi": "Tanya jawab"
        }]

    return {
        "nama_kitab": "Kitab dari Dokumen Unggahan",
        "fan_ilmu": "Umum / Diniyah",
        "tingkat_kelas": "Umum",
        "total_pertemuan": str(len(silabus_list)),
        "silabus": silabus_list
    }

# ---------------------------------------------------------
# Fungsi Helper: Membuat Dokumen Word (.docx) untuk Silabus
# ---------------------------------------------------------
def create_silabus_docx(data: dict) -> io.BytesIO:
    doc = Document()
    
    p_lembaga = doc.add_paragraph()
    p_lembaga.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_inst = p_lembaga.add_run("THE RUFIDZ INDONESIA\n")
    run_inst.bold = True
    run_inst.font.size = Pt(14)
    run_inst.font.name = 'Arial'
    
    run_title = p_lembaga.add_run("SILABUS PEMBELAJARAN KITAB")
    run_title.bold = True
    run_title.font.size = Pt(12)
    run_title.font.name = 'Arial'

    meta_text = (
        f"Nama Kitab    : {data.get('nama_kitab', '-')}\n"
        f"Fan Ilmu      : {data.get('fan_ilmu', '-')}\n"
        f"Kelas / Tahap : {data.get('tingkat_kelas', '-')}\n"
        f"Total Ptm     : {data.get('total_pertemuan', '-')}"
    )
    p_meta = doc.add_paragraph(meta_text)
    p_meta.style.font.name = 'Arial'
    p_meta.style.font.size = Pt(11)
    doc.add_paragraph("\n")

    silabus_items = data.get("silabus", [])
    if silabus_items:
        table = doc.add_table(rows=1, cols=8)
        table.style = 'Table Grid'
        headers = ["Ptm", "Bab / Fasal", "Halaman", "Alokasi Waktu", "Metode", "Capaian", "Indikator", "Evaluasi"]
        
        hdr_cells = table.rows[0].cells
        for i, h_text in enumerate(headers):
            hdr_cells[i].text = h_text
            hdr_cells[i].paragraphs[0].runs[0].font.bold = True
            hdr_cells[i].paragraphs[0].runs[0].font.size = Pt(9)
            hdr_cells[i].paragraphs[0].runs[0].font.name = 'Arial'

        for item in silabus_items:
            row_cells = table.add_row().cells
            row_cells[0].text = str(item.get("ptm", ""))
            row_cells[1].text = str(item.get("bab", ""))
            row_cells[2].text = str(item.get("halaman", ""))
            row_cells[3].text = str(item.get("alokasi_waktu", ""))
            row_cells[4].text = str(item.get("metode", ""))
            row_cells[5].text = str(item.get("capaian", ""))
            row_cells[6].text = str(item.get("indikator", ""))
            row_cells[7].text = str(item.get("evaluasi", ""))
            
            for cell in row_cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(8.5)
                        run.font.name = 'Arial'

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ---------------------------------------------------------
# Fungsi Helper: Membuat File Excel (.xlsx) untuk Jurnal Harian KBM
# ---------------------------------------------------------
def create_jurnal_excel(data: dict) -> io.BytesIO:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        fan_text = data.get('fan_ilmu', 'SHARAF').upper()
        kitab_text = data.get('nama_kitab', '')
        
        rows_data = []
        rows_data.append([f"JURNAL PEMBELAJARAN {fan_text} THE RUFIDZ INDONESIA", "", "", "", "", "", ""])
        rows_data.append([f"SEMESTER GANJIL T.A 2026/2027 (Kitab: {kitab_text})", "", "", "", "", "", ""])
        rows_data.append(["", "", "", "", "", "", ""])
        
        rows_data.append([
            "MATERI POKOK", 
            "ALOKASI WAKTU", 
            "Tanggal", 
            "NAMA PENGAJAR/PENGGANTI", 
            "CATATAN MATERI\n(Diisi jika diperlukan)", 
            "TTD", 
            "Keterangan"
        ])
        
        silabus_items = data.get("silabus", [])
        for item in silabus_items:
            rows_data.append([
                str(item.get("bab", "")),
                str(item.get("alokasi_waktu", "")),
                "......./....../2026",
                "",
                "",
                "",
                ""
            ])
            
        df_jurnal = pd.DataFrame(rows_data)
        df_jurnal.to_excel(writer, sheet_name='Jurnal KBM', index=False, header=False)
        
    output.seek(0)
    return output

# ---------------------------------------------------------
# Fungsi Helper: Format Dokumen Word (.docx) untuk RPP Tunggal
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

    p_tujuan_head = doc.add_paragraph()
    r_tujuan = p_tujuan_head.add_run("A. Tujuan Pembelajaran")
    r_tujuan.bold = True
    r_tujuan.font.size = Pt(11)
    r_tujuan.font.name = 'Arial'
    
    for t in rpp.get("tujuan_pembelajaran", []):
        p = doc.add_paragraph(f"• {t}")
        p.style.font.name = 'Arial'
        p.style.font.size = Pt(11)

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
    
    st_p3 = doc.add_paragraph(f"• Penilaian Keterampilan:
