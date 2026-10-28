import streamlit as strl
import indra_ocr
import os
import re
import io
from docx import Document
from datetime import datetime
from st_copy_to_clipboard import st_copy_to_clipboard
from pdf2image import convert_from_bytes
from PIL import Image

# 1. Page Configuration & Layout
strl.set_page_config(page_title="Aplikasi Indra OCR", page_icon="✨", layout="wide")

# Initialize Process History (Session State)
if "riwayat_proses" not in strl.session_state:
    strl.session_state.riwayat_proses = []

# 2. Multi-Language Interface Toggle
if "bahasa" not in strl.session_state:
    strl.session_state.bahasa = "ID"

strl.sidebar.title("🌐 Language / Bahasa")
pilihan_lang = strl.sidebar.radio(
    "Pilih Bahasa Antarmuka / Select Language:",
    ("Bahasa Indonesia", "English"),
    index=0 if strl.session_state.bahasa == "ID" else 1
)
strl.session_state.bahasa = "ID" if pilihan_lang == "Bahasa Indonesia" else "EN"

# Bilingual UI Translation Dictionary
teks_app = {
    "ID": {
        "title": "✨ 🔍 Aplikasi Indra OCR",
        "desc": "Alat bantu digital untuk memindai dokumen dan mengubah gambar atau file PDF menjadi teks siap pakai secara instan.",
        "uploader_label": "Pilih atau seret dokumen Anda ke sini (Mendukung file besar)",
        "ready": "File siap diproses: ",
        "btn_proses": "Mulai Ekstrak Teks ✨",
        "spinner": "Model AI sedang memindai dokumen Anda... Mohon tunggu...",
        "riwayat_title": "⏳ Riwayat Proses Sesi Ini",
        "riwayat_empty": "Belum ada file yang diproses.",
        "panel_title": "📝 Panel Hasil Ekstraksi & Fitur Tambahan",
        "format_label": "Pilih Format Tampilan Teks:",
        "format_asli": "Format Asli AI (Markdown / Tabel)",
        "format_bersih": "Format Bersih (Teks Polos / Plain Text)",
        "search_label": "🔍 Cari kata kunci tertentu (Hasil akan disorot kuning & diarahkan otomatis):",
        "search_found": "ditemukan sebanyak",
        "search_times": "kali di dalam dokumen!",
        "search_not_found": "tidak ditemukan di dalam dokumen.",
        "stats_title": "📊 Dashboard Statistik Teks",
        "stats_words": "Total Kata",
        "stats_chars": "Total Karakter",
        "stats_time": "Estimasi Waktu Baca",
        "stats_minute": "Menit",
        "btn_txt": "📥 Unduh sebagai File Teks (.txt)",
        "btn_word": "📥 Unduh sebagai Dokumen Word (.docx)",
        "copy_label": "📋 Salin Teks ke Clipboard:",
        "preview_title": "👁️ Pratinjau Dokumen Asli (Halaman 1)",
        "preview_not_supported": "Pratinjau tidak tersedia atau file tidak dapat dibaca.",
        "success_msg": "Sukses",
        "failed_msg": "Gagal"
    },
    "EN": {
        "title": "✨ 🔍 Aplikasi Indra OCR",
        "desc": "A digital utility to scan files and instantly convert your images or PDF documents into editable text blocks.",
        "uploader_label": "Choose or drag your document here (Supports large files)",
        "ready": "File ready to process: ",
        "btn_proses": "Start Extracting Text ✨",
        "spinner": "AI Model is scanning your document... Please wait...",
        "riwayat_title": "⏳ Process History This Session",
        "riwayat_empty": "No files processed yet.",
        "panel_title": "📝 Extraction Results & Additional Features",
        "format_label": "Choose Text View Format:",
        "format_asli": "Original AI Format (Markdown / Table)",
        "format_bersih": "Clean Format (Plain Text)",
        "search_label": "🔍 Search specific keywords (Will highlight yellow & auto-scroll):",
        "search_found": "found",
        "search_times": "times inside the document!",
        "search_not_found": "not found inside the document.",
        "stats_title": "📊 Text Statistics Dashboard",
        "stats_words": "Total Words",
        "stats_chars": "Total Characters",
        "stats_time": "Est. Reading Time",
        "stats_minute": "Min",
        "btn_txt": "📥 Download as Text File (.txt)",
        "btn_word": "📥 Download as Word Document (.docx)",
        "copy_label": "📋 Copy Text to Clipboard:",
        "preview_title": "👁️ Original Document Preview (Page 1)",
        "preview_not_supported": "Preview not supported or file unreadable.",
        "success_msg": "Success",
        "failed_msg": "Failed"
    }
}

L = teks_app[strl.session_state.bahasa]

# Render Process History Logs in Sidebar
strl.sidebar.markdown("---")
strl.sidebar.subheader(L["riwayat_title"])
if not strl.session_state.riwayat_proses:
    strl.sidebar.info(L["riwayat_empty"])
else:
    for item in strl.session_state.riwayat_proses:
        status_txt = L["success_msg"] if item['status'] == "Sukses" else L["failed_msg"]
        strl.sidebar.write(f"⏱️ {item['waktu']} - **{item['nama']}** ({status_txt})")

# 3. Load the AI Model Automatically
@strl.cache_resource
def muat_model_ocr():
    return indra_ocr.load()

try:
    pipeline = muat_model_ocr()
except Exception as e:
    strl.error(f"Error loading AI model: {e}")
    strl.stop()

# 4. Main Application View
strl.title(L["title"])
strl.write(L["desc"])

file_diunggah = strl.file_uploader(L["uploader_label"], type=["pdf", "png", "jpg", "jpeg"])

def bersihkan_ke_teks_polos(teks_md):
    teks = re.sub(r'#+\s+', '', teks_md)
    teks = re.sub(r'[*_~`]', '', teks)
    teks = re.sub(r'\[(.*?)\]\(.*?\)', r'\1', teks)
    teks = re.sub(r'^-\s+', '', teks, flags=re.MULTILINE)
    teks = re.sub(r'\|\s*', ' ', teks)
    return teks

def buat_file_word(teks_konten):
    doc = Document()
    doc.add_heading('Hasil Ekstraksi Teks - Indra OCR', level=1)
    doc.add_paragraph(f"Diproses pada: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    doc.add_paragraph(teks_konten)
    path_word = "hasil_ocr_sementara.docx"
    doc.save(path_word)
    with open(path_word, "rb") as f:
        byte_data = f.read()
    os.remove(path_word)
    return byte_data

# PERBAIKAN TOTAL: Mengubah PDF/Gambar langsung menjadi penampil elemen visual internal Streamlit yang aman
def tampilkan_pratinjau(file_data, mime_type):
    try:
        file_bytes = file_data.read()
        file_data.seek(0)
        
        if "pdf" in mime_type:
            # Mengonversi halaman pertama PDF menjadi gambar PNG di memori cloud
            gambar_halaman = convert_from_bytes(file_bytes, first_page=1, last_page=1)
            if gambar_halaman:
                strl.image(gambar_halaman[0], caption="Pratinjau Halaman 1", use_container_width=True)
            else:
                strl.info(L["preview_not_supported"])
        elif "image" in mime_type:
            # Jika masukan berupa gambar biasa (PNG/JPG) langsung tampilkan
            gambar_asli = Image.open(io.BytesIO(file_bytes))
            strl.image(gambar_asli, use_container_width=True)
        else:
            strl.info(L["preview_not_supported"])
    except Exception as e:
        strl.error(f"Gagal memuat pratinjau: {e}")

if file_diunggah is not None:
    strl.success(f"{L['ready']} **{file_diunggah.name}**")
    
    if strl.button(L["btn_proses"], type="primary"):
        path_sementara = os.path.join(".", file_diunggah.name)
        with open(path_sementara, "wb") as f:
            f.write(file_diunggah.getbuffer())
            
        teks_mentah_ocr = ""
        waktu_sekarang = datetime.now().strftime("%H:%M:%S")
        
        with strl.spinner(L["spinner"]):
            try:
                for index_halaman, halaman in pipeline.convert(path_sementara):
                    teks_mentah_ocr += f"=== HALAMAN {index_halaman + 1} ===\n{halaman.markdown}\n\n"
                
                strl.balloons()
                strl.session_state.teks_hasil = teks_mentah_ocr
                strl.session_state.nama_file_aktif = file_diunggah.name
                strl.session_state.riwayat_proses.append({"waktu": waktu_sekarang, "nama": file_diunggah.name, "status": "Sukses"})
                
            except Exception as e:
                strl.error(f"Error: {e}")
                strl.session_state.riwayat_proses.append({"waktu": waktu_sekarang, "nama": file_diunggah.name, "status": "Gagal"})
            finally:
                if os.path.exists(path_sementara):
                    os.remove(path_sementara)

    # Workspace Side-by-Side 
    if "teks_hasil" in strl.session_state and strl.session_state.nama_file_aktif == file_diunggah.name:
        strl.markdown("---")
        
        kolom_kiri, kolom_kanan = strl.columns(2)
        
        with kolom_kiri:
            strl.subheader(L["preview_title"])
            tampilkan_pratinjau(file_diunggah, file_diunggah.type)
            
        with kolom_kanan:
            strl.subheader(L["panel_title"])
            
            opsi_tampilan = strl.radio(L["format_label"], (L["format_asli"], L["format_bersih"]), horizontal=True)
            if opsi_tampilan == L["format_bersih"]:
                teks_final = bersihkan_ke_teks_polos(strl.session_state.teks_hasil)
            else:
                teks_final = strl.session_state.teks_hasil

            # Dashboard Statistik Teks
            strl.markdown(f"#### {L['stats_title']}")
            jumlah_kata = len(teks_final.split())
            jumlah_karakter = len(teks_final)
            waktu_baca = max(1, round(jumlah_kata / 200))
            
            c1, c2, c3 = strl.columns(3)
            c1.metric(L["stats_words"], f"{jumlah_kata:,}")
            c2.metric(L["stats_chars"], f"{jumlah_karakter:,}")
            c3.metric(L["stats_time"], f"{waktu_baca} {L['stats_minute']}")
            strl.markdown("---")

            # Pencarian Kata Kunci Canggih (Highlight & Auto-Scroll)
            kata_kunci = strl.text_input(L["search_label"])
            
            teks_tampilan_html = teks_final.replace("\n", "<br>")
            
            if kata_kunci:
                kemunculan = len(re.findall(re.escape(kata_kunci), teks_final, re.IGNORECASE))
                if kemunculan > 0:
import streamlit as strl
import indra_ocr
import os
import re
import io
from docx import Document
from datetime import datetime
from st_copy_to_clipboard import st_copy_to_clipboard
from pdf2image import convert_from_bytes
from PIL import Image

# 1. Page Configuration & Layout
strl.set_page_config(page_title="Aplikasi Indra OCR", page_icon="✨", layout="wide")

# Initialize Process History (Session State)
if "riwayat_proses" not in strl.session_state:
    strl.session_state.riwayat_proses = []

# 2. Multi-Language Interface Toggle
if "bahasa" not in strl.session_state:
    strl.session_state.bahasa = "ID"

strl.sidebar.title("🌐 Language / Bahasa")
pilihan_lang = strl.sidebar.radio(
    "Pilih Bahasa Antarmuka / Select Language:",
    ("Bahasa Indonesia", "English"),
    index=0 if strl.session_state.bahasa == "ID" else 1
)
strl.session_state.bahasa = "ID" if pilihan_lang == "Bahasa Indonesia" else "EN"

# Bilingual UI Translation Dictionary
teks_app = {
    "ID": {
        "title": "✨ 🔍 Aplikasi Indra OCR",
        "desc": "Alat bantu digital untuk memindai dokumen dan mengubah gambar atau file PDF menjadi teks siap pakai secara instan.",
        "uploader_label": "Pilih atau seret dokumen Anda ke sini (Mendukung file besar)",
        "ready": "File siap diproses: ",
        "btn_proses": "Mulai Ekstrak Teks ✨",
        "spinner": "Model AI sedang memindai dokumen Anda... Mohon tunggu...",
        "riwayat_title": "⏳ Riwayat Proses Sesi Ini",
        "riwayat_empty": "Belum ada file yang diproses.",
        "panel_title": "📝 Panel Hasil Ekstraksi & Fitur Tambahan",
        "format_label": "Pilih Format Tampilan Teks:",
        "format_asli": "Format Asli AI (Markdown / Tabel)",
        "format_bersih": "Format Bersih (Teks Polos / Plain Text)",
        "search_label": "🔍 Cari kata kunci tertentu (Hasil akan disorot kuning & diarahkan otomatis):",
        "search_found": "ditemukan sebanyak",
        "search_times": "kali di dalam dokumen!",
        "search_not_found": "tidak ditemukan di dalam dokumen.",
        "stats_title": "📊 Dashboard Statistik Teks",
        "stats_words": "Total Kata",
        "stats_chars": "Total Karakter",
        "stats_time": "Estimasi Waktu Baca",
        "stats_minute": "Menit",
        "btn_txt": "📥 Unduh sebagai File Teks (.txt)",
        "btn_word": "📥 Unduh sebagai Dokumen Word (.docx)",
        "copy_label": "📋 Salin Teks ke Clipboard:",
        "preview_title": "👁️ Pratinjau Dokumen Asli (Halaman 1)",
        "preview_not_supported": "Pratinjau tidak tersedia atau file tidak dapat dibaca.",
        "success_msg": "Sukses",
        "failed_msg": "Gagal"
    },
    "EN": {
        "title": "✨ 🔍 Aplikasi Indra OCR",
        "desc": "A digital utility to scan files and instantly convert your images or PDF documents into editable text blocks.",
        "uploader_label": "Choose or drag your document here (Supports large files)",
        "ready": "File ready to process: ",
        "btn_proses": "Start Extracting Text ✨",
        "spinner": "AI Model is scanning your document... Please wait...",
        "riwayat_title": "⏳ Process History This Session",
        "riwayat_empty": "No files processed yet.",
        "panel_title": "📝 Extraction Results & Additional Features",
        "format_label": "Choose Text View Format:",
        "format_asli": "Original AI Format (Markdown / Table)",
        "format_bersih": "Clean Format (Plain Text)",
        "search_label": "🔍 Search specific keywords (Will highlight yellow & auto-scroll):",
        "search_found": "found",
        "search_times": "times inside the document!",
        "search_not_found": "not found inside the document.",
        "stats_title": "📊 Text Statistics Dashboard",
        "stats_words": "Total Words",
        "stats_chars": "Total Characters",
        "stats_time": "Est. Reading Time",
        "stats_minute": "Min",
        "btn_txt": "📥 Download as Text File (.txt)",
        "btn_word": "📥 Download as Word Document (.docx)",
        "copy_label": "📋 Copy Text to Clipboard:",
        "preview_title": "👁️ Original Document Preview (Page 1)",
        "preview_not_supported": "Preview not supported or file unreadable.",
        "success_msg": "Success",
        "failed_msg": "Failed"
    }
}

L = teks_app[strl.session_state.bahasa]

# Render Process History Logs in Sidebar
strl.sidebar.markdown("---")
strl.sidebar.subheader(L["riwayat_title"])
if not strl.session_state.riwayat_proses:
    strl.sidebar.info(L["riwayat_empty"])
else:
    for item in strl.session_state.riwayat_proses:
        status_txt = L["success_msg"] if item['status'] == "Sukses" else L["failed_msg"]
        strl.sidebar.write(f"⏱️ {item['waktu']} - **{item['nama']}** ({status_txt})")

# 3. Load the AI Model Automatically
@strl.cache_resource
def muat_model_ocr():
    return indra_ocr.load()

try:
    pipeline = muat_model_ocr()
except Exception as e:
    strl.error(f"Error loading AI model: {e}")
    strl.stop()

# 4. Main Application View
strl.title(L["title"])
strl.write(L["desc"])

file_diunggah = strl.file_uploader(L["uploader_label"], type=["pdf", "png", "jpg", "jpeg"])

def bersihkan_ke_teks_polos(teks_md):
    teks = re.sub(r'#+\s+', '', teks_md)
    teks = re.sub(r'[*_~`]', '', teks)
    teks = re.sub(r'\[(.*?)\]\(.*?\)', r'\1', teks)
    teks = re.sub(r'^-\s+', '', teks, flags=re.MULTILINE)
    teks = re.sub(r'\|\s*', ' ', teks)
    return teks

def buat_file_word(teks_konten):
    doc = Document()
    doc.add_heading('Hasil Ekstraksi Teks - Indra OCR', level=1)
    doc.add_paragraph(f"Diproses pada: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    doc.add_paragraph(teks_konten)
    path_word = "hasil_ocr_sementara.docx"
    doc.save(path_word)
    with open(path_word, "rb") as f:
        byte_data = f.read()
    os.remove(path_word)
    return byte_data

# PERBAIKAN TOTAL: Mengubah PDF/Gambar langsung menjadi penampil elemen visual internal Streamlit yang aman
def tampilkan_pratinjau(file_data, mime_type):
    try:
        file_bytes = file_data.read()
        file_data.seek(0)
        
        if "pdf" in mime_type:
            # Mengonversi halaman pertama PDF menjadi gambar PNG di memori cloud
            gambar_halaman = convert_from_bytes(file_bytes, first_page=1, last_page=1)
            if gambar_halaman:
                strl.image(gambar_halaman[0], caption="Pratinjau Halaman 1", use_container_width=True)
            else:
                strl.info(L["preview_not_supported"])
        elif "image" in mime_type:
            # Jika masukan berupa gambar biasa (PNG/JPG) langsung tampilkan
            gambar_asli = Image.open(io.BytesIO(file_bytes))
            strl.image(gambar_asli, use_container_width=True)
        else:
            strl.info(L["preview_not_supported"])
    except Exception as e:
        strl.error(f"Gagal memuat pratinjau: {e}")

if file_diunggah is not None:
    strl.success(f"{L['ready']} **{file_diunggah.name}**")
    
    if strl.button(L["btn_proses"], type="primary"):
        path_sementara = os.path.join(".", file_diunggah.name)
        with open(path_sementara, "wb") as f:
            f.write(file_diunggah.getbuffer())
            
        teks_mentah_ocr = ""
        waktu_sekarang = datetime.now().strftime("%H:%M:%S")
        
        with strl.spinner(L["spinner"]):
            try:
                for index_halaman, halaman in pipeline.convert(path_sementara):
                    teks_mentah_ocr += f"=== HALAMAN {index_halaman + 1} ===\n{halaman.markdown}\n\n"
                
                strl.balloons()
                strl.session_state.teks_hasil = teks_mentah_ocr
                strl.session_state.nama_file_aktif = file_diunggah.name
                strl.session_state.riwayat_proses.append({"waktu": waktu_sekarang, "nama": file_diunggah.name, "status": "Sukses"})
                
            except Exception as e:
                strl.error(f"Error: {e}")
                strl.session_state.riwayat_proses.append({"waktu": waktu_sekarang, "nama": file_diunggah.name, "status": "Gagal"})
            finally:
                if os.path.exists(path_sementara):
                    os.remove(path_sementara)

    # Workspace Side-by-Side 
    if "teks_hasil" in strl.session_state and strl.session_state.nama_file_aktif == file_diunggah.name:
        strl.markdown("---")
        
        kolom_kiri, kolom_kanan = strl.columns(2)
        
        with kolom_kiri:
            strl.subheader(L["preview_title"])
            tampilkan_pratinjau(file_diunggah, file_diunggah.type)
            
        with kolom_kanan:
            strl.subheader(L["panel_title"])
            
            opsi_tampilan = strl.radio(L["format_label"], (L["format_asli"], L["format_bersih"]), horizontal=True)
            if opsi_tampilan == L["format_bersih"]:
                teks_final = bersihkan_ke_teks_polos(strl.session_state.teks_hasil)
            else:
                teks_final = strl.session_state.teks_hasil

            # Dashboard Statistik Teks
            strl.markdown(f"#### {L['stats_title']}")
            jumlah_kata = len(teks_final.split())
            jumlah_karakter = len(teks_final)
            waktu_baca = max(1, round(jumlah_kata / 200))
            
            c1, c2, c3 = strl.columns(3)
            c1.metric(L["stats_words"], f"{jumlah_kata:,}")
            c2.metric(L["stats_chars"], f"{jumlah_karakter:,}")
            c3.metric(L["stats_time"], f"{waktu_baca} {L['stats_minute']}")
            strl.markdown("---")

            # Pencarian Kata Kunci Canggih (Highlight & Auto-Scroll)
            kata_kunci = strl.text_input(L["search_label"])
            
            teks_tampilan_html = teks_final.replace("\n", "<br>")
            
            if kata_kunci:
                kemunculan = len(re.findall(re.escape(kata_kunci), teks_final, re.IGNORECASE))
                if kemunculan > 0:
