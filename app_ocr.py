import streamlit as strl
import indra_ocr
import os
import re
from docx import Document
from datetime import datetime

# 1. Page Configuration
strl.set_page_config(page_title="Indra OCR Super App", page_icon="🚀", layout="wide")

# Initialize Process History (Session State)
if "riwayat_proses" not in strl.session_state:
    strl.session_state.riwayat_proses = []

# 2. Sidebar Layout (History Logs Only)
strl.sidebar.title("⏳ Riwayat Proses Sesi Ini")
if not strl.session_state.riwayat_proses:
    strl.sidebar.info("Belum ada file yang diproses.")
else:
    for item in strl.session_state.riwayat_proses:
        strl.sidebar.write(f"⏱️ {item['waktu']} - **{item['nama']}** ({item['status']})")

# 3. Load the AI Model Automatically (Threads configuration removed)
@strl.cache_resource
def muat_model_ocr():
    # Model will now auto-manage resources safely on the cloud server
    return indra_ocr.load()

try:
    pipeline = muat_model_ocr()
except Exception as e:
    strl.error(f"Gagal memuat model AI: {e}")
    strl.stop()

# 4. Main App Interface
strl.title("🚀 Indra OCR Super App (Cloud Version)")
strl.write("Aplikasi ekstraksi dokumen canggih dengan fitur pembersih teks, pencarian kata, dan ekspor Word.")

# Drag-and-Drop File Uploader
file_diunggah = strl.file_uploader("Pilih atau seret dokumen Anda ke sini (Mendukung file besar)", type=["pdf", "png", "jpg", "jpeg"])

# Helper function to clean Markdown into Plain Text
def bersihkan_ke_teks_polos(teks_md):
    teks = re.sub(r'#+\s+', '', teks_md) # Remove headers
    teks = re.sub(r'[*_~`]', '', teks)   # Remove bold/italic styling
    teks = re.sub(r'\[(.*?)\]\(.*?\)', r'\1', teks) # Clean links
    teks = re.sub(r'^-\s+', '', teks, flags=re.MULTILINE) # Clean bullet points
    teks = re.sub(r'\|\s*', ' ', teks)    # Clean table boundaries
    return teks

# Helper function to build the Microsoft Word (.docx) file
def buat_file_word(teks_konten):
    doc = Document()
    doc.add_heading('Hasil Ekstraksi Teks - Indra OCR', level=1)
    doc.add_paragraph(f"Diproses pada: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    doc.add_paragraph(teks_konten)
    
    path_word = "hasil_ocr_sementara.docx"
    doc.save(path_word)
    with open(path_word, "rb") as f:
        byte_data = f.read()
    os.remove(path_word) # Clean up temporary file from server memory
    return byte_data

if file_diunggah is not None:
    strl.success(f"File siap diproses: **{file_diunggah.name}**")
    
    # Process Button
    if strl.button("Mulai Ekstrak Teks ✨", type="primary"):
        path_sementara = os.path.join(".", file_diunggah.name)
        with open(path_sementara, "wb") as f:
            f.write(file_diunggah.getbuffer())
            
        teks_mentah_ocr = ""
        waktu_sekarang = datetime.now().strftime("%H:%M:%S")
        
        with strl.spinner("Model AI sedang memindai dan membaca dokumen Anda... Mohon tunggu..."):
            try:
                for index_halaman, halaman in pipeline.convert(path_sementara):
                    teks_mentah_ocr += f"=== HALAMAN {index_halaman + 1} ===\n{halaman.markdown}\n\n"
                
                strl.balloons()
                
                # Keep output data in memory session state to protect against UI refreshes
                strl.session_state.teks_hasil = teks_mentah_ocr
                strl.session_state.nama_file_aktif = file_diunggah.name
                
                # Log success
                strl.session_state.riwayat_proses.append({"waktu": waktu_sekarang, "nama": file_diunggah.name, "status": "Sukses"})
                
            except Exception as e:
                strl.error(f"Terjadi kesalahan: {e}")
                strl.session_state.riwayat_proses.append({"waktu": waktu_sekarang, "nama": file_diunggah.name, "status": "Gagal"})
            finally:
                if os.path.exists(path_sementara):
                    os.remove(path_sementara)

    # Output panel display
    if "teks_hasil" in strl.session_state and strl.session_state.nama_file_aktif == file_diunggah.name:
        strl.markdown("---")
        strl.header("📝 Panel Hasil Ekstraksi & Fitur Tambahan")
        
        # Text view toggles
        opsi_tampilan = strl.radio(
            "Pilih Format Tampilan Teks:",
            ("Format Asli AI (Markdown / Tabel)", "Format Bersih (Teks Polos / Plain Text)"),
            horizontal=True
        )
        
        if opsi_tampilan == "Format Bersih (Teks Polos / Plain Text)":
            teks_final = bersihkan_ke_teks_polos(strl.session_state.teks_hasil)
        else:
            teks_final = strl.session_state.teks_hasil

        # Search Bar Tool
        kata_kunci = strl.text_input("🔍 Cari kata kunci tertentu di dalam teks hasil:")
        if kata_kunci:
            kemunculan = len(re.findall(re.escape(kata_kunci), teks_final, re.IGNORECASE))
            if kemunculan > 0:
                strl.success(f"Kata '{kata_kunci}' ditemukan sebanyak **{kemunculan} kali** di dalam dokumen!")
            else:
                strl.warning(f"Kata '{kata_kunci}' tidak ditemukan di dalam dokumen.")

        # Text Code Box Block
        strl.code(teks_final, language="markdown", line_numbers=True)
        
        # Download Controls
        kolom1, kolom2 = strl.columns(2)
        
        with kolom1:
            strl.download_button(
                label="📥 Unduh sebagai File Teks (.txt)",
                data=teks_final,
                file_name=f"hasil_{os.path.splitext(file_diunggah.name)[0]}.txt",
                mime="text/plain",
                use_container_width=True
            )
            
        with kolom2:
            data_word = buat_file_word(teks_final)
            strl.download_button(
                label="📥 Unduh sebagai Dokumen Word (.docx)",
                data=data_word,
                file_name=f"hasil_{os.path.splitext(file_diunggah.name)[0]}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
