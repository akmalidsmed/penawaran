import streamlit as st
from datetime import date
import io
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re

# Fungsi untuk format tanggal Indonesia
def format_tanggal_indonesia(tanggal):
    bulan_dict = {
        1: "Januari", 2: "Februari", 3: "Maret",
        4: "April", 5: "Mei", 6: "Juni",
        7: "Juli", 8: "Agustus", 9: "September",
        10: "Oktober", 11: "November", 12: "Desember"
    }
    hari = tanggal.day
    bulan = bulan_dict[tanggal.month]
    tahun = tanggal.year
    return f"{hari} {bulan} {tahun}"

# Fungsi untuk format mata uang
def format_uang(angka):
    return f"Rp {angka:,.0f}".replace(",", ".")

# Konfigurasi halaman Streamlit
st.set_page_config(page_title="Generator Penawaran Harga", layout="wide")

# Judul aplikasi
st.markdown("<h1 style='text-align: center;'>Penawaran Harga</h1>", unsafe_allow_html=True)

# Input Umum
col1, col2 = st.columns(2)
with col1:
    nama_customer = st.text_input("Nama Customer")
    alamat = st.text_area("Alamat Customer")
    nomor_penawaran = st.text_input("Nomor Penawaran")

with col2:
    tanggal = st.date_input("Tanggal", value=date.today())
    nama_unit = st.text_input("Nama Unit (Tipe dan Serial Number jika ada)")

# Data Item
st.markdown("<h3 style='text-align: center;'>Item yang ditawarkan</h3>", unsafe_allow_html=True)

items = []
jumlah_item = st.number_input("Jumlah item yang ditawarkan", min_value=1, max_value=5, value=1)

for i in range(jumlah_item):
    st.markdown(f"### Item {i+1}")
    col1, col2, col3 = st.columns([1, 3, 2])
    
    with col1:
        qty = st.text_input("Qty", value="", key=f"qty_{i}")
        uom = st.text_input("UOM", value="PC", key=f"uom_{i}")
    
    with col2:
        partnumber = st.text_input("Part Number", key=f"part_{i}")
        description = st.text_area("Description", key=f"desc_{i}")
    
    with col3:
        priceperitem = st.number_input("Harga per item (Rp)", min_value=0, value=0, key=f"harga_{i}")
        try:
            total = float(qty) * priceperitem if qty else 0.0
        except:
            total = 0.0
        st.text(f"Total: {format_uang(total)}")

    items.append({
        "qty": qty,
        "uom": uom,
        "partnumber": partnumber,
        "description": description,
        "priceperitem": priceperitem,
        "price": total
    })

# Input Diskon
diskontype = st.radio("Jenis Diskon", ["Tidak ada diskon", "Diskon Persentase"], horizontal=True)
diskon_value = 0
if diskontype == "Diskon Persentase":
    diskon_value = st.number_input("Diskon (%)", min_value=0, max_value=100, value=0)

# Data PIC
pic_options = {
    "Alamas Ramadhan": "0857 7376 2820",
    "Rully Candra": "0813 1515 4142",
    "Muhammad Lukmansyah": "0821 2291 1020",
    "Denny Firmansyah": "0821 1408 0011"
}
pic = st.selectbox("Nama PIC", list(pic_options.keys()))
pic_telp = pic_options[pic]

# Ketersediaan Barang
ketersediaan = st.selectbox("Ketersediaan Barang", ["Ready stock", "Ready jika persediaan masih ada", "Indent"])

# Generate Dokumen
if st.button("\U0001F4E5 Generate Dokumen Penawaran"):
    if not nama_customer or not alamat:
        st.error("Harap isi Nama Customer dan Alamat terlebih dahulu!")
    else:
        doc = Document()
        
        # Set default font
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Arial'
        font.size = Pt(11)
        
        # Header dokumen
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run("Kepada Yth")
        run.bold = True
        
        p = doc.add_paragraph()
        run = p.add_run(nama_customer)
        run.bold = True
        
        doc.add_paragraph(alamat)
        doc.add_paragraph()
        
        # Hal: Penawaran Harga
        p = doc.add_paragraph()
        run = p.add_run("Hal: Penawaran Harga")
        run.bold = True
        run.underline = True
        
        # Nomor dan tanggal
        doc.add_paragraph(f"No: {nomor_penawaran}/JKT/SRV/AA/25\t\t\tJakarta, {format_tanggal_indonesia(tanggal)}")
        doc.add_paragraph()
        
        # Konten utama
        doc.add_paragraph(
            f"Terima kasih atas kesempatan yang telah diberikan kepada kami. Bersama "
            f"ini kami mengajukan penawaran harga item untuk unit {nama_unit} di "
            f"{nama_customer}, Adapun penawaran harga adalah sebagai berikut:"
        )
        doc.add_paragraph()
        
        # Tabel item
        table = doc.add_table(rows=1, cols=5)
        table.style = 'Table Grid'
        hdr_cells = table.rows[0].cells
        
        # Header tabel
        headers = ['Qty', 'Part Number', 'Description', 'Price per item', 'Total Price']
        for i, header in enumerate(headers):
            hdr_cells[i].text = header
            hdr_cells[i].paragraphs[0].runs[0].bold = True
            hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Isi tabel
        subtotal = 0
        for item in items:
            row_cells = table.add_row().cells
            row_cells[0].text = f"{item['qty']} {item['uom']}" if item['uom'] else item['qty']
            row_cells[1].text = item['partnumber']
            row_cells[2].text = item['description']
            row_cells[3].text = format_uang(item['priceperitem'])
            row_cells[4].text = format_uang(item['price'])
            subtotal += item['price']
            
            for cell in row_cells:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Perhitungan PPN dan Total
        if diskontype == "Diskon Persentase":
            diskon = subtotal * (diskon_value / 100)
            subtotal_after_diskon = subtotal - diskon
            ppn = subtotal_after_diskon * 0.11
            total = subtotal_after_diskon + ppn
            
            # Baris diskon
            row_cells = table.add_row().cells
            row_cells[3].text = f"Diskon {diskon_value}%"
            row_cells[4].text = format_uang(-diskon)
            for i in range(5):
                if i < 3:
                    row_cells[i].text = ""
                row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            ppn = subtotal * 0.11
            total = subtotal + ppn
        
        # Baris PPN
        row_cells = table.add_row().cells
        row_cells[3].text = "PPN 11%"
        row_cells[4].text = format_uang(ppn)
        for i in range(5):
            if i < 3:
                row_cells[i].text = ""
            row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Baris TOTAL
        row_cells = table.add_row().cells
        row_cells[3].text = "TOTAL"
        row_cells[4].text = format_uang(total)
        for i in range(5):
            if i < 3:
                row_cells[i].text = ""
            row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if i == 3:
                row_cells[i].paragraphs[0].runs[0].bold = True
            if i == 4:
                row_cells[i].paragraphs[0].runs[0].bold = True
        
        # Syarat dan ketentuan
        doc.add_paragraph()
        p = doc.add_paragraph()
        run = p.add_run("Syarat dan kondisi penawaran kami adalah :")
        run.underline = True
        run.bold = True
        
        doc.add_paragraph(f"Harga : Sudah termasuk PPN 11%")
        doc.add_paragraph(f"Ketersediaan Barang : {ketersediaan}")
        doc.add_paragraph("Pembayaran : Tunai atau transfer")
        doc.add_paragraph("Masa berlaku : 2 minggu")
        doc.add_paragraph()
        
        # Footer
        doc.add_paragraph("Demikian penawaran harga ini kami ajukan. Sambil menunggu kabar baik dari Bapak / Ibu, kami mengucapkan terima kasih.")
        doc.add_paragraph()
        doc.add_paragraph("Hormat kami,")
        doc.add_paragraph("PT. IDS Medical Systems Indonesia")
        doc.add_paragraph()
        
        p = doc.add_paragraph()
        run = p.add_run("M. Athur Yassin")
        run.underline = True
        run.bold = True
        
        doc.add_paragraph("Manager II - Engineering")
        doc.add_paragraph()
        
        doc.add_paragraph(f"PIC: {pic}")
        doc.add_paragraph(pic_telp)
        
        # Simpan ke buffer
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        # Nama file
        nama_file = f"{nomor_penawaran} {nama_customer} - {nama_unit}, {items[0]['description'] if items else ''}.docx"
        nama_file = re.sub(r'[\\/*?:"<>|]', "", nama_file)
        
        # Download button
        st.download_button(
            label="⬇️ Download Penawaran",
            data=buffer,
            file_name=nama_file,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
