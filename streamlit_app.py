import streamlit as st
from datetime import date
import io
from docx import Document
from docx.shared import Inches, Pt
import os

def format_rupiah(angka):
    return "Rp. {:,.0f}".format(angka).replace(",", ".")

def format_tanggal_indonesia(tanggal):
    bulan_dict = {
        "January": "Januari", "February": "Februari", "March": "Maret",
        "April": "April", "May": "Mei", "June": "Juni",
        "July": "Juli", "August": "Agustus", "September": "September",
        "October": "Oktober", "November": "November", "December": "Desember"
    }
    hari = tanggal.day
    bulan = bulan_dict[tanggal.strftime('%B')]
    tahun = tanggal.year
    return f"{hari} {bulan} {tahun}"

# Konfigurasi Halaman
st.set_page_config(layout="wide")
st.markdown("<h1 style='text-align: center;'>Penawaran Harga</h1>", unsafe_allow_html=True)

# Data PIC
pic_options = {
    "Muhammad Lukmansyah": "0821 2291 1020",
    "Rully Candra": "0813 1515 4142",
    "Denny Firmansyah": "0821 1408 0011",
    "Alamas Ramadhan": "0857 7376 2820"
}

# Input Form di Sidebar
with st.sidebar:
    st.header("Informasi Pelanggan")
    nama_customer = st.text_input("Nama Customer")
    alamat = st.text_area("Alamat Customer")
    nomor_penawaran = st.text_input("Nomor Penawaran")
    tanggal = st.date_input("Tanggal", value=date.today())
    nama_unit = st.text_input("Nama Unit (Tipe dan Serial Number jika ada)")
    pic = st.selectbox("Nama PIC", list(pic_options.keys()))
    pic_telp = pic_options[pic]

# Input Item yang Ditawarkan
st.markdown("<h3 style='text-align: center;'>Item yang ditawarkan</h3>", unsafe_allow_html=True)
jumlah_item = st.number_input("Jumlah item yang ditawarkan", min_value=1, max_value=10, value=1, format="%d")

items = []
for i in range(jumlah_item):
    if jumlah_item > 1:
        st.markdown(f"### Item {i+1}")
        st.markdown("---")

    col1, col2 = st.columns(2)
    with col1:
        qty = st.text_input("Qty", key=f"qty_{i}")
    with col2:
        uom = st.text_input("UOM", key=f"uom_{i}")

    partnumber = st.text_input("Part Number", key=f"part_{i}")
    description = st.text_input("Description", key=f"desc_{i}")
    priceperitem = st.number_input("Harga per item (Rp)", value=0, key=f"harga_{i}", format="%d")

    try:
        total = float(qty) * priceperitem if qty else 0.0
    except:
        total = 0.0

    items.append({
        "qty": qty,
        "uom": uom,
        "partnumber": partnumber,
        "description": description,
        "priceperitem": priceperitem,
        "price": total
    })

# Keterangan Lain
st.markdown("---")
st.markdown("<h3 style='text-align: center;'>Keterangan lain-lain</h3>", unsafe_allow_html=True)

diskon_option = st.radio("Jenis Diskon", ["Tanpa diskon", "Diskon persentase (%)", "Diskon nominal (Rp)"])
diskon_value = 0
selected_items = []

if diskon_option != "Tanpa diskon":
    if jumlah_item > 1:
        diskon_scope = st.radio("Diskon berlaku untuk:", ["Semua item", "Pilih item tertentu"], index=0)
        if diskon_scope == "Pilih item tertentu":
            st.markdown("**Pilih item yang dapat diskon:**")
            cols = st.columns(3)
            for i in range(jumlah_item):
                with cols[i % 3]:
                    if st.checkbox(f"Item {i+1}", key=f"diskon_item_{i}"):
                        selected_items.append(i)
        else:
            selected_items = list(range(jumlah_item))
    else:
        selected_items = [0]

    if diskon_option == "Diskon persentase (%)":
        diskon_value = st.number_input("Besar diskon (%)", min_value=0, max_value=100, value=0, format="%d")
    else:
        diskon_value = st.number_input("Besar diskon (Rp)", min_value=0, value=0, format="%d")

opsi_ketersediaan = ["Jangan tampilkan", "Ready stock", "Ready jika persediaan masih ada", "Indent"]
ketersediaan = st.selectbox("Ketersediaan Barang", opsi_ketersediaan)

# Generate Dokumen
if st.button("\U0001F4E5 Generate Dokumen Penawaran"):
    doc = Document()

    # Header dengan logo (gunakan path relatif)
    section = doc.sections[0]
    header = section.header
    header_para = header.paragraphs[0]
    
    # Ganti dengan path gambar yang valid atau biarkan kosong
    image_path = "logo.png"  # Pastikan file ini ada di direktori yang sama
    if os.path.exists(image_path):
        try:
            header_para.add_run().add_picture(image_path, width=Inches(6.5))
        except Exception as e:
            st.warning(f"Gagal menambahkan logo: {e}")

    # Isi Dokumen
    doc.add_paragraph("Kepada Yth")
    doc.add_paragraph(nama_customer)
    doc.add_paragraph(alamat)
    
    hal = doc.add_paragraph()
    hal.add_run("Hal: Penawaran Harga").bold = True
    hal.add_run().underline = True
    
    doc.add_paragraph(f"No: {nomor_penawaran}/JKT/SRV/AA/25\t\t\tJakarta, {format_tanggal_indonesia(tanggal)}")
    doc.add_paragraph(f"Terima kasih atas kesempatan yang telah diberikan kepada kami. Bersama ini kami mengajukan penawaran harga item untuk unit {nama_unit} di {nama_customer}, sebagai berikut:\n")

    # Tabel Item
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Qty'
    hdr_cells[1].text = 'Part Number'
    hdr_cells[2].text = 'Description'
    hdr_cells[3].text = 'Price per item'
    hdr_cells[4].text = 'Total Price'

    subtotal1 = 0
    for item in items:
        row_cells = table.add_row().cells
        row_cells[0].text = f"{item['qty']}{item['uom']}"
        row_cells[1].text = item['partnumber']
        row_cells[2].text = item['description']
        row_cells[3].text = format_rupiah(item['priceperitem'])
        row_cells[4].text = format_rupiah(item['price'])
        subtotal1 += item['price']

    # Hitung Diskon
    price_diskon = 0
    if diskon_option != "Tanpa diskon" and selected_items:
        if diskon_option == "Diskon persentase (%)":
            for i in selected_items:
                price_diskon += items[i]['price'] * (diskon_value / 100)
        else:
            total_terdiskon = sum(items[i]['price'] for i in selected_items)
            for i in selected_items:
                price_diskon += (items[i]['price'] / total_terdiskon) * diskon_value
        price_diskon = round(price_diskon)

    subtotal2 = subtotal1 - price_diskon
    ppn = subtotal2 * 0.11
    total = subtotal2 + ppn

    # Tambahkan baris total
    for label, value in [("Sub Total I", subtotal1), ("Sub Total II", subtotal2), ("PPN 11%", ppn), ("TOTAL", total)]:
        row = table.add_row().cells
        row[3].text = label
        row[4].text = format_rupiah(value)

    if price_diskon > 0:
        row_disc = table.add_row().cells
        if diskon_option == "Diskon persentase (%)":
            row_disc[3].text = f"Diskon {round(diskon_value)}% ({', '.join(['Item ' + str(i+1) for i in selected_items])})"
        else:
            row_disc[3].text = f"Diskon (Rp) ({', '.join(['Item ' + str(i+1) for i in selected_items])})"
        row_disc[4].text = f"-{format_rupiah(price_diskon)}"

    # Footer
    doc.add_paragraph("\nSyarat dan ketentuan:")
    doc.add_paragraph("Harga: Sudah termasuk PPN 11%")
    doc.add_paragraph("Pembayaran: Tunai atau transfer")
    doc.add_paragraph("Masa berlaku: 2 minggu")

    if ketersediaan != "Jangan tampilkan":
        doc.add_paragraph(f"Ketersediaan Barang: {ketersediaan}")

    doc.add_paragraph("\nHormat kami,")
    doc.add_paragraph("PT. IDS Medical Systems Indonesia")
    doc.add_paragraph("M. Athur Yassin")
    doc.add_paragraph("Manager II - Engineering")
    doc.add_paragraph(pic)
    doc.add_paragraph(pic_telp)

    # Simpan ke buffer
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    # Preview dan Download
    st.success("Dokumen penawaran berhasil dibuat!")
    st.download_button(
        label="\u2B07\uFE0F Download Penawaran",
        data=buffer,
        file_name=f"Penawaran_{nama_customer}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
