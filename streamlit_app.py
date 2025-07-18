import streamlit as st
from datetime import date
import io
from docx import Document

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

st.markdown("<h1 style='text-align: center;'>Penawaran Harga</h1>", unsafe_allow_html=True)

# Data PIC
pic_options = {
    "Alamas Ramadhan": "0857 7376 2820",
    "Rully Candra": "0813 1515 4142",
    "Muhammad Lukmansyah": "0821 2291 1020",
    "Denny Firmansyah": "0821 1408 0011"
}

# Input Umum
nama_customer = st.text_input("Nama Customer")
alamat = st.text_area("Alamat Customer")
nomor_penawaran = st.text_input("Nomor Penawaran")
tanggal = st.date_input("Tanggal", value=date.today())
nama_unit = st.text_input("Nama Unit (Tipe dan Serial Number jika ada)")

# Input Item
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

# Keterangan Lain-lain
st.markdown("---")
st.markdown("<h3 style='text-align: center;'>Keterangan lain-lain</h3>", unsafe_allow_html=True)

# Input diskon
diskon_option = st.radio("Diskon", 
                        ["Tanpa diskon", "Diskon persentase (%)", "Diskon nominal (Rp)"], 
                        index=0)

diskon_value = 0
if diskon_option == "Diskon persentase (%)":
    diskon_value = st.number_input("Besar diskon (%)", min_value=0, max_value=100, value=0, format="%d")
elif diskon_option == "Diskon nominal (Rp)":
    diskon_value = st.number_input("Besar diskon (Rp)", min_value=0, value=0, format="%d")

# Input ketersediaan barang
opsi_ketersediaan = [
    "Jangan tampilkan",
    "Ready stock",
    "Ready jika persediaan masih ada",
    "Indent"
]
ketersediaan = st.selectbox("Ketersediaan Barang", opsi_ketersediaan)

# Input PIC
pic = st.selectbox("Nama PIC", list(pic_options.keys()))
pic_telp = pic_options[pic]

# Generate Dokumen dari Awal
if st.button("\U0001F4E5 Generate Dokumen Penawaran"):
    doc = Document()

    # Ganti spasi antar paragraf menjadi lebih rapat
    for para in doc.paragraphs:
        para.paragraph_format.space_after = 0

    doc.add_paragraph("Kepada Yth")
    doc.add_paragraph(nama_customer)
    doc.add_paragraph(alamat)

    p = doc.add_paragraph()
    run = p.add_run("Hal: Penawaran Harga")
    run.bold = True
    run.underline = True

    doc.add_paragraph(f"No: {nomor_penawaran}/JKT/SRV/AA/25            Jakarta, {format_tanggal_indonesia(tanggal)}")

    doc.add_paragraph(f"Terima kasih atas kesempatan yang telah diberikan kepada kami. Bersama ini kami mengajukan penawaran harga item untuk unit {nama_unit} di {nama_customer}, sebagai berikut:\n")

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

        for cell in row_cells:
            cell.paragraphs[0].alignment = 1  # Center alignment

    # Hitung diskon
    price_diskon = 0
    if diskon_option == "Diskon persentase (%)":
        price_diskon = subtotal1 * (diskon_value / 100)
    elif diskon_option == "Diskon nominal (Rp)":
        price_diskon = diskon_value
    
    subtotal2 = subtotal1 - price_diskon
    ppn = subtotal2 * 0.11
    total = subtotal2 + ppn

    row_subtotal1 = table.add_row().cells
    row_subtotal1[3].text = "Sub Total I"
    row_subtotal1[4].text = format_rupiah(subtotal1)
    for cell in row_subtotal1:
        cell.paragraphs[0].alignment = 1

    if price_diskon > 0:
        row_diskon = table.add_row().cells
        if diskon_option == "Diskon persentase (%)":
            row_diskon[3].text = f"Diskon {round(diskon_value)}%"
        else:
            row_diskon[3].text = "Diskon (Rp)"
        row_diskon[4].text = f"-{format_rupiah(price_diskon)}"
        for cell in row_diskon:
            cell.paragraphs[0].alignment = 1

    row_subtotal2 = table.add_row().cells
    row_subtotal2[3].text = "Sub Total II"
    row_subtotal2[4].text = format_rupiah(subtotal2)
    for cell in row_subtotal2:
        cell.paragraphs[0].alignment = 1

    row_ppn = table.add_row().cells
    row_ppn[3].text = "PPN 11%"
    row_ppn[4].text = format_rupiah(ppn)
    for cell in row_ppn:
        cell.paragraphs[0].alignment = 1

    row_total = table.add_row().cells
    row_total[3].text = "TOTAL"
    row_total[4].text = format_rupiah(total)
    for cell in row_total:
        cell.paragraphs[0].alignment = 1

    doc.add_paragraph("\nSyarat dan ketentuan:")
    doc.add_paragraph("Harga: Sudah termasuk PPN 11%")
    doc.add_paragraph("Pembayaran: Tunai atau transfer")
    doc.add_paragraph("Masa berlaku: 2 minggu")

    if price_diskon > 0:
        if diskon_option == "Diskon persentase (%)":
            doc.add_paragraph(f"Diskon: {round(diskon_value)}%")
        else:
            doc.add_paragraph(f"Diskon: {format_rupiah(price_diskon)}")
    
    if ketersediaan != "Jangan tampilkan":
        doc.add_paragraph(f"Ketersediaan Barang: {ketersediaan}")

    doc.add_paragraph("\nHormat kami,\n\nPT. IDS Medical Systems Indonesia\n\nM. Athur Yassin\nManager II - Engineering\n\n")
    
    # PIC dan nomor telepon di bagian paling bawah
    doc.add_paragraph(pic)
    doc.add_paragraph(pic_telp)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    # Preview dokumen
    buffer.seek(0)
    preview_doc = Document(buffer)
    preview_text = "\n".join([para.text for para in preview_doc.paragraphs])

    st.markdown("### Preview Penawaran")
    st.text_area("Isi Penawaran", value=preview_text, height=400)

    st.download_button(
        label="⬇️ Download Penawaran",
        data=buffer,
        file_name="Penawaran.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
