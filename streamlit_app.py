import streamlit as st
from datetime import date

# Fungsi untuk format tanggal Indonesia
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
import io
from docx import Document



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

# Data Item
opsi_ketersediaan = [
    "Jangan tampilkan",
    "Ready stock",
    "Ready jika persediaan masih ada",
    "Indent"
]
ketersediaan = st.selectbox("Ketersediaan Barang", opsi_ketersediaan)
st.markdown("<h3 style='text-align: center;'>Item yang ditawarkan</h3>", unsafe_allow_html=True)
items = []
jumlah_item = st.number_input("Jumlah item yang ditawarkan", min_value=1, max_value=10, value=1, format="%d")
diskon = st.number_input("Diskon (%)", value=0, key="diskon", format="%d")

for i in range(jumlah_item):
    if jumlah_item > 1:
        st.markdown(f"### Item {i+1}", key=f"judul{i}")
        st.markdown("---", key=f"garis{i}")

    qty = st.text_input("Qty", key=f"qty_{i}")
    uom = st.text_input("UOM", key=f"uom_{i}")
    partnumber = st.text_input("Part Number", key=f"part_{i}")
    description = st.text_input("Description", key=f"desc_{i}")
    priceperitem = st.number_input("Harga per item", value=0, key=f"harga_{i}", format="%d")
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

# Diskon & PIC
# (Sekarang bagian diskon digabung ke bawah form item)
pic = st.selectbox("Nama PIC", list(pic_options.keys()))
pic_telp = pic_options[pic]

# Generate Dokumen dari Awal
if st.button("\U0001F4E5 Generate Dokumen Penawaran"):
    doc = Document()

    doc.add_paragraph("Kepada Yth")
    doc.add_paragraph(nama_customer)
    doc.add_paragraph(alamat)

    p = doc.add_paragraph()
    run = p.add_run("Hal: Penawaran Harga")
    run.bold = True
    run.underline = True

    doc.add_paragraph(f"No: {nomor_penawaran}/JKT/SRV/AA/25			Jakarta, {format_tanggal_indonesia(tanggal)}")

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
        row_cells[3].text = f"{round(item['priceperitem'])}"
        row_cells[4].text = f"{round(item['price'])}"
        subtotal1 += item['price']

        for cell in row_cells:
            cell.paragraphs[0].alignment = 1  # Center alignment

    price_diskon = subtotal1 * (diskon / 100)
    subtotal2 = subtotal1 - price_diskon
    ppn = subtotal2 * 0.11
    total = subtotal2 + ppn

    row_subtotal1 = table.add_row().cells
    row_subtotal1[3].text = "Sub Total I"
    row_subtotal1[4].text = f"{round(subtotal1)}"
    for cell in row_subtotal1:
        cell.paragraphs[0].alignment = 1

    if diskon > 0:
        row_diskon = table.add_row().cells
        row_diskon[3].text = f"Diskon {round(diskon)}%"
        row_diskon[4].text = f"-{round(price_diskon)}"
        for cell in row_diskon:
            cell.paragraphs[0].alignment = 1

    row_subtotal2 = table.add_row().cells
    row_subtotal2[3].text = "Sub Total II"
    row_subtotal2[4].text = f"{round(subtotal2)}"
    for cell in row_subtotal2:
        cell.paragraphs[0].alignment = 1

    row_ppn = table.add_row().cells
    row_ppn[3].text = "PPN 11%"
    row_ppn[4].text = f"{round(ppn)}"
    for cell in row_ppn:
        cell.paragraphs[0].alignment = 1

    row_total = table.add_row().cells
    row_total[3].text = "TOTAL"
    row_total[4].text = f"{round(total)}"
    for cell in row_total:
        cell.paragraphs[0].alignment = 1

    doc.add_paragraph("\nSyarat dan ketentuan:")
    doc.add_paragraph("Harga: Sudah termasuk PPN 11%")
    doc.add_paragraph("Pembayaran: Tunai atau transfer")
    doc.add_paragraph("Masa berlaku: 2 minggu")

    doc.add_paragraph(f"Diskon: {round(diskon)}%")
    if ketersediaan != "Jangan tampilkan":
        doc.add_paragraph(f"Ketersediaan Barang: {ketersediaan}")
    doc.add_paragraph(f"PIC: {pic}")
    doc.add_paragraph(f"No. Telp PIC: {pic_telp}")

    doc.add_paragraph("\nHormat kami,\n\nPT. IDS Medical Systems Indonesia\n\nM. Athur Yassin\nManager II – Engineering")

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    buffer.seek(0)
    preview_doc = Document(buffer)
    preview_text = "
".join([para.text for para in preview_doc.paragraphs])  # fixed
".join([para.text for para in preview_doc.paragraphs])

    st.markdown("### Preview Penawaran")
    st.text_area("Isi Penawaran", value=preview_text, height=400)

    st.download_button("⬇️ Download Penawaran", buffer, file_name="Penawaran.docx")


