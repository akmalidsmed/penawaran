import streamlit as st
from datetime import date
import io
from docx import Document
from docx.shared import Inches, Pt
import os
from PIL import Image
from tempfile import NamedTemporaryFile

# ... (bagian awal tetap sama)

if st.button("\U0001F4E5 Generate Dokumen Penawaran"):
    doc = Document()

    # Set default font to Calibri
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)

    section = doc.sections[0]
    header = section.header
    header_para = header.paragraphs[0]

    image_path = "/mnt/data/92e028fb-3499-479f-a167-62ec17940b2d.png"
    if os.path.exists(image_path):
        try:
            header_para.add_run().add_picture(image_path, width=Inches(6.5))
        except Exception as e:
            st.warning(f"Gagal menambahkan kop surat: {e}")

    for text in ["Kepada Yth", nama_customer, alamat]:
        p = doc.add_paragraph(text)
        p.paragraph_format.space_after = Pt(0)

    p = doc.add_paragraph()
    run = p.add_run("Hal: Penawaran Harga")
    run.bold = True
    run.underline = True
    p.paragraph_format.space_after = Pt(0)

    p = doc.add_paragraph(f"No: {nomor_penawaran}/JKT/SRV/AA/25\t\t\t\tJakarta, {format_tanggal_indonesia(tanggal)}")
    p.paragraph_format.space_after = Pt(0)

    p = doc.add_paragraph(f"Terima kasih atas kesempatan yang telah diberikan kepada kami. Bersama ini kami mengajukan penawaran harga item untuk unit {nama_unit} di {nama_customer}, sebagai berikut:\n")
    p.paragraph_format.space_after = Pt(0)

    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Qty'
    hdr_cells[1].text = 'Part Number'
    hdr_cells[2].text = 'Description'
    hdr_cells[3].text = 'Price per item'
    hdr_cells[4].text = 'Total Price'

    subtotal1 = 0
    for i, item in enumerate(items):
        row_cells = table.add_row().cells
        row_cells[0].text = f"{item['qty']}{item['uom']}"
        row_cells[1].text = item['partnumber']
        row_cells[2].text = item['description']
        row_cells[3].text = format_rupiah(item['priceperitem'])
        row_cells[4].text = format_rupiah(item['price'])
        subtotal1 += item['price']
        for cell in row_cells:
            cell.paragraphs[0].alignment = 1
            cell.paragraphs[0].paragraph_format.space_after = Pt(0)

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

    for label, value in [("Sub Total I", subtotal1), ("Sub Total II", subtotal2), ("PPN 11%", ppn), ("TOTAL", total)]:
        row = table.add_row().cells
        row[3].text = label
        row[4].text = format_rupiah(value)
        for cell in row:
            cell.paragraphs[0].alignment = 1
            cell.paragraphs[0].paragraph_format.space_after = Pt(0)

    if price_diskon > 0:
        row_disc = table.add_row().cells
        if diskon_option == "Diskon persentase (%)":
            row_disc[3].text = f"Diskon {round(diskon_value)}% ({', '.join(['Item ' + str(i+1) for i in selected_items])})"
        else:
            row_disc[3].text = f"Diskon (Rp) ({', '.join(['Item ' + str(i+1) for i in selected_items])})"
        row_disc[4].text = f"-{format_rupiah(price_diskon)}"
        for cell in row_disc:
            cell.paragraphs[0].alignment = 1
            cell.paragraphs[0].paragraph_format.space_after = Pt(0)

    for text in ["\nSyarat dan ketentuan:", "Harga: Sudah termasuk PPN 11%", "Pembayaran: Tunai atau transfer", "Masa berlaku: 2 minggu"]:
        p = doc.add_paragraph(text)
        p.paragraph_format.space_after = Pt(0)

    if ketersediaan != "Jangan tampilkan":
        p = doc.add_paragraph(f"Ketersediaan Barang: {ketersediaan}")
        p.paragraph_format.space_after = Pt(0)

    for text in ["\nHormat kami,", "PT. IDS Medical Systems Indonesia", "M. Athur Yassin", "Manager II - Engineering", pic, pic_telp]:
        p = doc.add_paragraph(text)
        p.paragraph_format.space_after = Pt(0)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    preview_doc = Document(buffer)
    preview_text = "\n".join([para.text for para in preview_doc.paragraphs])

    st.markdown("### Preview Penawaran")
    st.text_area("Isi Penawaran", value=preview_text, height=400)

    st.download_button(
        label="⬇️ Download Penawaran (Word)",
        data=buffer,
        file_name="Penawaran.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
