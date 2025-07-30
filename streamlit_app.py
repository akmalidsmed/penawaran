import streamlit as st
from datetime import date
import io
from docx import Document
from docx.shared import Inches, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.shape import WD_INLINE_SHAPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
from PIL import Image

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

# ... Streamlit UI input code as previously provided ...

    if st.button("\U0001F4E5 Generate Dokumen Penawaran"):
        doc = Document()

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
        p.add_run("\n")  # Enter sekali
        run = p.add_run("Hal: Penawaran Harga")
        run.bold = True
        run.underline = True
        p.paragraph_format.space_after = Pt(0)

        p = doc.add_paragraph()
        p.add_run(f"No: {nomor_penawaran}/JKT/SRV/AA/25\t\t\t\tJakarta, {format_tanggal_indonesia(tanggal)}")
        p.paragraph_format.space_after = Pt(0)

        p = doc.add_paragraph("\nTerima kasih atas kesempatan yang telah diberikan kepada kami. Bersama ini kami mengajukan penawaran harga item untuk unit {} di {}, sebagai berikut:\n".format(nama_unit, nama_customer))
        p.paragraph_format.space_after = Pt(0)

        table = doc.add_table(rows=1, cols=5)
        table.style = 'Table Grid'
        hdr_cells = table.rows[0].cells
        headers = ['Qty', 'Part Number', 'Description', 'Price per item', 'Total Price']
        for i, text in enumerate(headers):
            para = hdr_cells[i].paragraphs[0]
            run = para.add_run(text)
            run.bold = True
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER

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
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
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
                para = cell.paragraphs[0]
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if label == "TOTAL":
                    para.runs[0].bold = True
                para.paragraph_format.space_after = Pt(0)

        if price_diskon > 0:
            row_disc = table.add_row().cells
            if diskon_option == "Diskon persentase (%)":
                row_disc[3].text = f"Diskon {round(diskon_value)}% ({', '.join(['Item ' + str(i+1) for i in selected_items])})"
            else:
                row_disc[3].text = f"Diskon (Rp) ({', '.join(['Item ' + str(i+1) for i in selected_items])})"
            row_disc[4].text = f"-{format_rupiah(price_diskon)}"
            for cell in row_disc:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                cell.paragraphs[0].paragraph_format.space_after = Pt(0)

        doc.add_paragraph("\nSyarat dan ketentuan:")

        p1 = doc.add_paragraph()
        p1.add_run("Harga          : ").bold = True
        p1.add_run("Sudah termasuk PPN 11%")

        if ketersediaan != "Jangan tampilkan":
            p2 = doc.add_paragraph()
            p2.add_run("Ketersediaan   : ").bold = True
            p2.add_run(ketersediaan)

        p3 = doc.add_paragraph()
        p3.add_run("Pembayaran     : ").bold = True
        p3.add_run("Tunai atau transfer")

        p4 = doc.add_paragraph()
        p4.add_run("Masa berlaku   : ").bold = True
        p4.add_run("2 minggu")

        for text in ["\nHormat kami:", "PT. IDS Medical Systems Indonesia"]:
            doc.add_paragraph(text)

        doc.add_paragraph("\n")
        underline = doc.add_paragraph()
        underline_run = underline.add_run("M. Athur Yassin")
        underline_run.bold = True
        underline_run.underline = True

        p = doc.add_paragraph("Manager II - Engineering")
        p.runs[0].bold = True

        doc.add_paragraph(pic)
        doc.add_paragraph(pic_telp)

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        preview_doc = Document(buffer)
        preview_text = "\n".join([para.text for para in preview_doc.paragraphs])

        st.markdown("### Preview Penawaran")
        st.text_area("Isi Penawaran", value=preview_text, height=400)

        st.download_button(
            label="\u2B07\uFE0F Download Penawaran",
            data=buffer,
            file_name="Penawaran.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
