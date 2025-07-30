import streamlit as st
from datetime import date
import io
from docx import Document
from docx.shared import Inches, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.shape import WD_INLINE_SHAPE
import os
from PIL import Image

def format_rupiah(angka):
    return "Rp. {:,.0f}".format(angka).replace(",", ".")

def format_tanggal_indonesia(tanggal):
    bulan_list = [
        "Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ]
    hari = tanggal.day
    bulan = bulan_list[tanggal.month - 1]
    tahun = tanggal.year
    return f"{hari} {bulan} {tahun}"

st.title("Penawaran Harga")

# Upload header image
uploaded_header = st.file_uploader("Upload header image (PNG/JPG)", type=["png", "jpg", "jpeg"])

# Customer information
col1, col2 = st.columns(2)
with col1:
    nama_customer = st.text_input("Nama Customer", placeholder="Masukkan nama customer")
with col2:
    pic = st.selectbox("Nama PIC", ["Muhammad Lukmansyah", "Rully Candra", "Denny Firmansyah", "Alamas Ramadhan"])
    pic_telp = {"Muhammad Lukmansyah": "0821 2291 1020",
                "Rully Candra": "0813 1515 4142",
                "Denny Firmansyah": "0821 1408 0011",
                "Alamas Ramadhan": "0857 7376 2820"}[pic]

alamat = st.text_area("Alamat Customer", placeholder="Masukkan alamat lengkap")
nomor_penawaran = st.text_input("Nomor Penawaran", placeholder="Contoh: OFFER/2023/001")
tanggal = st.date_input("Tanggal", value=date.today())
nama_unit = st.text_input("Nama Unit", placeholder="Tipe dan Serial Number jika ada")

# Item details
st.subheader("Item yang ditawarkan")
jumlah_item = st.number_input("Jumlah item", min_value=1, max_value=10, value=1)

items = []
for i in range(int(jumlah_item)):
    st.markdown(f"**Item {i+1}**")
    col1, col2 = st.columns(2)
    with col1:
        qty = st.number_input("Qty", min_value=1, value=1, key=f"qty_{i}")
        partnumber = st.text_input("Part Number", key=f"part_{i}")
    with col2:
        uom = st.text_input("UOM", key=f"uom_{i}")
        description = st.text_input("Description", key=f"desc_{i}")
    
    priceperitem = st.number_input("Harga per item (Rp)", min_value=0, value=0, key=f"harga_{i}")
    
    items.append({
        "qty": qty,
        "uom": uom,
        "partnumber": partnumber,
        "description": description,
        "priceperitem": priceperitem,
        "price": qty * priceperitem
    })

# Discount options
st.subheader("Diskon")
diskon_option = st.radio("Jenis Diskon", ["Tanpa diskon", "Diskon persentase (%)", "Diskon nominal (Rp)"])
diskon_value = 0
selected_items = []

if diskon_option != "Tanpa diskon":
    if jumlah_item > 1:
        diskon_scope = st.radio("Diskon berlaku untuk:", ["Semua item", "Pilih item tertentu"])
        if diskon_scope == "Pilih item tertentu":
            selected_items = []
            cols = st.columns(3)
            for i in range(int(jumlah_item)):
                with cols[i % 3]:
                    if st.checkbox(f"Item {i+1}", key=f"diskon_item_{i}"):
                        selected_items.append(i)
    else:
        selected_items = [0]

    if diskon_option == "Diskon persentase (%)":
        diskon_value = st.number_input("Besar diskon (%)", min_value=0.0, max_value=100.0, value=0.0)
    else:
        diskon_value = st.number_input("Besar diskon (Rp)", min_value=0, value=0)

# Other details
ketersediaan = st.selectbox("Ketersediaan Barang", 
                           ["Jangan tampilkan", "Ready stock", "Ready jika persediaan masih ada", "Indent"])

# Generate document
if st.button("\U0001F4E5 Generate Dokumen Penawaran"):
    # Validation checks
    if not all([nama_customer, alamat, nomor_penawaran]):
        st.error("Mohon isi semua field yang diperlukan!")
    else:
        doc = Document()
        
        # Add header image
        section = doc.sections[0]
        header = section.header
        header_para = header.paragraphs[0]
        
        if uploaded_header is not None:
            image_bytes = uploaded_header.getbuffer().tobytes()
            image_buffer = io.BytesIO(image_bytes)
            try:
                header_para.add_run().add_picture(image_buffer, width=Inches(6.5))
            except Exception as e:
                st.warning(f"Gagal menambahkan header: {e}")
        else:
            image_path = "/mnt/data/92e028fb-349e-479f-a167-62ec17940b2d.png"
            if os.path.exists(image_path):
                try:
                    header_para.add_run().add_picture(image_path, width=Inches(6.5))
                except Exception as e:
                    st.warning(f"Gagal menambahkan header default: {e}")

        # Customer info
        doc.add_paragraph(nama_customer).paragraph_format.space_after = Pt(0)
        doc.add_paragraph(alamat).paragraph_format.space_after = Pt(0)
        
        p = doc.add_paragraph("Hal: Penawaran Harga")
        p.runs[0].bold = True
        p.runs[0].underline = True
        
        doc.add_paragraph(f"No: {nomor_penawaran}/JKT/SRV/AA/25\t\t\tJakarta, {format_tanggal_indonesia(tanggal)}")
        
        # Items table
        table = doc.add_table(rows=1, cols=5)
        table.style = 'Table Grid'
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'Qty'
        hdr_cells[1].text = 'Part Number'
        hdr_cells[2].text = 'Description'
        hdr_cells[3].text = 'Price per item'
        hdr_cells[4].text = 'Total Price'

        subtotal1 = sum(item['price'] for item in items)
        
        for item in items:
            row = table.add_row().cells
            row[0].text = f"{item['qty']} {item['uom']}"
            row[1].text = item['partnumber']
            row[2].text = item['description']
            row[3].text = format_rupiah(item['priceperitem'])
            row[4].text = format_rupiah(item['price'])
            for cell in row:
                cell.paragraphs[0].alignment = 1

        # Discount calculation
        price_diskon = 0
        if diskon_option != "Tanpa diskon" and selected_items:
            if diskon_option == "Diskon persentase (%)":
                total_selected = sum(items[i]['price'] for i in selected_items)
                price_diskon = total_selected * (diskon_value / 100)
            else:
                total_selected = sum(items[i]['price'] for i in selected_items)
                price_diskon = diskon_value if total_selected >= diskon_value else total_selected

        subtotal2 = subtotal1 - price_diskon
        ppn = subtotal2 * 0.11
        total = subtotal2 + ppn

        # Add totals
        rows = [
            ("Sub Total I", subtotal1),
            ("Sub Total II", subtotal2),
            ("PPN 11%", ppn),
            ("TOTAL", total)
        ]
        
        if price_diskon > 0:
            disc_text = f"Diskon {diskon_value}{'%' if diskon_option == 'Diskon百分比 (%)' else ''}"
            if selected_items:
                disc_text += f" (Item {' '.join(str(i+1) for i in selected_items)})"
            row = table.add_row().cells
            row[3].text = disc_text
            row[4].text = f"-{format_rupiah(price_diskon)}"
            for cell in row:
                cell.paragraphs[0].alignment = 1

        for label, value in rows:
            row = table.add_row().cells
            row[3].text = label
            row[4].text = format_rupiah(value)
            for cell in row:
                cell.paragraphs[0].alignment = 1

        # Terms and conditions
        doc.add_paragraph("\nSyarat dan ketentuan:")
        doc.add_paragraph("Harga: Sudah termasuk PPN 11%")
        doc.add_paragraph("Pembayaran: Tunai atau transfer")
        doc.add_paragraph("Masa berlaku: 2 minggu")
        
        if ketersediaan != "Jangan tampilkan":
            doc.add_paragraph(f"Ketersediaan Barang: {ketersediaan}")

        # Signature
        doc.add_paragraph("\nHormat kami,")
        doc.add_paragraph("PT. IDS Medical Systems Indonesia")
        doc.add_paragraph("M. Athur Yassin")
        doc.add_paragraph("Manager II - Engineering")
        doc.add_paragraph(pic)
        doc.add_paragraph(pic_telp)

        # Save and offer download
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        st.success("Dokumen berhasil dibuat!")
        st.download_button(
            label="Download Penawaran",
            data=buffer,
            file_name="Penawaran_Harga.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

        # Preview
        st.subheader("Preview Dokumen")
        preview_doc = Document(buffer)
        preview_text = "\n".join([para.text for para in preview_doc.paragraphs])
        st.text_area("", value=preview_text, height=300)
