import streamlit as st
from datetime import date
import io
from docx import Document
from docx.shared import Inches, Pt
import os

def format_rupiah(angka):
    """Format currency with Indonesian Rupiah style"""
    return "Rp. {:,.0f}".format(angka).replace(",", ".")

def format_tanggal_indonesia(tanggal):
    """Convert date to Indonesian format"""
    bulan_dict = {
        1: "Januari", 2: "Februari", 3: "Maret",
        4: "April", 5: "Mei", 6: "Juni",
        7: "Juli", 8: "Agustus", 9: "September",
        10: "Oktober", 11: "November", 12: "Desember"
    }
    return f"{tanggal.day} {bulan_dict[tanggal.month]} {tanggal.year}"

# ========== UI CONFIGURATION ==========
st.set_page_config(layout="wide", page_title="Generator Penawaran Harga")
st.markdown("<h1 style='text-align: center;'>Penawaran Harga</h1>", unsafe_allow_html=True)

# ========== DATA ==========
PIC_OPTIONS = {
    "Muhammad Lukmansyah": "0821 2291 1020",
    "Rully Candra": "0813 1515 4142",
    "Denny Firmansyah": "0821 1408 0011",
    "Alamas Ramadhan": "0857 7376 2820"
}

KETERSEDIAAN_OPTIONS = [
    "Jangan tampilkan", 
    "Ready stock", 
    "Ready jika persediaan masih ada", 
    "Indent"
]

# ========== SIDEBAR INPUTS ==========
with st.sidebar:
    st.header("Informasi Perusahaan")
    uploaded_logo = st.file_uploader("Upload Logo Perusahaan", type=['png', 'jpg'])
    
    st.header("Konfigurasi Dokumen")
    masa_berlaku = st.number_input("Masa Berlaku (hari)", min_value=1, value=14)
    default_filename = st.text_input("Nama File Default", value="Penawaran")

# ========== MAIN FORM ==========
with st.form("penawaran_form"):
    col1, col2 = st.columns(2)
    with col1:
        nama_customer = st.text_input("Nama Customer*", placeholder="PT. Contoh Indonesia")
        alamat = st.text_area("Alamat Customer*", height=100)
    with col2:
        nomor_penawaran = st.text_input("Nomor Penawaran*", placeholder="001/ABC/2023")
        tanggal = st.date_input("Tanggal*", value=date.today())
        nama_unit = st.text_input("Nama Unit (Tipe dan Serial Number jika ada)")

    # Item Section
    st.markdown("---")
    st.markdown("<h3 style='text-align: center;'>Item yang ditawarkan</h3>", unsafe_allow_html=True)
    
    jumlah_item = st.number_input("Jumlah item yang ditawarkan*", min_value=1, max_value=20, value=1)
    items = []
    
    for i in range(jumlah_item):
        st.markdown(f"### Item {i+1}")
        cols = st.columns([1, 1, 2, 3, 2])
        with cols[0]:
            qty = st.text_input(f"Qty #{i+1}", value="1", key=f"qty_{i}")
        with cols[1]:
            uom = st.text_input(f"Satuan #{i+1}", value="pcs", key=f"uom_{i}")
        with cols[2]:
            partnumber = st.text_input(f"Part Number #{i+1}", key=f"part_{i}")
        with cols[3]:
            description = st.text_input(f"Description #{i+1}", key=f"desc_{i}")
        with cols[4]:
            priceperitem = st.number_input(f"Harga per item (Rp) #{i+1}", value=0, key=f"harga_{i}")

        try:
            total = float(qty) * priceperitem if qty else 0.0
        except ValueError:
            total = 0.0

        items.append({
            "qty": qty,
            "uom": uom,
            "partnumber": partnumber,
            "description": description,
            "priceperitem": priceperitem,
            "price": total
        })

    # Additional Options
    st.markdown("---")
    st.markdown("<h3 style='text-align: center;'>Keterangan lain-lain</h3>", unsafe_allow_html=True)
    
    diskon_option = st.radio("Jenis Diskon", ["Tanpa diskon", "Diskon persentase (%)", "Diskon nominal (Rp)"])
    diskon_value = 0
    selected_items = []

    if diskon_option != "Tanpa diskon":
        if jumlah_item > 1:
            diskon_scope = st.radio("Diskon berlaku untuk:", ["Semua item", "Pilih item tertentu"])
            if diskon_scope == "Pilih item tertentu":
                selected_items = st.multiselect("Pilih item yang dapat diskon", 
                                              options=[f"Item {i+1}" for i in range(jumlah_item)],
                                              default=[f"Item {i+1}" for i in range(jumlah_item)])
                selected_items = [int(item.split()[1])-1 for item in selected_items]
            else:
                selected_items = list(range(jumlah_item))
        else:
            selected_items = [0]

        if diskon_option == "Diskon persentase (%)":
            diskon_value = st.slider("Besar diskon (%)", min_value=0, max_value=100, value=10)
        else:
            diskon_value = st.number_input("Besar diskon (Rp)", min_value=0, value=0)

    ketersediaan = st.selectbox("Ketersediaan Barang", KETERSEDIAAN_OPTIONS)
    pic = st.selectbox("Nama PIC", list(PIC_OPTIONS.keys()))
    
    # Form Submission
    submitted = st.form_submit_button("🚀 Generate Dokumen Penawaran")

# ========== DOCUMENT GENERATION ==========
if submitted:
    # Input Validation
    required_fields = {
        "Nama Customer": nama_customer,
        "Alamat": alamat,
        "Nomor Penawaran": nomor_penawaran
    }
    
    missing_fields = [field for field, value in required_fields.items() if not value]
    if missing_fields:
        st.error(f"Mohon lengkapi field yang wajib diisi: {', '.join(missing_fields)}")
        st.stop()
    
    try:
        doc = Document()
        
        # Add Logo if uploaded
        if uploaded_logo:
            try:
                header = doc.sections[0].header
                header_para = header.paragraphs[0]
                header_para.alignment = 1  # Center alignment
                header_para.add_run().add_picture(io.BytesIO(uploaded_logo.read()), width=Inches(6.5))
            except Exception as e:
                st.warning(f"Gagal menambahkan logo: {str(e)}")

        # Document Content
        doc.add_paragraph("Kepada Yth:")
        doc.add_paragraph(nama_customer)
        doc.add_paragraph(alamat)
        
        # Header with underline
        hal = doc.add_paragraph()
        hal_run = hal.add_run("Hal: Penawaran Harga")
        hal_run.bold = True
        hal_run.underline = True
        
        doc.add_paragraph(f"No: {nomor_penawaran}\t\t\tJakarta, {format_tanggal_indonesia(tanggal)}")
        doc.add_paragraph(f"Terima kasih atas kesempatan yang telah diberikan kepada kami. Bersama ini kami mengajukan penawaran harga item untuk unit {nama_unit if nama_unit else '-'} di {nama_customer}, sebagai berikut:\n")

        # Create Items Table
        table = doc.add_table(rows=1, cols=5)
        table.style = 'Table Grid'
        
        # Table Header
        headers = ['Qty', 'Part Number', 'Description', 'Price per item', 'Total Price']
        for i, header in enumerate(headers):
            table.rows[0].cells[i].text = header
            table.rows[0].cells[i].paragraphs[0].runs[0].bold = True

        # Add Items
        subtotal1 = 0
        for item in items:
            row_cells = table.add_row().cells
            row_cells[0].text = f"{item['qty']} {item['uom']}"
            row_cells[1].text = item['partnumber'] if item['partnumber'] else '-'
            row_cells[2].text = item['description']
            row_cells[3].text = format_rupiah(item['priceperitem'])
            row_cells[4].text = format_rupiah(item['price'])
            subtotal1 += item['price']

        # Calculate Discounts
        price_diskon = 0
        if diskon_option != "Tanpa diskon" and selected_items:
            if diskon_option == "Diskon persentase (%)":
                price_diskon = sum(items[i]['price'] for i in selected_items) * (diskon_value / 100)
            else:
                price_diskon = diskon_value
            
            price_diskon = round(price_diskon)
            
            # Add Discount Row
            row_disc = table.add_row().cells
            disc_text = f"Diskon {diskon_value}% " if diskon_option == "Diskon persentase (%)" else "Diskon (Rp) "
            disc_text += f"({', '.join([f'Item {i+1}' for i in selected_items])})"
            
            row_disc[3].text = disc_text
            row_disc[4].text = f"-{format_rupiah(price_diskon)}"

        subtotal2 = subtotal1 - price_diskon
        ppn = subtotal2 * 0.11
        total = subtotal2 + ppn

        # Add Summary Rows
        summary_data = [
            ("Sub Total", subtotal1),
            ("Sub Total Setelah Diskon", subtotal2),
            ("PPN 11%", ppn),
            ("TOTAL", total)
        ]
        
        for label, value in summary_data:
            row = table.add_row().cells
            row[3].text = label
            row[3].paragraphs[0].runs[0].bold = True
            row[4].text = format_rupiah(value)
            row[4].paragraphs[0].runs[0].bold = True

        # Add Terms and Conditions
        doc.add_paragraph("\nSyarat dan ketentuan:")
        doc.add_paragraph("Harga: Sudah termasuk PPN 11%")
        doc.add_paragraph("Pembayaran: Tunai atau transfer")
        doc.add_paragraph(f"Masa berlaku: {masa_berlaku} hari")

        if ketersediaan != "Jangan tampilkan":
            doc.add_paragraph(f"Ketersediaan Barang: {ketersediaan}")

        # Add Closing
        doc.add_paragraph("\nHormat kami,")
        doc.add_paragraph("PT. IDS Medical Systems Indonesia")
        doc.add_paragraph("M. Athur Yassin")
        doc.add_paragraph("Manager II - Engineering")
        doc.add_paragraph(pic)
        doc.add_paragraph(PIC_OPTIONS[pic])

        # Save to buffer
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        # Download Button
        st.success("Dokumen penawaran berhasil dibuat!")
        st.download_button(
            label="⬇️ Download Penawaran",
            data=buffer,
            file_name=f"{default_filename}_{nama_customer[:20]}.docx".replace(" ", "_"),
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        st.error(f"Terjadi kesalahan saat membuat dokumen: {str(e)}")
