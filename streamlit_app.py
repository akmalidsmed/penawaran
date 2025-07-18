import streamlit as st
from datetime import date
import io
from docx import Document
import re

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

# Fungsi untuk format mata uang
def format_uang(angka):
    return f"Rp {angka:,.0f}".replace(",", ".")

# Fungsi untuk parse input uang
def parse_uang(uang_str):
    try:
        return float(re.sub(r'[^\d]', '', uang_str))
    except:
        return 0

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
jumlah_item = st.number_input("Jumlah item yang ditawarkan", min_value=1, max_value=20, value=1)

for i in range(jumlah_item):
    st.markdown(f"### Item {i+1}")
    col1, col2, col3 = st.columns([1, 3, 2])
    
    with col1:
        qty = st.text_input("Qty", value="", key=f"qty_{i}")
        uom = st.text_input("UOM", value="", key=f"uom_{i}")
    
    with col2:
        partnumber = st.text_input("Part Number", key=f"part_{i}")
        description = st.text_area("Description", key=f"desc_{i}")
    
    with col3:
        price_input = st.text_input("Harga per item (Rp)", value="", key=f"harga_input_{i}")
        priceperitem = parse_uang(price_input)
        
        try:
            qty_val = float(qty) if qty else 0
            total = qty_val * priceperitem
        except:
            total = 0
        st.text(f"Total: {format_uang(total)}")

    items.append({
        "qty": qty,
        "uom": uom,
        "partnumber": partnumber,
        "description": description,
        "priceperitem": priceperitem,
        "price": total
    })

# Bagian Keterangan Lain-lain
st.markdown("---")
st.markdown("<h3 style='text-align: center;'>Keterangan Lain-lain</h3>", unsafe_allow_html=True)

# Input Diskon
st.markdown("**Diskon**")
diskontype = st.radio("Jenis Diskon", 
                     ["Tidak ada diskon", "Diskon Persentase", "Diskon Nominal"], 
                     horizontal=True)

diskon_value = 0
diskon_items = []

if diskontype == "Diskon Persentase":
    diskon_persen = st.number_input("Diskon (%)", min_value=0, max_value=100, value=0)
    diskon_value = diskon_persen
    
    # Pilih item yang dapat diskon
    st.markdown("**Berlaku untuk item:**")
    all_items = st.checkbox("Semua item", value=True, key="all_items_diskon")
    if not all_items:
        diskon_items = st.multiselect("Pilih item yang dapat diskon", 
                                    [f"Item {i+1}" for i in range(jumlah_item)],
                                    default=[f"Item {i+1}" for i in range(jumlah_item)])

elif diskontype == "Diskon Nominal":
    diskon_nominal_input = st.text_input("Diskon Nominal (Rp)", value="")
    diskon_value = parse_uang(diskon_nominal_input)
    
    # Pilih item yang dapat diskon
    st.markdown("**Berlaku untuk item:**")
    all_items = st.checkbox("Semua item", value=True, key="all_items_diskon_nominal")
    if not all_items:
        diskon_items = st.multiselect("Pilih item yang dapat diskon", 
                                    [f"Item {i+1}" for i in range(jumlah_item)],
                                    default=[f"Item {i+1}" for i in range(jumlah_item)])

# Ketersediaan Barang
opsi_ketersediaan = [
    "Jangan tampilkan",
    "Ready stock",
    "Ready jika persediaan masih ada",
    "Indent"
]
ketersediaan = st.selectbox("Ketersediaan Barang", opsi_ketersediaan)

# Data PIC (dengan urutan baru)
pic_options = {
    "Muhammad Lukmansyah": "0821 2291 1020",
    "Rully Candra": "0813 1515 4142",
    "Denny Firmansyah": "0821 1408 0011",
    "Alamas Ramadhan": "0857 7376 2820"
}
pic = st.selectbox("Nama PIC", list(pic_options.keys()))
pic_telp = pic_options[pic]

# Generate Dokumen
if st.button("\U0001F4E5 Generate Dokumen Penawaran"):
    if not nama_customer or not alamat:
        st.error("Harap isi Nama Customer dan Alamat terlebih dahulu!")
    else:
        doc = Document()

        # Header dokumen
        doc.add_paragraph("Kepada Yth")
        doc.add_paragraph(nama_customer)
        doc.add_paragraph(alamat)

        p = doc.add_paragraph()
        run = p.add_run("Hal: Penawaran Harga")
        run.bold = True
        run.underline = True

        doc.add_paragraph(f"No: {nomor_penawaran}/JKT/SRV/AA/25\t\t\tJakarta, {format_tanggal_indonesia(tanggal)}")

        # Konten utama
        doc.add_paragraph(
            f"Terima kasih atas kesempatan yang telah diberikan kepada kami. "
            f"Bersama ini kami mengajukan penawaran harga item untuk unit {nama_unit} "
            f"di {nama_customer}, sebagai berikut:\n"
        )

        # Tabel item
        table = doc.add_table(rows=1, cols=5)
        table.style = 'Table Grid'
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'Qty'
        hdr_cells[1].text = 'Part Number'
        hdr_cells[2].text = 'Description'
        hdr_cells[3].text = 'Price per item'
        hdr_cells[4].text = 'Total Price'

        subtotal1 = 0
        for idx, item in enumerate(items):
            row_cells = table.add_row().cells
            row_cells[0].text = f"{item['qty']} {item['uom']}" if item['uom'] else item['qty']
            row_cells[1].text = item['partnumber']
            row_cells[2].text = item['description']
            row_cells[3].text = format_uang(item['priceperitem'])
            row_cells[4].text = format_uang(item['price'])
            subtotal1 += item['price']

            for cell in row_cells:
                cell.paragraphs[0].alignment = 1  # Center alignment

        # Perhitungan harga
        if diskontype == "Diskon Persentase":
            if diskon_items == [] or f"Semua item" in diskon_items:
                price_diskon = subtotal1 * (diskon_value / 100)
            else:
                price_diskon = sum(items[i]['price'] * (diskon_value / 100) 
                            for i in range(jumlah_item) if f"Item {i+1}" in diskon_items)
        elif diskontype == "Diskon Nominal":
            if diskon_items == [] or f"Semua item" in diskon_items:
                price_diskon = diskon_value
            else:
                total_selected = sum(items[i]['price'] for i in range(jumlah_item) if f"Item {i+1}" in diskon_items)
                price_diskon = (diskon_value / subtotal1) * total_selected if subtotal1 > 0 else 0
        else:
            price_diskon = 0

        subtotal2 = subtotal1 - price_diskon
        ppn = subtotal2 * 0.11
        total = subtotal2 + ppn

        # Baris subtotal dan total
        def add_total_row(table, label, value):
            row = table.add_row().cells
            row[3].text = label
            row[4].text = format_uang(value)
            for cell in row:
                cell.paragraphs[0].alignment = 1
            return row

        add_total_row(table, "Sub Total I", subtotal1)
        if diskontype != "Tidak ada diskon":
            if diskontype == "Diskon Persentase":
                add_total_row(table, f"Diskon {round(diskon_value)}%", -price_diskon)
            else:
                add_total_row(table, "Diskon Nominal", -price_diskon)
        add_total_row(table, "Sub Total II", subtotal2)
        add_total_row(table, "PPN 11%", ppn)
        add_total_row(table, "TOTAL", total)

        # Syarat dan ketentuan
        doc.add_paragraph("\nSyarat dan ketentuan:")
        doc.add_paragraph("1. Harga: Sudah termasuk PPN 11%")
        
        # Ketersediaan barang di urutan kedua jika ditampilkan
        if ketersediaan != "Jangan tampilkan":
            doc.add_paragraph(f"2. Ketersediaan Barang: {ketersediaan}")
        
        doc.add_paragraph("3. Pembayaran: Tunai atau transfer")
        doc.add_paragraph("4. Masa berlaku: 2 minggu")

        if diskontype != "Tidak ada diskon":
            if diskontype == "Diskon Persentase":
                doc.add_paragraph(f"5. Diskon: {round(diskon_value)}%")
            else:
                doc.add_paragraph(f"5. Diskon Nominal: {format_uang(diskon_value)}")

        doc.add_paragraph(f"\nPIC: {pic}")
        doc.add_paragraph(pic_telp)  # Hanya menampilkan nomor telepon saja

        # Footer
        doc.add_paragraph("\n\nHormat kami,\n\nPT. IDS Medical Systems Indonesia\n\nM. Athur Yassin\nManager II – Engineering")

        # Simpan ke buffer
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        # Preview dokumen
        preview_doc = Document(buffer)
        preview_text = "\n".join([para.text for para in preview_doc.paragraphs])

        st.markdown("### Preview Penawaran")
        st.text_area("Isi Penawaran", value=preview_text, height=400)

        # Download button
        buffer.seek(0)
        st.download_button(
            label="⬇️ Download Penawaran",
            data=buffer,
            file_name=f"Penawaran_{nama_customer.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

st.markdown("---")
st.markdown(
    "<div style='text-align: center;'>Generator Penawaran Harga © 2023 - PT. IDS Medical Systems Indonesia</div>", 
    unsafe_allow_html=True
)
