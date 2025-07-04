
import streamlit as st
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
from io import BytesIO
import datetime

st.set_page_config(page_title="Penawaran Otomatis", layout="centered")
st.title("📄 Aplikasi Penawaran Otomatis (Modifikasi)")

with st.form("form_penawaran"):
    nama_klinik = st.text_input("Nama Klinik", "Klinik Medika Plaza")
    alamat = st.text_area("Alamat", "Jakarta")
    kode_nomor = st.text_input("Nomor Awal Surat (misal: 971)", "971")
    bulan_romawi = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"][datetime.datetime.now().month - 1]
    tahun = datetime.datetime.now().year % 100
    no_surat = f"{kode_nomor}/JKT/SRV/AA/{bulan_romawi}/{tahun}"
    nama_unit = st.text_input("Nama Unit", "Fukuda Denshi ECG FCP-7101")

    st.markdown("### Tabel Penawaran")
    num_items = st.number_input("Jumlah Item", min_value=1, max_value=20, value=1, step=1)
    st.info(f"Total item yang akan dimasukkan: {num_items}")

    rows = []
    for i in range(num_items):
        st.markdown(f"**Item {i+1}**")
        col1, col2, col3, col4 = st.columns([1, 2, 4, 2])
        with col1: qty = st.number_input(f"Qty {i+1}", value=1, key=f"qty{i}")
        with col2: part_number = st.text_input(f"Part Number {i+1}", key=f"pn{i}")
        with col3: description = st.text_input(f"Deskripsi {i+1}", key=f"desc{i}")
        with col4: harga = st.number_input(f"Harga {i+1}", value=0, step=1000, key=f"harga{i}")
        if description:
            rows.append({ "qty": qty, "part_number": part_number, "description": description, "harga": harga })

    diskon = st.number_input("Diskon (jika ada)", value=0, step=1000)
    pic = st.text_input("Nama PIC", "Denny Firmansyah")
    no_pic = st.text_input("No. Telepon PIC", "0821 1408 001")
    submitted = st.form_submit_button("📄 Generate PDF")

if submitted:
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    y = height - 50

    c.setFont("Helvetica", 12)
    c.drawString(50, y, "Kepada Yth,")
    y -= 20
    for line in nama_klinik.split("\n"):
        c.drawString(50, y, line)
        y -= 15
    for line in alamat.split("\n"):
        c.drawString(50, y, line)
        y -= 15

    y -= 10
    c.drawString(50, y, f"Hal: Penawaran Harga")
    y -= 15
    c.drawString(50, y, f"No: {no_surat}")
    y -= 15
    c.drawString(50, y, f"Jakarta, {datetime.datetime.now().strftime('%-d %B %Y')}")
    y -= 30

    c.drawString(50, y, f"Terima kasih atas kesempatan yang telah diberikan kepada kami.")
    y -= 15
    c.drawString(50, y, f"Bersama ini kami berikan penawaran harga untuk unit {nama_unit}:")
    y -= 20

    # Table header and data
    data = [["Qty", "Part Number", "Description", "Harga"]]
    subtotal = 0
    for row in rows:
        data.append([
            str(row['qty']),
            row['part_number'],
            row['description'],
            f"Rp {row['harga']:,.0f}".replace(",", ".")
        ])
        subtotal += row['harga']

    table = Table(data, colWidths=[50, 120, 220, 100])
    table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 1, colors.black),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold')
    ]))
    y -= len(data)*20
    table.wrapOn(c, width, height)
    table.drawOn(c, 50, y)

    y -= 20
    ppn = int(subtotal * 0.11)
    total = subtotal + ppn
    if diskon > 0:
        total -= diskon

    y -= 60
    c.setFont("Helvetica-Bold", 11)
    c.drawString(50, y, f"Sub Total: Rp {subtotal:,.0f}".replace(",", "."))
    y -= 15
    c.drawString(50, y, f"PPN 11%: Rp {ppn:,.0f}".replace(",", "."))
    y -= 15
    if diskon > 0:
        c.drawString(50, y, f"Diskon: Rp {diskon:,.0f}".replace(",", "."))
        y -= 15
    c.drawString(50, y, f"TOTAL: Rp {total:,.0f}".replace(",", "."))
    y -= 30

    c.setFont("Helvetica", 11)
    c.drawString(50, y, "Syarat dan kondisi penawaran kami adalah:")
    y -= 15
    c.drawString(50, y, "- Harga sudah termasuk PPN 11%")
    y -= 15
    c.drawString(50, y, "- Pembayaran: Tunai atau Transfer sesuai kesepakatan")
    y -= 15
    c.drawString(50, y, "- Masa berlaku: 2 (Dua) Minggu")
    y -= 30

    c.drawString(50, y, "Hormat kami,")
    y -= 40
    c.drawString(50, y, "PT. IDS Medical Systems Indonesia")
    y -= 20
    c.drawString(50, y, "M. Athur Yasin")
    y -= 15
    c.drawString(50, y, "Manager II – Engineering")
    y -= 30
    c.drawString(50, y, f"PIC: {pic}")
    y -= 15
    c.drawString(50, y, f"{no_pic}")

    c.showPage()
    c.save()
    buffer.seek(0)

    st.success("PDF berhasil dibuat!")
    st.download_button("📥 Download PDF", data=buffer, file_name="penawaran.pdf")
