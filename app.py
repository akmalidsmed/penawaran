
import streamlit as st
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from io import BytesIO
import datetime

st.set_page_config(page_title="Penawaran Otomatis", layout="centered")
st.title("📄 Aplikasi Penawaran Otomatis (Versi ReportLab)")

with st.form("form_penawaran"):
    nama_klinik = st.text_input("Nama Klinik", "Klinik Medika Plaza")
    alamat = st.text_area("Alamat", "Jakarta")
    kode_nomor = st.text_input("Nomor Awal Surat (misal: 971)", "971")
    bulan_romawi = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"][datetime.datetime.now().month - 1]
    tahun = datetime.datetime.now().year % 100
    no_surat = f"{kode_nomor}/JKT/SRV/AA/{bulan_romawi}/{tahun}"
    nama_unit = st.text_input("Nama Unit", "Fukuda Denshi ECG FCP-7101")

    st.markdown("### Tabel Penawaran")
    rows = []
    for i in range(1, 4):
        col1, col2, col3, col4 = st.columns([1, 2, 4, 2])
        with col1: qty = st.number_input(f"Qty {i}", value=1, key=f"qty{i}")
        with col2: part_number = st.text_input(f"Part Number {i}", key=f"pn{i}")
        with col3: description = st.text_input(f"Deskripsi {i}", key=f"desc{i}")
        with col4: harga = st.number_input(f"Harga {i}", value=0, step=1000, key=f"harga{i}")
        if description:
            rows.append({ "qty": qty, "part_number": part_number, "description": description, "harga": harga })

    pic = st.text_input("Nama PIC", "Denny Firmansyah")
    no_pic = st.text_input("No. Telepon PIC", "0821 1408 001")
    submitted = st.form_submit_button("📄 Generate PDF")

if submitted:
    buffer = BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    y = height - 50

    pdf.setFont("Helvetica", 12)
    pdf.drawString(50, y, "Kepada Yth,")
    y -= 20
    for line in nama_klinik.split("\n"):
        pdf.drawString(50, y, line)
        y -= 15
    for line in alamat.split("\n"):
        pdf.drawString(50, y, line)
        y -= 15

    y -= 10
    pdf.drawString(50, y, f"Hal: Penawaran Harga")
    y -= 15
    pdf.drawString(50, y, f"No: {no_surat}")
    y -= 15
    pdf.drawString(50, y, f"Jakarta, {datetime.datetime.now().strftime('%-d %B %Y')}")
    y -= 30

    pdf.drawString(50, y, f"Terima kasih atas kesempatan yang telah diberikan kepada kami.")
    y -= 15
    pdf.drawString(50, y, f"Bersama ini kami berikan penawaran harga untuk unit {nama_unit}:")
    y -= 25

    pdf.setFont("Helvetica-Bold", 11)
    pdf.drawString(50, y, "Qty")
    pdf.drawString(90, y, "Part Number")
    pdf.drawString(200, y, "Description")
    pdf.drawString(420, y, "Harga")
    y -= 15
    pdf.setFont("Helvetica", 11)

    subtotal = 0
    for row in rows:
        pdf.drawString(50, y, str(row['qty']))
        pdf.drawString(90, y, row['part_number'])
        pdf.drawString(200, y, row['description'])
        pdf.drawString(420, y, f"Rp {row['harga']:,.0f}".replace(',', '.'))
        subtotal += row['harga']
        y -= 15

    ppn = int(subtotal * 0.11)
    total = subtotal + ppn
    y -= 10
    pdf.setFont("Helvetica-Bold", 11)
    pdf.drawString(50, y, f"Sub Total: Rp {subtotal:,.0f}".replace(',', '.'))
    y -= 15
    pdf.drawString(50, y, f"PPN 11%: Rp {ppn:,.0f}".replace(',', '.'))
    y -= 15
    pdf.drawString(50, y, f"TOTAL: Rp {total:,.0f}".replace(',', '.'))
    y -= 30

    pdf.setFont("Helvetica", 11)
    pdf.drawString(50, y, "Syarat dan kondisi penawaran kami adalah:")
    y -= 15
    pdf.drawString(50, y, "- Harga sudah termasuk PPN 11%")
    y -= 15
    pdf.drawString(50, y, "- Pembayaran: Tunai atau Transfer sesuai kesepakatan")
    y -= 15
    pdf.drawString(50, y, "- Masa berlaku: 2 (Dua) Minggu")
    y -= 30

    pdf.drawString(50, y, "Hormat kami,")
    y -= 40
    pdf.drawString(50, y, "PT. IDS Medical Systems Indonesia")
    y -= 20
    pdf.drawString(50, y, "M. Athur Yasin")
    y -= 15
    pdf.drawString(50, y, "Manager II – Engineering")
    y -= 30
    pdf.drawString(50, y, f"PIC: {pic}")
    y -= 15
    pdf.drawString(50, y, f"{no_pic}")

    pdf.showPage()
    pdf.save()
    buffer.seek(0)

    st.success("PDF berhasil dibuat!")
    st.download_button("📥 Download PDF", data=buffer, file_name="penawaran.pdf")
