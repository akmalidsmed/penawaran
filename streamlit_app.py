import streamlit as st
import pandas as pd
from datetime import date

# ---------- KONFIGURASI ----------
st.set_page_config(
    page_title="Penawaran Harga",
    page_icon="💼",
    layout="centered"
)

# ---------- STYLE ----------
st.markdown("""
    <style>
        body { font-family: "Times New Roman", serif; }
        .letter {
            background-color: white;
            padding: 30px 50px;
            border: 1px solid #ddd;
            max-width: 900px;
            margin: auto;
        }
        .header-table td {
            font-size: 12px;
            line-height: 1.2;
        }
        .table-header {
            background-color: #e6e6e6;
            font-weight: bold;
            text-align: center;
        }
        .subtotal-row td { font-weight: bold; }
        .total-row td {
            font-weight: bold;
            background-color: #e6e6e6;
        }
        .right { text-align: right; }
        .center { text-align: center; }
        .bold { font-weight: bold; }
        .underline { text-decoration: underline; }
    </style>
""", unsafe_allow_html=True)

st.title("Generator Penawaran Harga")

# ---------- STATE ----------
if "items" not in st.session_state:
    st.session_state.items = []

# ---------- FUNGSI ----------
def format_rupiah(x: float) -> str:
    return "Rp. {:,.0f},-".format(x).replace(",", ".")


def parse_chat_text(text: str):
    """
    Parsing sederhana dari teks chat menjadi list item.
    Format per baris yang didukung:
      1 pc MSP369106 LI-ION BATTERY 22610000
      1 set 160216 HEPA FILTER 5100000
    Angka terakhir = harga.
    """
    items = []
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    for line in lines:
        tokens = line.split()
        if len(tokens) < 4:
            continue
        qty = tokens[0]
        unit = tokens[1]
        price_token = tokens[-1].replace(".", "").replace(",", "")
        if not price_token.isdigit():
            continue
        price = float(price_token)
        part_number = tokens[2]
        desc_tokens = tokens[3:-1]
        desc = " ".join(desc_tokens)
        items.append({
            "qty": qty,
            "unit": unit,
            "part_number": part_number,
            "desc": desc,
            "price": price
        })
    return items

# ---------- DATA UMUM ----------
st.subheader("Data Umum Penawaran")

col1, col2 = st.columns(2)
with col1:
    nama_rs = st.text_input("Nama Rumah Sakit / Pelanggan", "Mary Cileungsi General Hospital")
    alamat1 = st.text_input("Alamat Baris 1", "Jl. Raya Narogong No.KM. 21, Cileungsi")
    alamat2 = st.text_input("Alamat Baris 2", "Kec. Cileungsi, Kabupaten Bogor, Jawa Barat 16820")
with col2:
    nomor_penawaran = st.text_input("Nomor Penawaran", "1916/JKT/SRV/AA/XI/25")
    kota_tanggal = st.text_input("Kota & Tanggal", f"Jakarta, {date.today().strftime('%d %B %Y')}")
    perihal = st.text_input("Perihal", "Penawaran Harga")

nama_contact = st.text_input("Kepada Yth (nama penerima)", nama_rs)
nama_grup = st.text_input("Judul group item (contoh: Hamilton C2)", "Hamilton C2")
ppn_persen = st.number_input("PPN (%)", min_value=0.0, max_value=20.0, value=11.0, step=0.5)

# ---------- PILIH MODE ----------
st.subheader("Mode Input Item")
mode = st.radio(
    "Pilih cara input barang:",
    ["Mode Chat (mirip ChatGPT)", "Mode Form Manual"],
    index=0
)

# ---------- MODE CHAT ----------
if mode == "Mode Chat (mirip ChatGPT)":
    st.markdown("Tulis daftar barang, 1 baris 1 item.")
    st.markdown("""
    **Contoh:**
    ```text
    1 pc MSP369106 LI-ION BATTERY 22610000
    1 pc 160216 HEPA INLET FILTER C2 5100000
    1 set 160215 DUST FILTER C2 4590000
    1 pc 160497 HPO INLET FILTER SERVICE KIT 5270000
    ```
    """)
    chat_input = st.text_area("Tulis pesan / daftar item di sini", height=200)

    if st.button("Proses Pesan Menjadi Item"):
        parsed_items = parse_chat_text(chat_input)
        if parsed_items:
            st.session_state.items = parsed_items
            st.success(f"Berhasil membaca {len(parsed_items)} item dari pesan.")
        else:
            st.warning("Tidak ada baris yang berhasil diparsing. Pastikan format seperti contoh.")

    if st.session_state.items:
        st.markdown("**Hasil parsing (boleh diedit):**")
        new_items = []
        for i, it in enumerate(st.session_state.items):
            st.markdown(f"**Item {i+1}**")
            c1, c2, c3, c4 = st.columns([1, 2, 4, 2])
            with c1:
                qty = st.text_input("Qty", it["qty"], key=f"qty_chat_{i}")
                unit = st.text_input("Satuan", it["unit"], key=f"unit_chat_{i}")
            with c2:
                part_number = st.text_input("Part Number", it["part_number"], key=f"pn_chat_{i}")
            with c3:
                desc = st.text_input("Deskripsi", it["desc"], key=f"desc_chat_{i}")
            with c4:
                price = st.number_input("Harga (Rp)", min_value=0.0,
                                        value=float(it["price"]),
                                        step=1000.0, key=f"price_chat_{i}")
            new_items.append({
                "qty": qty,
                "unit": unit,
                "part_number": part_number,
                "desc": desc,
                "price": price
            })
            st.markdown("---")
        st.session_state.items = new_items

# ---------- MODE FORM MANUAL ----------
else:
    jumlah_item = st.number_input(
        "Jumlah item",
        min_value=1,
        max_value=20,
        value=max(1, len(st.session_state.items) or 4),
        step=1
    )
    items = []
    for i in range(int(jumlah_item)):
        st.markdown(f"**Item {i+1}**")
        c1, c2, c3, c4 = st.columns([1, 2, 4, 2])
        old = st.session_state.items[i] if i < len(st.session_state.items) else \
              {"qty": "1", "unit": "Pc", "part_number": "", "desc": "", "price": 0.0}
        with c1:
            qty = st.text_input("Qty", old["qty"], key=f"qty_{i}")
            unit = st.text_input("Satuan", old["unit"], key=f"unit_{i}")
        with c2:
            part_number = st.text_input("Part Number", old["part_number"], key=f"pn_{i}")
        with c3:
            desc = st.text_input("Deskripsi", old["desc"], key=f"desc_{i}")
        with c4:
            price = st.number_input("Harga (Rp)", min_value=0.0,
                                    value=float(old["price"]),
                                    step=1000.0, key=f"price_{i}")
        items.append({
            "qty": qty,
            "unit": unit,
            "part_number": part_number,
            "desc": desc,
            "price": price
        })
        st.markdown("---")
    st.session_state.items = items

# ---------- SYARAT & KETENTUAN ----------
st.subheader("Syarat & Ketentuan Penawaran")
harga_ket = st.text_input("Harga (keterangan)", "Franco RS, Sudah Termasuk PPN 11 %")
pembayaran_ket = st.text_input("Pembayaran", "Tunai atau transfer")
masa_berlaku_ket = st.text_input("Masa berlaku", "2 (dua) minggu")

nama_perusahaan = st.text_input("Nama Perusahaan", "PT. IDS Medical Systems Indonesia")
nama_penanda_tangan = st.text_input("Nama Penanda Tangan", "M. Athur Yassin")
jabatan_penanda_tangan = st.text_input("Jabatan Penanda Tangan", "Manager II - Engineering")
pic_nama = st.text_input("PIC", "Denny Firmansyah")
pic_hp = st.text_input("HP PIC", "0821 1408 0011")

st.markdown("---")

# ---------- HITUNG TOTAL ----------
items = st.session_state.items
subtotal = sum(i["price"] for i in items)
ppn = subtotal * ppn_persen / 100
total = subtotal + ppn

# ================== PREVIEW DOKUMEN ===================
st.header("Preview Dokumen Penawaran Harga")

st.markdown("<div class='letter'>", unsafe_allow_html=True)

st.markdown("""
<table class="header-table" width="100%">
<tr>
<td style="font-size:18px; font-weight:bold;">PT. IDS Medical Systems Indonesia</td>
<td style="text-align:right; font-size:24px; font-weight:bold; color:#004f9e;">idsMED</td>
</tr>
<tr><td colspan="2">Wisma 76, 17th & 22nd Floor, Jl. Letjen. S. Parman Kav. 76</td></tr>
<tr><td colspan="2">Slipi - Jakarta 11410, Indonesia</td></tr>
<tr><td colspan="2">T : +62 21 2567 8989&nbsp;&nbsp;&nbsp;F : +62 21 5366 1038</td></tr>
<tr><td colspan="2">E : idninfo@idsmed.com&nbsp;&nbsp;&nbsp;www.idsMED.com</td></tr>
</table>
<br>
""", unsafe_allow_html=True)

st.markdown(f"""
<p><b>Kepada Yth,</b><br>
{nama_contact}<br>
{alamat1}<br>
{alamat2}</p>
""", unsafe_allow_html=True)

st.markdown(f"""
<table width="100%">
<tr>
<td width="10%">Hal</td><td width="3%">:</td><td><b>{perihal}</b></td>
<td style="text-align:right;">{kota_tanggal}</td>
</tr>
<tr>
<td>No</td><td>:</td><td>{nomor_penawaran}</td>
<td></td>
</tr>
</table>
<br>
""", unsafe_allow_html=True)

st.markdown(f"""
<p>Dengan hormat,</p>
<p>Terima kasih atas kesempatan yang telah diberikan kepada kami, bersama ini kami ingin memberikan penawaran harga item untuk unit <b>{nama_grup}</b> di <b>{nama_rs}</b>, adapun penawaran harga sebagai berikut:</p>
""", unsafe_allow_html=True)

table_html = """
<table width="100%" border="1" cellspacing="0" cellpadding="4" style="border-collapse:collapse; font-size:12px;">
<tr class="table-header">
    <td width="8%">Qty</td>
    <td width="18%">Part Number</td>
    <td>Description</td>
    <td width="18%">Price</td>
</tr>
<tr>
    <td colspan="4" class="bold">""" + nama_grup + """</td>
</tr>
"""

for it in items:
    if (it["part_number"] or it["desc"] or it["price"] > 0) and it["qty"]:
        qty_text = f'{it["qty"]} {it["unit"]}'.strip()
        table_html += f"""
        <tr>
            <td class="center">{qty_text}</td>
            <td>{it["part_number"]}</td>
            <td>{it["desc"]}</td>
            <td class="right">{format_rupiah(it["price"]) if it["price"] > 0 else ""}</td>
        </tr>
        """

table_html += f"""
<tr class="subtotal-row">
    <td colspan="3" class="right">Sub Total</td>
    <td class="right">{format_rupiah(subtotal)}</td>
</tr>
<tr class="subtotal-row">
    <td colspan="3" class="right">PPN {ppn_persen:.0f}%</td>
    <td class="right">{format_rupiah(ppn)}</td>
</tr>
<tr class="total-row">
    <td colspan="3" class="right">TOTAL</td>
    <td class="right">{format_rupiah(total)}</td>
</tr>
</table>
"""

st.markdown(table_html, unsafe_allow_html=True)

st.markdown("""
<br>
<p class="underline bold">Syarat dan kondisi penawaran kami adalah :</p>
<table style="font-size:12px;">
<tr><td width="90">Harga</td><td width="10">:</td><td>""" + harga_ket + """</td></tr>
<tr><td>Pembayaran</td><td>:</td><td>""" + pembayaran_ket + """</td></tr>
<tr><td>Masa berlaku</td><td>:</td><td>""" + masa_berlaku_ket + """</td></tr>
</table>
<br>
<p>Demikian penawaran harga ini kami ajukan. Sambil menunggu kabar baik dari Bapak / Ibu, kami mengucapkan terima kasih.</p>
<p>Hormat kami,</p>
<p><b>""" + nama_perusahaan + """</b></p>
<br><br><br>
<p><b>""" + nama_penanda_tangan + """</b><br>
""" + jabatan_penanda_tangan + """</p>
<br>
<p>PIC: """ + pic_nama + """<br>
""" + pic_hp + """</p>
""", unsafe_allow_html=True)

st.markdown("</div>", unsafe_allow_html=True)
