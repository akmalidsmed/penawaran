# ... (previous code remains the same until the document generation part)

    doc.add_paragraph("\nSyarat dan ketentuan:")
    doc.add_paragraph("Harga: Sudah termasuk PPN 11%")
    doc.add_paragraph("Pembayaran: Tunai atau transfer")
    doc.add_paragraph("Masa berlaku: 2 minggu")

    if price_diskon > 0:
        if diskon_option == "Diskon persentase (%)":
            doc.add_paragraph(f"Diskon: {round(diskon_value)}%")
        else:
            doc.add_paragraph(f"Diskon: Rp {round(price_diskon)}")
    
    if ketersediaan != "Jangan tampilkan":
        doc.add_paragraph(f"Ketersediaan Barang: {ketersediaan}")

    doc.add_paragraph("\nHormat kami,\n\nPT. IDS Medical Systems Indonesia\n\nM. Athur Yassin\nManager II - Engineering\n\n")  # Added extra \n here
    
    # Moved PIC info to bottom
    doc.add_paragraph(f"{pic}")
    doc.add_paragraph(pic_telp)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

# ... (rest of the code remains the same)
