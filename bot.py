def generate_struk_pdf(outlet, no_nota, waktu, kasir, capster, nama_customer,
                       hp_customer, keranjang, grand_total, tunai, kembalian, metode):
    """Generate struk PDF siap print thermal 58mm, return BytesIO."""
    try:
        from fpdf import FPDF

        W   = 58   # lebar kertas thermal 58mm
        PAD = 2    # margin kiri & kanan

        pdf = FPDF(orientation='P', unit='mm', format=(W, 297))
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=3)
        pdf.set_margins(PAD, PAD, PAD)
        pdf.set_text_color(0, 0, 0)

        CW = W - PAD * 2   # content width = 54mm

        # ── Helpers ──────────────────────────────────────────

        def ln_gap(h=1):
            pdf.ln(h)

        def center(text, size=8, bold=False):
            pdf.set_font('Helvetica', 'B' if bold else '', size)
            pdf.set_x(PAD)
            pdf.cell(CW, size * 0.45, str(text), align='C', ln=True)

        def separator(char='─', size=7):
            line = char * 32
            pdf.set_font('Courier', '', size)
            pdf.set_x(PAD)
            pdf.cell(CW, 4, line, align='C', ln=True)

        def row_lr(label, value, size=8, bold_label=False, bold_val=True):
            """Row dengan label kiri & value kanan."""
            label_w = 22
            val_w   = CW - label_w
            pdf.set_x(PAD)
            pdf.set_font('Helvetica', 'B' if bold_label else '', size)
            pdf.cell(label_w, 5, str(label))
            pdf.set_font('Helvetica', 'B' if bold_val else '', size)
            pdf.cell(val_w, 5, str(value), align='R', ln=True)

        def full_row(text, size=8, bold=False, align='L'):
            pdf.set_font('Helvetica', 'B' if bold else '', size)
            pdf.set_x(PAD)
            pdf.cell(CW, 5, str(text), align=align, ln=True)

        # ── HEADER ───────────────────────────────────────────
        brand = (outlet.upper()
                 .replace('BARBERSHOP POS', '')
                 .replace(' POS', '')
                 .strip())

        ln_gap(2)
        separator('=')
        center(brand, size=14, bold=True)
        ln_gap(1)
        center('BARBERSHOP', size=8)
        separator('=')
        ln_gap(1)
        center(no_nota, size=8)
        center(waktu,   size=8)
        ln_gap(1)
        separator()

        # ── INFO ─────────────────────────────────────────────
        ln_gap(1)
        row_lr('Kasir',   str(kasir),   bold_val=False)
        row_lr('Capster', str(capster), bold_val=False)
        if nama_customer and nama_customer not in ('-', ''):
            row_lr('Customer', str(nama_customer), bold_val=False)
        if hp_customer and hp_customer not in ('-', ''):
            row_lr('HP', str(hp_customer), bold_val=False)
        ln_gap(1)
        separator()

        # ── ITEMS ─────────────────────────────────────────────
        ln_gap(1)
        for item in keranjang:
            # Nama produk — bold
            full_row(item['nama'], size=9, bold=True)

            # Qty x harga (kiri) | subtotal (kanan)
            label_w = 28
            val_w   = CW - label_w
            pdf.set_x(PAD)
            pdf.set_font('Helvetica', '', 8)
            pdf.cell(label_w, 5, f"  {item['qty']} x {fmt_rupiah(item['harga'])}")
            pdf.set_font('Helvetica', 'B', 8)
            pdf.cell(val_w, 5, fmt_rupiah(item['subtotal']), align='R', ln=True)
            ln_gap(1)

        separator()

        # ── TUNAI & KEMBALI ───────────────────────────────────
        if tunai > 0:
            ln_gap(1)
            row_lr('Tunai',   fmt_rupiah(tunai),    size=8, bold_val=False)
            row_lr('Kembali', fmt_rupiah(kembalian), size=8, bold_val=False)

        # ── TOTAL ─────────────────────────────────────────────
        ln_gap(1)
        separator('-')
        label_w = 20
        val_w   = CW - label_w
        pdf.set_x(PAD)
        pdf.set_font('Helvetica', 'B', 13)
        pdf.cell(label_w, 8, 'TOTAL')
        pdf.set_font('Helvetica', 'B', 13)
        pdf.cell(val_w, 8, fmt_rupiah(grand_total), align='R', ln=True)
        separator('-')

        # ── METODE BAYAR ──────────────────────────────────────
        ln_gap(1)
        row_lr('Metode Bayar', str(metode), size=8, bold_val=False)
        ln_gap(1)

        # ── FOOTER ────────────────────────────────────────────
        separator('=')
        ln_gap(1)
        center('Terima kasih!', size=10, bold=True)
        ln_gap(1)
        center('kasbot.id', size=7)
        ln_gap(1)
        separator('=')
        ln_gap(3)

        buf = io.BytesIO(pdf.output())
        buf.seek(0)
        return buf

    except Exception as e:
        logger.error(f"[PDF] Gagal generate: {e}", exc_info=True)
        return None
