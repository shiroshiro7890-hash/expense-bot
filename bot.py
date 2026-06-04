import os
import json
import base64
import logging
import re
from datetime import datetime

import anthropic
import gspread
from google.oauth2.service_account import Credentials
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

logging.basicConfig(format="%(asctime)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

BULAN = ["Januari","Februari","Maret","April","Mei","Juni",
         "Juli","Agustus","September","Oktober","November","Desember"]

def get_sheet_name_besar():
    now = datetime.now()
    return "Rekap " + BULAN[now.month-1] + " " + str(now.year)

def get_sheet_name_kecil():
    now = datetime.now()
    return "Petty Cash " + BULAN[now.month-1] + " " + str(now.year)

def get_group_type(chat_title):
    t = (chat_title or "").lower()
    if "kecil" in t or "petty" in t:
        return "kecil"
    return "besar"

def get_gspread_client():
    raw = os.environ["GOOGLE_CREDENTIALS_JSON"]
    info = json.loads(raw)
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)

def get_sheet(group):
    gc = get_gspread_client()
    if group == "kecil":
        sid = os.environ["SPREADSHEET_ID_KAS_KECIL"]
        sname = get_sheet_name_kecil()
    else:
        sid = os.environ["SPREADSHEET_ID_KAS_BESAR"]
        sname = get_sheet_name_besar()
    sh = gc.open_by_key(sid)
    try:
        ws = sh.worksheet(sname)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=sname, rows=1000, cols=15)
        if group == "kecil":
            # 12 kolom: A-L sesuai struktur sheet asli
            ws.append_row([
                "Tanggal","Admin","Deskripsi","Kategori",
                "Debet (Keluar)","Kredit (Masuk)","Saldo","Keterangan",
                "Tanggal Invoice","Rekening Tujuan","Nama Penerima","Status"
            ])
        else:
            # 11 kolom sesuai struktur sheet kas besar
            ws.append_row([
                "Tanggal","Admin","Vendor","Nominal","Jenis",
                "Kategori","Deskripsi","Tanggal Invoice",
                "Rekening","Penerima","Status"
            ])
        ws.format("A1:L1", {
            "backgroundColor": {"red": 0.18, "green": 0.33, "blue": 0.59},
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}}
        })
    return ws

def parse_json_safe(raw_text):
    text = raw_text.strip()
    text = text.replace("```json", "").replace("```", "").strip()
    try:
        return json.loads(text)
    except Exception:
        pass
    match = re.search(r'\{[\s\S]*\}', text)
    if match:
        try:
            return json.loads(match.group())
        except Exception:
            pass
    result = {}
    patterns = {
        "tanggal": r'tanggal["\s:]+([0-9]{4}-[0-9]{2}-[0-9]{2})',
        "vendor": r'vendor["\s:]+([^\n,}"]+)',
        "deskripsi": r'deskripsi["\s:]+([^\n,}"]+)',
        "kategori": r'kategori["\s:]+([^\n,}"]+)',
        "jumlah": r'jumlah["\s:]+([0-9]+)',
        "metode": r'metode["\s:]+([^\n,}"]+)',
        "keterangan": r'keterangan["\s:]+([^\n,}"]*)',
    }
    for key, pattern in patterns.items():
        m = re.search(pattern, text, re.IGNORECASE)
        result[key] = m.group(1).strip().strip('"').strip("'") if m else ""
    return result

def analyze_image(image_bytes):
    client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])
    b64 = base64.standard_b64encode(image_bytes).decode("utf-8")

    resp = client.messages.create(
        model="claude-haiku-4-5",
        max_tokens=500,
        system='Kamu adalah asisten keuangan. Ekstrak data struk. Kembalikan HANYA JSON valid. Contoh: {"tanggal":"2026-06-04","vendor":"Alfamart","deskripsi":"Belanja harian","kategori":"Konsumsi","jumlah":62500,"metode":"QRIS","keterangan":""}',
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "image",
                    "source": {"type": "base64", "media_type": "image/jpeg", "data": b64}
                },
                {
                    "type": "text",
                    "text": 'Analisa struk ini. Kembalikan HANYA JSON (tidak ada teks lain):\n{"tanggal":"YYYY-MM-DD","vendor":"nama toko","deskripsi":"deskripsi singkat","kategori":"Konsumsi","jumlah":0,"metode":"Tunai","keterangan":""}\n\nKategori: Operational, Perlengkapan, Konsumsi, Transportasi, Lainnya\nMetode: Tunai, Transfer Bank, QRIS, E-Wallet, Kartu Debit, Kartu Kredit, Lainnya\nJumlah = angka bulat'
                }
            ]
        }]
    )

    raw = resp.content[0].text
    logger.info("Claude response: " + raw)
    data = parse_json_safe(raw)

    tanggal = str(data.get("tanggal") or "").strip()
    if not re.match(r'^\d{4}-\d{2}-\d{2}$', tanggal):
        tanggal = datetime.now().strftime("%Y-%m-%d")

    jumlah_raw = re.sub(r'[^0-9]', '', str(data.get("jumlah") or "0"))
    jumlah = int(jumlah_raw) if jumlah_raw else 0

    return {
        "tanggal": tanggal,
        "vendor": str(data.get("vendor") or "-").strip(),
        "deskripsi": str(data.get("deskripsi") or "-").strip(),
        "kategori": str(data.get("kategori") or "Lainnya").strip(),
        "jumlah": jumlah,
        "metode": str(data.get("metode") or "Lainnya").strip(),
        "keterangan": str(data.get("keterangan") or "").strip(),
    }

def get_saldo_kecil(ws):
    rows = ws.get_all_values()
    saldo = 0
    for row in rows[1:]:
        if not row or not row[0]:
            continue
        try:
            debet = float(re.sub(r'[^0-9.]', '', str(row[4]))) if len(row) > 4 and row[4] else 0
            kredit = float(re.sub(r'[^0-9.]', '', str(row[5]))) if len(row) > 5 and row[5] else 0
            saldo = saldo + kredit - debet
        except Exception:
            continue
    return saldo

def append_kas_besar(ws, data, dicatat_oleh):
    tgl = datetime.now().strftime("%d/%m/%Y")
    # 11 kolom: Tanggal, Admin, Vendor, Nominal, Jenis, Kategori, Deskripsi, Tgl Invoice, Rekening, Penerima, Status
    row = [
        tgl,
        dicatat_oleh,
        data["vendor"],
        data["jumlah"],
        "Dana Keluar",
        data["kategori"],
        data["deskripsi"],
        data["tanggal"],
        "",
        "",
        "Bot - " + datetime.now().strftime("%d/%m/%Y %H:%M")
    ]
    ws.append_row(row)

def append_kas_kecil(ws, data, dicatat_oleh):
    saldo = get_saldo_kecil(ws)
    saldo_baru = saldo - data["jumlah"]
    tgl = datetime.now().strftime("%d/%m/%Y")
    # 12 kolom: Tanggal, Admin, Deskripsi, Kategori, Debet, Kredit, Saldo, Keterangan, Tgl Invoice, Rekening, Penerima, Status
    row = [
        tgl,                                          # A: Tanggal
        dicatat_oleh,                                 # B: Admin
        data["deskripsi"],                            # C: Deskripsi
        data["kategori"],                             # D: Kategori
        data["jumlah"],                               # E: Debet (Keluar)
        0,                                            # F: Kredit (Masuk)
        saldo_baru,                                   # G: Saldo
        data["keterangan"] or data["vendor"],         # H: Keterangan
        data["tanggal"],                              # I: Tanggal Invoice
        "",                                           # J: Rekening Tujuan
        "",                                           # K: Nama Penerima
        "Bot - " + datetime.now().strftime("%d/%m/%Y %H:%M")  # L: Status
    ]
    ws.append_row(row)

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat = update.effective_chat
    group = get_group_type(chat.title or "")
    label = "Kas Besar" if group == "besar" else "Kas Kecil"
    sheet_name = get_sheet_name_besar() if group == "besar" else get_sheet_name_kecil()
    await update.message.reply_text(
        "Halo! Saya bot pencatat untuk " + label + "\n"
        "Sheet aktif: " + sheet_name + "\n\n"
        "Kirim foto struk -> otomatis tercatat ke Google Sheets\n\n"
        "/cek - 5 transaksi terakhir\n"
        "/total - Total bulan ini"
    )

async def cmd_cek(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat = update.effective_chat
    group = get_group_type(chat.title or "")
    label = "Kas Besar" if group == "besar" else "Kas Kecil"
    try:
        ws = get_sheet(group)
        rows = ws.get_all_values()
        data_rows = [r for r in rows[1:] if r and r[0]]
        if not data_rows:
            await update.message.reply_text("Belum ada data transaksi.")
            return
        last5 = data_rows[-5:]
        text = "5 Transaksi Terakhir - " + label + "\n\n"
        for row in reversed(last5):
            try:
                if group == "besar":
                    nominal = int(float(re.sub(r'[^0-9.]', '', str(row[3] or 0))))
                    text += "- " + str(row[0]) + " | " + str(row[5]) + "\n  " + str(row[2]) + "\n  Rp " + "{:,}".format(nominal) + "\n\n"
                else:
                    debet = int(float(re.sub(r'[^0-9.]', '', str(row[4] or 0))))
                    text += "- " + str(row[0]) + " | " + str(row[3]) + "\n  " + str(row[2]) + "\n  Rp " + "{:,}".format(debet) + "\n\n"
            except Exception:
                text += "- " + str(row[0]) + "\n\n"
        await update.message.reply_text(text)
    except Exception as e:
        logger.error("cmd_cek error: " + str(e))
        await update.message.reply_text("Gagal mengambil data: " + str(e))

async def cmd_total(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat = update.effective_chat
    group = get_group_type(chat.title or "")
    label = "Kas Besar" if group == "besar" else "Kas Kecil"
    try:
        ws = get_sheet(group)
        rows = ws.get_all_values()
        data_rows = [r for r in rows[1:] if r and r[0]]
        total = 0
        count = 0
        for row in data_rows:
            try:
                if group == "besar":
                    total += int(float(re.sub(r'[^0-9.]', '', str(row[3] or 0))))
                else:
                    total += int(float(re.sub(r'[^0-9.]', '', str(row[4] or 0))))
                count += 1
            except Exception:
                continue
        sheet_name = get_sheet_name_besar() if group == "besar" else get_sheet_name_kecil()
        await update.message.reply_text(
            "Total " + label + "\n"
            "Sheet: " + sheet_name + "\n\n"
            "Total pengeluaran: Rp " + "{:,}".format(total) + "\n"
            "Jumlah transaksi: " + str(count)
        )
    except Exception as e:
        logger.error("cmd_total error: " + str(e))
        await update.message.reply_text("Gagal mengambil data: " + str(e))

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat = update.effective_chat
    group = get_group_type(chat.title or "")
    label = "Kas Besar" if group == "besar" else "Kas Kecil"
    dicatat_oleh = update.effective_user.first_name or "Bot"
    msg = await update.message.reply_text("Menganalisa struk, mohon tunggu...")
    try:
        photo = update.message.photo[-1]
        file = await context.bot.get_file(photo.file_id)
        image_bytes = bytes(await file.download_as_bytearray())
        data = analyze_image(image_bytes)
        ws = get_sheet(group)
        if group == "besar":
            append_kas_besar(ws, data, dicatat_oleh)
        else:
            append_kas_kecil(ws, data, dicatat_oleh)
        await msg.edit_text(
            "Berhasil dicatat ke " + label + "!\n\n"
            "Tanggal: " + data["tanggal"] + "\n"
            "Vendor: " + data["vendor"] + "\n"
            "Deskripsi: " + data["deskripsi"] + "\n"
            "Kategori: " + data["kategori"] + "\n"
            "Jumlah: Rp " + "{:,}".format(data["jumlah"]) + "\n"
            "Metode: " + data["metode"]
        )
    except Exception as e:
        logger.error("handle_photo error: " + str(e), exc_info=True)
        await msg.edit_text("Terjadi error: " + str(e))

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Kirim foto struk untuk mencatat pengeluaran.\n"
        "Perintah: /cek /total /start"
    )

def main():
    token = os.environ["TELEGRAM_BOT_TOKEN"]
    app = Application.builder().token(token).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("cek", cmd_cek))
    app.add_handler(CommandHandler("total", cmd_total))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    logger.info("Bot started!")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
