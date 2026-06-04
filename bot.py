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
    return f"Rekap {BULAN[now.month-1]} {now.year}"

def get_sheet_name_kecil():
    now = datetime.now()
    return f"Petty Cash {BULAN[now.month-1]} {now.year}"

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
            ws.append_row(["Tanggal","Admin","Deskripsi","Kategori","Debet (Keluar)","Kredit (Masuk)","Saldo","Keterangan","","","","Status"])
        else:
            ws.append_row(["Tanggal","Admin","Vendor","Nominal","Jenis","Kategori","Deskripsi","Tanggal Invoice","Rekening","Penerima","Status"])
        ws.format("A1:L1", {
            "backgroundColor": {"red": 0.18, "green": 0.33, "blue": 0.59},
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}}
        })
    return ws

def analyze_image(image_bytes):
    client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])
    b64 = base64.standard_b64encode(image_bytes).decode("utf-8")

    system = "Kamu adalah asisten keuangan. Ekstrak data dari struk. Kembalikan HANYA JSON valid, tanpa teks lain, tanpa markdown."

    prompt = """Analisa struk ini. Kembalikan JSON PERSIS seperti ini (tidak ada teks lain):
{"tanggal":"YYYY-MM-DD","vendor":"nama toko","deskripsi":"deskripsi singkat","kategori":"Kategori","jumlah":0,"metode":"Metode","keterangan":""}

Pilihan kategori: Operational, Perlengkapan, Konsumsi, Transportasi, Lainnya
Pilihan metode: Tunai, Transfer Bank, Kartu Kredit, Kartu Debit, QRIS, E-Wallet, Lainnya

- tanggal: YYYY-MM-DD, jika tidak ada pakai hari ini
- jumlah: angka bulat saja
- Hanya JSON, tidak ada teks lain"""

    resp = client.messages.create(
        model="claude-haiku-4-5",
        max_tokens=300,
        system=system,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": b64}},
                {"type": "text", "text": prompt}
            ]
        }]
    )

    raw = resp.content[0].text.strip()
    logger.info(f"Claude response: {raw}")
    raw = raw.replace("```json", "").replace("```", "").strip()
    match = re.search(r'\{[^{}]*\}', raw, re.DOTALL)
    if match:
        raw = match.group()
    data = json.loads(raw)

    return {
        "tanggal": data.get("tanggal") or datetime.now().strftime("%Y-%m-%d"),
        "vendor": str(data.get("vendor") or "-"),
        "deskripsi": str(data.get("deskripsi") or "-"),
        "kategori": str(data.get("kategori") or "Lainnya"),
        "jumlah": int(float(re.sub(r'[^0-9.]', '', str(data.get("jumlah") or 0)))),
        "metode": str(data.get("metode") or "Lainnya"),
        "keterangan": str(data.get("keterangan") or ""),
    }

def append_kas_besar(ws, data, dicatat_oleh):
    tgl = datetime.now().strftime("%d/%m/%Y")
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
        f"Bot - {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ]
    ws.append_row(row)

def append_kas_kecil(ws, data, dicatat_oleh):
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

    saldo_baru = saldo - data["jumlah"]
    tgl = datetime.now().strftime("%d/%m/%Y")
    row = [
        tgl,
        dicatat_oleh,
        data["deskripsi"],
        data["kategori"],
        data["jumlah"],
        0,
        saldo_baru,
        data["keterangan"] or data["vendor"],
        "", "", "",
        f"Bot - {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ]
    ws.append_row(row)

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat = update.effective_chat
    group = get_group_type(chat.title or "")
    label = "Kas Besar" if group == "besar" else "Kas Kecil"
    sheet_name = get_sheet_name_besar() if group == "besar" else get_sheet_name_kecil()
    await update.message.reply_text(
        f"Halo! Saya bot pencatat untuk *{label}*\n"
        f"Sheet aktif: _{sheet_name}_\n\n"
        f"Kirim foto struk -> otomatis tercatat ke Google Sheets\n\n"
        f"/cek - 5 transaksi terakhir\n"
        f"/total - Total bulan ini",
        parse_mode="Markdown"
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
        text = f"*5 Transaksi Terakhir - {label}*\n\n"
        for row in reversed(last5):
            try:
                if group == "besar":
                    nominal = int(float(re.sub(r'[^0-9.]', '', str(row[3] or 0))))
                    text += f"- {row[0]} | {row[5]}\n  {row[2]}\n  Rp {nominal:,}\n\n"
                else:
                    debet = int(float(re.sub(r'[^0-9.]', '', str(row[4] or 0))))
                    text += f"- {row[0]} | {row[3]}\n  {row[2]}\n  Rp {debet:,}\n\n"
            except Exception:
                text += f"- {row[0]} | {row[2] if len(row) > 2 else ''}\n\n"
        await update.message.reply_text(text, parse_mode="Markdown")
    except Exception as e:
        logger.error(f"cmd_cek error: {e}")
        await update.message.reply_text(f"Gagal mengambil data: {str(e)}")

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
            f"*Total {label}*\n"
            f"Sheet: _{sheet_name}_\n\n"
            f"Total pengeluaran: Rp {total:,}\n"
            f"Jumlah transaksi: {count}",
            parse_mode="Markdown"
        )
    except Exception as e:
        logger.error(f"cmd_total error: {e}")
        await update.message.reply_text(f"Gagal mengambil data: {str(e)}")

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
            f"Berhasil dicatat ke {label}!\n\n"
            f"Tanggal: {data['tanggal']}\n"
            f"Vendor: {data['vendor']}\n"
            f"Deskripsi: {data['deskripsi']}\n"
            f"Kategori: {data['kategori']}\n"
            f"Jumlah: Rp {data['jumlah']:,}\n"
            f"Metode: {data['metode']}"
        )
    except json.JSONDecodeError:
        await msg.edit_text("Gagal membaca struk. Coba foto lebih jelas.")
    except Exception as e:
        logger.error(f"handle_photo error: {e}", exc_info=True)
        await msg.edit_text(f"Terjadi error: {str(e)}")

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
