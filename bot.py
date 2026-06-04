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

# ── Logging ──
logging.basicConfig(
    format="%(asctime)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# ── Constants ──
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

HEADER = ["No", "Tanggal", "Deskripsi", "Kategori", "Jumlah (Rp)", "Metode", "Keterangan", "Dicatat"]

# ── Helpers ──

def get_group_type(chat_title: str) -> str:
    t = (chat_title or "").lower()
    if "kecil" in t or "petty" in t:
        return "kecil"
    return "besar"

def get_spreadsheet_id(group: str) -> str:
    if group == "kecil":
        return os.environ["SPREADSHEET_ID_KAS_KECIL"]
    return os.environ["SPREADSHEET_ID_KAS_BESAR"]

def get_sheet_name(group: str) -> str:
    if group == "kecil":
        return "Expense Bot Kas Kecil"
    return "Expense Bot Kas Besar"

def get_gspread_client():
    raw = os.environ["GOOGLE_CREDENTIALS_JSON"]
    info = json.loads(raw)
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)

def get_or_create_sheet(group: str):
    gc = get_gspread_client()
    sid = get_spreadsheet_id(group)
    sname = get_sheet_name(group)
    sh = gc.open_by_key(sid)
    try:
        ws = sh.worksheet(sname)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=sname, rows=1000, cols=10)
        ws.append_row(HEADER)
        ws.format("A1:H1", {
            "backgroundColor": {"red": 0.18, "green": 0.33, "blue": 0.59},
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}}
        })
    return ws

def analyze_image(image_bytes: bytes) -> dict:
    client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])
    b64 = base64.standard_b64encode(image_bytes).decode("utf-8")

    system = "Kamu adalah asisten keuangan yang mengekstrak data dari struk belanja. Selalu kembalikan HANYA JSON valid tanpa teks lain, tanpa markdown, tanpa penjelasan."

    prompt = """\
Analisa struk/foto ini dan kembalikan JSON dengan format PERSIS seperti ini:
{"tanggal":"YYYY-MM-DD","deskripsi":"deskripsi singkat","kategori":"Kategori","jumlah":0,"metode":"Metode","keterangan":"info tambahan"}

Pilihan kategori: Transportasi, Makan & Minum, Belanja, Operasional Kantor, Hiburan, Lainnya
Pilihan metode: Tunai, Transfer Bank, Kartu Kredit, Kartu Debit, QRIS, E-Wallet, Lainnya

Aturan:
- tanggal: format YYYY-MM-DD, jika tidak ada gunakan hari ini
- jumlah: angka bulat tanpa titik/koma
- Kembalikan HANYA JSON, tidak ada teks lain"""

    resp = client.messages.create(
        model="claude-haiku-4-5",
        max_tokens=300,
        system=system,
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "image",
                    "source": {"type": "base64", "media_type": "image/jpeg", "data": b64}
                },
                {"type": "text", "text": prompt}
            ]
        }]
    )

    raw = resp.content[0].text.strip()
    logger.info(f"Claude raw response: {raw}")

    # Bersihkan response
    raw = raw.replace("```json", "").replace("```", "").strip()

    # Cari JSON object
    match = re.search(r'\{[^{}]*\}', raw, re.DOTALL)
    if match:
        raw = match.group()

    data = json.loads(raw)

    # Validasi field
    return {
        "tanggal": data.get("tanggal") or datetime.now().strftime("%Y-%m-%d"),
        "deskripsi": str(data.get("deskripsi") or "-"),
        "kategori": str(data.get("kategori") or "Lainnya"),
        "jumlah": int(float(str(data.get("jumlah") or 0).replace(",", "").replace(".", ""))),
        "metode": str(data.get("metode") or "Lainnya"),
        "keterangan": str(data.get("keterangan") or ""),
    }

# ── Handlers ──

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat = update.effective_chat
    group = get_group_type(chat.title or "")
    label = "Kas Besar" if group == "besar" else "Kas Kecil"

    text = (
        f"Halo! Saya bot pencatat expense untuk *{label}*\n\n"
        f"Kirim foto struk -> otomatis tercatat ke Google Sheets\n\n"
        f"Perintah:\n"
        f"/start - Info bot\n"
        f"/cek - 5 transaksi terakhir\n"
        f"/total - Total pengeluaran bulan ini"
    )
    await update.message.reply_text(text, parse_mode="Markdown")

async def cmd_cek(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat = update.effective_chat
    group = get_group_type(chat.title or "")
    label = "Kas Besar" if group == "besar" else "Kas Kecil"

    try:
        ws = get_or_create_sheet(group)
        rows = ws.get_all_values()
        data_rows = [r for r in rows[1:] if r and r[0] and r[0] != "No"]

        if not data_rows:
            await update.message.reply_text("Belum ada data transaksi.")
            return

        last5 = data_rows[-5:]
        text = f"*5 Transaksi Terakhir - {label}:*\n\n"

        for row in reversed(last5):
            try:
                jumlah = int(float(str(row[4]).replace(",", "")))
                text += f"- {row[1]} | {row[3]}\n  {row[2]}\n  Rp {jumlah:,} | {row[5]}\n\n"
            except Exception:
                text += f"- {row[1]} | {row[2]}\n\n"

        await update.message.reply_text(text, parse_mode="Markdown")

    except Exception as e:
        logger.error(f"cmd_cek error: {e}")
        await update.message.reply_text(f"Gagal mengambil data: {str(e)}")

async def cmd_total(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat = update.effective_chat
    group = get_group_type(chat.title or "")
    label = "Kas Besar" if group == "besar" else "Kas Kecil"
    bulan_ini = datetime.now().strftime("%Y-%m")

    try:
        ws = get_or_create_sheet(group)
        rows = ws.get_all_values()
        data_rows = [r for r in rows[1:] if r and r[0] and r[0] != "No"]

        total = 0
        count = 0
        for row in data_rows:
            try:
                if str(row[1]).startswith(bulan_ini):
                    total += int(float(str(row[4]).replace(",", "")))
                    count += 1
            except Exception:
                continue

        bulan_label = datetime.now().strftime("%B %Y")
        text = (
            f"*Total {label} - {bulan_label}*\n\n"
            f"Total: Rp {total:,}\n"
            f"Transaksi: {count}"
        )
        await update.message.reply_text(text, parse_mode="Markdown")

    except Exception as e:
        logger.error(f"cmd_total error: {e}")
        await update.message.reply_text(f"Gagal mengambil data: {str(e)}")

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat = update.effective_chat
    group = get_group_type(chat.title or "")
    label = "Kas Besar" if group == "besar" else "Kas Kecil"

    msg = await update.message.reply_text("Menganalisa struk, mohon tunggu...")

    try:
        photo = update.message.photo[-1]
        file = await context.bot.get_file(photo.file_id)
        image_bytes = bytes(await file.download_as_bytearray())

        data = analyze_image(image_bytes)

        ws = get_or_create_sheet(group)
        all_rows = ws.get_all_values()
        no = len(all_rows)

        row = [
            no,
            data["tanggal"],
            data["deskripsi"],
            data["kategori"],
            data["jumlah"],
            data["metode"],
            data["keterangan"],
            datetime.now().strftime("%Y-%m-%d %H:%M")
        ]
        ws.append_row(row)

        await msg.edit_text(
            f"Berhasil dicatat! ({label})\n\n"
            f"Tanggal: {data['tanggal']}\n"
            f"Deskripsi: {data['deskripsi']}\n"
            f"Kategori: {data['kategori']}\n"
            f"Jumlah: Rp {data['jumlah']:,}\n"
            f"Metode: {data['metode']}\n"
            f"Keterangan: {data['keterangan'] or '-'}"
        )

    except json.JSONDecodeError as e:
        logger.error(f"JSON parse error: {e}")
        await msg.edit_text("Gagal membaca struk. Coba foto lebih jelas atau lebih dekat.")
    except Exception as e:
        logger.error(f"handle_photo error: {e}", exc_info=True)
        await msg.edit_text(f"Terjadi error: {str(e)}")

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Kirim foto struk untuk mencatat pengeluaran.\n\n"
        "Perintah: /cek /total /start"
    )

# ── Main ──

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
