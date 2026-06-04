import os
import json
import base64
import logging
import anthropic
import gspread
from datetime import datetime
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from google.oauth2.service_account import Credentials

logging.basicConfig(format="%(asctime)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

def get_credentials():
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    creds_dict = json.loads(creds_json)
    return Credentials.from_service_account_info(creds_dict, scopes=SCOPES)

def detect_group(chat_title: str) -> str:
    """Deteksi group berdasarkan nama: 'besar' atau 'petty/kecil'"""
    title = (chat_title or "").lower()
    if "besar" in title:
        return "besar"
    elif "petty" in title or "kecil" in title:
        return "kecil"
    return "besar"  # default

def get_sheet(group: str):
    """Ambil worksheet berdasarkan group"""
    creds = get_credentials()
    gc = gspread.authorize(creds)

    if group == "kecil":
        spreadsheet_id = os.environ["SPREADSHEET_ID_KAS_KECIL"]
        sheet_title = "Expense Bot Kas Kecil"
    else:
        spreadsheet_id = os.environ["SPREADSHEET_ID_KAS_BESAR"]
        sheet_title = "Expense Bot Kas Besar"

    sh = gc.open_by_key(spreadsheet_id)

    try:
        ws = sh.worksheet(sheet_title)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=sheet_title, rows=1000, cols=10)
        ws.append_row(["No", "Tanggal", "Deskripsi", "Kategori", "Jumlah (Rp)", "Metode Pembayaran", "Keterangan", "Dicatat"])
        ws.format("A1:H1", {
            "backgroundColor": {"red": 0.18, "green": 0.33, "blue": 0.59},
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}}
        })
    return ws

def extract_expense_from_image(image_bytes: bytes, mime_type: str = "image/jpeg") -> dict:
    client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])
    image_data = base64.standard_b64encode(image_bytes).decode("utf-8")

    prompt = """Kamu adalah asisten keuangan. Analisa struk/foto expense ini dan ekstrak informasi berikut dalam format JSON.

Kategori yang tersedia:
- Transportasi
- Makan & Minum
- Akomodasi
- Operasional Kantor
- Hiburan / Entertainment
- Lainnya

Metode pembayaran yang tersedia:
- Tunai
- Transfer Bank
- Kartu Kredit
- Kartu Debit
- E-Wallet
- Lainnya

Kembalikan HANYA JSON ini (tanpa teks lain):
{
  "tanggal": "YYYY-MM-DD",
  "deskripsi": "deskripsi singkat expense",
  "kategori": "salah satu kategori di atas",
  "jumlah": 0,
  "metode_pembayaran": "salah satu metode di atas",
  "keterangan": "info tambahan jika ada"
}

Jika tanggal tidak terlihat, gunakan tanggal hari ini.
Jika jumlah tidak jelas, isi 0.
Jumlah harus berupa angka tanpa titik/koma."""

    response = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=500,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": mime_type, "data": image_data}},
                {"type": "text", "text": prompt}
            ]
        }]
    )

    text = response.content[0].text.strip()
text = text.replace("```json", "").replace("```", "").strip()
# Cari JSON di dalam text
import re
match = re.search(r'\{.*\}', text, re.DOTALL)
if match:
    text = match.group()
return json.loads(text)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat = update.effective_chat
    group = detect_group(chat.title or "")
    label = "Kas Besar" if group == "besar" else "Kas Kecil"

    await update.message.reply_text(
        f"👋 Halo! Gw bot pencatat expense untuk *{label}*.\n\n"
        f"📸 Kirim foto struk/bon → gw otomatis catat ke Google Sheets!\n\n"
        f"Perintah:\n"
        f"/start - Mulai\n"
        f"/cek - Lihat 5 expense terakhir",
        parse_mode="Markdown"
    )

async def cek(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        chat = update.effective_chat
        group = detect_group(chat.title or "")
        ws = get_sheet(group)
        data = ws.get_all_values()

        if len(data) <= 1:
            await update.message.reply_text("📭 Belum ada data expense.")
            return

        rows = data[1:][-5:]  # 5 terakhir, skip header
        label = "Kas Besar" if group == "besar" else "Kas Kecil"
        text = f"📋 *5 Expense Terakhir ({label}):*\n\n"

        for row in rows:
            if len(row) >= 5 and row[0] != "No":
                try:
                    jumlah = int(float(row[4]))
                    text += f"• {row[1]} | {row[3]} | Rp {jumlah:,}\n  _{row[2]}_\n\n"
                except:
                    text += f"• {row[1]} | {row[3]} | {row[4]}\n  _{row[2]}_\n\n"

        await update.message.reply_text(text, parse_mode="Markdown")
    except Exception as e:
        logger.error(f"Cek error: {e}")
        await update.message.reply_text("❌ Gagal mengambil data.")

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("⏳ Lagi analisa struk-nya, tunggu sebentar...")
    try:
        chat = update.effective_chat
        group = detect_group(chat.title or "")
        label = "Kas Besar" if group == "besar" else "Kas Kecil"

        photo = update.message.photo[-1]
        file = await context.bot.get_file(photo.file_id)
        image_bytes = await file.download_as_bytearray()

        data = extract_expense_from_image(bytes(image_bytes))

        ws = get_sheet(group)
        all_rows = ws.get_all_values()
        no = len(all_rows)

        jumlah = int(data.get("jumlah", 0))
        row = [
            no,
            data.get("tanggal", datetime.now().strftime("%Y-%m-%d")),
            data.get("deskripsi", "-"),
            data.get("kategori", "Lainnya"),
            jumlah,
            data.get("metode_pembayaran", "Lainnya"),
            data.get("keterangan", ""),
            datetime.now().strftime("%Y-%m-%d %H:%M")
        ]
        ws.append_row(row)

        konfirmasi = (
            f"✅ *Expense berhasil dicatat! ({label})*\n\n"
            f"📅 Tanggal: {row[1]}\n"
            f"📝 Deskripsi: {row[2]}\n"
            f"🏷️ Kategori: {row[3]}\n"
            f"💰 Jumlah: Rp {jumlah:,}\n"
            f"💳 Metode: {row[5]}\n"
            f"📌 Keterangan: {row[6] or '-'}"
        )
        await update.message.reply_text(konfirmasi, parse_mode="Markdown")

    except json.JSONDecodeError:
        await update.message.reply_text("❌ Gagal membaca struk. Coba foto yang lebih jelas ya!")
    except Exception as e:
        logger.error(f"Photo error: {e}")
        await update.message.reply_text(f"❌ Terjadi error: {str(e)}")

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📸 Kirim foto struk/bon untuk mencatat expense.\n"
        "Atau ketik /cek untuk melihat expense terakhir."
    )

def main():
    token = os.environ["TELEGRAM_BOT_TOKEN"]
    app = Application.builder().token(token).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("cek", cek))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    logger.info("Bot started!")
    app.run_polling()

if __name__ == "__main__":
    main()
