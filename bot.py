import os
import json
import base64
import logging
import re
import hashlib
import io
from datetime import datetime

import urllib.request
import urllib.parse

import anthropic
import gspread
from google.oauth2.service_account import Credentials
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ConversationHandler, filters, ContextTypes
)

logging.basicConfig(format="%(asctime)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

BULAN = ["Januari","Februari","Maret","April","Mei","Juni",
         "Juli","Agustus","September","Oktober","November","Desember"]

KONFIRMASI_TANGGAL, INPUT_TANGGAL, KONFIRMASI_NOMINAL, INPUT_NOMINAL, PILIH_KATEGORI, TULIS_DESKRIPSI, INPUT_REKENING, INPUT_PENERIMA = range(8)
EDIT_PILIH_TRANSAKSI, EDIT_PILIH_FIELD, EDIT_INPUT_NILAI = range(8, 11)
DELETE_PILIH_TRANSAKSI, DELETE_KONFIRMASI = range(11, 13)
RESET_KONFIRMASI = 13

POS_PILIH_PRODUK, POS_INPUT_QTY, POS_PILIH_BAYAR, POS_INPUT_TUNAI, POS_INPUT_CAPSTER, POS_INPUT_NAMA_CUSTOMER, POS_INPUT_HP_CUSTOMER, POS_PILIH_CAPSTER = range(14, 22)
POS_SETUP_NAMA, POS_SETUP_HARGA, POS_SETUP_KONFIRMASI = range(30, 33)
POS_EDIT_PILIH, POS_EDIT_FIELD, POS_EDIT_NILAI = range(33, 36)
POS_HAPUS_PILIH, POS_HAPUS_KONFIRMASI = range(36, 38)

ADMIN_IDS = [5418153944]

KATEGORI = [
    "Operational", "Perlengkapan",
    "Pendapatan", "Modal Usaha",
    "Beban Bunga", "Petty Cash",
    "Marketing", "Gaji"
]

PAGE_SIZE = 10
processed_hashes = set()
user_data_temp = {}

def fmt_rupiah(angka):
    try:
        return "Rp " + f"{int(angka):,}".replace(",", ".")
    except Exception:
        return "Rp 0"

# ─────────────────────────────────────────
# Google Auth
# ─────────────────────────────────────────

def get_credentials():
    raw = os.environ["GOOGLE_CREDENTIALS_JSON"]
    info = json.loads(raw)
    return Credentials.from_service_account_info(info, scopes=SCOPES)

def get_gspread_client():
    return gspread.authorize(get_credentials())

# ─────────────────────────────────────────
# Cloudinary Upload (FIXED)
# ─────────────────────────────────────────

def upload_foto_to_drive(image_bytes, filename):
    """Upload foto ke Cloudinary, return secure_url. Return '' jika gagal."""
    try:
        import time

        cloud_name = os.environ.get("CLOUDINARY_CLOUD_NAME", "").strip()
        api_key    = os.environ.get("CLOUDINARY_API_KEY", "").strip()
        api_secret = os.environ.get("CLOUDINARY_API_SECRET", "").strip()

        if not all([cloud_name, api_key, api_secret]):
            logger.error("[CLOUDINARY] Credentials kosong!")
            return ""

        logger.info(f"[CLOUDINARY] Mulai upload: {filename}")

        timestamp = str(int(time.time()))
        public_id = filename.replace(".jpg", "").replace(".", "_")
        folder    = "kasbot_struk"

        # Signature harus sorted: folder, public_id, timestamp
        sign_str  = f"folder={folder}&public_id={public_id}&timestamp={timestamp}{api_secret}"
        signature = hashlib.sha1(sign_str.encode("utf-8")).hexdigest()

        boundary = "----KasBotBoundary" + hashlib.md5(image_bytes[:16]).hexdigest()[:8]
        body = b""

        fields = {
            "api_key":   api_key,
            "timestamp": timestamp,
            "public_id": public_id,
            "folder":    folder,
            "signature": signature,
        }
        for key, val in fields.items():
            body += (
                f"--{boundary}\r\n"
                f"Content-Disposition: form-data; name=\"{key}\"\r\n\r\n"
                f"{val}\r\n"
            ).encode("utf-8")

        body += (
            f"--{boundary}\r\n"
            f"Content-Disposition: form-data; name=\"file\"; filename=\"{filename}\"\r\n"
            f"Content-Type: image/jpeg\r\n\r\n"
        ).encode("utf-8")
        body += image_bytes
        body += f"\r\n--{boundary}--\r\n".encode("utf-8")

        url = f"https://api.cloudinary.com/v1_1/{cloud_name}/image/upload"
        req = urllib.request.Request(
            url, data=body, method="POST",
            headers={"Content-Type": f"multipart/form-data; boundary={boundary}"}
        )

        with urllib.request.urlopen(req, timeout=30) as resp:
            result = json.loads(resp.read().decode("utf-8"))

        link = result.get("secure_url", "")
        if link:
            logger.info(f"[CLOUDINARY] Upload berhasil: {link}")
        else:
            logger.error(f"[CLOUDINARY] Tidak ada secure_url: {result}")
        return link

    except Exception as e:
        logger.error(f"[CLOUDINARY] Upload error - {type(e).__name__}: {e}", exc_info=True)
        return ""

# ─────────────────────────────────────────
# Sheet Helpers
# ─────────────────────────────────────────

def get_sheet_name_besar(dt=None):
    d = dt or datetime.now()
    return f"Rekap {BULAN[d.month-1]} {d.year}"

def get_sheet_name_kecil(dt=None):
    d = dt or datetime.now()
    return f"Petty Cash {BULAN[d.month-1]} {d.year}"

def get_group_type(chat_title):
    t = (chat_title or "").lower()
    if "kecil" in t or "petty" in t:
        return "kecil"
    if any(w in t for w in ["pos", "outlet", "barbershop", "salon", "toko", "jualan", "kasir", "penjualan"]):
        return "pos"
    return "besar"

def get_sheet(group, dt=None):
    gc = get_gspread_client()
    if group == "kecil":
        sid = os.environ["SPREADSHEET_ID_KAS_KECIL"]
        sname = get_sheet_name_kecil(dt)
    else:
        sid = os.environ["SPREADSHEET_ID_KAS_BESAR"]
        sname = get_sheet_name_besar(dt)

    sh = gc.open_by_key(sid)
    try:
        ws = sh.worksheet(sname)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=sname, rows=1000, cols=15)
        ws.append_row([
            "Tanggal", "Admin", "Deskripsi", "Kategori",
            "Debet (Keluar)", "Kredit (Masuk)", "Saldo", "Keterangan",
            "Tanggal Invoice", "Rekening Tujuan", "Nama Penerima", "Status", "Link Foto"
        ])
        ws.format("A1:M1", {
            "backgroundColor": {"red": 0.18, "green": 0.33, "blue": 0.59},
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}}
        })
    return ws

# ─────────────────────────────────────────
# Append ke Sheet
# ─────────────────────────────────────────

def append_kas_besar(ws, data, dicatat_oleh, foto_link=""):
    saldo = get_saldo_besar(ws)
    jenis_masuk = data["kategori"] in ["Pendapatan", "Modal Usaha"]
    if jenis_masuk:
        debet, kredit = 0, data["jumlah"]
        saldo_baru = saldo + data["jumlah"]
    else:
        debet, kredit = data["jumlah"], 0
        saldo_baru = saldo - data["jumlah"]
    tgl = datetime.now().strftime("%d/%m/%Y")
    row = [
        tgl, dicatat_oleh, data["deskripsi"], data["kategori"],
        fmt_rupiah(debet), fmt_rupiah(kredit), fmt_rupiah(saldo_baru),
        data["vendor"], format_tanggal_invoice(data["tanggal"]),
        data.get("rekening", ""), data.get("penerima", ""),
        "Bot - " + datetime.now().strftime("%d/%m/%Y %H:%M"), foto_link
    ]
    logger.info(f"[SHEET] Append kas besar, foto_link: '{foto_link}'")
    ws.append_row(row)

def recalculate_saldo(ws):
    try:
        rows = ws.get_all_values()
        saldo = 0
        updates = []
        for i, row in enumerate(rows[1:], start=2):
            if not row or not row[0]:
                continue
            debet  = float(re.sub(r'[^0-9]', '', str(row[4]))) if len(row) > 4 and row[4] else 0
            kredit = float(re.sub(r'[^0-9]', '', str(row[5]))) if len(row) > 5 and row[5] else 0
            saldo  = saldo + kredit - debet
            updates.append({"range": f"G{i}", "values": [[fmt_rupiah(saldo)]]})
        if updates:
            ws.batch_update(updates)
        logger.info(f"[SALDO] Recalculate selesai, {len(updates)} baris diupdate")
    except Exception as e:
        logger.error(f"[SALDO] Recalculate gagal: {e}", exc_info=True)

def get_saldo_besar(ws):
    rows = ws.get_all_values()
    saldo = 0
    for row in rows[1:]:
        if not row or not row[0]:
            continue
        try:
            debet  = float(re.sub(r'[^0-9]', '', str(row[4]))) if len(row) > 4 and row[4] else 0
            kredit = float(re.sub(r'[^0-9]', '', str(row[5]))) if len(row) > 5 and row[5] else 0
            saldo  = saldo + kredit - debet
        except Exception:
            continue
    return saldo

def append_kas_kecil(ws, data, dicatat_oleh, foto_link=""):
    saldo = get_saldo_kecil(ws)
    jenis_masuk = data["kategori"] in ["Pendapatan", "Modal Usaha"]
    if jenis_masuk:
        debet, kredit = 0, data["jumlah"]
        saldo_baru = saldo + data["jumlah"]
    else:
        debet, kredit = data["jumlah"], 0
        saldo_baru = saldo - data["jumlah"]
    tgl = datetime.now().strftime("%d/%m/%Y")
    row = [
        tgl, dicatat_oleh, data["deskripsi"], data["kategori"],
        fmt_rupiah(debet), fmt_rupiah(kredit), fmt_rupiah(saldo_baru),
        data["vendor"], format_tanggal_invoice(data["tanggal"]),
        data.get("rekening", ""), data.get("penerima", ""),
        "Bot - " + datetime.now().strftime("%d/%m/%Y %H:%M"), foto_link
    ]
    logger.info(f"[SHEET] Append kas kecil, foto_link: '{foto_link}'")
    ws.append_row(row)

def get_saldo_kecil(ws):
    rows = ws.get_all_values()
    saldo = 0
    for row in rows[1:]:
        if not row or not row[0]:
            continue
        try:
            debet  = float(re.sub(r'[^0-9]', '', str(row[4]))) if len(row) > 4 and row[4] else 0
            kredit = float(re.sub(r'[^0-9]', '', str(row[5]))) if len(row) > 5 and row[5] else 0
            saldo  = saldo + kredit - debet
        except Exception:
            continue
    return saldo

# ─────────────────────────────────────────
# Claude AI - Analisa Struk
# ─────────────────────────────────────────

def parse_json_safe(raw_text):
    text = raw_text.strip().replace("```json", "").replace("```", "").strip()
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
        "jumlah": r'jumlah["\s:]+([0-9]+)',
        "metode": r'metode["\s:]+([^\n,}"]+)',
    }
    for key, pattern in patterns.items():
        m = re.search(pattern, text, re.IGNORECASE)
        result[key] = m.group(1).strip().strip('"').strip("'") if m else ""
    return result

def parse_tanggal(tanggal):
    tanggal = str(tanggal or "").strip()
    if re.match(r'^\d{4}-\d{2}-\d{2}$', tanggal):
        tahun = int(tanggal[:4])
        if 2020 <= tahun <= 2030:
            return tanggal
    elif re.match(r'^\d{2}[-/]\d{2}[-/]\d{4}$', tanggal):
        sep = '-' if '-' in tanggal else '/'
        parts = tanggal.split(sep)
        return f"{parts[2]}-{parts[1]}-{parts[0]}"
    return datetime.now().strftime("%Y-%m-%d")

def format_tanggal_display(tanggal_str):
    try:
        dt = datetime.strptime(tanggal_str, "%Y-%m-%d")
        return dt.strftime("%d %B %Y")
    except Exception:
        return tanggal_str

def format_tanggal_invoice(tanggal_str):
    try:
        dt = datetime.strptime(tanggal_str, "%Y-%m-%d")
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return tanggal_str

def analyze_image(image_bytes):
    try:
        client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"], timeout=30.0)
        b64 = base64.standard_b64encode(image_bytes).decode("utf-8")
        resp = client.messages.create(
            model="claude-haiku-4-5",
            max_tokens=300,
            system='Kamu adalah asisten keuangan. Ekstrak data struk. Kembalikan HANYA JSON valid.',
            messages=[{
                "role": "user",
                "content": [
                    {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": b64}},
                    {"type": "text", "text": (
                        'Analisa struk ini. Kembalikan HANYA JSON:\n'
                        '{"tanggal":"YYYY-MM-DD","vendor":"nama toko","jumlah":0,"metode":"Tunai"}\n\n'
                        'ATURAN TANGGAL:\n'
                        '- Format output WAJIB YYYY-MM-DD\n'
                        '- Struk tulis 01-06-2026 atau 01/06/2026 = output 2026-06-01\n'
                        '- Tahun HARUS 4 digit penuh\n\n'
                        'Metode: Tunai, Transfer Bank, QRIS, E-Wallet, Kartu Debit, Kartu Kredit, Lainnya\n'
                        'Jumlah = angka bulat tanpa titik/koma\n'
                        'Hanya JSON, tidak ada teks lain'
                    )}
                ]
            }]
        )
        raw = resp.content[0].text
        logger.info(f"[CLAUDE] Response: {raw}")
        data = parse_json_safe(raw)
        tanggal = parse_tanggal(data.get("tanggal"))
        jumlah_raw = re.sub(r'[^0-9]', '', str(data.get("jumlah") or "0"))
        jumlah = int(jumlah_raw) if jumlah_raw else 0
        return {
            "tanggal": tanggal,
            "vendor": str(data.get("vendor") or "-").strip(),
            "jumlah": jumlah,
            "metode": str(data.get("metode") or "Lainnya").strip(),
        }
    except anthropic.APITimeoutError:
        raise Exception("Claude AI timeout. Coba kirim foto ulang.")
    except anthropic.APIStatusError as e:
        raise Exception(f"Claude AI error ({e.status_code}). Coba lagi.")
    except Exception as e:
        logger.error(f"[CLAUDE] Error: {e}", exc_info=True)
        raise

# ─────────────────────────────────────────
# Keyboard Helpers
# ─────────────────────────────────────────

def build_konfirmasi_keyboard(tipe):
    return InlineKeyboardMarkup([
        [
            InlineKeyboardButton("✅ Benar", callback_data=f"{tipe}_benar"),
            InlineKeyboardButton("✏️ Ubah", callback_data=f"{tipe}_ubah")
        ],
        [InlineKeyboardButton("❌ Batalkan", callback_data=f"{tipe}_batal")]
    ])

def build_kategori_keyboard():
    keyboard = []
    row = []
    for i, kat in enumerate(KATEGORI):
        row.append(InlineKeyboardButton(kat, callback_data=f"kat_{kat}"))
        if len(row) == 2:
            keyboard.append(row)
            row = []
    if row:
        keyboard.append(row)
    keyboard.append([InlineKeyboardButton("❌ Batalkan", callback_data="kat_BATAL")])
    return InlineKeyboardMarkup(keyboard)

def get_all_transactions(ws, group):
    rows = ws.get_all_values()
    data_rows = []
    for i, row in enumerate(rows[1:], start=2):
        if not row or not row[0]:
            continue
        data_rows.append({
            "row_idx": i,
            "tanggal": row[0] if len(row) > 0 else "",
            "vendor": row[7] if len(row) > 7 else "",
            "nominal": row[4] if len(row) > 4 else "0",
            "kategori": row[3] if len(row) > 3 else "",
            "deskripsi": row[2] if len(row) > 2 else "",
            "status": row[11] if len(row) > 11 else "",
        })
    return list(reversed(data_rows))

def build_transaksi_keyboard(transactions, page, group):
    start = page * PAGE_SIZE
    end = start + PAGE_SIZE
    page_data = transactions[start:end]
    total_pages = (len(transactions) + PAGE_SIZE - 1) // PAGE_SIZE
    keyboard = []
    for trx in page_data:
        try:
            nominal = int(float(re.sub(r'[^0-9.]', '', str(trx["nominal"] or 0))))
            label_btn = f"{trx['tanggal']} | {(trx['deskripsi'] or trx['vendor'])[:18]} | {fmt_rupiah(nominal)}"
        except Exception:
            label_btn = f"{trx['tanggal']} | {(trx['deskripsi'] or '')[:25]}"
        keyboard.append([InlineKeyboardButton(label_btn, callback_data=f"edit_trx_{trx['row_idx']}")])
    nav = []
    if page > 0:
        nav.append(InlineKeyboardButton("◀ Prev", callback_data=f"edit_page_{page - 1}"))
    if end < len(transactions):
        nav.append(InlineKeyboardButton("Next ▶", callback_data=f"edit_page_{page + 1}"))
    if nav:
        keyboard.append(nav)
    keyboard.append([InlineKeyboardButton("❌ Batalkan", callback_data="edit_batal")])
    return InlineKeyboardMarkup(keyboard), total_pages

# ─────────────────────────────────────────
# Command Handlers
# ─────────────────────────────────────────

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat = update.effective_chat
    group = get_group_type(chat.title or "")
    if group == "pos":
        await cmd_pos_start(update, context)
        return
    label = "Kas Besar" if group == "besar" else "Kas Kecil"
    sheet_name = get_sheet_name_besar() if group == "besar" else get_sheet_name_kecil()
    await update.message.reply_text(
        f"Halo! Saya bot pencatat {label}\n"
        f"Sheet aktif: {sheet_name}\n\n"
        f"Perintah:\n"
        f"/cek - 10 transaksi terbaru\n"
        f"/total - Total bulan ini\n"
        f"/edit - Edit transaksi\n"
        f"/delete - Hapus transaksi (admin)\n"
        f"/reset_bulan - Reset semua data bulan ini (admin)\n"
        f"/batal - Batalkan proses"
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
        last10 = data_rows[-10:]
        total_semua = len(data_rows)
        text = f"📋 10 Transaksi Terbaru — {label}\n(Total: {total_semua} transaksi)\n\n"
        for row in reversed(last10):
            try:
                nominal = int(float(re.sub(r'[^0-9]', '', str(row[4] or 0))))
                deskripsi = (row[2] or '-')[:25]
                text += f"• {row[0]} | {row[3]}\n  {deskripsi}\n  {fmt_rupiah(nominal)}\n\n"
            except Exception:
                text += f"• {row[0]}\n\n"
        if len(text) > 3800:
            text = text[:3800] + "\n\n... (lihat kasbot.id)"
        await update.message.reply_text(text)
    except Exception as e:
        await update.message.reply_text(f"Gagal mengambil data: {e}")

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
                total += int(float(re.sub(r'[^0-9.]', '', str(row[4] or 0))))
                count += 1
            except Exception:
                continue
        sheet_name = get_sheet_name_besar() if group == "besar" else get_sheet_name_kecil()
        await update.message.reply_text(
            f"Total {label}\nSheet: {sheet_name}\n\nTotal: {fmt_rupiah(total)}\nTransaksi: {count}"
        )
    except Exception as e:
        await update.message.reply_text(f"Gagal mengambil data: {e}")

# ─────────────────────────────────────────
# Foto Handler
# ─────────────────────────────────────────

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat = update.effective_chat
    group = get_group_type(chat.title or "")
    user_id = update.effective_user.id
    msg = await update.message.reply_text("⏳ Menganalisa struk, mohon tunggu...")
    try:
        photo = update.message.photo[-1]
        file = await context.bot.get_file(photo.file_id)
        image_bytes = bytes(await file.download_as_bytearray())

        foto_hash = hashlib.md5(image_bytes).hexdigest()
        if foto_hash in processed_hashes:
            await msg.edit_text("⚠️ Foto ini sudah pernah dicatat!\nKirim foto berbeda untuk transaksi baru.")
            return ConversationHandler.END

        data = analyze_image(image_bytes)

        filename = f"struk_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{foto_hash[:8]}.jpg"
        foto_link = upload_foto_to_drive(image_bytes, filename)
        logger.info(f"[PHOTO] foto_link: '{foto_link}'")

        user_data_temp[user_id] = {
            "data": data, "group": group, "foto_hash": foto_hash,
            "foto_link": foto_link,
            "dicatat_oleh": update.effective_user.first_name or "Bot",
        }

        tgl_display = format_tanggal_display(data["tanggal"])
        foto_status = "✅ Foto terupload ke Cloudinary" if foto_link else "⚠️ Foto gagal upload"

        await msg.edit_text(
            f"📋 Struk terdeteksi!\n\n🏪 Vendor: {data['vendor']}\n{foto_status}\n\n"
            f"📅 Konfirmasi tanggal:\nTanggal: {tgl_display} ({data['tanggal']})\n\nApakah tanggal ini benar?",
            reply_markup=build_konfirmasi_keyboard("tgl")
        )
        return KONFIRMASI_TANGGAL

    except Exception as e:
        logger.error(f"[PHOTO] Error: {e}", exc_info=True)
        await msg.edit_text(f"❌ Terjadi error: {e}")
        return ConversationHandler.END

async def handle_konfirmasi_tanggal(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id

    if query.data == "tgl_batal":
        user_data_temp.pop(user_id, None)
        await query.edit_message_text("❌ Transaksi dibatalkan.")
        return ConversationHandler.END

    if user_id not in user_data_temp:
        await query.edit_message_text("⚠️ Session expired. Kirim foto ulang.")
        return ConversationHandler.END

    data = user_data_temp[user_id]["data"]

    if query.data == "tgl_benar":
        await query.edit_message_text(
            f"✅ Tanggal: {data['tanggal']}\n\n💰 Konfirmasi nominal:\nNominal: {fmt_rupiah(data['jumlah'])}\n\nApakah nominal ini benar?",
            reply_markup=build_konfirmasi_keyboard("nom")
        )
        return KONFIRMASI_NOMINAL
    elif query.data == "tgl_ubah":
        await query.edit_message_text("✏️ Ketik tanggal yang benar:\nFormat: DD/MM/YYYY\nContoh: 04/06/2026")
        return INPUT_TANGGAL

async def handle_input_tanggal(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text.strip()
    if user_id not in user_data_temp:
        await update.message.reply_text("⚠️ Session expired.")
        return ConversationHandler.END
    tanggal = parse_tanggal(text)
    if not re.match(r'^\d{4}-\d{2}-\d{2}$', tanggal):
        await update.message.reply_text("❌ Format tidak valid.\nGunakan DD/MM/YYYY\nCoba lagi:")
        return INPUT_TANGGAL
    user_data_temp[user_id]["data"]["tanggal"] = tanggal
    data = user_data_temp[user_id]["data"]
    await update.message.reply_text(
        f"✅ Tanggal: {tanggal}\n\n💰 Konfirmasi nominal:\nNominal: {fmt_rupiah(data['jumlah'])}\n\nApakah benar?",
        reply_markup=build_konfirmasi_keyboard("nom")
    )
    return KONFIRMASI_NOMINAL

async def handle_konfirmasi_nominal(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    if query.data == "nom_batal":
        user_data_temp.pop(user_id, None)
        await query.edit_message_text("❌ Transaksi dibatalkan.")
        return ConversationHandler.END
    if user_id not in user_data_temp:
        await query.edit_message_text("⚠️ Session expired.")
        return ConversationHandler.END
    data = user_data_temp[user_id]["data"]
    if query.data == "nom_benar":
        await query.edit_message_text(
            f"✅ Tanggal: {data['tanggal']}\n✅ Nominal: {fmt_rupiah(data['jumlah'])}\n\n📂 Pilih kategori:",
            reply_markup=build_kategori_keyboard()
        )
        return PILIH_KATEGORI
    elif query.data == "nom_ubah":
        await query.edit_message_text("✏️ Ketik nominal yang benar (angka saja):\nContoh: 62500")
        return INPUT_NOMINAL

async def handle_input_nominal(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text.strip()
    if user_id not in user_data_temp:
        await update.message.reply_text("⚠️ Session expired.")
        return ConversationHandler.END
    jumlah_raw = re.sub(r'[^0-9]', '', text)
    if not jumlah_raw:
        await update.message.reply_text("❌ Format tidak valid. Ketik angka saja:\nCoba lagi:")
        return INPUT_NOMINAL
    jumlah = int(jumlah_raw)
    user_data_temp[user_id]["data"]["jumlah"] = jumlah
    data = user_data_temp[user_id]["data"]
    await update.message.reply_text(
        f"✅ Nominal: {fmt_rupiah(jumlah)}\n\n📂 Pilih kategori:",
        reply_markup=build_kategori_keyboard()
    )
    return PILIH_KATEGORI

async def handle_kategori(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    if query.data == "kat_BATAL":
        user_data_temp.pop(user_id, None)
        await query.edit_message_text("❌ Transaksi dibatalkan.")
        return ConversationHandler.END
    if user_id not in user_data_temp:
        await query.edit_message_text("⚠️ Session expired.")
        return ConversationHandler.END
    kategori = query.data.replace("kat_", "")
    user_data_temp[user_id]["kategori"] = kategori
    data = user_data_temp[user_id]["data"]
    await query.edit_message_text(
        f"✅ Tanggal: {data['tanggal']}\n✅ Nominal: {fmt_rupiah(data['jumlah'])}\n✅ Kategori: {kategori}\n\n"
        f"✏️ Tulis deskripsi:\nContoh: bayar gaji, beli ATK"
    )
    return TULIS_DESKRIPSI

async def handle_deskripsi(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    deskripsi = update.message.text.strip()
    if user_id not in user_data_temp:
        await update.message.reply_text("⚠️ Session expired.")
        return ConversationHandler.END
    user_data_temp[user_id]["data"]["deskripsi"] = deskripsi
    await update.message.reply_text(
        f"📝 Deskripsi: {deskripsi}\n\n🏦 Masukkan nomor rekening tujuan:\n(Ketik '-' jika tunai)"
    )
    return INPUT_REKENING

async def handle_input_rekening(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    rekening = update.message.text.strip()
    if user_id not in user_data_temp:
        await update.message.reply_text("⚠️ Session expired.")
        return ConversationHandler.END
    user_data_temp[user_id]["rekening"] = rekening
    await update.message.reply_text(f"🏦 Rekening: {rekening}\n\n👤 Masukkan nama penerima:\n(Ketik '-' jika tidak ada)")
    return INPUT_PENERIMA

async def handle_input_penerima(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    penerima = update.message.text.strip()
    if user_id not in user_data_temp:
        await update.message.reply_text("⚠️ Session expired.")
        return ConversationHandler.END

    temp = user_data_temp[user_id]
    data = temp["data"]
    data["kategori"] = temp["kategori"]
    data["rekening"] = temp["rekening"]
    data["penerima"] = penerima
    group = temp["group"]
    dicatat_oleh = temp["dicatat_oleh"]
    foto_hash = temp["foto_hash"]
    foto_link = temp.get("foto_link", "")
    label = "Kas Besar" if group == "besar" else "Kas Kecil"

    try:
        dt = datetime.strptime(data["tanggal"], "%Y-%m-%d")
    except Exception:
        dt = datetime.now()

    try:
        ws = get_sheet(group, dt)
        if group == "besar":
            append_kas_besar(ws, data, dicatat_oleh, foto_link)
        else:
            append_kas_kecil(ws, data, dicatat_oleh, foto_link)

        processed_hashes.add(foto_hash)
        user_data_temp.pop(user_id, None)
        sheet_name = get_sheet_name_besar(dt) if group == "besar" else get_sheet_name_kecil(dt)
        foto_info = f"\n🔗 Foto: {foto_link}" if foto_link else "\n⚠️ Foto tidak terupload"

        await update.message.reply_text(
            f"✅ Berhasil dicatat ke {label}!\n📊 Sheet: {sheet_name}\n\n"
            f"📅 Tanggal: {data['tanggal']}\n🏪 Vendor: {data['vendor']}\n"
            f"📝 Deskripsi: {data['deskripsi']}\n📂 Kategori: {data['kategori']}\n"
            f"💰 Jumlah: {fmt_rupiah(data['jumlah'])}\n💳 Metode: {data['metode']}\n"
            f"🏦 Rekening: {data['rekening']}\n👤 Penerima: {data['penerima']}{foto_info}"
        )
    except Exception as e:
        logger.error(f"[PENERIMA] Error: {e}", exc_info=True)
        await update.message.reply_text(f"❌ Gagal menyimpan.\nError: {e}")

    return ConversationHandler.END

# ─────────────────────────────────────────
# Edit Handlers
# ─────────────────────────────────────────

async def cmd_edit(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat = update.effective_chat
    group = get_group_type(chat.title or "")
    user_id = update.effective_user.id
    try:
        ws = get_sheet(group)
        transactions = get_all_transactions(ws, group)
        if not transactions:
            await update.message.reply_text("Belum ada transaksi.")
            return ConversationHandler.END
        user_data_temp[user_id] = {"group": group, "transactions": transactions, "edit_page": 0}
        keyboard, total_pages = build_transaksi_keyboard(transactions, 0, group)
        await update.message.reply_text(
            f"Pilih transaksi:\nHalaman 1/{total_pages} ({len(transactions)} transaksi)",
            reply_markup=keyboard
        )
        return EDIT_PILIH_TRANSAKSI
    except Exception as e:
        await update.message.reply_text(f"❌ Gagal: {e}")
        return ConversationHandler.END

async def handle_edit_pilih_transaksi(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id

    if query.data == "edit_batal":
        user_data_temp.pop(user_id, None)
        await query.edit_message_text("❌ Edit dibatalkan.")
        return ConversationHandler.END

    if user_id not in user_data_temp:
        await query.edit_message_text("⚠️ Session expired.")
        return ConversationHandler.END

    if query.data.startswith("edit_page_"):
        page = int(query.data.replace("edit_page_", ""))
        transactions = user_data_temp[user_id]["transactions"]
        group = user_data_temp[user_id]["group"]
        keyboard, total_pages = build_transaksi_keyboard(transactions, page, group)
        await query.edit_message_text(
            f"Pilih transaksi:\nHalaman {page + 1}/{total_pages}",
            reply_markup=keyboard
        )
        return EDIT_PILIH_TRANSAKSI

    row_idx = int(query.data.replace("edit_trx_", ""))
    transactions = user_data_temp[user_id]["transactions"]
    trx = next((t for t in transactions if t["row_idx"] == row_idx), None)
    if not trx:
        await query.edit_message_text("❌ Tidak ditemukan.")
        return ConversationHandler.END

    user_data_temp[user_id]["edit_row_idx"] = row_idx
    user_data_temp[user_id]["edit_trx"] = trx

    try:
        nominal = int(float(re.sub(r'[^0-9.]', '', str(trx["nominal"] or 0))))
    except Exception:
        nominal = 0

    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("✏️ Deskripsi", callback_data="editfield_deskripsi")],
        [InlineKeyboardButton("📂 Kategori", callback_data="editfield_kategori")],
        [InlineKeyboardButton("💰 Nominal", callback_data="editfield_nominal")],
        [InlineKeyboardButton("📅 Tanggal", callback_data="editfield_tanggal")],
        [InlineKeyboardButton("🏦 Rekening", callback_data="editfield_rekening")],
        [InlineKeyboardButton("👤 Penerima", callback_data="editfield_penerima")],
        [InlineKeyboardButton("❌ Batalkan", callback_data="editfield_batal")],
    ])
    await query.edit_message_text(
        f"📅 {trx['tanggal']} | 📂 {trx['kategori']}\n📝 {trx['deskripsi']}\n💰 {fmt_rupiah(nominal)}\n\nField mana?",
        reply_markup=keyboard
    )
    return EDIT_PILIH_FIELD

async def handle_edit_pilih_field(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id

    if query.data == "editfield_batal":
        user_data_temp.pop(user_id, None)
        await query.edit_message_text("❌ Edit dibatalkan.")
        return ConversationHandler.END

    if user_id not in user_data_temp:
        await query.edit_message_text("⚠️ Session expired.")
        return ConversationHandler.END

    field = query.data.replace("editfield_", "")
    user_data_temp[user_id]["edit_field"] = field

    if field == "kategori":
        await query.edit_message_text("📂 Pilih kategori baru:", reply_markup=build_kategori_keyboard())
        return EDIT_INPUT_NILAI

    prompts = {
        "deskripsi": "✏️ Ketik deskripsi baru:",
        "nominal": "💰 Ketik nominal baru (angka saja):",
        "tanggal": "📅 Ketik tanggal baru (DD/MM/YYYY):",
        "rekening": "🏦 Ketik rekening baru:",
        "penerima": "👤 Ketik nama penerima baru:",
    }
    await query.edit_message_text(prompts.get(field, "Ketik nilai baru:"))
    return EDIT_INPUT_NILAI

async def handle_edit_input_nilai(update: Update, context: ContextTypes.DEFAULT_TYPE):
    is_callback = update.callback_query is not None
    user_id = update.callback_query.from_user.id if is_callback else update.effective_user.id

    if is_callback:
        query = update.callback_query
        await query.answer()
        if query.data == "kat_BATAL":
            user_data_temp.pop(user_id, None)
            await query.edit_message_text("❌ Edit dibatalkan.")
            return ConversationHandler.END
        nilai_baru = query.data.replace("kat_", "")
        reply = query.edit_message_text
    else:
        nilai_baru = update.message.text.strip()
        reply = update.message.reply_text

    if user_id not in user_data_temp:
        await reply("⚠️ Session expired.")
        return ConversationHandler.END

    temp = user_data_temp[user_id]
    field = temp["edit_field"]
    row_idx = temp["edit_row_idx"]
    group = temp["group"]
    editor = update.effective_user.first_name or "Admin"

    if field == "nominal":
        clean = re.sub(r'[^0-9]', '', nilai_baru)
        if not clean:
            await reply("❌ Angka tidak valid. Coba lagi:")
            return EDIT_INPUT_NILAI
        nilai_baru = fmt_rupiah(int(clean))
    elif field == "tanggal":
        nilai_baru = parse_tanggal(nilai_baru)
        if not re.match(r'^\d{4}-\d{2}-\d{2}$', nilai_baru):
            await reply("❌ Format tidak valid. DD/MM/YYYY:")
            return EDIT_INPUT_NILAI

    try:
        ws = get_sheet(group)
        field_col_map = {"deskripsi": 3, "kategori": 4, "nominal": 5, "tanggal": 9, "rekening": 10, "penerima": 11}
        col_idx = field_col_map.get(field)
        if col_idx:
            ws.update_cell(row_idx, col_idx, nilai_baru)
        ws.update_cell(row_idx, 12, f"EDITED - {datetime.now().strftime('%d/%m/%Y %H:%M')} by {editor}")
        if field == "nominal":
            recalculate_saldo(ws)
        user_data_temp.pop(user_id, None)
        await reply(f"✅ Berhasil diedit!\nField: {field}\nNilai baru: {nilai_baru}")
    except Exception as e:
        await reply(f"❌ Gagal: {e}")

    return ConversationHandler.END

# ─────────────────────────────────────────
# Delete Handlers
# ─────────────────────────────────────────

def catat_log_delete(group, trx, deleted_by):
    try:
        gc = get_gspread_client()
        sid = os.environ["SPREADSHEET_ID_KAS_KECIL"] if group == "kecil" else os.environ["SPREADSHEET_ID_KAS_BESAR"]
        sh = gc.open_by_key(sid)
        try:
            ws_log = sh.worksheet("Log Delete")
        except gspread.WorksheetNotFound:
            ws_log = sh.add_worksheet(title="Log Delete", rows=1000, cols=9)
            ws_log.append_row(["Waktu Delete", "Dihapus Oleh", "Group", "Tanggal",
                               "Deskripsi", "Kategori", "Nominal", "Vendor", "Status"])
        ws_log.append_row([
            datetime.now().strftime("%d/%m/%Y %H:%M:%S"), deleted_by,
            "Kas Besar" if group == "besar" else "Kas Kecil",
            trx.get("tanggal",""), trx.get("deskripsi",""), trx.get("kategori",""),
            trx.get("nominal",""), trx.get("vendor",""), trx.get("status",""),
        ])
    except Exception as e:
        logger.error(f"[DELETE] Log gagal: {e}")

async def cmd_delete(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id not in ADMIN_IDS:
        await update.message.reply_text("⛔ Hanya admin yang bisa menghapus.")
        return ConversationHandler.END
    chat = update.effective_chat
    group = get_group_type(chat.title or "")
    try:
        ws = get_sheet(group)
        transactions = get_all_transactions(ws, group)
        if not transactions:
            await update.message.reply_text("Belum ada transaksi.")
            return ConversationHandler.END
        user_data_temp[user_id] = {"group": group, "transactions": transactions, "delete_page": 0}
        keyboard, total_pages = build_transaksi_keyboard(transactions, 0, group)
        await update.message.reply_text(
            f"🗑️ Pilih transaksi:\nHalaman 1/{total_pages} ({len(transactions)} transaksi)",
            reply_markup=keyboard
        )
        return DELETE_PILIH_TRANSAKSI
    except Exception as e:
        await update.message.reply_text(f"❌ Gagal: {e}")
        return ConversationHandler.END

async def handle_delete_pilih_transaksi(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id

    if query.data == "edit_batal":
        user_data_temp.pop(user_id, None)
        await query.edit_message_text("❌ Dibatalkan.")
        return ConversationHandler.END

    if user_id not in user_data_temp:
        await query.edit_message_text("⚠️ Session expired.")
        return ConversationHandler.END

    if query.data.startswith("edit_page_"):
        page = int(query.data.replace("edit_page_", ""))
        transactions = user_data_temp[user_id]["transactions"]
        group = user_data_temp[user_id]["group"]
        keyboard, total_pages = build_transaksi_keyboard(transactions, page, group)
        await query.edit_message_text(f"🗑️ Pilih transaksi:\nHalaman {page+1}/{total_pages}", reply_markup=keyboard)
        return DELETE_PILIH_TRANSAKSI

    row_idx = int(query.data.replace("edit_trx_", ""))
    transactions = user_data_temp[user_id]["transactions"]
    trx = next((t for t in transactions if t["row_idx"] == row_idx), None)
    if not trx:
        await query.edit_message_text("❌ Tidak ditemukan.")
        return ConversationHandler.END

    user_data_temp[user_id]["delete_row_idx"] = row_idx
    user_data_temp[user_id]["delete_trx"] = trx

    try:
        nominal = int(float(re.sub(r'[^0-9.]', '', str(trx["nominal"] or 0))))
    except Exception:
        nominal = 0

    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("🗑️ Ya, Hapus", callback_data="delete_konfirmasi_ya")],
        [InlineKeyboardButton("❌ Batal", callback_data="delete_konfirmasi_batal")],
    ])
    await query.edit_message_text(
        f"⚠️ Hapus?\n📅 {trx['tanggal']}\n📝 {trx['deskripsi']}\n💰 {fmt_rupiah(nominal)}",
        reply_markup=keyboard
    )
    return DELETE_KONFIRMASI

async def handle_delete_konfirmasi(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id

    if query.data == "delete_konfirmasi_batal":
        user_data_temp.pop(user_id, None)
        await query.edit_message_text("❌ Dibatalkan.")
        return ConversationHandler.END

    if user_id not in user_data_temp:
        await query.edit_message_text("⚠️ Session expired.")
        return ConversationHandler.END

    temp = user_data_temp[user_id]
    row_idx = temp["delete_row_idx"]
    trx = temp["delete_trx"]
    group = temp["group"]
    deleted_by = update.effective_user.first_name or "Admin"

    try:
        ws = get_sheet(group)
        catat_log_delete(group, trx, deleted_by)
        ws.delete_rows(row_idx)
        recalculate_saldo(ws)
        user_data_temp.pop(user_id, None)
        await query.edit_message_text(f"✅ Terhapus!\n📝 {trx['deskripsi']}\n👤 Oleh: {deleted_by}")
    except Exception as e:
        await query.edit_message_text(f"❌ Gagal: {e}")

    return ConversationHandler.END

# ─────────────────────────────────────────
# Reset Bulan
# ─────────────────────────────────────────

async def cmd_reset_bulan(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id not in ADMIN_IDS:
        await update.message.reply_text("⛔ Hanya admin.")
        return ConversationHandler.END
    chat = update.effective_chat
    group = get_group_type(chat.title or "")
    sheet_name = get_sheet_name_besar() if group == "besar" else get_sheet_name_kecil()
    user_data_temp[user_id] = {"group": group}
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("🗑️ Ya, Hapus Semua", callback_data="reset_ya")],
        [InlineKeyboardButton("❌ Batal", callback_data="reset_batal")],
    ])
    await update.message.reply_text(
        f"⚠️ Reset Data Bulan Ini\nSheet: {sheet_name}\n\nSemua data akan dihapus permanen!",
        reply_markup=keyboard
    )
    return RESET_KONFIRMASI

async def handle_reset_konfirmasi(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id

    if query.data == "reset_batal":
        user_data_temp.pop(user_id, None)
        await query.edit_message_text("❌ Reset dibatalkan.")
        return ConversationHandler.END

    if user_id not in user_data_temp:
        await query.edit_message_text("⚠️ Session expired.")
        return ConversationHandler.END

    group = user_data_temp[user_id]["group"]
    deleted_by = update.effective_user.first_name or "Admin"
    sheet_name = get_sheet_name_besar() if group == "besar" else get_sheet_name_kecil()

    try:
        ws = get_sheet(group)
        rows = ws.get_all_values()
        data_rows = [r for r in rows[1:] if r and r[0]]
        total = len(data_rows)

        if total == 0:
            user_data_temp.pop(user_id, None)
            await query.edit_message_text("ℹ️ Tidak ada data.")
            return ConversationHandler.END

        gc = get_gspread_client()
        sid = os.environ["SPREADSHEET_ID_KAS_KECIL"] if group == "kecil" else os.environ["SPREADSHEET_ID_KAS_BESAR"]
        sh = gc.open_by_key(sid)
        try:
            ws_log = sh.worksheet("Log Delete")
        except gspread.WorksheetNotFound:
            ws_log = sh.add_worksheet(title="Log Delete", rows=1000, cols=9)
            ws_log.append_row(["Waktu Delete", "Dihapus Oleh", "Group", "Tanggal",
                               "Deskripsi", "Kategori", "Nominal", "Vendor", "Status"])

        log_rows = []
        label = "Kas Besar" if group == "besar" else "Kas Kecil"
        for row in data_rows:
            log_rows.append([
                datetime.now().strftime("%d/%m/%Y %H:%M:%S"), f"{deleted_by} (RESET)", label,
                row[0] if len(row) > 0 else "", row[2] if len(row) > 2 else "",
                row[3] if len(row) > 3 else "", row[4] if len(row) > 4 else "",
                row[7] if len(row) > 7 else "", row[11] if len(row) > 11 else "",
            ])
        if log_rows:
            ws_log.append_rows(log_rows)

        if len(rows) > 1:
            ws.delete_rows(2, len(rows))

        user_data_temp.pop(user_id, None)
        await query.edit_message_text(
            f"✅ Reset selesai!\n📊 {sheet_name}\n🗑️ {total} transaksi dihapus\n👤 Oleh: {deleted_by}"
        )
    except Exception as e:
        await query.edit_message_text(f"❌ Gagal reset: {e}")

    return ConversationHandler.END

async def handle_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_data_temp.pop(user_id, None)
    await update.message.reply_text("❌ Proses dibatalkan.")
    return ConversationHandler.END

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat = update.effective_chat
    group = get_group_type(chat.title or "")
    if group == "pos":
        await update.message.reply_text("Perintah POS: /jual /omzet /laporan /produk\nAdmin: /tambah_produk /batal")
    else:
        await update.message.reply_text("Kirim foto struk untuk mencatat.\nPerintah: /cek /total /edit /delete /reset_bulan /batal")

# ═══════════════════════════════════════════
# POS MODULE
# ═══════════════════════════════════════════

def get_sheet_name_pos(dt=None):
    d = dt or datetime.now()
    return f"Penjualan {BULAN[d.month-1]} {d.year}"

def get_pos_spreadsheet_id():
    return os.environ.get("SPREADSHEET_ID_POS", os.environ.get("SPREADSHEET_ID_KAS_BESAR"))

def get_pos_sheet(dt=None):
    gc = get_gspread_client()
    sid = get_pos_spreadsheet_id()
    sh = gc.open_by_key(sid)
    sname = get_sheet_name_pos(dt)
    try:
        ws = sh.worksheet(sname)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=sname, rows=2000, cols=15)
        ws.append_row([
            "No Nota", "Waktu", "Outlet", "Kasir", "Capster",
            "Nama Customer", "No HP Customer", "Produk", "Qty",
            "Harga Satuan", "Total", "Metode Bayar", "Tunai", "Kembalian", "Status"
        ])
        ws.format("A1:O1", {
            "backgroundColor": {"red": 0.18, "green": 0.33, "blue": 0.59},
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}}
        })
    return ws

def get_produk_sheet():
    gc = get_gspread_client()
    sid = get_pos_spreadsheet_id()
    sh = gc.open_by_key(sid)
    try:
        ws = sh.worksheet("Master Produk")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title="Master Produk", rows=200, cols=6)
        ws.append_row(["Kode", "Nama Produk", "Harga", "Kategori", "Aktif", "Dibuat"])
        ws.format("A1:F1", {
            "backgroundColor": {"red": 0.1, "green": 0.5, "blue": 0.3},
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}}
        })
    return ws

def get_all_produk():
    try:
        ws = get_produk_sheet()
        rows = ws.get_all_values()
        produk = []
        for i, row in enumerate(rows[1:], start=2):
            if not row or not row[0]:
                continue
            aktif = str(row[4]).strip().upper() if len(row) > 4 else "YA"
            if aktif in ["YA", "YES", "1", "TRUE", ""]:
                produk.append({
                    "row_idx": i, "kode": row[0],
                    "nama": row[1] if len(row) > 1 else "",
                    "harga": int(re.sub(r"[^0-9]", "", str(row[2]))) if len(row) > 2 and row[2] else 0,
                    "kategori": row[3] if len(row) > 3 else "",
                })
        return produk
    except Exception as e:
        logger.error(f"[POS] get_all_produk error: {e}")
        return []

def generate_no_nota(outlet):
    now = datetime.now()
    outlet_code = "".join([c for c in (outlet or "OUT").upper() if c.isalpha()])[:3]
    return f"{outlet_code}/{now.strftime('%y%m%d')}/{now.strftime('%H%M%S')}"

def get_omzet_hari_ini(outlet=None):
    try:
        ws = get_pos_sheet()
        rows = ws.get_all_values()
        today = datetime.now().strftime("%d/%m/%Y")
        total = 0
        count = 0
        for row in rows[1:]:
            if not row or not row[0]:
                continue
            if today not in (row[1] if len(row) > 1 else ""):
                continue
            try:
                total += int(re.sub(r"[^0-9]", "", str(row[10]))) if len(row) > 10 and row[10] else 0
                count += 1
            except Exception:
                continue
        return total, count
    except Exception as e:
        logger.error(f"[POS] omzet error: {e}")
        return 0, 0

def get_all_capster():
    """Ambil semua capster aktif dari sheet Master Capster"""
    try:
        gc = get_gspread_client()
        sid = get_pos_spreadsheet_id()
        sh = gc.open_by_key(sid)

        # Coba semua kemungkinan nama sheet
        ws = None
        sheet_names = [s.title for s in sh.worksheets()]
        logger.info(f"[CAPSTER] Sheet tersedia: {sheet_names}")

        for name in sheet_names:
            if "capster" in name.lower():
                ws = sh.worksheet(name)
                logger.info(f"[CAPSTER] Pakai sheet: {name}")
                break

        if not ws:
            logger.error(f"[CAPSTER] Sheet capster tidak ditemukan! Available: {sheet_names}")
            return []

        rows = ws.get_all_values()
        logger.info(f"[CAPSTER] Total rows: {len(rows)}, data: {rows}")
        capster = []
        for row in rows[1:]:
            if not row or not row[0]:
                continue
            aktif = str(row[1]).strip().upper() if len(row) > 1 else "YA"
            if aktif in ["YA", "YES", "1", "TRUE", ""]:
                capster.append(row[0].strip())
        logger.info(f"[CAPSTER] Capster aktif: {capster}")
        return capster
    except Exception as e:
        logger.error(f"[POS] get_all_capster error: {e}", exc_info=True)
        return []

def build_capster_keyboard(capster_list):
    keyboard = []
    for nama in capster_list:
        keyboard.append([InlineKeyboardButton(f"✂️ {nama}", callback_data=f"pos_capster_{nama}")])
    keyboard.append([InlineKeyboardButton("❌ Batalkan", callback_data="pos_batal")])
    return InlineKeyboardMarkup(keyboard)

def build_produk_keyboard(produk_list, selected=None):
    keyboard = []
    for p in produk_list:
        label = f"{'✅ ' if selected and p['kode'] == selected else ''}{p['nama']} - {fmt_rupiah(p['harga'])}"
        keyboard.append([InlineKeyboardButton(label, callback_data=f"pos_produk_{p['kode']}")])
    keyboard.append([InlineKeyboardButton("❌ Batalkan", callback_data="pos_batal")])
    return InlineKeyboardMarkup(keyboard)

def build_bayar_keyboard():
    methods = [
        ("💵 Tunai", "pos_bayar_tunai"),
        ("🏦 Transfer Bank", "pos_bayar_transfer"),
        ("📱 QRIS", "pos_bayar_qris"),
        ("💳 Kartu Debit", "pos_bayar_debit"),
        ("💳 Kartu Kredit", "pos_bayar_kredit"),
    ]
    keyboard = [[InlineKeyboardButton(label, callback_data=cb)] for label, cb in methods]
    keyboard.append([InlineKeyboardButton("❌ Batalkan", callback_data="pos_batal")])
    return InlineKeyboardMarkup(keyboard)

async def cmd_pos_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat = update.effective_chat
    produk = get_all_produk()
    await update.message.reply_text(
        f"🏪 KasBot POS — {chat.title}\nMenu: {len(produk)} produk\n\n"
        f"/jual - Input transaksi\n/omzet - Omzet hari ini\n"
        f"/laporan - Rekap bulanan\n/produk - Daftar produk\n"
        f"/tambah_produk - Tambah produk (admin)\n/batal - Batalkan"
    )

async def cmd_produk(update: Update, context: ContextTypes.DEFAULT_TYPE):
    produk = get_all_produk()
    if not produk:
        await update.message.reply_text("Belum ada produk.\nTambah dengan /tambah_produk")
        return
    text = "📋 Daftar Produk\n\n"
    for i, p in enumerate(produk, 1):
        text += f"{i}. {p['nama']} — {fmt_rupiah(p['harga'])}\n"
    await update.message.reply_text(text)

async def cmd_tambah_produk(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id not in ADMIN_IDS:
        await update.message.reply_text("⛔ Hanya admin.")
        return ConversationHandler.END
    user_data_temp[user_id] = {"pos_action": "tambah_produk"}
    await update.message.reply_text("➕ Nama produk/layanan:\nContoh: Hair Cut")
    return POS_SETUP_NAMA

async def handle_pos_setup_nama(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    nama = update.message.text.strip()
    if user_id not in user_data_temp:
        await update.message.reply_text("⚠️ Session expired.")
        return ConversationHandler.END
    user_data_temp[user_id]["produk_nama"] = nama
    await update.message.reply_text(f"Nama: {nama}\n\n💰 Harga (angka saja):\nContoh: 60000")
    return POS_SETUP_HARGA

async def handle_pos_setup_harga(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text.strip()
    if user_id not in user_data_temp:
        await update.message.reply_text("⚠️ Session expired.")
        return ConversationHandler.END
    harga_raw = re.sub(r"[^0-9]", "", text)
    if not harga_raw:
        await update.message.reply_text("❌ Format tidak valid:")
        return POS_SETUP_HARGA
    harga = int(harga_raw)
    user_data_temp[user_id]["produk_harga"] = harga
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ Simpan", callback_data="pos_setup_simpan")],
        [InlineKeyboardButton("❌ Batalkan", callback_data="pos_setup_batal")],
    ])
    await update.message.reply_text(
        f"Nama: {user_data_temp[user_id]['produk_nama']}\nHarga: {fmt_rupiah(harga)}\n\nSimpan?",
        reply_markup=keyboard
    )
    return POS_SETUP_KONFIRMASI

async def handle_pos_setup_konfirmasi(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    if query.data == "pos_setup_batal":
        user_data_temp.pop(user_id, None)
        await query.edit_message_text("❌ Dibatalkan.")
        return ConversationHandler.END
    if user_id not in user_data_temp:
        await query.edit_message_text("⚠️ Session expired.")
        return ConversationHandler.END
    temp = user_data_temp[user_id]
    nama = temp["produk_nama"]
    harga = temp["produk_harga"]
    try:
        ws = get_produk_sheet()
        rows = ws.get_all_values()
        kode = f"P{len(rows):03d}"
        ws.append_row([kode, nama, fmt_rupiah(harga), "", "YA", datetime.now().strftime("%d/%m/%Y %H:%M")])
        user_data_temp.pop(user_id, None)
        await query.edit_message_text(f"✅ Produk ditambahkan!\nKode: {kode}\nNama: {nama}\nHarga: {fmt_rupiah(harga)}")
    except Exception as e:
        await query.edit_message_text(f"❌ Gagal: {e}")
    return ConversationHandler.END

async def cmd_jual(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    chat = update.effective_chat
    produk = get_all_produk()
    if not produk:
        await update.message.reply_text("⚠️ Belum ada produk! Minta admin tambah dulu.")
        return ConversationHandler.END
    user_data_temp[user_id] = {
        "pos_mode": True, "outlet": chat.title or "Outlet",
        "kasir": update.effective_user.first_name or "Kasir",
        "produk_list": produk, "keranjang": [],
    }
    await update.message.reply_text("🛒 Transaksi Baru\n\nPilih produk:", reply_markup=build_produk_keyboard(produk))
    return POS_PILIH_PRODUK

async def handle_pos_pilih_produk(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    if query.data == "pos_batal":
        user_data_temp.pop(user_id, None)
        await query.edit_message_text("❌ Dibatalkan.")
        return ConversationHandler.END
    if user_id not in user_data_temp:
        await query.edit_message_text("⚠️ Session expired.")
        return ConversationHandler.END
    kode = query.data.replace("pos_produk_", "")
    produk = next((p for p in user_data_temp[user_id]["produk_list"] if p["kode"] == kode), None)
    if not produk:
        await query.edit_message_text("❌ Produk tidak ditemukan.")
        return ConversationHandler.END
    user_data_temp[user_id]["produk_dipilih"] = produk
    await query.edit_message_text(f"✅ {produk['nama']} — {fmt_rupiah(produk['harga'])}\n\nKetik qty:")
    return POS_INPUT_QTY

async def handle_pos_input_qty(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text.strip()
    if user_id not in user_data_temp:
        await update.message.reply_text("⚠️ Session expired.")
        return ConversationHandler.END
    qty_raw = re.sub(r"[^0-9]", "", text)
    if not qty_raw or int(qty_raw) < 1:
        await update.message.reply_text("❌ Qty minimal 1:")
        return POS_INPUT_QTY
    qty = int(qty_raw)
    produk = user_data_temp[user_id]["produk_dipilih"]
    subtotal = produk["harga"] * qty
    user_data_temp[user_id]["keranjang"].append({
        "kode": produk["kode"], "nama": produk["nama"],
        "harga": produk["harga"], "qty": qty, "subtotal": subtotal,
    })
    keranjang = user_data_temp[user_id]["keranjang"]
    grand_total = sum(i["subtotal"] for i in keranjang)
    text_keranjang = "🛒 Keranjang:\n"
    for item in keranjang:
        text_keranjang += f"• {item['nama']} x{item['qty']} = {fmt_rupiah(item['subtotal'])}\n"
    text_keranjang += f"\n💰 Total: {fmt_rupiah(grand_total)}"
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("➕ Tambah Produk", callback_data="pos_tambah_lagi")],
        [InlineKeyboardButton("💳 Lanjut Bayar", callback_data="pos_lanjut_bayar")],
        [InlineKeyboardButton("❌ Batalkan", callback_data="pos_batal")],
    ])
    await update.message.reply_text(text_keranjang, reply_markup=keyboard)
    return POS_PILIH_PRODUK

async def handle_pos_keranjang_action(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    if query.data == "pos_batal":
        user_data_temp.pop(user_id, None)
        await query.edit_message_text("❌ Dibatalkan.")
        return ConversationHandler.END
    if user_id not in user_data_temp:
        await query.edit_message_text("⚠️ Session expired.")
        return ConversationHandler.END
    if query.data == "pos_tambah_lagi":
        produk_list = user_data_temp[user_id]["produk_list"]
        await query.edit_message_text("Pilih produk:", reply_markup=build_produk_keyboard(produk_list))
        return POS_PILIH_PRODUK
    if query.data == "pos_lanjut_bayar":
        keranjang = user_data_temp[user_id]["keranjang"]
        grand_total = sum(i["subtotal"] for i in keranjang)
        user_data_temp[user_id]["grand_total"] = grand_total
        capster_list = get_all_capster()
        if not capster_list:
            await query.edit_message_text(f"💰 Total: {fmt_rupiah(grand_total)}\n\n✂️ Ketik nama capster:")
            return POS_INPUT_CAPSTER
        user_data_temp[user_id]["capster_list"] = capster_list
        await query.edit_message_text(
            f"💰 Total: {fmt_rupiah(grand_total)}\n\n✂️ Pilih capster:",
            reply_markup=build_capster_keyboard(capster_list)
        )
        return POS_PILIH_CAPSTER

async def handle_pos_pilih_capster(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id

    if query.data == "pos_batal":
        user_data_temp.pop(user_id, None)
        await query.edit_message_text("❌ Dibatalkan.")
        return ConversationHandler.END

    if user_id not in user_data_temp:
        await query.edit_message_text("⚠️ Session expired.")
        return ConversationHandler.END

    capster = query.data.replace("pos_capster_", "")
    user_data_temp[user_id]["capster"] = capster
    await query.edit_message_text(
        f"✂️ Capster: {capster}\n\n👤 Nama customer:\n(Ketik '-' jika tidak mau isi)"
    )
    return POS_INPUT_NAMA_CUSTOMER

async def handle_pos_input_capster(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    capster = update.message.text.strip()
    if user_id not in user_data_temp:
        await update.message.reply_text("⚠️ Session expired.")
        return ConversationHandler.END
    if not capster:
        await update.message.reply_text("❌ Nama capster tidak boleh kosong.")
        return POS_INPUT_CAPSTER
    user_data_temp[user_id]["capster"] = capster
    await update.message.reply_text(f"✂️ Capster: {capster}\n\n👤 Nama customer:\n(Ketik '-' jika tidak mau isi)")
    return POS_INPUT_NAMA_CUSTOMER

async def handle_pos_input_nama_customer(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    nama = update.message.text.strip()
    if user_id not in user_data_temp:
        await update.message.reply_text("⚠️ Session expired.")
        return ConversationHandler.END
    user_data_temp[user_id]["nama_customer"] = nama
    await update.message.reply_text(f"👤 Customer: {nama}\n\n📱 No HP customer:\n(Ketik '-' jika tidak mau isi)")
    return POS_INPUT_HP_CUSTOMER

async def handle_pos_input_hp_customer(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    hp = update.message.text.strip()
    if user_id not in user_data_temp:
        await update.message.reply_text("⚠️ Session expired.")
        return ConversationHandler.END
    user_data_temp[user_id]["hp_customer"] = hp
    grand_total = user_data_temp[user_id]["grand_total"]
    capster = user_data_temp[user_id].get("capster", "-")
    nama = user_data_temp[user_id].get("nama_customer", "-")
    await update.message.reply_text(
        f"✂️ {capster} | 👤 {nama} | 📱 {hp}\n💰 Total: {fmt_rupiah(grand_total)}\n\nMetode bayar:",
        reply_markup=build_bayar_keyboard()
    )
    return POS_PILIH_BAYAR

async def handle_pos_pilih_bayar(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    if query.data == "pos_batal":
        user_data_temp.pop(user_id, None)
        await query.edit_message_text("❌ Dibatalkan.")
        return ConversationHandler.END
    if user_id not in user_data_temp:
        await query.edit_message_text("⚠️ Session expired.")
        return ConversationHandler.END
    metode_map = {
        "pos_bayar_tunai": "Tunai", "pos_bayar_transfer": "Transfer Bank",
        "pos_bayar_qris": "QRIS", "pos_bayar_debit": "Kartu Debit", "pos_bayar_kredit": "Kartu Kredit",
    }
    metode = metode_map.get(query.data, "Tunai")
    user_data_temp[user_id]["metode_bayar"] = metode
    grand_total = user_data_temp[user_id]["grand_total"]
    if metode == "Tunai":
        await query.edit_message_text(f"💵 Tunai\n💰 Total: {fmt_rupiah(grand_total)}\n\nJumlah uang diterima:")
        return POS_INPUT_TUNAI
    else:
        return await simpan_transaksi_pos(query, user_id, grand_total, 0, 0)

async def handle_pos_input_tunai(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text.strip()
    if user_id not in user_data_temp:
        await update.message.reply_text("⚠️ Session expired.")
        return ConversationHandler.END
    tunai_raw = re.sub(r"[^0-9]", "", text)
    if not tunai_raw:
        await update.message.reply_text("❌ Format tidak valid.")
        return POS_INPUT_TUNAI
    tunai = int(tunai_raw)
    grand_total = user_data_temp[user_id]["grand_total"]
    if tunai < grand_total:
        await update.message.reply_text(
            f"❌ Uang kurang!\nTotal: {fmt_rupiah(grand_total)}\nKurang: {fmt_rupiah(grand_total - tunai)}\nCoba lagi:"
        )
        return POS_INPUT_TUNAI
    return await simpan_transaksi_pos(update.message, user_id, grand_total, tunai, tunai - grand_total)

def generate_struk_image(outlet, no_nota, waktu, kasir, capster, nama_customer,
                         hp_customer, keranjang, grand_total, tunai, kembalian, metode):
    """Generate struk PNG modern minimal, return BytesIO."""
    from PIL import Image, ImageDraw, ImageFont

    W       = 640
    PAD     = 36
    WHITE   = (255, 255, 255)
    BLACK   = (26,  26,  26)
    DARK    = (26,  26,  26)
    DGRAY   = (100, 100, 100)
    LGRAY   = (200, 200, 200)
    LLGRAY  = (245, 245, 243)

    # Font
    FONT_PATHS = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    ]
    def load_font(size, bold=False):
        paths = [p for p in FONT_PATHS if ("Bold" in p) == bold]
        for p in paths:
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                continue
        return ImageFont.load_default()

    f_brand   = load_font(52, bold=True)
    f_sub     = load_font(18, bold=False)
    f_label   = load_font(20, bold=False)
    f_value   = load_font(20, bold=True)
    f_item    = load_font(22, bold=True)
    f_item_sm = load_font(19, bold=False)
    f_total_l = load_font(19, bold=False)
    f_total_v = load_font(34, bold=True)
    f_footer  = load_font(18, bold=False)

    def tw(text, font):
        bb = ImageDraw.Draw(Image.new("RGB", (1,1))).textbbox((0,0), text, font=font)
        return bb[2] - bb[0]

    def th(text, font):
        bb = ImageDraw.Draw(Image.new("RGB", (1,1))).textbbox((0,0), text, font=font)
        return bb[3] - bb[1]

    # ── Hitung tinggi ──
    HEADER_H  = 160
    INFO_H    = 14 + (4 * 32) + (1 if nama_customer != "-" else 0)*32 + (1 if hp_customer != "-" else 0)*32 + 14
    DIVIDER_H = 1
    ITEMS_H   = sum([32 + 28 + 10 for _ in keranjang]) + 16
    TOTAL_H   = 120
    TUNAI_H   = 64 if tunai > 0 else 0
    METODE_H  = 48
    FOOTER_H  = 80
    H = HEADER_H + INFO_H + DIVIDER_H + ITEMS_H + TOTAL_H + TUNAI_H + METODE_H + FOOTER_H + 20

    img  = Image.new("RGB", (W, H), WHITE)
    draw = ImageDraw.Draw(img)

    # ── HEADER (dark bg) ──
    draw.rectangle([(0, 0), (W, HEADER_H)], fill=DARK)

    # Sub label
    sub_text = "BARBERSHOP"
    draw.text(((W - tw(sub_text, f_sub)) // 2, 22), sub_text, fill=(136, 136, 136), font=f_sub)

    # Nama outlet besar
    brand = outlet.upper().replace(" BARBERSHOP POS", "").replace(" POS", "").strip()
    draw.text(((W - tw(brand, f_brand)) // 2, 44), brand, fill=WHITE, font=f_brand)

    # Divider tipis di header
    draw.line([(PAD, 118), (W - PAD, 118)], fill=(60, 60, 60), width=1)

    # No nota & waktu
    draw.text(((W - tw(no_nota, f_item_sm)) // 2, 126), no_nota, fill=(100, 100, 100), font=f_item_sm)

    y = HEADER_H + 14

    # ── INFO ──
    def draw_row(y, label, value):
        draw.text((PAD, y), label, fill=DGRAY, font=f_label)
        draw.text((W - PAD - tw(value, f_value), y), value, fill=BLACK, font=f_value)
        return y + 32

    y = draw_row(y, "Kasir",    str(kasir))
    y = draw_row(y, "Capster",  str(capster))
    if nama_customer != "-":
        y = draw_row(y, "Customer", str(nama_customer))
    if hp_customer != "-":
        y = draw_row(y, "HP",       str(hp_customer))
    y = draw_row(y, "Waktu",    str(waktu))

    # ── DIVIDER ──
    y += 10
    draw.line([(PAD, y), (W - PAD, y)], fill=LGRAY, width=1)
    y += 16

    # ── ITEMS ──
    for item in keranjang:
        draw.text((PAD, y), item["nama"], fill=BLACK, font=f_item)
        draw.text((W - PAD - tw(fmt_rupiah(item["subtotal"]), f_item), y),
                  fmt_rupiah(item["subtotal"]), fill=BLACK, font=f_item)
        y += 32
        qty_txt = f"{item['qty']} x {fmt_rupiah(item['harga'])}"
        draw.text((PAD, y), qty_txt, fill=DGRAY, font=f_item_sm)
        y += 28 + 10

    # ── DIVIDER ──
    draw.line([(PAD, y), (W - PAD, y)], fill=LGRAY, width=1)
    y += 20

    # ── TUNAI & KEMBALI ──
    if tunai > 0:
        draw.text((PAD, y), "Tunai", fill=DGRAY, font=f_total_l)
        draw.text((W - PAD - tw(fmt_rupiah(tunai), f_total_l), y),
                  fmt_rupiah(tunai), fill=DGRAY, font=f_total_l)
        y += 30
        draw.text((PAD, y), "Kembali", fill=DGRAY, font=f_total_l)
        draw.text((W - PAD - tw(fmt_rupiah(kembalian), f_total_l), y),
                  fmt_rupiah(kembalian), fill=DGRAY, font=f_total_l)
        y += 34

    # ── TOTAL ──
    draw.text((PAD, y + 6), "TOTAL", fill=DGRAY, font=f_total_l)
    draw.text((W - PAD - tw(fmt_rupiah(grand_total), f_total_v), y),
              fmt_rupiah(grand_total), fill=BLACK, font=f_total_v)
    y += 52

    # ── METODE ──
    draw.line([(PAD, y), (W - PAD, y)], fill=LGRAY, width=1)
    y += 12
    draw.text((PAD, y), "Metode bayar", fill=DGRAY, font=f_label)
    pill_text = metode
    pill_w = tw(pill_text, f_label) + 24
    pill_x = W - PAD - pill_w
    draw.rounded_rectangle([(pill_x, y - 2), (W - PAD, y + 26)], radius=12, fill=LLGRAY)
    draw.text((pill_x + 12, y), pill_text, fill=BLACK, font=f_label)
    y += 44

    # ── FOOTER ──
    draw.line([(PAD, y), (W - PAD, y)], fill=LGRAY, width=1)
    y += 16
    thanks = "Terima kasih!"
    draw.text(((W - tw(thanks, f_item_sm)) // 2, y), thanks, fill=BLACK, font=f_item_sm)
    y += 28
    kasbot = "kasbot.id"
    draw.text(((W - tw(kasbot, f_footer)) // 2, y), kasbot, fill=DGRAY, font=f_footer)

    buf = io.BytesIO()
    img.save(buf, format="PNG", dpi=(203, 203))
    buf.seek(0)
    return buf


async def simpan_transaksi_pos(reply_target, user_id, grand_total, tunai, kembalian):
    if user_id not in user_data_temp:
        return ConversationHandler.END
    temp = user_data_temp[user_id]
    keranjang = temp["keranjang"]
    outlet = temp["outlet"]
    kasir = temp["kasir"]
    metode = temp["metode_bayar"]
    capster = temp.get("capster", "-")
    nama_customer = temp.get("nama_customer", "-")
    hp_customer = temp.get("hp_customer", "-")
    no_nota = generate_no_nota(outlet)
    waktu = datetime.now().strftime("%d/%m/%Y %H:%M")

    # Tentukan chat_id untuk kirim gambar
    if hasattr(reply_target, "message"):
        chat_id = reply_target.message.chat_id
        bot = reply_target.get_bot()
    else:
        chat_id = reply_target.chat_id
        bot = reply_target.get_bot()

    try:
        ws = get_pos_sheet()
        for item in keranjang:
            ws.append_row([
                no_nota, waktu, outlet, kasir, capster, nama_customer, hp_customer,
                item["nama"], item["qty"], fmt_rupiah(item["harga"]), fmt_rupiah(item["subtotal"]),
                metode,
                fmt_rupiah(tunai) if tunai > 0 else "-",
                fmt_rupiah(kembalian) if kembalian > 0 else "-",
                "Lunas"
            ])
        user_data_temp.pop(user_id, None)

        # Konfirmasi singkat dulu
        if hasattr(reply_target, "edit_message_text"):
            await reply_target.edit_message_text("✅ Transaksi berhasil! Membuat struk...")
        else:
            await reply_target.reply_text("✅ Transaksi berhasil! Membuat struk...")

        # Generate & kirim struk PNG
        struk_img = generate_struk_image(
            outlet, no_nota, waktu, kasir, capster,
            nama_customer, hp_customer, keranjang,
            grand_total, tunai, kembalian, metode
        )
        await bot.send_photo(
            chat_id=chat_id,
            photo=struk_img,
            caption=f"🧾 Struk {no_nota}\n💰 {fmt_rupiah(grand_total)} — {metode}"
        )

    except Exception as e:
        logger.error(f"[POS] Gagal simpan: {e}", exc_info=True)
        if hasattr(reply_target, "edit_message_text"):
            await reply_target.edit_message_text(f"❌ Gagal: {e}")
        else:
            await reply_target.reply_text(f"❌ Gagal: {e}")
    return ConversationHandler.END

async def cmd_omzet(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat = update.effective_chat
    total, count = get_omzet_hari_ini(chat.title)
    today = datetime.now().strftime("%d/%m/%Y")
    await update.message.reply_text(
        f"📊 Omzet Hari Ini — {today}\n🏪 {chat.title}\n\n💰 {fmt_rupiah(total)}\n🧾 {count} transaksi"
    )

async def cmd_laporan(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat = update.effective_chat
    try:
        ws = get_pos_sheet()
        rows = ws.get_all_values()
        data = [r for r in rows[1:] if r and r[0]]
        total_bulan = 0
        count_bulan = 0
        by_metode = {}
        by_produk = {}
        for row in data:
            if len(row) < 8:
                continue
            try:
                nominal = int(re.sub(r"[^0-9]", "", str(row[10]))) if len(row) > 10 and row[10] else 0
                total_bulan += nominal
                count_bulan += 1
                metode = row[11] if len(row) > 11 else "Lainnya"
                by_metode[metode] = by_metode.get(metode, 0) + nominal
                pname = row[7] if len(row) > 7 else "Lainnya"
                by_produk[pname] = by_produk.get(pname, 0) + nominal
            except Exception:
                continue
        text = f"📊 {get_sheet_name_pos()}\n🏪 {chat.title}\n{'─'*25}\n"
        text += f"💰 {fmt_rupiah(total_bulan)}\n🧾 {count_bulan} transaksi\n"
        if count_bulan > 0:
            text += f"📈 Rata-rata: {fmt_rupiah(total_bulan // count_bulan)}\n"
        text += f"{'─'*25}\n💳 Per Metode:\n"
        for m, v in sorted(by_metode.items(), key=lambda x: -x[1]):
            text += f"  {m}: {fmt_rupiah(v)}\n"
        text += f"{'─'*25}\n🏷️ Top 5 Produk:\n"
        for p, v in sorted(by_produk.items(), key=lambda x: -x[1])[:5]:
            text += f"  {p}: {fmt_rupiah(v)}\n"
        await update.message.reply_text(text)
    except Exception as e:
        await update.message.reply_text(f"❌ Gagal: {e}")

# ─────────────────────────────────────────
# Main
# ─────────────────────────────────────────

def main():
    token = os.environ["TELEGRAM_BOT_TOKEN"]
    app = Application.builder().token(token).build()

    input_handler = ConversationHandler(
        entry_points=[MessageHandler(filters.PHOTO, handle_photo)],
        states={
            KONFIRMASI_TANGGAL: [CallbackQueryHandler(handle_konfirmasi_tanggal, pattern="^tgl_")],
            INPUT_TANGGAL: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_input_tanggal)],
            KONFIRMASI_NOMINAL: [CallbackQueryHandler(handle_konfirmasi_nominal, pattern="^nom_")],
            INPUT_NOMINAL: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_input_nominal)],
            PILIH_KATEGORI: [CallbackQueryHandler(handle_kategori, pattern="^kat_")],
            TULIS_DESKRIPSI: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_deskripsi)],
            INPUT_REKENING: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_input_rekening)],
            INPUT_PENERIMA: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_input_penerima)],
        },
        fallbacks=[CommandHandler("batal", handle_cancel)],
        per_chat=False, per_user=True,
    )

    edit_handler = ConversationHandler(
        entry_points=[CommandHandler("edit", cmd_edit)],
        states={
            EDIT_PILIH_TRANSAKSI: [CallbackQueryHandler(handle_edit_pilih_transaksi, pattern="^edit_")],
            EDIT_PILIH_FIELD: [CallbackQueryHandler(handle_edit_pilih_field, pattern="^editfield_")],
            EDIT_INPUT_NILAI: [
                CallbackQueryHandler(handle_edit_input_nilai, pattern="^kat_"),
                MessageHandler(filters.TEXT & ~filters.COMMAND, handle_edit_input_nilai),
            ],
        },
        fallbacks=[CommandHandler("batal", handle_cancel)],
        per_chat=False, per_user=True,
    )

    delete_handler = ConversationHandler(
        entry_points=[CommandHandler("delete", cmd_delete)],
        states={
            DELETE_PILIH_TRANSAKSI: [CallbackQueryHandler(handle_delete_pilih_transaksi, pattern="^edit_")],
            DELETE_KONFIRMASI: [CallbackQueryHandler(handle_delete_konfirmasi, pattern="^delete_konfirmasi_")],
        },
        fallbacks=[CommandHandler("batal", handle_cancel)],
        per_chat=False, per_user=True,
    )

    reset_handler = ConversationHandler(
        entry_points=[CommandHandler("reset_bulan", cmd_reset_bulan)],
        states={
            RESET_KONFIRMASI: [CallbackQueryHandler(handle_reset_konfirmasi, pattern="^reset_")],
        },
        fallbacks=[CommandHandler("batal", handle_cancel)],
        per_chat=False, per_user=True,
    )

    jual_handler = ConversationHandler(
        entry_points=[CommandHandler("jual", cmd_jual)],
        states={
            POS_PILIH_PRODUK: [
                CallbackQueryHandler(handle_pos_keranjang_action, pattern="^pos_tambah_lagi$|^pos_lanjut_bayar$|^pos_batal$"),
                CallbackQueryHandler(handle_pos_pilih_produk, pattern="^pos_produk_"),
            ],
            POS_INPUT_QTY: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_pos_input_qty)],
            POS_PILIH_CAPSTER: [CallbackQueryHandler(handle_pos_pilih_capster, pattern="^pos_capster_|^pos_batal$")],
            POS_INPUT_CAPSTER: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_pos_input_capster)],
            POS_INPUT_NAMA_CUSTOMER: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_pos_input_nama_customer)],
            POS_INPUT_HP_CUSTOMER: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_pos_input_hp_customer)],
            POS_PILIH_BAYAR: [CallbackQueryHandler(handle_pos_pilih_bayar, pattern="^pos_bayar_|^pos_batal$")],
            POS_INPUT_TUNAI: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_pos_input_tunai)],
        },
        fallbacks=[CommandHandler("batal", handle_cancel)],
        per_chat=False, per_user=True,
    )

    tambah_produk_handler = ConversationHandler(
        entry_points=[CommandHandler("tambah_produk", cmd_tambah_produk)],
        states={
            POS_SETUP_NAMA: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_pos_setup_nama)],
            POS_SETUP_HARGA: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_pos_setup_harga)],
            POS_SETUP_KONFIRMASI: [CallbackQueryHandler(handle_pos_setup_konfirmasi, pattern="^pos_setup_")],
        },
        fallbacks=[CommandHandler("batal", handle_cancel)],
        per_chat=False, per_user=True,
    )

    app.add_handler(input_handler)
    app.add_handler(edit_handler)
    app.add_handler(delete_handler)
    app.add_handler(reset_handler)
    app.add_handler(jual_handler)
    app.add_handler(tambah_produk_handler)
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("cek", cmd_cek))
    app.add_handler(CommandHandler("total", cmd_total))
    app.add_handler(CommandHandler("produk", cmd_produk))
    app.add_handler(CommandHandler("omzet", cmd_omzet))
    app.add_handler(CommandHandler("laporan", cmd_laporan))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    logger.info("Bot started!")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
