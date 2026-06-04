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

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets"
]

BULAN = ["Januari","Februari","Maret","April","Mei","Juni",
         "Juli","Agustus","September","Oktober","November","Desember"]

KONFIRMASI_TANGGAL, INPUT_TANGGAL, KONFIRMASI_NOMINAL, INPUT_NOMINAL, PILIH_KATEGORI, TULIS_DESKRIPSI, INPUT_REKENING, INPUT_PENERIMA = range(8)
EDIT_PILIH_TRANSAKSI, EDIT_PILIH_FIELD, EDIT_INPUT_NILAI = range(8, 11)
DELETE_PILIH_TRANSAKSI, DELETE_KONFIRMASI = range(11, 13)

# Admin yang boleh delete transaksi
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
    """Format angka ke Rp 62.500 (format Indonesia)"""
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
# ImgBB Upload
# ─────────────────────────────────────────

def upload_foto_to_drive(image_bytes, filename):
    """Upload foto ke ImgBB, return link. Return '' jika gagal."""
    try:
        api_key = os.environ.get("IMGBB_API_KEY", "").strip()
        if not api_key:
            logger.error("[IMGBB] IMGBB_API_KEY kosong!")
            return ""

        logger.info(f"[IMGBB] Mulai upload: {filename}")

        b64 = base64.b64encode(image_bytes).decode("utf-8")
        data = urllib.parse.urlencode({
            "key": api_key,
            "image": b64,
            "name": filename,
        }).encode("utf-8")

        req = urllib.request.Request(
            "https://api.imgbb.com/1/upload",
            data=data,
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=30) as resp:
            result = json.loads(resp.read().decode("utf-8"))

        if result.get("success"):
            link = result["data"]["url"]  # direct image link
            logger.info(f"[IMGBB] Upload berhasil: {link}")
            return link
        else:
            logger.error(f"[IMGBB] Upload gagal: {result}")
            return ""

    except Exception as e:
        logger.error(f"[IMGBB] Upload error - {type(e).__name__}: {e}", exc_info=True)
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
        if group == "kecil":
            ws.append_row([
                "Tanggal", "Admin", "Deskripsi", "Kategori",
                "Debet (Keluar)", "Kredit (Masuk)", "Saldo", "Keterangan",
                "Tanggal Invoice", "Rekening Tujuan", "Nama Penerima", "Status", "Link Foto"
            ])
        else:
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
        tgl,                    # A - Tanggal catat
        dicatat_oleh,           # B - Admin
        data["deskripsi"],      # C - Deskripsi
        data["kategori"],       # D - Kategori
        fmt_rupiah(debet),      # E - Debet (Keluar)
        fmt_rupiah(kredit),     # F - Kredit (Masuk)
        fmt_rupiah(saldo_baru), # G - Saldo
        data["vendor"],         # H - Keterangan/Vendor
        data["tanggal"],        # I - Tanggal Invoice
        data.get("rekening", ""),  # J - Rekening
        data.get("penerima", ""),  # K - Penerima
        "Bot - " + datetime.now().strftime("%d/%m/%Y %H:%M"),  # L - Status
        foto_link               # M - Link Foto
    ]
    logger.info(f"[SHEET] Append kas besar, foto_link: '{foto_link}'")
    ws.append_row(row)

def recalculate_saldo(ws):
    """Hitung ulang semua saldo di sheet dari atas ke bawah"""
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
            updates.append({
                "range": f"G{i}",
                "values": [[fmt_rupiah(saldo)]]
            })
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
        tgl,                    # A - Tanggal catat
        dicatat_oleh,           # B - Admin
        data["deskripsi"],      # C - Deskripsi
        data["kategori"],       # D - Kategori
        fmt_rupiah(debet),      # E - Debet
        fmt_rupiah(kredit),     # F - Kredit
        fmt_rupiah(saldo_baru), # G - Saldo
        data["vendor"],         # H - Keterangan/Vendor
        data["tanggal"],        # I - Tanggal Invoice
        data.get("rekening", ""),  # J - Rekening
        data.get("penerima", ""),  # K - Penerima
        "Bot - " + datetime.now().strftime("%d/%m/%Y %H:%M"),  # L - Status
        foto_link               # M - Link Foto
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

def analyze_image(image_bytes):
    try:
        client = anthropic.Anthropic(
            api_key=os.environ["ANTHROPIC_API_KEY"],
            timeout=30.0  # timeout 30 detik
        )
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
        logger.error("[CLAUDE] Timeout — API tidak respon dalam 30 detik")
        raise Exception("Claude AI timeout. Coba kirim foto ulang.")
    except anthropic.APIStatusError as e:
        logger.error(f"[CLAUDE] API error: {e.status_code} — {e.message}")
        raise Exception(f"Claude AI error ({e.status_code}). Coba lagi.")
    except Exception as e:
        logger.error(f"[CLAUDE] Error tidak terduga: {e}", exc_info=True)
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
        if group == "besar":
            data_rows.append({
                "row_idx": i,
                "tanggal": row[0] if len(row) > 0 else "",
                "vendor": row[7] if len(row) > 7 else "",
                "nominal": row[4] if len(row) > 4 else "0",
                "kategori": row[3] if len(row) > 3 else "",
                "deskripsi": row[2] if len(row) > 2 else "",
                "status": row[11] if len(row) > 11 else "",
            })
        else:
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
    label = "Kas Besar" if group == "besar" else "Kas Kecil"
    sheet_name = get_sheet_name_besar() if group == "besar" else get_sheet_name_kecil()
    await update.message.reply_text(
        f"Halo! Saya bot pencatat {label}\n"
        f"Sheet aktif: {sheet_name}\n\n"
        f"Perintah:\n"
        f"/cek - Semua transaksi bulan ini\n"
        f"/total - Total bulan ini\n"
        f"/edit - Edit transaksi\n"
        f"/delete - Hapus transaksi (admin)\n"
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

        # Ambil 10 transaksi terbaru saja (hindari pesan terlalu panjang)
        last10 = data_rows[-10:]
        total_semua = len(data_rows)
        text = f"📋 10 Transaksi Terbaru — {label}\n"
        text += f"(Total bulan ini: {total_semua} transaksi)\n\n"

        for row in reversed(last10):
            try:
                if group == "besar":
                    nominal = int(float(re.sub(r'[^0-9]', '', str(row[4] or 0))))
                    deskripsi = (row[2] or '-')[:25]
                    text += f"• {row[0]} | {row[3]}\n  {deskripsi}\n  {fmt_rupiah(nominal)}\n\n"
                else:
                    debet = int(float(re.sub(r'[^0-9]', '', str(row[4] or 0))))
                    deskripsi = (row[2] or '-')[:25]
                    text += f"• {row[0]} | {row[3]}\n  {deskripsi}\n  {fmt_rupiah(debet)}\n\n"
            except Exception:
                text += f"• {row[0]}\n\n"

        # Potong jika melebihi limit Telegram 4096 karakter
        if len(text) > 3800:
            text = text[:3800] + "\n\n... (terpotong, lihat dashboard kasbot.id)"

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
                if group == "besar":
                    total += int(float(re.sub(r'[^0-9.]', '', str(row[4] or 0))))
                else:
                    total += int(float(re.sub(r'[^0-9.]', '', str(row[4] or 0))))
                count += 1
            except Exception:
                continue
        sheet_name = get_sheet_name_besar() if group == "besar" else get_sheet_name_kecil()
        await update.message.reply_text(
            f"Total {label}\n"
            f"Sheet: {sheet_name}\n\n"
            f"Total: {fmt_rupiah(total)}\n"
            f"Transaksi: {count}"
        )
    except Exception as e:
        await update.message.reply_text(f"Gagal mengambil data: {e}")

# ─────────────────────────────────────────
# Foto Handler (ConversationHandler)
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

        # Analisa dengan Claude
        data = analyze_image(image_bytes)

        # Upload ke Drive
        filename = f"struk_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{foto_hash[:8]}.jpg"
        foto_link = upload_foto_to_drive(image_bytes, filename)
        logger.info(f"[PHOTO] foto_link setelah upload: '{foto_link}'")

        user_data_temp[user_id] = {
            "data": data,
            "group": group,
            "foto_hash": foto_hash,
            "foto_link": foto_link,
            "dicatat_oleh": update.effective_user.first_name or "Bot",
        }

        tgl_display = format_tanggal_display(data["tanggal"])
        foto_status = "✅ Foto terupload ke Drive" if foto_link else "⚠️ Foto gagal upload ke Drive"

        await msg.edit_text(
            f"📋 Struk terdeteksi!\n\n"
            f"🏪 Vendor: {data['vendor']}\n"
            f"{foto_status}\n\n"
            f"📅 Konfirmasi tanggal transaksi:\n"
            f"Tanggal: {tgl_display} ({data['tanggal']})\n\n"
            f"Apakah tanggal ini benar?",
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
            f"✅ Tanggal: {data['tanggal']} (dikonfirmasi)\n\n"
            f"💰 Konfirmasi nominal transaksi:\n"
            f"Nominal: {fmt_rupiah(data['jumlah'])}\n\n"
            f"Apakah nominal ini benar?",
            reply_markup=build_konfirmasi_keyboard("nom")
        )
        return KONFIRMASI_NOMINAL

    elif query.data == "tgl_ubah":
        await query.edit_message_text(
            "✏️ Ketik tanggal yang benar:\n"
            "Format: DD/MM/YYYY\n"
            "Contoh: 04/06/2026"
        )
        return INPUT_TANGGAL

async def handle_input_tanggal(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text.strip()

    if user_id not in user_data_temp:
        await update.message.reply_text("⚠️ Session expired. Kirim foto ulang.")
        return ConversationHandler.END

    tanggal = parse_tanggal(text)
    if not re.match(r'^\d{4}-\d{2}-\d{2}$', tanggal):
        await update.message.reply_text(
            "❌ Format tidak valid.\nGunakan DD/MM/YYYY\nContoh: 04/06/2026\n\nCoba lagi:"
        )
        return INPUT_TANGGAL

    user_data_temp[user_id]["data"]["tanggal"] = tanggal
    data = user_data_temp[user_id]["data"]
    await update.message.reply_text(
        f"✅ Tanggal diperbarui: {tanggal}\n\n"
        f"💰 Konfirmasi nominal:\n"
        f"Nominal: {fmt_rupiah(data['jumlah'])}\n\n"
        f"Apakah nominal ini benar?",
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
        await query.edit_message_text("⚠️ Session expired. Kirim foto ulang.")
        return ConversationHandler.END

    data = user_data_temp[user_id]["data"]

    if query.data == "nom_benar":
        await query.edit_message_text(
            f"✅ Tanggal: {data['tanggal']}\n"
            f"✅ Nominal: {fmt_rupiah(data['jumlah'])}\n\n"
            f"📂 Pilih kategori:",
            reply_markup=build_kategori_keyboard()
        )
        return PILIH_KATEGORI

    elif query.data == "nom_ubah":
        await query.edit_message_text(
            "✏️ Ketik nominal yang benar (angka saja):\nContoh: 62500"
        )
        return INPUT_NOMINAL

async def handle_input_nominal(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text.strip()

    if user_id not in user_data_temp:
        await update.message.reply_text("⚠️ Session expired. Kirim foto ulang.")
        return ConversationHandler.END

    jumlah_raw = re.sub(r'[^0-9]', '', text)
    if not jumlah_raw:
        await update.message.reply_text(
            "❌ Format tidak valid. Ketik angka saja.\nContoh: 62500\n\nCoba lagi:"
        )
        return INPUT_NOMINAL

    jumlah = int(jumlah_raw)
    user_data_temp[user_id]["data"]["jumlah"] = jumlah
    data = user_data_temp[user_id]["data"]
    await update.message.reply_text(
        f"✅ Nominal diperbarui: {fmt_rupiah(jumlah)}\n\n"
        f"📂 Pilih kategori:",
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
        await query.edit_message_text("⚠️ Session expired. Kirim foto ulang.")
        return ConversationHandler.END

    kategori = query.data.replace("kat_", "")
    user_data_temp[user_id]["kategori"] = kategori
    data = user_data_temp[user_id]["data"]
    await query.edit_message_text(
        f"✅ Tanggal: {data['tanggal']}\n"
        f"✅ Nominal: {fmt_rupiah(data['jumlah'])}\n"
        f"✅ Kategori: {kategori}\n\n"
        f"✏️ Tulis deskripsi transaksi:\n"
        f"Contoh: bayar gaji, beli ATK, iuran RT"
    )
    return TULIS_DESKRIPSI

async def handle_deskripsi(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    deskripsi = update.message.text.strip()

    if user_id not in user_data_temp:
        await update.message.reply_text("⚠️ Session expired. Kirim foto ulang.")
        return ConversationHandler.END

    user_data_temp[user_id]["data"]["deskripsi"] = deskripsi

    await update.message.reply_text(
        f"📝 Deskripsi: {deskripsi}\n\n"
        f"🏦 Masukkan nomor rekening tujuan:\n"
        f"Contoh: 1234567890\n"
        f"(Ketik '-' jika tunai/tidak ada rekening)"
    )
    return INPUT_REKENING

async def handle_input_rekening(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    rekening = update.message.text.strip()

    if user_id not in user_data_temp:
        await update.message.reply_text("⚠️ Session expired. Kirim foto ulang.")
        return ConversationHandler.END

    user_data_temp[user_id]["rekening"] = rekening

    await update.message.reply_text(
        f"🏦 Rekening: {rekening}\n\n"
        f"👤 Masukkan nama penerima:\n"
        f"Contoh: PT Sumber Makmur\n"
        f"(Ketik '-' jika tidak ada)"
    )
    return INPUT_PENERIMA

async def handle_input_penerima(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    penerima = update.message.text.strip()

    if user_id not in user_data_temp:
        await update.message.reply_text("⚠️ Session expired. Kirim foto ulang.")
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

    logger.info(f"[PENERIMA] Menyimpan ke sheet, foto_link: '{foto_link}'")

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
        foto_info = f"\n🔗 Link foto: {foto_link}" if foto_link else "\n⚠️ Foto tidak terupload"

        await update.message.reply_text(
            f"✅ Berhasil dicatat ke {label}!\n"
            f"📊 Sheet: {sheet_name}\n\n"
            f"📅 Tanggal: {data['tanggal']}\n"
            f"🏪 Vendor: {data['vendor']}\n"
            f"📝 Deskripsi: {data['deskripsi']}\n"
            f"📂 Kategori: {data['kategori']}\n"
            f"💰 Jumlah: {fmt_rupiah(data['jumlah'])}\n"
            f"💳 Metode: {data['metode']}\n"
            f"🏦 Rekening: {data['rekening']}\n"
            f"👤 Penerima: {data['penerima']}"
            f"{foto_info}"
        )

    except Exception as e:
        logger.error(f"[PENERIMA] Error simpan sheet: {e}", exc_info=True)
        await update.message.reply_text(f"❌ Gagal menyimpan ke sheet.\nError: {e}")

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
            await update.message.reply_text("Belum ada transaksi yang bisa diedit.")
            return ConversationHandler.END

        user_data_temp[user_id] = {
            "group": group,
            "edit_mode": True,
            "transactions": transactions,
            "edit_page": 0,
        }
        keyboard, total_pages = build_transaksi_keyboard(transactions, 0, group)
        await update.message.reply_text(
            f"Pilih transaksi yang ingin diedit:\n"
            f"Halaman 1/{total_pages} ({len(transactions)} transaksi)",
            reply_markup=keyboard
        )
        return EDIT_PILIH_TRANSAKSI

    except Exception as e:
        logger.error(f"[EDIT] cmd_edit error: {e}")
        await update.message.reply_text(f"❌ Gagal memuat transaksi: {e}")
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
        user_data_temp[user_id]["edit_page"] = page
        transactions = user_data_temp[user_id]["transactions"]
        group = user_data_temp[user_id]["group"]
        keyboard, total_pages = build_transaksi_keyboard(transactions, page, group)
        await query.edit_message_text(
            f"Pilih transaksi yang ingin diedit:\n"
            f"Halaman {page + 1}/{total_pages} ({len(transactions)} transaksi)",
            reply_markup=keyboard
        )
        return EDIT_PILIH_TRANSAKSI

    row_idx = int(query.data.replace("edit_trx_", ""))
    transactions = user_data_temp[user_id]["transactions"]
    trx = next((t for t in transactions if t["row_idx"] == row_idx), None)
    if not trx:
        await query.edit_message_text("❌ Transaksi tidak ditemukan.")
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
        [InlineKeyboardButton("⛔ Saldo (otomatis, tidak bisa diedit)", callback_data="editfield_saldo")],
        [InlineKeyboardButton("❌ Batalkan", callback_data="editfield_batal")],
    ])

    await query.edit_message_text(
        f"Transaksi dipilih:\n\n"
        f"📅 Tanggal: {trx['tanggal']}\n"
        f"🏪 Vendor: {trx['vendor']}\n"
        f"📝 Deskripsi: {trx['deskripsi']}\n"
        f"📂 Kategori: {trx['kategori']}\n"
        f"💰 Nominal: {fmt_rupiah(nominal)}\n\n"
        f"Field mana yang ingin diedit?",
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

    # Saldo tidak boleh diedit manual
    if field == "saldo":
        await query.edit_message_text(
            "⛔ Kolom Saldo tidak bisa diedit manual!\n\n"
            "Saldo dihitung otomatis oleh bot berdasarkan Debet & Kredit.\n"
            "Jika ingin koreksi saldo, gunakan cara berikut:\n\n"
            "1. Edit nominal transaksi yang salah via tombol 💰 Nominal\n"
            "2. Atau tambah transaksi koreksi baru lewat bot\n\n"
            "Kirim /edit untuk kembali ke menu edit."
        )
        return ConversationHandler.END

    if field == "kategori":
        await query.edit_message_text("📂 Pilih kategori baru:", reply_markup=build_kategori_keyboard())
        return EDIT_INPUT_NILAI

    prompts = {
        "deskripsi": "✏️ Ketik deskripsi baru:\nContoh: bayar gaji bulan Juni",
        "nominal": "💰 Ketik nominal baru (angka saja):\nContoh: 150000",
        "tanggal": "📅 Ketik tanggal baru:\nFormat: DD/MM/YYYY\nContoh: 04/06/2026",
        "rekening": "🏦 Ketik nomor rekening baru:\nContoh: 1234567890\n(Ketik '-' jika tunai)",
        "penerima": "👤 Ketik nama penerima baru:\nContoh: PT Sumber Makmur\n(Ketik '-' jika tidak ada)",
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
        nilai_baru_clean = re.sub(r'[^0-9]', '', nilai_baru)
        if not nilai_baru_clean:
            await reply("❌ Format tidak valid. Ketik angka saja.\nCoba lagi:")
            return EDIT_INPUT_NILAI
        # Simpan dengan format Rp supaya konsisten
        nilai_baru = fmt_rupiah(int(nilai_baru_clean))

    elif field == "tanggal":
        nilai_baru = parse_tanggal(nilai_baru)
        if not re.match(r'^\d{4}-\d{2}-\d{2}$', nilai_baru):
            await reply("❌ Format tidak valid.\nGunakan DD/MM/YYYY\nCoba lagi:")
            return EDIT_INPUT_NILAI

    try:
        ws = get_sheet(group)
        edit_label = f"EDITED - {datetime.now().strftime('%d/%m/%Y, %H.%M.%S')} by {editor}"

        if group == "besar":
            field_col_map = {"deskripsi": 3, "kategori": 4, "nominal": 5, "tanggal": 9, "rekening": 10, "penerima": 11}
            status_col = 12
        else:
            field_col_map = {"deskripsi": 3, "kategori": 4, "nominal": 5, "tanggal": 9, "rekening": 10, "penerima": 11}
            status_col = 12

        col_idx = field_col_map.get(field)
        if col_idx:
            ws.update_cell(row_idx, col_idx, nilai_baru)
        ws.update_cell(row_idx, status_col, edit_label)

        # Recalculate saldo semua baris setelah edit nominal
        if field == "nominal":
            recalculate_saldo(ws)

        user_data_temp.pop(user_id, None)
        await reply(
            f"✅ Berhasil diedit!\n\n"
            f"Field: {field.capitalize()}\n"
            f"Nilai baru: {nilai_baru}\n"
            f"Diedit oleh: {editor}\n"
            f"Waktu: {datetime.now().strftime('%d/%m/%Y %H:%M')}\n\n"
            f"Kolom Status & Saldo di sheet sudah diperbarui."
        )

    except Exception as e:
        logger.error(f"[EDIT] handle_edit_input_nilai error: {e}", exc_info=True)
        await reply(f"❌ Gagal mengedit: {e}")

    return ConversationHandler.END

# ─────────────────────────────────────────
# Delete Handlers
# ─────────────────────────────────────────

def catat_log_delete(group, trx, deleted_by):
    """Catat transaksi yang didelete ke sheet Log Delete"""
    try:
        gc = get_gspread_client()
        if group == "kecil":
            sid = os.environ["SPREADSHEET_ID_KAS_KECIL"]
        else:
            sid = os.environ["SPREADSHEET_ID_KAS_BESAR"]

        sh = gc.open_by_key(sid)
        try:
            ws_log = sh.worksheet("Log Delete")
        except gspread.WorksheetNotFound:
            ws_log = sh.add_worksheet(title="Log Delete", rows=1000, cols=10)
            ws_log.append_row([
                "Waktu Delete", "Dihapus Oleh", "Group",
                "Tanggal Transaksi", "Deskripsi", "Kategori",
                "Nominal", "Vendor", "Status Lama"
            ])
            ws_log.format("A1:I1", {
                "backgroundColor": {"red": 0.8, "green": 0.1, "blue": 0.1},
                "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}}
            })

        ws_log.append_row([
            datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            deleted_by,
            "Kas Besar" if group == "besar" else "Kas Kecil",
            trx.get("tanggal", ""),
            trx.get("deskripsi", ""),
            trx.get("kategori", ""),
            trx.get("nominal", ""),
            trx.get("vendor", ""),
            trx.get("status", ""),
        ])
        logger.info(f"[DELETE] Log delete berhasil dicatat")
    except Exception as e:
        logger.error(f"[DELETE] Gagal catat log: {e}", exc_info=True)

async def cmd_delete(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id

    # Cek apakah user adalah admin
    if user_id not in ADMIN_IDS:
        await update.message.reply_text(
            "⛔ Kamu tidak punya akses untuk menghapus transaksi.\n"
            "Fitur ini hanya untuk admin."
        )
        return ConversationHandler.END

    chat = update.effective_chat
    group = get_group_type(chat.title or "")

    try:
        ws = get_sheet(group)
        transactions = get_all_transactions(ws, group)
        if not transactions:
            await update.message.reply_text("Belum ada transaksi yang bisa dihapus.")
            return ConversationHandler.END

        user_data_temp[user_id] = {
            "group": group,
            "delete_mode": True,
            "transactions": transactions,
            "delete_page": 0,
        }
        keyboard, total_pages = build_transaksi_keyboard(transactions, 0, group)
        await update.message.reply_text(
            f"🗑️ Pilih transaksi yang ingin dihapus:\n"
            f"Halaman 1/{total_pages} ({len(transactions)} transaksi)\n\n"
            f"⚠️ Transaksi yang dihapus akan dicatat di Log Delete.",
            reply_markup=keyboard
        )
        return DELETE_PILIH_TRANSAKSI

    except Exception as e:
        logger.error(f"[DELETE] cmd_delete error: {e}")
        await update.message.reply_text(f"❌ Gagal memuat transaksi: {e}")
        return ConversationHandler.END

async def handle_delete_pilih_transaksi(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id

    if query.data == "edit_batal":
        user_data_temp.pop(user_id, None)
        await query.edit_message_text("❌ Hapus dibatalkan.")
        return ConversationHandler.END

    if user_id not in user_data_temp:
        await query.edit_message_text("⚠️ Session expired.")
        return ConversationHandler.END

    if query.data.startswith("edit_page_"):
        page = int(query.data.replace("edit_page_", ""))
        user_data_temp[user_id]["delete_page"] = page
        transactions = user_data_temp[user_id]["transactions"]
        group = user_data_temp[user_id]["group"]
        keyboard, total_pages = build_transaksi_keyboard(transactions, page, group)
        await query.edit_message_text(
            f"🗑️ Pilih transaksi yang ingin dihapus:\n"
            f"Halaman {page + 1}/{total_pages} ({len(transactions)} transaksi)",
            reply_markup=keyboard
        )
        return DELETE_PILIH_TRANSAKSI

    row_idx = int(query.data.replace("edit_trx_", ""))
    transactions = user_data_temp[user_id]["transactions"]
    trx = next((t for t in transactions if t["row_idx"] == row_idx), None)
    if not trx:
        await query.edit_message_text("❌ Transaksi tidak ditemukan.")
        return ConversationHandler.END

    user_data_temp[user_id]["delete_row_idx"] = row_idx
    user_data_temp[user_id]["delete_trx"] = trx

    try:
        nominal = int(float(re.sub(r'[^0-9.]', '', str(trx["nominal"] or 0))))
    except Exception:
        nominal = 0

    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("🗑️ Ya, Hapus Transaksi Ini", callback_data="delete_konfirmasi_ya")],
        [InlineKeyboardButton("❌ Batal", callback_data="delete_konfirmasi_batal")],
    ])

    await query.edit_message_text(
        f"⚠️ Konfirmasi Hapus Transaksi\n\n"
        f"📅 Tanggal: {trx['tanggal']}\n"
        f"📝 Deskripsi: {trx['deskripsi']}\n"
        f"📂 Kategori: {trx['kategori']}\n"
        f"💰 Nominal: {fmt_rupiah(nominal)}\n\n"
        f"Transaksi ini akan dihapus dari sheet dan dicatat di Log Delete.\n"
        f"Yakin ingin menghapus?",
        reply_markup=keyboard
    )
    return DELETE_KONFIRMASI

async def handle_delete_konfirmasi(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id

    if query.data == "delete_konfirmasi_batal":
        user_data_temp.pop(user_id, None)
        await query.edit_message_text("❌ Hapus dibatalkan.")
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

        # Catat ke log delete dulu sebelum hapus
        catat_log_delete(group, trx, deleted_by)

        # Hapus baris dari sheet
        ws.delete_rows(row_idx)

        # Recalculate saldo semua baris setelah delete
        recalculate_saldo(ws)

        user_data_temp.pop(user_id, None)
        await query.edit_message_text(
            f"✅ Transaksi berhasil dihapus!\n\n"
            f"📝 Deskripsi: {trx['deskripsi']}\n"
            f"📂 Kategori: {trx['kategori']}\n"
            f"🕐 Dihapus oleh: {deleted_by}\n"
            f"📋 Log tersimpan di sheet 'Log Delete'\n"
            f"📊 Saldo semua baris sudah diupdate otomatis."
        )

    except Exception as e:
        logger.error(f"[DELETE] Gagal hapus: {e}", exc_info=True)
        await query.edit_message_text(f"❌ Gagal menghapus transaksi: {e}")

    return ConversationHandler.END

async def handle_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_data_temp.pop(user_id, None)
    await update.message.reply_text("❌ Proses dibatalkan.")
    return ConversationHandler.END

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Kirim foto struk untuk mencatat pengeluaran.\n"
        "Perintah: /cek /total /edit /delete /start /batal"
    )

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
        per_chat=False,
        per_user=True,
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
        per_chat=False,
        per_user=True,
    )

    delete_handler = ConversationHandler(
        entry_points=[CommandHandler("delete", cmd_delete)],
        states={
            DELETE_PILIH_TRANSAKSI: [CallbackQueryHandler(handle_delete_pilih_transaksi, pattern="^edit_")],
            DELETE_KONFIRMASI: [CallbackQueryHandler(handle_delete_konfirmasi, pattern="^delete_konfirmasi_")],
        },
        fallbacks=[CommandHandler("batal", handle_cancel)],
        per_chat=False,
        per_user=True,
    )

    app.add_handler(input_handler)
    app.add_handler(edit_handler)
    app.add_handler(delete_handler)
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("cek", cmd_cek))
    app.add_handler(CommandHandler("total", cmd_total))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    logger.info("Bot started!")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
