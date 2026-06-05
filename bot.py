# ═══════════════════════════════════════════════════════
# PATCH 3A: Fix fallback bahaya di get_pos_spreadsheet_id()
# Ganti fungsi yang lama dengan ini
# ═══════════════════════════════════════════════════════

def get_pos_spreadsheet_id(chat_title=None):
    """Ambil Spreadsheet ID berdasar nama group Telegram (multi-outlet)"""
    if chat_title:
        key = re.sub(r'[^A-Z0-9]', '_', chat_title.upper().strip())
        key = re.sub(r'_+', '_', key).strip('_')
        env_key = f"OUTLET_{key}"
        sid = os.environ.get(env_key)
        if sid:
            logger.info(f"[OUTLET] Match: {chat_title} -> {env_key}")
            return sid
        logger.warning(f"[OUTLET] Tidak ada match untuk '{chat_title}' (key: {env_key})")

    # Fallback ke SPREADSHEET_ID_POS — JANGAN fallback ke KAS BESAR!
    sid = os.environ.get("SPREADSHEET_ID_POS")
    if sid:
        logger.info(f"[OUTLET] Pakai SPREADSHEET_ID_POS sebagai default")
        return sid

    # Tidak ada fallback ke KAS BESAR — raise error supaya ketahuan
    raise ValueError(
        f"[OUTLET] Spreadsheet ID tidak ditemukan untuk outlet '{chat_title}'. "
        f"Tambahkan env var OUTLET_{re.sub(r'[^A-Z0-9]', '_', (chat_title or '').upper()).strip('_')} "
        f"atau SPREADSHEET_ID_POS di Railway."
    )


# ═══════════════════════════════════════════════════════
# PATCH 3B: processed_hashes persistent ke file
# Tambahkan di bagian atas file, setelah semua import
# ═══════════════════════════════════════════════════════

import json
import os

HASHES_FILE = "processed_hashes.json"

def load_processed_hashes():
    """Load processed hashes dari file saat bot start."""
    try:
        if os.path.exists(HASHES_FILE):
            with open(HASHES_FILE, 'r') as f:
                data = json.load(f)
                hashes = set(data.get('hashes', []))
                logger.info(f"[HASH] Loaded {len(hashes)} processed hashes dari file")
                return hashes
    except Exception as e:
        logger.error(f"[HASH] Gagal load hashes: {e}")
    return set()

def save_processed_hashes(hashes):
    """Simpan processed hashes ke file."""
    try:
        with open(HASHES_FILE, 'w') as f:
            json.dump({'hashes': list(hashes)}, f)
    except Exception as e:
        logger.error(f"[HASH] Gagal save hashes: {e}")

# ── Ganti inisialisasi processed_hashes ──────────────────
# SEBELUM: processed_hashes = set()
# SESUDAH:
processed_hashes = load_processed_hashes()


# ═══════════════════════════════════════════════════════
# PATCH 3C: Update handle_photo() — simpan hash ke file
# Cari baris: processed_hashes.add(foto_hash)
# Tambahkan baris berikutnya:
# ═══════════════════════════════════════════════════════

# SEBELUM:
#   processed_hashes.add(foto_hash)
#   user_data_temp.pop(user_id, None)

# SESUDAH:
#   processed_hashes.add(foto_hash)
#   save_processed_hashes(processed_hashes)  # ← tambahkan ini
#   user_data_temp.pop(user_id, None)
