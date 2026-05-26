# 🤖 Expense Bot - Telegram → Google Sheets

Bot Telegram yang otomatis baca struk dari foto dan catat ke Google Sheets menggunakan Claude AI.

## Cara Deploy ke Railway

### 1. Upload ke GitHub
1. Buat repo baru di github.com
2. Upload semua file ini (bot.py, requirements.txt, railway.toml)

### 2. Deploy di Railway
1. Buka railway.app → New Project → Deploy from GitHub repo
2. Pilih repo yang baru dibuat
3. Klik "Add Variables" dan isi environment variables berikut:

### 3. Environment Variables (wajib diisi di Railway)

| Variable | Nilai |
|----------|-------|
| `TELEGRAM_BOT_TOKEN` | Token dari BotFather |
| `ANTHROPIC_API_KEY` | sk-ant-xxx... |
| `SPREADSHEET_ID` | 1V2SszCht1qUZ9PWOdh6PiheLzw5bchSDtp3QolGdNCE |
| `GOOGLE_CREDENTIALS_JSON` | Isi seluruh isi file JSON service account (copy-paste) |

### 4. Deploy
Setelah variables diisi, Railway otomatis deploy. Bot langsung jalan!

## Cara Pakai Bot
- Kirim foto struk/bon ke bot
- Bot otomatis analisa dan catat ke Google Sheets
- Ketik /cek untuk lihat 5 expense terakhir
