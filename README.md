# xscrape

Scraper Twitter/X via Microsoft Edge CDP. Output langsung ke Excel (`.xlsx`) dengan 186 kolom standar dataset Apify.

## Fitur Utama

- **CDP Session:** Memakai Microsoft Edge profil lokal (`.edge-profile/`), login cukup sekali.
- **Safety Limit:** Dibatasi maksimal 500 tweet per proses untuk keamanan akun.
- **End-of-Timeline Fallback:** Berhenti otomatis jika tweet pada rentang waktu sudah habis (tanpa error).
- **Anti-Duplikasi Real-time:** Otomatis mendeteksi tweet yang sudah pernah di-scrape dari proses sebelumnya agar tidak tersimpan ganda.
- **Folder Output Rapi:** Semua file Excel disimpan ke folder `output/`.
- **Merge & Deduplikasi:** Tersedia tool untuk menggabungkan banyak file harian menjadi 1 file Excel master unik.

## Requirements

- Node.js 18+
- Microsoft Edge

## Setup

```bash
npm install
```

## Cara Pakai

### 1. Buka Browser & Login

Jalankan sekali untuk membuka Edge dan login ke akun X:

```bash
npm run launch
# atau: node launch.js
```

Sesi login tersimpan di `.edge-profile/`. Untuk memeriksa apakah status login sudah aktif:

```bash
npm run status
# atau: node remote.js
```

### 2. Jalankan Scraper

```bash
# Basic (query + limit, maksimal 500)
node scrape.js "prabowo" 100

# Multi-kata biasa
node scrape.js "islam ala prabowo" 300

# Exact phrase (frasa utuh)
node scrape.js "islam ala prabowo" 300 --exact

# Dengan filter rentang tanggal
node scrape.js "islam ala prabowo" 300 --exact --since 2026-09-01 --until 2026-09-02
```

> **Catatan Tanggal:** Di Twitter, `until:` bersifat *eksklusif* (sampai jam 00:00 UTC hari tersebut). Untuk scrape tepat tanggal 1 September, gunakan `--since 2026-09-01 --until 2026-09-02`.

File hasil scrape otomatis tersimpan di folder `output/`:
```text
output/tweets_islam_ala_prabowo_2026-09-01_80_2026-09-15.xlsx
```

### 3. Gabungkan File Hasil Scrape (Opsional)

Jika Anda melakukan scraping harian secara bertahap dan ingin menggabungkan semua file di folder `output/` menjadi 1 file Excel bersih tanpa duplikat:

```bash
npm run merge
# atau: node merge.js
```
File hasil penggabungan akan otomatis disimpan di folder `output/`:
```text
output/merged_unique_<total_tweet>_<timestamp>.xlsx
```
