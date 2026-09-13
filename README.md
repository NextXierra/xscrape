# xscrape

Scraper Twitter/X via Microsoft Edge CDP. Output langsung ke Excel (.xlsx) dengan 186 kolom standar dataset Apify.

## Requirements

- Node.js 18+
- Microsoft Edge

## Setup

```bash
npm install
```

## Cara Pakai

### 1. Buka Browser & Login

Jalankan sekali untuk buka Edge dan login akun X:

```bash
node launch.js
```

Sesi login tersimpan di `.edge-profile/`, jadi cukup login sekali.

### 2. Jalankan Scraper

```bash
# Basic (query + limit)
node scrape.js "prabowo" 100

# Multi-kata
node scrape.js "islam ala prabowo" 100

# Exact phrase
node scrape.js "\"islam ala prabowo\"" 100

# Range tanggal
node scrape.js "prabowo" 100 --since 2024-01-01 --until 2024-02-01
```

File hasil scrape (`.xlsx`) otomatis disimpan di direktori project.
