# Panduan mingguan — ambil harga tiket

> Scraping otomatis tidak lagi bisa: tiket.com memasang proteksi bot Cloudflare
> (403) di endpoint pencarian sejak September 2026. Data sekarang diambil
> semi-manual — Anda yang membuka halaman, skrip yang membaca isinya.

## Tiap Senin (boleh geser 1–2 hari)

### 1. Buka worksheet

```bash
cd "/Users/haris/Library/Mobile Documents/com~apple~CloudDocs/Haris Eko Faruddin/claude/tiketcom-pesawat" && .venv/bin/python input_manual.py --html
```

Otomatis: tanggal disetel ke 5 Jumat terdekat, `kumpulan.txt` dikosongkan
(cadangan disimpan ke `kumpulan_lalu.txt`), worksheet terbuka di browser.

### 2. Halaman pertama

1. Klik tombol **Salin skrip panen** di worksheet
2. Klik link nomor **1**
3. **Cmd+Option+J** → klik baris `>` → **Cmd+V** → **Enter**

Kalau Chrome minta, ketik `allow pasting` lalu Enter, baru tempel lagi.

### 3. Halaman 2–35

Klik link → **Cmd+Option+J** → **↑** → **Enter** → tunggu `✅ SELESAI` → **Cmd+W**

Yang dilirik tiap halaman:
- `✔ sudah mencapai dasar daftar` → aman
- `⚠ BELUM sampai dasar` → ulangi halaman itu (↑ lalu Enter)
- `📦 Total terkumpul` → harus terus naik

Tip: **Cmd+klik** beberapa link sekaligus, lalu kerjakan tab per tab.

### 4. Halaman terakhir

1. Ketik `copy(HASIL)` → **Enter**
2. Buka `kumpulan.txt` di TextEdit → **Cmd+V** → **Cmd+S**

### 5. Serahkan ke Claude

Bilang **"sudah tersimpan"**. Claude akan memproses jadi Excel,
memperbarui rata-rata bulanan + Rp/km, lalu commit & push.

Atau jalankan sendiri:
```bash
.venv/bin/python input_manual.py --panen kumpulan.txt
.venv/bin/python hitung_rata_rata.py
```

## Catatan

- Tumpukan di browser **direset otomatis** tiap ganti minggu — tidak perlu apa-apa
- Tab yang sudah dijalankan **boleh ditutup**; data tersimpan di localStorage
- **Jangan** pakai Incognito, dan jangan hapus data situs tiket.com sebelum selesai
- Ingin lebih ringan? Ubah jumlah Jumat: `--n 3` (21 halaman) atau `--n 2` (14 halaman)
