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

### 2. Pasang tombol (sekali saja)

1. Tampilkan bookmark bar Chrome: **Cmd+Shift+B**
2. Dari worksheet, **seret** tombol **📦 Panen** dan **📋 Salin hasil** ke bookmark bar

Kalau skripnya diperbarui, hapus tombol lama lalu seret ulang dari worksheet terbaru.

### 3. Tiap halaman (59×: 7 rute domestik × 5 Jumat + 8 rute internasional × 3 Jumat)

Klik link → tunggu daftar penerbangan muncul → klik **📦 Panen** →
tunggu kotak **hijau ✅** di pojok kanan bawah → **Cmd+W**

- Kotak hijau `✔ sudah sampai dasar daftar` → aman
- Kotak **merah** → baca pesannya, biasanya cukup klik **📦 Panen** lagi
- `📦 Total terkumpul` → harus terus naik

Tip: **Cmd+klik** beberapa link sekaligus, lalu kerjakan tab per tab.

### 4. Setelah halaman terakhir

1. Di tab tiket.com mana pun, klik **📋 Salin hasil**
2. Buka `kumpulan.txt` di TextEdit → **Cmd+V** → **Cmd+S**

(Cadangan tanpa tombol: Cmd+Option+J → tempel isi `panen_browser.js` → Enter;
di akhir ketik `copy(HASIL)`.)

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
