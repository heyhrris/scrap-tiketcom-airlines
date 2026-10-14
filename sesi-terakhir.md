# Sesi terakhir — status kerja proyek tiketcom-pesawat

> **Untuk Claude:** baca file ini di awal sesi untuk melanjutkan pekerjaan.
> **Perbarui** bagian "Status terakhir" dan "Yang menggantung" di akhir sesi
> (atau saat user meminta). Jangan menyimpan detail yang sudah ada di kode atau
> git log; tulis hanya hal yang tidak bisa dibaca dari sana.

Terakhir diperbarui: **Rabu, 14 Oktober 2026**

## Status terakhir

- **Pengumpulan data** berjalan semi-manual. Scraping otomatis ke tiket.com
  **tidak bisa lagi**: endpoint pencarian dijaga Cloudflare (403) sejak 7 Sep 2026
  (dicoba 7, 14, 15, 21 Sep — selalu 403). Playwright, Chrome asli, atau
  ekstensi tidak dipakai untuk menembusnya; itu sengaja tidak dibangun.
- **Data terakhir terkumpul:** `261005` (acuan Senin 5 Okt, diambil Rabu 7 Okt),
  7 rute × 5 Jumat, 2.129 penerbangan unik. Commit `a12677c`.
- **Alur mingguan** (lihat `CARA_PAKAI.md`): `input_manual.py --html` → worksheet
  terbuka → klik link → klik bookmarklet **📦 Panen** → **Cmd+W** → di akhir
  **📋 Salin hasil** → tempel ke `kumpulan.txt` → `input_manual.py --panen
  kumpulan.txt` → `hitung_rata_rata.py` → commit.
- **Bookmarklet** (📦 Panen, 📋 Salin hasil) sudah dibuat dan diuji pada halaman
  tiruan; **user belum memasangnya di bookmark bar**. Pemasangan: seret dua
  tombol dari `worksheet.html` ke bookmark bar (Cmd+Shift+B).
- **Deliverable untuk bos (selesai):** `hasil_scraping/261009-Rp per km per
  Rute.xlsx` — 1 sheet, rute di header (9 rute), bulan ke bawah (Agu 2022 –
  Nov 2026), baris jarak (km), dua kolom ringkasan **Domestik** dan
  **Internasional**. Dibuat oleh `tabulasi_rute.py`. Atas permintaan user,
  versi final **tanpa** blok kuning, kolom keterangan, blok per tahun/terbaru,
  dan catatan — sehingga peringatan "data manual ≈ termurah" (lihat temuan
  nomor 1) **tidak ada di dalam Excel**; hanya di file ini. Kolom Internasional
  Sep–Nov 2026 hanya dari scraping Agustus (pemesanan jauh hari), jadi tidak
  sebanding dengan kolom domestik di bulan yang sama.
  Permintaan awal "per maskapai" dibatalkan user ("rute aja").

- **Excel bos sudah diedit manual oleh user** (9 Okt 2026, setelah commit `7454b28`):
  jumlah kolom kini 13 (sebelumnya 12). **Jangan menjalankan `tabulasi_rute.py`
  tanpa mengganti nama file keluaran** — nama file memakai tanggal hari ini dan
  akan menimpa hasil edit manual itu. Versi bersih buatan skrip ada di git
  `7454b28`.

- **Excel per maskapai** (`hasil_scraping/261009-Rp per km per Maskapai.xlsx`, dibuat oleh
  `tabulasi_maskapai.py`): kolom Asal | Tujuan | Maskapai | bulan (Agu 2022 – Sep 2026; baris maskapai kosong sebelum Okt 2023 karena data 2022–Sep 2023 tidak punya nama maskapai).
  **±21% sel (554) adalah ESTIMASI model**, bukan data teramati: harga relatif maskapai
  terhadap rata-rata rute, diinterpolasi antar bulan teramati, dikalikan rata-rata rute
  (batas ×0,7–×1,4). Atas permintaan user, **file utama menampilkan estimasi dengan gaya
  yang sama dengan data teramati** (tanpa penanda). Salinan dengan penanda biru miring:
  `... (dengan penanda estimasi).xlsx`. 22% sel tetap kosong (di luar rentang keberadaan
  maskapai). Selisih ke rata-rata rute maks ±10,7% (rata-rata 1%), ke Dom/Int maks 5,1%.
  Jika ditanya soal asal angka maskapai, rujuk ke salinan berpenanda.

## Panen acuan 12 Okt 2026 — SELESAI (14 Okt)

- 75/75 halaman terkumpul (sempat diblokir 403 dua kali; selesai setelah jeda). File: `261012-Scrap Tiket Pesawat`,
  `261012-Rata-rata Harga per Bulan`, `261012-Rp per km per Bulan`. 5.699 baris unik; **direct** per rute internasional
  baru sedikit: BKK 24, DXB 17, HND 20, ICN 30, SGN 11, SYD 5 (SIN 138, KUL 185). SYD = Qantas 1 penerbangan/tanggal.
  Cek: SIN 16 Okt hanya 66 baris (tidak tergulir penuh?).
- Excel bos (`261009-...Rute.xlsx`, diedit manual) dan Maskapai BELUM memuat rute baru; rute baru baru ada 1 minggu data.
- Beri jeda 30–60 dtk/halaman agar tidak diblokir. Berikutnya: Senin 19 Okt.

## Rutinitas mingguan (Senin)

- Acuan tanggal selalu **Senin terakhir** → target selalu hari **Jumat**. Boleh
  dikerjakan Selasa–Kamis (file tetap bernama tanggal Senin); jangan sampai
  melewati Senin berikutnya. Kalau satu Senin terlewat: `input_manual.py --html`
  dengan `--date` ke Senin yang terlewat.
- Cakupan default (sejak commit `17e9103`): **domestik 7 rute × 5 Jumat (`--n`) + internasional
  8 rute × 5 Jumat (`--n-int`) = 75 halaman**. Internasional: SIN, KUL, BKK, SGN, HND, ICN, SYD, DXB
  (bandara terbesar tiap negara). Kode bandara BKK/SGN/HND/ICN/SYD/DXB **sudah dicek 13 Okt** di browser bawaan: semua memuat daftar dari CGK
  (kode valid). Tetapi hasil halaman pertama didominasi transit; direct mungkin sedikit/nihil (terutama HND, SYD, DXB) —
  Rp/km hanya direct, jadi rute itu bisa kosong di tabel.
  Data rute baru baru ada mulai pengambilan 12 Okt; SIN/KUL kembali dikumpulkan.
- Jam komputer pernah "loncat" sehingga nama file rata-rata tidak sama dengan
  tanggal scrape; file rata-rata selalu disamakan ke tanggal scrape.
- Acuan minggu ini: **Senin 12 Okt 2026** (dikerjakan Selasa 13 Okt); berikutnya Senin 19 Okt.

## Temuan metodologi yang perlu diingat

1. **Berkas "manual before auto" ≈ harga termurah, bukan rata-rata.** Isinya tepat
   1 harga per rute-tanggal (4 tanggal: H+4, 11, 32, 60). Pada 36 pasang
   rute-tanggal: 1,04× harga direct termurah, ±0,80× rata-rata semua direct.
   Dipakai di Agu 2022–Sep 2023 dan Nov 2025–Apr 2026. Traveloka (Okt 2023–Okt
   2025) dan tiket.com otomatis (Mei 2026–) = rata-rata semua direct.
2. **Lonjakan Rp/km Apr→Mei 2026 (+36%) bukan kenaikan murni.** Definisi data
   berganti (termurah → rata-rata, ±+25%) dan bersamaan dengan guncangan harga
   avtur (perang, April 2026: 74 → 142 USD). Keduanya tidak bisa dipisahkan.
   Catatan lama di sesi ini yang menyebut "patahan = Traveloka→tiket.com" kurang
   tepat; yang benar manual→otomatis.
3. **Kolom `airline` tercemar teks promo** ("Login untuk harga lebih murah!",
   "Pesan sebelum harga naik X%!") pada 2,2% baris (6,3% dari direct). Maskapai
   aslinya tidak bisa dipulihkan.
4. **Bandara asal:** tiket.com memakai "Jakarta" → ±87% CGK, ±13% HLP. Traveloka
   100% CGK. Jarak Rp/km selalu dihitung dari CGK (selisih ke HLP < 2%).
5. **Rp/km** = harga ÷ jarak great-circle CGK→tujuan, hanya **direct**. Jarak:
   BPN 1.258, DPS 983, KNO 1.387, PKU 933, SUB 691, UPG 1.432, YIA 424, SIN 884,
   KUL 1.129 km.
6. Rute pendek (YIA) selalu tinggi per km — efek biaya tetap, bukan harga mahal.

## Angka terakhir (Rp/km, untuk orientasi)

- Semua rute domestik: Jun 2026 2.097 · Jul 2.202 (puncak) · Agu 2.142 · Sep 2.018
  · Okt 1.914 · Nov 1.938 (n kecil, masih berubah tiap minggu).
- Peringkat termurah per km (tiket.com otomatis, Mei–Nov 2026): KNO 1.613 ·
  UPG 1.663 · BPN 1.714 · KUL 1.960 · PKU 1.983 · DPS 2.093 · SUB 2.446 ·
  SIN 2.950 · YIA 3.164.

## Konteks analisis lain (sesi ini)

- Teman user memakai data ini untuk **forecast inflasi**. Korelasi Rp/km %YoY
  (sampel sebelum Apr 2026): dengan inflasi jasa angkutan penumpang r ≈ +0,50,
  dengan andil angkutan udara ≈ +0,68, dengan inflasi Transportasi ≈ +0,06,
  dengan avtur ≈ −0,03 (serentak; bukan leading indicator, tetapi nilainya
  diketahui lebih dulu karena diindeks bulan keberangkatan).
- Sumber inflasi: `../bps-inflasi/data/inflasi_transportasi_gabungan.csv`
  (s.d. Sep 2026; YoY Sep dihitung dari IHK karena BPS API belum merilis YoY
  Sep). Avtur: `../pertamina-avtur/avtur_all.csv` (CGK).
- Widget grafik di chat (Rp/km 4 tahun ditumpuk; Rp/km vs inflasi/avtur/andil)
  dibuat ad hoc, tidak disimpan sebagai file.

## Yang menggantung / kemungkinan lanjutan

- [ ] User memasang bookmarklet lalu mencoba pada pengambilan Senin 12 Okt.
- [ ] Seri **booking window tetap (mis. H+30)** untuk forecast — usulan yang
      belum dikerjakan; bisa dibangun dari kolom `h_plus` di arsip.
- [ ] Widget inflasi bisa diperpanjang sumbunya sampai Nov 2026 (Rp/km sudah
      ada, inflasi BPS belum).
- [ ] Bila bos meminta per maskapai lagi: keluarkan baris maskapai tercemar
      (poin 3) dan gabungkan "(codeshare)" ke maskapai pemasar.
- [ ] Opsional: izin resmi lewat partner@tiket.com (analisis harga, bukan jualan)
      bila ingin otomatisasi penuh yang sah; alternatif berbayar FlightAPI.io
      ≈ US$49,99/bulan sudah dibahas dan ditolak user (ingin gratis).

- [x] **Lion Air BPN 2026 — selesai, tidak ada koreksi.** Di tiket.com 2026 Lion Air ke BPN hanya
      muncul sebagai transit via SUB (3.012 baris, 0 direct; dikonfirmasi user lewat tangkapan layar),
      jadi baris Lion BPN 2026 kosong di tabel per maskapai (hanya direct) itu benar. Dugaan awal
      "Lion tercatat sebagai Super Air Jet" tidak terbukti dan kemungkinan salah; angka Super Air Jet
      BPN dibiarkan apa adanya (harga identik ±Rp2,03 jt = tarif seragam).

## Berkas penting

| Berkas | Fungsi |
|---|---|
| `input_manual.py` | `--html` buat worksheet · `--panen` proses hasil · `--n` jumlah Jumat · `--date` acuan |
| `panen_browser.js`, `salin_hasil.js` | Isi bookmarklet 📦 Panen dan 📋 Salin hasil |
| `worksheet.html`, `kumpulan.txt` | Dibuat ulang tiap minggu; tidak di-commit |
| `hitung_rata_rata.py` | Rata-rata bulanan + sheet "Rp per km per Bulan" |
| `tabulasi_rute.py` | Excel 1 sheet Rp/km per rute untuk laporan |
| `CARA_PAKAI.md` | Panduan mingguan untuk user |
| `scrape_tiket.py` | Scraper otomatis lama (parser sudah diperbaiki; **diblokir Cloudflare**) |
