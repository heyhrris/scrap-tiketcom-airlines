#!/usr/bin/env python3
"""
Input manual harga tiket — pengganti scraping setelah tiket.com memasang
proteksi bot (403 Cloudflare) di endpoint pencarian.

Hasil disimpan ke format yang SAMA dengan output scraping
(hasil_scraping/{yymmdd}-Scrap Tiket Pesawat.csv/xlsx), sehingga
hitung_rata_rata.py langsung memprosesnya tanpa perubahan apa pun.

Baris manual ditandai airline = "(manual)" agar selalu bisa dibedakan
dari data hasil scraping bila suatu saat perlu dipisahkan.

CARA PAKAI
  1) Cetak daftar yang perlu dicek di browser:
        python input_manual.py --worksheet
  2) Simpan hasilnya — format per baris: KODE: h1, h2, ... (13 angka, urut H+4..H+88)
        BPN: 1.200.000, 1250000, ...
        DPS: 980000, -, 1010000, ...        ("-" = lewati/tidak ada)
     lalu:
        python input_manual.py --input data.txt
     (atau pipe:  cat data.txt | python input_manual.py --input - )
"""
import argparse
import re
import subprocess
import sys
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

OUTPUT_DIR = Path(__file__).parent / "hasil_scraping"
DAYS_AHEAD = list(range(4, 89, 7))          # sama dengan scraper: 13 Jumat
ROUTES_DOM = {                               # 7 rute domestik (yang dipakai analisis)
    "BPN": "Balikpapan", "DPS": "Denpasar-Bali", "KNO": "Medan", "PKU": "Pekanbaru",
    "SUB": "Surabaya",   "UPG": "Makassar",     "YIA": "Yogyakarta",
}
ROUTES_INT = {                               # internasional: bandara terbesar tiap negara
    "SIN": "Singapore",  "KUL": "Kuala-Lumpur", "BKK": "Bangkok",   # Singapura, Malaysia, Thailand
    "SGN": "Ho-Chi-Minh", "HND": "Tokyo",       "ICN": "Seoul",     # Vietnam, Jepang, Korea Selatan
    "SYD": "Sydney",     "DXB": "Dubai",                            # Australia, Uni Emirat Arab
}
ROUTES = {**ROUTES_DOM, **ROUTES_INT}
N_DOM, N_INT = 5, 3                          # jumlah Jumat per rute (diubah lewat --n / --n-int)


def hari(kode: str) -> list:
    """Daftar H+ untuk satu rute (internasional memakai lebih sedikit Jumat)."""
    return DAYS_AHEAD[:N_INT] if kode in ROUTES_INT else DAYS_AHEAD[:N_DOM]


COLS = ["scrape_date", "scrape_time", "destination", "h_plus", "travel_date",
        "airline", "jam_berangkat", "jam_tiba", "bandara_asal", "bandara_tujuan",
        "duration", "harga_display", "harga_angka"]


def senin_terakhir(d: datetime) -> datetime:
    """Senin pada/paling dekat sebelum d — agar tanggal target jatuh di Jumat."""
    senin = d - timedelta(days=d.weekday())
    return senin.replace(hour=0, minute=0, second=0, microsecond=0)


def build_url(dest: str, date_str: str) -> str:
    kode = {"BPN": ("BPNC", "CITY"), "DPS": ("DPSC", "CITY"), "KNO": ("KNO", "AIRPORT"),
            "PKU": ("PKU", "AIRPORT"), "SUB": ("SUBC", "CITY"), "UPG": ("UPGC", "CITY"),
            "YIA": ("YIA", "AIRPORT")}.get(dest, (dest, "AIRPORT"))   # internasional: kode bandara apa adanya
    return ("https://www.tiket.com/id-id/flights/search"
            f"?d=JKTC&dType=CITY&a={kode[0]}&aType={kode[1]}"
            f"&class=economy&adult=1&type=depart&date={date_str}"
            f"&dLabel=Jakarta&aLabel={ROUTES[dest].replace(' ', '-')}")


def worksheet(ref: datetime) -> None:
    total = sum(len(hari(k)) for k in ROUTES)
    print(f"\nAcuan : {ref.strftime('%A, %d %B %Y')}")
    print(f"Target: domestik {len(ROUTES_DOM)} rute x {N_DOM} Jumat + internasional {len(ROUTES_INT)} rute x {N_INT} Jumat = {total} halaman\n")
    print("\nBuka tiap URL di Chrome, klik 📦 Panen, lalu 📋 Salin hasil di akhir.\n")
    n = 0
    for kode in ROUTES:
        print(f"--- {kode} ({ROUTES[kode]}){' [internasional]' if kode in ROUTES_INT else ''} ---")
        for d in hari(kode):
            n += 1
            t = (ref + timedelta(days=d)).strftime("%Y-%m-%d")
            print(f"{n:3}. H+{d:<3} {t}  {build_url(kode, t)}")
        print()


def parse_input(teks: str) -> dict:
    hasil = {}
    for baris in teks.splitlines():
        baris = baris.strip()
        if not baris or ":" not in baris:
            continue
        kode, sisa = baris.split(":", 1)
        kode = kode.strip().upper()
        if kode not in ROUTES:
            continue
        angka = []
        for bagian in re.split(r"[,\s;|]+", sisa.strip()):
            if not bagian:
                continue
            if bagian in {"-", "_", "x", "X"}:
                angka.append(None)
                continue
            bersih = re.sub(r"[^\d]", "", bagian)
            angka.append(int(bersih) if bersih else None)
        hasil[kode] = angka
    return hasil


def simpan(data: dict, ref: datetime) -> None:
    ts = ref.strftime("%y%m%d")
    csv_path = OUTPUT_DIR / f"{ts}-Scrap Tiket Pesawat.csv"
    xlsx_path = OUTPUT_DIR / f"{ts}-Scrap Tiket Pesawat.xlsx"
    OUTPUT_DIR.mkdir(exist_ok=True)

    baris = []
    now = datetime.now()
    for kode, angka in data.items():
        if len(angka) != len(hari(kode)):
            print(f"⚠  {kode}: {len(angka)} angka (diharapkan {len(hari(kode))}) — "
                  f"dipetakan berurutan sejauh yang ada")
        for d, harga in zip(hari(kode), angka):
            if harga is None:
                continue
            td = (ref + timedelta(days=d)).strftime("%Y-%m-%d")
            baris.append({
                "scrape_date": ref.strftime("%Y-%m-%d"),
                "scrape_time": now.strftime("%H:%M:%S"),
                "destination": kode,
                "h_plus": f"H+{d}",
                "travel_date": td,
                "airline": "(manual)",
                "jam_berangkat": "", "jam_tiba": "",
                "bandara_asal": "CGK", "bandara_tujuan": kode,
                "duration": "Langsung",          # wajib: filter direct di hitung_rata_rata
                "harga_display": f"IDR {harga:,}".replace(",", "."),
                "harga_angka": harga,
            })

    if not baris:
        print("Tidak ada data untuk disimpan.")
        return

    df_baru = pd.DataFrame(baris)[COLS]
    if csv_path.exists():
        lama = pd.read_csv(csv_path, encoding="utf-8-sig")
        gabung = pd.concat([lama, df_baru], ignore_index=True)
        gabung = gabung.drop_duplicates(subset=["destination", "travel_date", "airline",
                                                "harga_angka"], keep="last")
    else:
        gabung = df_baru
    gabung.to_csv(csv_path, index=False, encoding="utf-8-sig")
    gabung.to_excel(xlsx_path, index=False, sheet_name="Semua Rute")

    print(f"\n✅ {len(df_baru)} baris ditambahkan (total {len(gabung)})")
    print(f"📄 {csv_path.name}\n📊 {xlsx_path.name}")
    print(f"\nLangkah berikutnya:  .venv/bin/python hitung_rata_rata.py")



def parse_panen(teks: str) -> list[dict]:
    """Parse keluaran panen_browser.js — satu baris per penerbangan, dipisah '|':
    dest|travel_date|airline|jam_berangkat|jam_tiba|kode_asal|kode_tujuan|durasi|harga
    """
    baris = []
    for ln in teks.splitlines():
        ln = ln.strip()
        if not ln or "|" not in ln:
            continue
        f = [x.strip() for x in ln.split("|")]
        if len(f) < 9:
            continue
        try:
            harga = int(re.sub(r"[^\d]", "", f[8]))
        except ValueError:
            continue
        if not harga or f[0].upper() not in ROUTES:
            continue
        baris.append({
            "destination": f[0].upper(), "travel_date": f[1], "airline": f[2],
            "jam_berangkat": f[3], "jam_tiba": f[4],
            "bandara_asal": f[5], "bandara_tujuan": f[6],
            "duration": f[7], "harga_angka": harga,
        })
    return baris


def simpan_panen(baris: list[dict], ref: datetime) -> None:
    """Simpan hasil panen ke format identik dengan output scraping."""
    if not baris:
        print("Tidak ada baris yang dikenali. Pastikan menempel keluaran panen_browser.js")
        sys.exit(1)

    ts = ref.strftime("%y%m%d")
    csv_path = OUTPUT_DIR / f"{ts}-Scrap Tiket Pesawat.csv"
    xlsx_path = OUTPUT_DIR / f"{ts}-Scrap Tiket Pesawat.xlsx"
    OUTPUT_DIR.mkdir(exist_ok=True)

    now = datetime.now()
    rows = []
    for b in baris:
        try:
            td = datetime.strptime(b["travel_date"], "%Y-%m-%d")
        except ValueError:
            continue
        rows.append({
            "scrape_date": ref.strftime("%Y-%m-%d"),
            "scrape_time": now.strftime("%H:%M:%S"),
            "destination": b["destination"],
            "h_plus": f"H+{(td - ref).days}",
            "travel_date": b["travel_date"],
            "airline": b["airline"],
            "jam_berangkat": b["jam_berangkat"], "jam_tiba": b["jam_tiba"],
            "bandara_asal": b["bandara_asal"], "bandara_tujuan": b["bandara_tujuan"],
            "duration": b["duration"],
            "harga_display": f"IDR {b['harga_angka']:,}".replace(",", "."),
            "harga_angka": b["harga_angka"],
        })

    df_baru = pd.DataFrame(rows)[COLS]
    if csv_path.exists():
        lama = pd.read_csv(csv_path, encoding="utf-8-sig")
        gabung = pd.concat([lama, df_baru], ignore_index=True)
    else:
        gabung = df_baru
    gabung = gabung.drop_duplicates(
        subset=["destination", "travel_date", "airline", "jam_berangkat", "harga_angka"],
        keep="last")
    gabung.to_csv(csv_path, index=False, encoding="utf-8-sig")
    gabung.to_excel(xlsx_path, index=False, sheet_name="Semua Rute")

    rute = df_baru["destination"].nunique()
    tgl = df_baru["travel_date"].nunique()
    print(f"\n✅ {len(df_baru)} penerbangan ditambahkan ({rute} rute, {tgl} tanggal)")
    print(f"   Total di file: {len(gabung)} baris")
    print(f"📄 {csv_path.name}\n📊 {xlsx_path.name}")
    sisa = sorted(set(ROUTES) - set(df_baru["destination"]))
    if sisa:
        print(f"⚠  Belum ada data: {', '.join(sisa)}")
    print(f"\nLangkah berikutnya:  .venv/bin/python hitung_rata_rata.py")



def bookmarklet(nama_file: str) -> str:
    """Ubah skrip JS menjadi URL bookmarklet (javascript:...)."""
    from urllib.parse import quote
    kode = (Path(__file__).parent / nama_file).read_text()
    return "javascript:" + quote(kode, safe="")


def buat_html(ref: datetime) -> Path:
    """Tulis worksheet.html: tombol bookmarklet + semua link rute x tanggal."""
    import html as _h
    bm_panen = _h.escape(bookmarklet("panen_browser.js"))
    bm_salin = _h.escape(bookmarklet("salin_hasil.js"))
    skrip = (Path(__file__).parent / "panen_browser.js").read_text()
    rows, n = [], 0
    for kode, nama in ROUTES.items():
        label = " · internasional" if kode in ROUTES_INT else ""
        rows.append(f'<h3>{kode} — {nama.replace("-", " ")}{label}</h3><ol start="{n+1}">')
        for d in hari(kode):
            t = (ref + timedelta(days=d)).strftime("%Y-%m-%d")
            n += 1
            rows.append(f'<li><a href="{_h.escape(build_url(kode, t))}" target="_blank">'
                        f'{t} &nbsp;<small>H+{d}</small></a></li>')
        rows.append("</ol>")
    doc = f"""<!doctype html><meta charset="utf-8"><title>Worksheet {ref:%d %b %Y}</title>
<style>
body{{font:15px/1.6 -apple-system,system-ui,sans-serif;max-width:860px;margin:2rem auto;padding:0 1rem;color:#222}}
h1{{font-size:20px}} h3{{margin:1.4rem 0 .3rem;font-size:15px;color:#05a}}
ol{{margin:.2rem 0 .2rem 1.2rem;padding:0}} li{{margin:.15rem 0}} ol.l li{{margin:.35rem 0}}
a{{color:#06c}} a:visited{{color:#999}}
.box{{background:#f6f7f9;border:1px solid #e2e4e8;border-radius:8px;padding:1rem 1.2rem;margin:1rem 0}}
.bm{{display:inline-block;margin:.3rem .6rem .3rem 0;padding:.5rem 1rem;border:2px dashed #06c;border-radius:8px;
     background:#fff;color:#06c !important;font-weight:600;text-decoration:none;cursor:grab}}
details{{margin-top:.6rem}} summary{{cursor:pointer;color:#555}}
textarea{{width:100%;height:110px;font:12px/1.4 ui-monospace,Menlo,monospace;border:1px solid #ccc;border-radius:6px;padding:.6rem}}
</style>
<h1>Worksheet panen harga — acuan Senin {ref:%d %B %Y}</h1>
<p><b>{n} halaman</b> (domestik {N_DOM} Jumat × {len(ROUTES_DOM)} rute; internasional {N_INT} Jumat × {len(ROUTES_INT)} rute). Link yang sudah dikunjungi jadi abu-abu.</p>
<p><small>Rute internasional BARU (BKK, SGN, HND, ICN, SYD, DXB): buka halaman pertama tiap rute dulu dan pastikan daftar penerbangan muncul. Kalau kosong atau kotak 📦 Panen merah, kode bandara di tiket.com perlu disesuaikan — kabari Claude.</small></p>

<div class="box">
<b>Pasang tombol (sekali saja, lewati kalau sudah terpasang):</b>
<ol class="l">
<li>Tampilkan bookmark bar Chrome: <b>Cmd+Shift+B</b></li>
<li><b>Seret</b> kedua tombol di bawah ini ke bookmark bar:</li>
</ol>
<a class="bm" href="{bm_panen}">📦 Panen</a>
<a class="bm" href="{bm_salin}">📋 Salin hasil</a>
<p><small>Jangan diklik di halaman ini — diseret ke bookmark bar. Kalau skripnya diperbarui,
hapus tombol lama di bookmark bar lalu seret ulang dari worksheet terbaru.</small></p>
</div>

<div class="box">
<b>Tiap halaman (2 klik):</b>
<ol class="l">
<li>Klik link di bawah → tunggu daftar penerbangan muncul</li>
<li>Klik <b>📦 Panen</b> di bookmark bar → tunggu kotak hijau <b>✅</b> di pojok kanan bawah → <b>Cmd+W</b></li>
</ol>
<b>Setelah halaman terakhir:</b>
<ol class="l">
<li>Di tab tiket.com mana pun, klik <b>📋 Salin hasil</b></li>
<li>Buka kumpulan.txt → <b>Cmd+V</b> → <b>Cmd+S</b> → bilang ke Claude "sudah tersimpan"</li>
</ol>
<p><small>Kotak <b style="color:#c00">merah</b> = ada yang perlu diulang; baca pesannya.</small></p>
<details><summary>Cara lama lewat Console (cadangan)</summary>
<p>Cmd+Option+J → tempel skrip → Enter. Di akhir ketik <code>copy(HASIL)</code>.</p>
<textarea readonly>{_h.escape(skrip)}</textarea></details>
</div>
{''.join(rows)}
"""
    out = Path(__file__).parent / "worksheet.html"
    out.write_text(doc)
    return out


def main():
    ap = argparse.ArgumentParser(description="Input manual harga tiket")
    ap.add_argument("--worksheet", action="store_true", help="cetak daftar & URL yang perlu dicek")
    ap.add_argument("--html", action="store_true", help="buat worksheet.html & buka di browser")
    ap.add_argument("--input", help="file berisi data (atau '-' untuk stdin)")
    ap.add_argument("--panen", help="file berisi keluaran panen_browser.js (atau '-' untuk stdin)")
    ap.add_argument("--n", type=int, default=5, help="jumlah Jumat untuk rute domestik (default 5, maks 13)")
    ap.add_argument("--n-int", type=int, default=3, dest="n_int", help="jumlah Jumat untuk rute internasional (default 3)")
    ap.add_argument("--date", help="tanggal acuan YYYY-MM-DD (default: Senin terakhir)")
    a = ap.parse_args()

    global N_DOM, N_INT
    N_DOM = max(1, min(a.n, len(DAYS_AHEAD)))
    N_INT = max(1, min(a.n_int, len(DAYS_AHEAD)))

    ref = (datetime.strptime(a.date, "%Y-%m-%d") if a.date
           else senin_terakhir(datetime.now()))

    if a.html:
        # Kosongkan tampungan minggu lalu (disalin dulu sebagai cadangan),
        # supaya data antar-minggu tidak tercampur tanpa disadari.
        kum = Path(__file__).parent / "kumpulan.txt"
        if kum.exists() and kum.stat().st_size > 0:
            (Path(__file__).parent / "kumpulan_lalu.txt").write_text(kum.read_text())
            kum.write_text("")
            print("🔄 kumpulan.txt dikosongkan (cadangan: kumpulan_lalu.txt)")
        else:
            kum.write_text("")

        out = buat_html(ref)
        print(f'✅ {out.name} dibuat (acuan {ref:%A, %d %B %Y})')
        subprocess.run(['open', str(out)], check=False)
        return

    if a.panen:
        teks = sys.stdin.read() if a.panen == "-" else Path(a.panen).read_text()
        simpan_panen(parse_panen(teks), ref)
        return

    if a.worksheet or not a.input:
        worksheet(ref)
        if not a.input:
            return
    teks = sys.stdin.read() if a.input == "-" else Path(a.input).read_text()
    data = parse_input(teks)
    if not data:
        print("Tidak ada baris yang dikenali. Format: 'BPN: 1200000, 1250000, ...'")
        sys.exit(1)
    simpan(data, ref)


if __name__ == "__main__":
    main()
