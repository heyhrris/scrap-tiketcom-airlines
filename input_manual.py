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
import sys
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

OUTPUT_DIR = Path(__file__).parent / "hasil_scraping"
DAYS_AHEAD = list(range(4, 89, 7))          # sama dengan scraper: 13 Jumat
ROUTES = {                                   # 7 rute domestik (yang dipakai analisis)
    "BPN": "Balikpapan", "DPS": "Denpasar-Bali", "KNO": "Medan", "PKU": "Pekanbaru",
    "SUB": "Surabaya",   "UPG": "Makassar",     "YIA": "Yogyakarta",
}
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
            "YIA": ("YIA", "AIRPORT")}[dest]
    return ("https://www.tiket.com/id-id/flights/search"
            f"?d=JKTC&dType=CITY&a={kode[0]}&aType={kode[1]}"
            f"&class=economy&adult=1&type=depart&date={date_str}"
            f"&dLabel=Jakarta&aLabel={ROUTES[dest].replace(' ', '-')}")


def worksheet(ref: datetime) -> None:
    print(f"\nAcuan: {ref.strftime('%A, %d %B %Y')}  →  {len(DAYS_AHEAD)} tanggal Jumat\n")
    tgl = [(ref + timedelta(days=d)).strftime("%Y-%m-%d") for d in DAYS_AHEAD]
    print("Tanggal (urut H+4..H+88):")
    for d, t in zip(DAYS_AHEAD, tgl):
        print(f"   H+{d:<3} {t}  ({(ref + timedelta(days=d)).strftime('%a')})")
    print("\nBuka satu halaman per rute — strip 'harga termurah' di atas hasil")
    print("biasanya sudah menampilkan beberapa tanggal sekaligus:\n")
    for kode in ROUTES:
        print(f"  {kode}: {build_url(kode, tgl[0])}")
    print("\nFormat jawaban (13 angka per rute, urut seperti daftar tanggal di atas):")
    for kode in ROUTES:
        print(f"  {kode}: ")
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
        if len(angka) != len(DAYS_AHEAD):
            print(f"⚠  {kode}: {len(angka)} angka (diharapkan {len(DAYS_AHEAD)}) — "
                  f"dipetakan berurutan sejauh yang ada")
        for d, harga in zip(DAYS_AHEAD, angka):
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


def main():
    ap = argparse.ArgumentParser(description="Input manual harga tiket")
    ap.add_argument("--worksheet", action="store_true", help="cetak daftar & URL yang perlu dicek")
    ap.add_argument("--input", help="file berisi data (atau '-' untuk stdin)")
    ap.add_argument("--panen", help="file berisi keluaran panen_browser.js (atau '-' untuk stdin)")
    ap.add_argument("--date", help="tanggal acuan YYYY-MM-DD (default: Senin terakhir)")
    a = ap.parse_args()

    ref = (datetime.strptime(a.date, "%Y-%m-%d") if a.date
           else senin_terakhir(datetime.now()))

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
