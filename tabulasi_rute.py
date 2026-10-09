#!/usr/bin/env python3
"""Tabulasi Rp/km per rute -> Excel (untuk laporan).

Rp/km = harga / jarak great-circle CGK->tujuan; hanya penerbangan DIRECT.
Sumber: seluruh arsip hasil_scraping (Traveloka historis + tiket.com).
"""
import glob, math
from datetime import datetime
from pathlib import Path
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DIR = Path(__file__).parent / "hasil_scraping"
CGK = (-6.1256, 106.6559)
BANDARA = {"BPN": (-1.2683, 116.8945), "DPS": (-8.7482, 115.1671), "KNO": (3.6422, 98.8854),
           "PKU": (0.4609, 101.4445), "SUB": (-7.3798, 112.7869), "UPG": (-5.0617, 119.5547),
           "YIA": (-7.9055, 110.0573), "SIN": (1.3644, 103.9915), "KUL": (2.7456, 101.7099)}
NAMA = {"BPN": "Balikpapan", "DPS": "Denpasar", "KNO": "Medan", "PKU": "Pekanbaru", "SUB": "Surabaya",
        "UPG": "Makassar", "YIA": "Yogyakarta", "SIN": "Singapura", "KUL": "Kuala Lumpur"}
DOM = ["BPN", "DPS", "KNO", "PKU", "SUB", "UPG", "YIA"]
LN = ["SIN", "KUL"]
NORM = {"MES": "KNO", "JOG": "YIA"}


def hav(a, b):
    R = 6371; p1, l1, p2, l2 = map(math.radians, [a[0], a[1], b[0], b[1]])
    return round(2 * R * math.asin(math.sqrt(math.sin((p2-p1)/2)**2 + math.cos(p1)*math.cos(p2)*math.sin((l2-l1)/2)**2)))
KM = {k: hav(CGK, v) for k, v in BANDARA.items()}


def tgl(s):
    d = pd.to_datetime(s, format="%d/%m/%y", errors="coerce")
    for f in ("%d-%m-%Y", "%Y-%m-%d"):
        d = d.fillna(pd.to_datetime(s, format=f, errors="coerce"))
    return d


def baca_traveloka(path):
    with open(path, encoding="utf-8") as fh: head = fh.readline()
    sep = ";" if head.count(";") > head.count(",") else ","
    d = pd.read_csv(path, sep=sep, encoding="utf-8", on_bad_lines="skip")
    if not {"Date", "Price", "Schedule"}.issubset(d.columns) or d.empty: return None
    return pd.DataFrame({
        "tanggal": tgl(d["Date"]),
        "rute": d["Schedule"].astype(str).str[-3:].str.upper().replace(NORM),
        "direct": d["Schedule"].astype(str).str.contains("Direct", case=False, na=False),
        "harga": pd.to_numeric(d["Price"].astype(str).str.replace(r"[^\d]", "", regex=True), errors="coerce"),
        "sumber": "Traveloka"})


def baca_tiket(path):
    d = pd.read_csv(path, encoding="utf-8-sig", on_bad_lines="skip")
    if "destination" not in d.columns: return None
    if "travel_date" in d.columns:
        t = pd.to_datetime(d["travel_date"], errors="coerce")
    else:
        t = pd.to_datetime(d["scrape_date"], errors="coerce") + pd.to_timedelta(
            pd.to_numeric(d["h_plus"].astype(str).str.replace(r"[^\d]", "", regex=True), errors="coerce"), unit="D")
    direct = d["duration"].astype(str).str.contains("Langsung", case=False, na=False) if "duration" in d.columns else True
    return pd.DataFrame({"tanggal": t, "rute": d["destination"].astype(str).str.upper().replace(NORM),
                         "direct": direct, "harga": pd.to_numeric(d["harga_angka"], errors="coerce"),
                         "sumber": "tiket.com"})


frames = []
for f in sorted(glob.glob(str(DIR / "*.csv"))):
    nm = Path(f).name
    if "Rata-rata" in nm or "Rp per km" in nm: continue
    try:
        x = baca_traveloka(f) if "traveloka" in nm.lower() else baca_tiket(f)
    except Exception:
        x = None
    if x is not None and len(x): frames.append(x)
df = pd.concat(frames, ignore_index=True)
df = df[df["rute"].isin(KM) & df["direct"] & (df["harga"] > 0)].dropna(subset=["tanggal"]).copy()
df["km"] = df["rute"].map(KM); df["rpkm"] = df["harga"] / df["km"]
df["bulan"] = df["tanggal"].dt.to_period("M").astype(str)

# ---- ringkasan terbaru: tiket.com semua periode & per bulan terbaru
tk = df[df["sumber"] == "tiket.com"]
tk = tk[tk["tanggal"] >= "2026-05-01"]
ring = tk.groupby("rute").agg(n=("rpkm", "size"), harga=("harga", "mean"), rp=("rpkm", "mean"),
                              med=("rpkm", "median"), mn=("rpkm", "min"), mx=("rpkm", "max"),
                              dari=("tanggal", "min"), sampai=("tanggal", "max")).reset_index()

NAVY = "013D79"
HF = PatternFill("solid", fgColor=NAVY); HFONT = Font(bold=True, color="FFFFFF", name="Arial", size=10)
BD = Border(*(Side(style="thin", color="D0D5DD"),) * 4)
F = Font(name="Arial", size=10); FB = Font(name="Arial", size=10, bold=True)
GREY = Font(name="Arial", size=10, italic=True, color="98A2B3")
RP = "#,##0"
ASAL = "Jakarta"
KET_ASAL = "Asal: Jakarta (CGK dan HLP); jarak dihitung dari CGK. Semua rute sekali jalan Jakarta → tujuan."


def hdr(ws, r, kolom, c0=1):
    for j, t in enumerate(kolom, c0):
        c = ws.cell(r, j, t); c.font = HFONT; c.fill = HF; c.border = BD
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def lebar(ws, w):
    for i, x in enumerate(w, 1): ws.column_dimensions[get_column_letter(i)].width = x


wb = Workbook()
# ---- Lembar 1: ringkasan per rute
ws = wb.active; ws.title = "Ringkasan per rute"
ws["A1"] = "Harga tiket pesawat per rute — Rp per km (penerbangan direct, Jakarta → tujuan)"
ws["A1"].font = Font(bold=True, name="Arial", size=13, color=NAVY)
ws["A2"] = (f"Sumber: tiket.com, keberangkatan {tk['tanggal'].min():%d %b %Y} – {tk['tanggal'].max():%d %b %Y}. "
            "Diurutkan dari Rp/km termurah.")
ws["A2"].font = Font(italic=True, name="Arial", size=9, color="667085")
ws["A3"] = KET_ASAL; ws["A3"].font = Font(italic=True, name="Arial", size=9, color="667085")
hdr(ws, 4, ["Peringkat", "Asal", "Tujuan", "Kota tujuan", "Jenis", "Jarak (km)", "Jumlah penerbangan", "Harga rata-rata (Rp)",
            "Rp/km rata-rata", "Rp/km median", "Rp/km minimum", "Rp/km maksimum"])
ws.row_dimensions[4].height = 32
r0 = 5
for jenis, daftar in (("Domestik", DOM), ("Internasional", LN)):
    sub = ring[ring.rute.isin(daftar)].sort_values("rp")
    for k, r in enumerate(sub.itertuples(), 1):
        v = [k, ASAL, r.rute, NAMA[r.rute], jenis, KM[r.rute], r.n, round(r.harga), round(r.rp), round(r.med), round(r.mn), round(r.mx)]
        for j, x in enumerate(v, 1):
            c = ws.cell(r0, j, x); c.border = BD; c.font = FB if j in (3, 9) else F
            if j >= 6: c.number_format = RP
            if j == 1: c.alignment = Alignment(horizontal="center")
        r0 += 1
    r0 += 1
ws.cell(r0, 1, "Rute internasional (SIN, KUL) hanya terambil sampai Agustus 2026 saat scraping otomatis masih berjalan; "
               "harga rata-rata yang lebih tinggi dipengaruhi maskapai full-service.").font = Font(italic=True, name="Arial", size=9, color="667085")
lebar(ws, [10, 10, 9, 15, 14, 11, 14, 16, 14, 13, 13, 13]); ws.freeze_panes = "A5"

# ---- Lembar 2 & 3: Rp/km & n per bulan (domestik)
dm = df[df.rute.isin(DOM)]
rk = dm.pivot_table(index="bulan", columns="rute", values="rpkm", aggfunc="mean")[DOM]
nn = dm.pivot_table(index="bulan", columns="rute", values="rpkm", aggfunc="size")[DOM]
semua = dm.groupby("bulan")["rpkm"].mean()
ns = dm.groupby("bulan").size()
src = dm.groupby("bulan")["sumber"].agg(lambda s: "tiket.com" if (s == "tiket.com").all()
                                         else ("Traveloka" if (s == "Traveloka").all() else "campuran"))
MIN_N = 30
for judul, tabel, extra, fmt in (("Rp per km per bulan", rk, semua, "rp"), ("Jumlah sampel per bulan", nn, ns, "n")):
    ws = wb.create_sheet(judul)
    ws["A1"] = ("Rata-rata Rp/km per rute per bulan keberangkatan (direct, domestik)" if fmt == "rp"
                else "Jumlah penerbangan direct yang dihitung per rute per bulan")
    ws["A1"].font = Font(bold=True, name="Arial", size=12, color=NAVY)
    if fmt == "rp":
        ws["A2"] = f"Abu-abu miring = sampel < {MIN_N} penerbangan. Kolom 'Sumber' menunjukkan asal data bulan itu."
        ws["A2"].font = Font(italic=True, name="Arial", size=9, color="667085")
    ws["A3"] = KET_ASAL; ws["A3"].font = Font(italic=True, name="Arial", size=9, color="667085")
    hdr(ws, 4, ["Bulan"] + [f"Jakarta →\n{r} {NAMA[r]}" for r in DOM] + ["Semua rute (domestik)", "Sumber"])
    ws.row_dimensions[4].height = 32
    for i, b in enumerate(tabel.index, 5):
        ws.cell(i, 1, b).font = FB; ws.cell(i, 1).border = BD
        for j, r in enumerate(DOM, 2):
            v = tabel.loc[b, r]; c = ws.cell(i, j); c.border = BD
            if pd.isna(v):
                c.value = "–"; c.font = GREY; c.alignment = Alignment(horizontal="center")
            else:
                c.value = round(float(v)) if fmt == "rp" else int(v); c.number_format = RP
                c.font = GREY if (fmt == "rp" and nn.loc[b, r] < MIN_N) else F
        c = ws.cell(i, len(DOM) + 2, round(float(extra[b])) if fmt == "rp" else int(extra[b]))
        c.number_format = RP; c.font = FB; c.border = BD
        c = ws.cell(i, len(DOM) + 3, src[b]); c.font = F; c.border = BD
    lebar(ws, [10] + [13] * len(DOM) + [20, 12]); ws.freeze_panes = "B5"

# ---- Lembar 4: catatan
ws = wb.create_sheet("Catatan metode")
ws["A1"] = "Catatan metode & batasan data"; ws["A1"].font = Font(bold=True, name="Arial", size=12, color=NAVY)
cat = [
 ("Asal & arah", "Semua rute sekali jalan dari Jakarta ke kota tujuan (tidak ada arah sebaliknya). Data Traveloka seluruhnya berangkat dari CGK; "
   "data tiket.com mencakup semua bandara Jakarta (sekitar 87% CGK, 13% HLP/Halim). Jarak selalu dihitung dari CGK; selisihnya ke HLP kurang dari 2% sehingga tidak mengubah peringkat."),
 ("Rumus", "Rp/km = harga tiket ÷ jarak lurus (great-circle) Jakarta (CGK) → bandara tujuan, dirata-rata per kelompok."),
 ("Cakupan", "Hanya penerbangan DIRECT. Penerbangan transit dikeluarkan karena menempuh jarak jauh lebih panjang dari jarak lurus, sehingga Rp/km-nya menyesatkan."),
 ("Jarak (km)", "; ".join(f"{r} {KM[r]:,}".replace(",", ".") for r in DOM + LN) + ". Jarak lurus, bukan jarak rute terbang."),
 ("Mengapa rute pendek mahal per km", "Biaya tetap per penerbangan (pajak bandara, ground handling) tersebar ke jarak yang lebih pendek, "
   "sehingga rute seperti Yogyakarta (424 km) selalu tinggi per km walaupun harga totalnya tidak mahal. Bandingkan antar rute dengan hati-hati; "
   "perubahan antar waktu pada rute yang sama lebih bermakna."),
 ("Sumber & periode", "Arsip Traveloka (Agu 2022 – 2025) dan tiket.com (2026). Dua sumber memiliki cakupan dan komposisi harga berbeda, "
   "sehingga ada patahan seri di sekitar Mei 2026 — lihat kolom 'Sumber'. Jangan membaca selisih Apr→Mei 2026 sebagai kenaikan harga murni."),
 ("Kenaikan 2026", "Lonjakan Rp/km sejak Mei 2026 bertepatan dengan dua hal: pergantian sumber data dan kenaikan harga avtur akibat guncangan April 2026. "
   "Dengan data yang ada keduanya tidak bisa dipisahkan sepenuhnya."),
 ("Sampel tipis", f"Angka dengan jumlah observasi < {MIN_N} ditulis abu-abu miring. Periode Okt 2025 – Apr 2026 memang bersampel tipis (≈110–140 penerbangan per bulan)."),
 ("Harga", "Harga tampil untuk 1 penumpang ekonomi sekali jalan; bukan harga tiket yang benar-benar terjual. Penerbangan yang sama dapat terambil di beberapa minggu; semua observasi dirata-rata."),
]
for i, (k, v) in enumerate(cat, 3):
    a = ws.cell(i, 1, k); a.font = FB; a.alignment = Alignment(vertical="top", wrap_text=True)
    b = ws.cell(i, 2, v); b.font = F; b.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[i].height = max(30, 15 * (len(v) // 100 + 1))
lebar(ws, [26, 110])

out = DIR / f"{datetime.now():%y%m%d}-Rp per km per Rute.xlsx"
wb.save(out)
print("OK:", out.name)
print(ring.sort_values("rp")[["rute", "n", "rp", "med"]].round(0).to_string(index=False))
print("\nRp/km 'Semua rute' 2026-06..11 (cek vs tabel sebelumnya):")
print(semua["2026-06":].round(0).to_string())
