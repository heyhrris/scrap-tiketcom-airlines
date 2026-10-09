#!/usr/bin/env python3
"""Tabulasi Rp/km per rute x maskapai x bulan -> Excel 1 sheet.

Kolom 1 = rute, kolom 2 = maskapai, header = bulan keberangkatan.
Sumber: Traveloka (ada nama maskapai) + tiket.com scraping otomatis. Hanya penerbangan direct.
Data manual lama (tanpa nama maskapai) tidak ikut, sehingga bulan-bulan itu kosong ("–").
"""
import glob, math, re
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
           "YIA": (-7.9055, 110.0573), "SIN": (1.3644, 103.9915), "KUL": (2.7456, 101.7099),
           "BKK": (13.6900, 100.7501), "SGN": (10.8188, 106.6520), "HND": (35.5494, 139.7798),
           "ICN": (37.4602, 126.4407), "SYD": (-33.9399, 151.1753), "DXB": (25.2532, 55.3657)}
NAMA = {"BPN": "Balikpapan", "DPS": "Denpasar", "KNO": "Medan", "PKU": "Pekanbaru", "SUB": "Surabaya",
        "UPG": "Makassar", "YIA": "Yogyakarta", "SIN": "Singapura", "KUL": "Kuala Lumpur",
        "BKK": "Bangkok", "SGN": "Ho Chi Minh", "HND": "Tokyo", "ICN": "Seoul", "SYD": "Sydney", "DXB": "Dubai"}
URUT = ["BPN", "DPS", "KNO", "PKU", "SUB", "UPG", "YIA", "SIN", "KUL", "BKK", "SGN", "HND", "ICN", "SYD", "DXB"]
NORM = {"MES": "KNO", "JOG": "YIA"}
MIN_N = 5          # sel bulanan dengan < 5 penerbangan tidak ditampilkan ("–")
MIN_TOTAL = 30     # kombinasi rute-maskapai dengan total < 30 penerbangan tidak dimuat
SAMPAI = "2026-09"


def hav(a, b):
    R = 6371; p1, l1, p2, l2 = map(math.radians, [a[0], a[1], b[0], b[1]])
    return round(2 * R * math.asin(math.sqrt(math.sin((p2-p1)/2)**2 + math.cos(p1)*math.cos(p2)*math.sin((l2-l1)/2)**2)))
KM = {k: hav(CGK, v) for k, v in BANDARA.items()}


def promo(x):
    return bool(re.search(r"Login untuk|Pesan sebelum", str(x))) or str(x).strip() in ("", "nan")


def bersih(x):
    x = re.sub(r"\s*\(codeshare\)\s*$", "", str(x).strip(), flags=re.I)
    return {"Batik Air": "Batik Air Indonesia", "Batik Air (Malaysia)": "Batik Air Malaysia"}.get(x, x)


def tgl(s):
    d = pd.to_datetime(s, format="%d/%m/%y", errors="coerce")
    for f in ("%d-%m-%Y", "%Y-%m-%d"):
        d = d.fillna(pd.to_datetime(s, format=f, errors="coerce"))
    return d


# ---------- tiket.com otomatis (+ pemulihan nama maskapai dari jadwal yang sama) ----------
fr = []
for f in sorted(glob.glob(str(DIR / "26*-Scrap Tiket Pesawat.csv"))):
    d = pd.read_csv(f, encoding="utf-8-sig", on_bad_lines="skip")
    if "duration" in d.columns and "airline" in d.columns: fr.append(d)
tk = pd.concat(fr, ignore_index=True)
tk = tk[tk["destination"].isin(KM) & (tk["harga_angka"] > 0)
        & tk["duration"].astype(str).str.contains("Langsung", case=False, na=False)].copy()
tk["promo"] = tk["airline"].map(promo)
tk["maskapai"] = tk["airline"].where(~tk["promo"]).map(lambda x: bersih(x) if isinstance(x, str) else None)
kunci = ["destination", "jam_berangkat", "jam_tiba"]
ident = tk[tk["maskapai"].notna()]
mode = ident.groupby(kunci)["maskapai"].agg(lambda s: s.value_counts().index[0])
share = ident.groupby(kunci)["maskapai"].agg(lambda s: s.value_counts(normalize=True).iloc[0])
peta = mode[share >= 0.9]                                    # jadwal yang hampir pasti milik satu maskapai
idx = pd.MultiIndex.from_frame(tk.loc[tk["promo"], kunci])
pulih = pd.Series(peta.reindex(idx).values, index=tk.index[tk["promo"]])
n_promo = int(tk["promo"].sum()); n_pulih = int(pulih.notna().sum())
tk.loc[pulih.index, "maskapai"] = pulih
tk = tk.assign(rute=tk["destination"], harga=tk["harga_angka"],
               bulan=pd.to_datetime(tk["travel_date"]).dt.to_period("M").astype(str))[["rute", "maskapai", "harga", "bulan"]]

# ---------- Traveloka ----------
tv = []
for f in sorted(glob.glob(str(DIR / "*traveloka*.csv"))):
    with open(f, encoding="utf-8") as fh: head = fh.readline()
    d = pd.read_csv(f, sep=";" if head.count(";") > head.count(",") else ",", encoding="utf-8", on_bad_lines="skip")
    if {"Date", "Price", "Schedule", "Carrier"}.issubset(d.columns) and len(d):
        tv.append(pd.DataFrame({
            "tanggal": tgl(d["Date"]), "rute": d["Schedule"].astype(str).str[-3:].str.upper().replace(NORM),
            "direct": d["Schedule"].astype(str).str.contains("Direct", case=False, na=False),
            "harga": pd.to_numeric(d["Price"].astype(str).str.replace(r"[^\d]", "", regex=True), errors="coerce"),
            "maskapai": d["Carrier"].map(bersih)}))
tv = pd.concat(tv, ignore_index=True)
tv = tv[tv["rute"].isin(KM) & tv["direct"] & (tv["harga"] > 0)].dropna(subset=["tanggal"]).copy()
tv["bulan"] = tv["tanggal"].dt.to_period("M").astype(str)
tv = tv[["rute", "maskapai", "harga", "bulan"]]

df = pd.concat([tv, tk], ignore_index=True)
df = df[df["bulan"] <= SAMPAI].copy()
df["rpkm"] = df["harga"] / df["rute"].map(KM)
ident = df[df["maskapai"].notna()]
bulan = [str(p) for p in pd.period_range(ident["bulan"].min(), SAMPAI, freq="M")]
ada = [r for r in URUT if r in set(df["rute"])]
pm = ident.groupby(["rute", "maskapai", "bulan"])["rpkm"].agg(["mean", "size"])
tot = ident.groupby(["rute", "maskapai"]).size()
sm = df.groupby(["rute", "bulan"])["rpkm"].agg(["mean", "size"])      # semua direct, termasuk yang tak teridentifikasi

# ---------- Excel ----------
NAVY = "013D79"
HF = PatternFill("solid", fgColor=NAVY); HFONT = Font(bold=True, color="FFFFFF", name="Arial", size=10)
BD = Border(*(Side(style="thin", color="D0D5DD"),) * 4)
F = Font(name="Arial", size=10); FB = Font(name="Arial", size=10, bold=True)
BLUE = PatternFill("solid", fgColor="E8EEF7"); RP = "#,##0"
wb = Workbook(); ws = wb.active; ws.title = "Rp per km per maskapai"
ws["A1"] = "Harga tiket pesawat per rute per maskapai — Rp per km per bulan (penerbangan direct, CGK Jakarta → tujuan)"
ws["A1"].font = Font(bold=True, name="Arial", size=13, color=NAVY)
HDR = 3
for j, t in enumerate(["Rute", "Maskapai"] + bulan, 1):
    c = ws.cell(HDR, j, t); c.font = HFONT; c.fill = HF; c.border = BD
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ws.row_dimensions[HDR].height = 30


def sel(r, j, mean, n, tebal=False, fill=None):
    c = ws.cell(r, j); c.border = BD
    if mean is not None and n >= MIN_N:
        c.value = round(float(mean)); c.number_format = RP; c.font = FB if tebal else F
    else:
        c.value = "–"; c.font = F; c.alignment = Alignment(horizontal="center")
    if fill: c.fill = fill


r = HDR + 1
for rt in ada:
    label = f"{rt} {NAMA[rt]} ({KM[rt]:,} km)".replace(",", ".")
    mas = [m for m in tot[rt].sort_values(ascending=False).index if tot[rt][m] >= MIN_TOTAL] if rt in tot.index.get_level_values(0) else []
    for m in ["Semua maskapai"] + mas:
        semua = m == "Semua maskapai"
        for j, v in ((1, label), (2, m)):
            c = ws.cell(r, j, v); c.font = FB if (semua or j == 1) else F; c.border = BD
            if semua: c.fill = BLUE
        for j, b in enumerate(bulan, 3):
            src = sm if semua else pm
            k = (rt, b) if semua else (rt, m, b)
            if k in src.index: sel(r, j, src.loc[k, "mean"], src.loc[k, "size"], tebal=semua, fill=BLUE if semua else None)
            else: sel(r, j, None, 0, fill=BLUE if semua else None)
        r += 1

ws.column_dimensions["A"].width = 30; ws.column_dimensions["B"].width = 28
for j in range(3, len(bulan) + 3): ws.column_dimensions[get_column_letter(j)].width = 9
ws.freeze_panes = ws.cell(HDR + 1, 3)
out = DIR / f"{datetime.now():%y%m%d}-Rp per km per Maskapai.xlsx"
k = 2
while out.exists():
    out = DIR / f"{datetime.now():%y%m%d}-Rp per km per Maskapai ({k}).xlsx"; k += 1
wb.save(out)
print("OK:", out.name, "| baris data:", ws.max_row - HDR, "| bulan:", bulan[0], "..", bulan[-1])
print(f"tiket.com: {n_promo:,} penerbangan direct bernama promo; dipulihkan dari jadwal sama: {n_pulih:,} ({n_pulih/n_promo*100:.0f}%)".replace(",", "."))
print("rute:", ada)
