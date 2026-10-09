#!/usr/bin/env python3
"""Tabulasi Rp/km per rute -> Excel 1 sheet (untuk laporan).

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
           "YIA": (-7.9055, 110.0573), "SIN": (1.3644, 103.9915), "KUL": (2.7456, 101.7099),
           "BKK": (13.6900, 100.7501), "SGN": (10.8188, 106.6520), "HND": (35.5494, 139.7798),
           "ICN": (37.4602, 126.4407), "SYD": (-33.9399, 151.1753), "DXB": (25.2532, 55.3657)}
NAMA = {"BPN": "Balikpapan", "DPS": "Denpasar", "KNO": "Medan", "PKU": "Pekanbaru", "SUB": "Surabaya",
        "UPG": "Makassar", "YIA": "Yogyakarta", "SIN": "Singapura", "KUL": "Kuala Lumpur",
        "BKK": "Bangkok", "SGN": "Ho Chi Minh", "HND": "Tokyo", "ICN": "Seoul", "SYD": "Sydney", "DXB": "Dubai"}
DOM = ["BPN", "DPS", "KNO", "PKU", "SUB", "UPG", "YIA"]
LN = ["SIN", "KUL", "BKK", "SGN", "HND", "ICN", "SYD", "DXB"]
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
                         "sumber": "Otomatis" if "duration" in d.columns else "Manual (≈termurah)"})


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
tk = df[df["sumber"] == "Otomatis"]
tk = tk[tk["tanggal"] >= "2026-05-01"]
ring = tk.groupby("rute").agg(n=("rpkm", "size"), harga=("harga", "mean"), rp=("rpkm", "mean"),
                              med=("rpkm", "median"), mn=("rpkm", "min"), mx=("rpkm", "max"),
                              dari=("tanggal", "min"), sampai=("tanggal", "max")).reset_index()

NAVY = "013D79"; MIN_N = 30
HF = PatternFill("solid", fgColor=NAVY); HFONT = Font(bold=True, color="FFFFFF", name="Arial", size=10)
BD = Border(*(Side(style="thin", color="D0D5DD"),) * 4)
F = Font(name="Arial", size=10); FB = Font(name="Arial", size=10, bold=True)
GREY = Font(name="Arial", size=10)          # sel kosong ("–") dan sampel tipis: tetap hitam, tidak miring
BLUE = PatternFill("solid", fgColor="E8EEF7")
RP = "#,##0"
URUT = DOM + LN
NC = len(URUT) + 3                                    # Bulan + 9 rute + Domestik + Internasional

df = df[df["bulan"] <= "2026-09"]                  # data dipotong s.d. September 2026
rk = df.pivot_table(index="bulan", columns="rute", values="rpkm", aggfunc="mean").reindex(columns=URUT)
nn = df.pivot_table(index="bulan", columns="rute", values="rpkm", aggfunc="size").reindex(columns=URUT)
dom = df[df.rute.isin(DOM)].groupby("bulan")["rpkm"].mean()
intl = df[df.rute.isin(LN)].groupby("bulan")["rpkm"].mean()

wb = Workbook(); ws = wb.active; ws.title = "Rp per km per rute"
ws["A1"] = "Harga tiket pesawat per rute — Rp per km (penerbangan direct, CGK Jakarta → tujuan, sekali jalan)"
ws["A1"].font = Font(bold=True, name="Arial", size=13, color=NAVY)

HDR = 3
for j, t in enumerate(["Bulan"] + [f"CGK →\n{x}\n{NAMA[x]}" for x in URUT] + ["Domestik", "Internasional"], 1):
    c = ws.cell(HDR, j, t); c.font = HFONT; c.fill = HF; c.border = BD
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ws.row_dimensions[HDR].height = 52


def sel(r, j, v, n=None, tebal=False):
    c = ws.cell(r, j); c.border = BD
    if v is None or (isinstance(v, float) and pd.isna(v)):
        c.value = "–"; c.font = GREY; c.alignment = Alignment(horizontal="center")
    else:
        c.value = round(float(v)); c.number_format = RP
        c.font = FB if tebal else F


r = HDR + 1
c = ws.cell(r, 1, "Jarak (km)"); c.font = FB; c.border = BD; c.fill = BLUE
for j, x in enumerate(URUT, 2):
    c = ws.cell(r, j, KM[x]); c.number_format = RP; c.font = FB; c.border = BD; c.fill = BLUE
for j in (NC - 1, NC): ws.cell(r, j).border = BD; ws.cell(r, j).fill = BLUE
r += 1
for b in rk.index:
    c = ws.cell(r, 1, b); c.font = FB; c.border = BD
    for j, x in enumerate(URUT, 2): sel(r, j, rk.loc[b, x], nn.loc[b, x])
    sel(r, NC - 1, dom.get(b), tebal=True)
    sel(r, NC, intl.get(b), tebal=True)
    r += 1

for i, w in enumerate([13] + [13] * len(URUT) + [13, 15], 1): ws.column_dimensions[get_column_letter(i)].width = w
ws.freeze_panes = ws.cell(HDR + 2, 2)
out = DIR / f"{datetime.now():%y%m%d}-Rp per km per Rute.xlsx"
k = 2
while out.exists():                               # jangan menimpa (mis. hasil edit manual)
    out = DIR / f"{datetime.now():%y%m%d}-Rp per km per Rute ({k}).xlsx"; k += 1
wb.save(out)
print("OK:", out.name, "| baris terakhir:", ws.max_row, "| sheet:", wb.sheetnames)
