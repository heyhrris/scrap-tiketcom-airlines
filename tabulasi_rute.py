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
GREY = Font(name="Arial", size=10, italic=True, color="98A2B3")
KET = Font(name="Arial", size=9, italic=True, color="667085")
AMBER = PatternFill("solid", fgColor="FFF4D6"); BLUE = PatternFill("solid", fgColor="E8EEF7")
RP = "#,##0"
BLN = {1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "Mei", 6: "Jun", 7: "Jul", 8: "Agu", 9: "Sep", 10: "Okt", 11: "Nov", 12: "Des"}
URUT = DOM + LN
NC = len(URUT) + 3                                    # A + 9 rute + Semua + Keterangan

df["tahun"] = df["tanggal"].dt.year
dm = df[df.rute.isin(DOM)]
rk = df.pivot_table(index="bulan", columns="rute", values="rpkm", aggfunc="mean").reindex(columns=URUT)
nn = df.pivot_table(index="bulan", columns="rute", values="rpkm", aggfunc="size").reindex(columns=URUT)
semua = dm.groupby("bulan")["rpkm"].mean()
src = dm.groupby("bulan")["sumber"].agg(lambda x: "+".join(sorted(set(x))))
ty = df.pivot_table(index="tahun", columns="rute", values="rpkm", aggfunc="mean").reindex(columns=URUT)
ny = df.pivot_table(index="tahun", columns="rute", values="rpkm", aggfunc="size").reindex(columns=URUT)
ty_dom = dm.groupby("tahun")["rpkm"].mean()
per = df.groupby("tahun")["bulan"].agg(lambda b: f"{b.min()} s.d. {b.max()}")
rg = ring.set_index("rute")

wb = Workbook(); ws = wb.active; ws.title = "Rp per km per rute"
ws["A1"] = "Harga tiket pesawat per rute — Rp per km (penerbangan direct, CGK Jakarta → tujuan, sekali jalan)"
ws["A1"].font = Font(bold=True, name="Arial", size=13, color=NAVY)
ws["A2"] = ("Rp/km = harga ÷ jarak lurus CGK→tujuan. Per bulan keberangkatan. Abu-abu miring = sampel < 30 penerbangan; '–' = tidak ada data. "
            "Baris KUNING = bulan yang memuat data Manual (≈ harga termurah) — tidak sebanding langsung dengan baris lain; lihat catatan di bawah.")
ws["A2"].font = KET; ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=NC); ws.row_dimensions[2].height = 28


def kepala_tabel(r, judul):
    for j, t in enumerate([judul] + [f"CGK →\n{x} {NAMA[x]}" for x in URUT] + ["Semua rute\n(domestik)", "Keterangan"], 1):
        c = ws.cell(r, j, t); c.font = HFONT; c.fill = HF; c.border = BD
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[r].height = 34


def sel(r, j, v, n=None, tebal=False, fill=None):
    c = ws.cell(r, j); c.border = BD
    if v is None or (isinstance(v, float) and pd.isna(v)):
        c.value = "–"; c.font = GREY; c.alignment = Alignment(horizontal="center")
    else:
        c.value = round(float(v)); c.number_format = RP
        c.font = FB if tebal else (GREY if (n is not None and n < MIN_N) else F)
    if fill: c.fill = fill


# --- blok 1: bulanan
r = 4; kepala_tabel(r, "Bulan"); r += 1
ws.cell(r, 1, "Jarak (km)").font = FB; ws.cell(r, 1).border = BD; ws.cell(r, 1).fill = BLUE
for j, x in enumerate(URUT, 2):
    c = ws.cell(r, j, KM[x]); c.number_format = RP; c.font = FB; c.border = BD; c.fill = BLUE
for j in (NC - 1, NC): ws.cell(r, j).border = BD; ws.cell(r, j).fill = BLUE
r += 1; baris_data_awal = r
for b in rk.index:
    manual = "Manual" in src[b]; fl = AMBER if manual else None
    c = ws.cell(r, 1, b); c.font = FB; c.border = BD
    if fl: c.fill = fl
    for j, x in enumerate(URUT, 2): sel(r, j, rk.loc[b, x], nn.loc[b, x], fill=fl)
    sel(r, NC - 1, semua.get(b), tebal=True, fill=fl)
    c = ws.cell(r, NC, src[b]); c.font = F; c.border = BD
    if fl: c.fill = fl
    r += 1
r += 1

# --- blok 2: per tahun
kepala_tabel(r, "Rata-rata\nper tahun"); ws.cell(r, NC, "Periode data"); r += 1
for t in ty.index:
    ws.cell(r, 1, int(t)).font = FB; ws.cell(r, 1).border = BD; ws.cell(r, 1).alignment = Alignment(horizontal="left")
    for j, x in enumerate(URUT, 2): sel(r, j, ty.loc[t, x], ny.loc[t, x])
    sel(r, NC - 1, ty_dom.get(t), tebal=True)
    c = ws.cell(r, NC, per[t]); c.font = Font(name="Arial", size=9, color="667085"); c.border = BD
    r += 1
r += 1

# --- blok 3: terbaru (otomatis)
kepala_tabel(r, "Terbaru"); ws.cell(r, NC, "Periode"); r += 1
awal, akhir = tk["tanggal"].min(), tk["tanggal"].max()
dom_rp = tk[tk.rute.isin(DOM)]["rpkm"].mean()
rank = rg["rp"].rank(method="min").astype(int)
for nama, fn in (("Rp/km rata-rata", lambda x: rg.loc[x, "rp"] if x in rg.index else None),
                 ("Rp/km median", lambda x: rg.loc[x, "med"] if x in rg.index else None),
                 ("Jumlah penerbangan", lambda x: rg.loc[x, "n"] if x in rg.index else None)):
    ws.cell(r, 1, nama).font = FB; ws.cell(r, 1).border = BD
    for j, x in enumerate(URUT, 2):
        sel(r, j, fn(x), tebal=(nama == "Rp/km rata-rata"))
    if nama == "Rp/km rata-rata": sel(r, NC - 1, dom_rp, tebal=True)
    else: ws.cell(r, NC - 1).border = BD
    c = ws.cell(r, NC, f"tiket.com otomatis, {BLN[awal.month]} {awal.year} – {BLN[akhir.month]} {akhir.year}"); c.font = Font(name="Arial", size=9, color="667085"); c.border = BD
    r += 1
ws.cell(r, 1, "Peringkat (1 = termurah)").font = FB; ws.cell(r, 1).border = BD
for j, x in enumerate(URUT, 2):
    c = ws.cell(r, j, int(rank[x]) if x in rank.index else "–"); c.font = FB; c.border = BD; c.alignment = Alignment(horizontal="center")
ws.cell(r, NC - 1).border = BD; ws.cell(r, NC).border = BD
r += 2

# --- catatan
ws.cell(r, 1, "Catatan").font = Font(bold=True, name="Arial", size=11, color=NAVY); r += 1
catatan = [
 "Asal & arah: semua rute sekali jalan dari Jakarta (kode CGK). Data Traveloka seluruhnya CGK; data tiket.com ±87% CGK dan ±13% HLP (Halim). Jarak dihitung dari CGK; selisih ke HLP < 2% sehingga tidak mengubah peringkat.",
 "Cakupan: hanya penerbangan DIRECT. Transit dikeluarkan karena menempuh jarak jauh lebih panjang dari jarak lurus sehingga Rp/km-nya menyesatkan.",
 "Jenis data: (1) Traveloka (Okt 2023 – Okt 2025) dan (2) tiket.com otomatis (Mei 2026 –) = rata-rata semua penerbangan direct. (3) tiket.com Manual (Agu 2022 – Sep 2023 dan Nov 2025 – Apr 2026) = satu harga per rute-tanggal yang setara harga TERMURAH (1,04× termurah, ±0,80× rata-rata semua direct, dari perbandingan 36 pasang rute-tanggal). Periode Manual karenanya lebih rendah secara struktural.",
 "Lonjakan Apr → Mei 2026 (≈ +36%) bukan kenaikan harga murni: definisi data berganti dari Manual (≈ termurah) ke otomatis (rata-rata), yang saja menaikkan angka ±25%. Pada saat yang sama harga avtur melonjak (guncangan April 2026), sehingga porsi kenaikan riil tidak dapat dipisahkan. Bandingkan hanya bulan-bulan dalam jenis data yang sama.",
 "Rute pendek selalu tampak mahal per km (mis. Yogyakarta 424 km): biaya tetap per penerbangan (pajak bandara, ground handling) tersebar ke jarak yang lebih pendek. Perubahan satu rute antar waktu lebih bermakna daripada perbandingan antar rute.",
 "SIN & KUL hanya terambil sampai Agustus 2026 (scraping otomatis); setelah itu hanya rute domestik. 'Semua rute (domestik)' hanya merata-ratakan 7 rute domestik. Harga = tampilan tiket.com untuk 1 penumpang ekonomi sekali jalan, bukan harga tiket terjual; penerbangan yang sama dapat terambil di beberapa minggu dan semuanya dirata-rata.",
]
for t in catatan:
    c = ws.cell(r, 1, "• " + t); c.font = F; c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NC)
    ws.row_dimensions[r].height = 15 * (len(t) // 150 + 1) + 4
    r += 1

for i, w in enumerate([22] + [13] * len(URUT) + [14, 24], 1): ws.column_dimensions[get_column_letter(i)].width = w
ws.freeze_panes = ws.cell(6, 2)
out = DIR / f"{datetime.now():%y%m%d}-Rp per km per Rute.xlsx"
wb.save(out)
print("OK:", out.name, "| baris:", ws.max_row, "| sheet:", wb.sheetnames)
