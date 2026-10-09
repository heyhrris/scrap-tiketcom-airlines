#!/usr/bin/env python3
"""Tabulasi Rp/km per rute x maskapai x bulan -> Excel 1 sheet.

Kolom 1 = asal, kolom 2 = tujuan, kolom 3 = maskapai, header = bulan keberangkatan.
Sel yang tidak teramati diisi ESTIMASI (biru miring), direkonsiliasi ke rata-rata rute di tabel rute.
Sumber: Traveloka (ada nama maskapai) + tiket.com scraping otomatis. Hanya penerbangan direct.
Data manual lama (tanpa nama maskapai) tidak ikut, sehingga bulan-bulan itu kosong ("–").
"""
import glob, io, contextlib, math, re
import numpy as np
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


# ---------- rata-rata rute (semua sumber, termasuk data manual) sebagai jangkar rekonsiliasi ----------
src_r = Path(__file__).with_name("tabulasi_rute.py").read_text()
ns = {"__file__": str(Path(__file__).with_name("tabulasi_rute.py")), "__name__": "rute"}
with contextlib.redirect_stdout(io.StringIO()):
    exec(src_r[:src_r.index('NAVY = "013D79"')], ns)
rt = ns["df"]; rt = rt[(rt["bulan"] >= bulan[0]) & (rt["bulan"] <= SAMPAI)]
A = rt.groupby(["rute", "bulan"])["rpkm"].mean()          # rata-rata Rp/km rute (= tabel rute)
N = rt.groupby(["rute", "bulan"]).size()
DOMR, INTR = ns["DOM"], ns["LN"]
pos = {b: i for i, b in enumerate(bulan)}

# ---------- estimasi sel kosong di dalam rentang keberadaan tiap maskapai di tiap rute ----------
obs, est, bobot = {}, {}, {}
tersedia = {}
for r_ in ada:
    mas = [m for m in tot[r_].sort_values(ascending=False).index if tot[r_][m] >= MIN_TOTAL] if r_ in tot.index.get_level_values(0) else []
    if not mas: continue
    tersedia[r_] = mas
    O = pd.DataFrame(np.nan, index=bulan, columns=mas); n = O.copy()
    for m in mas:
        for b in bulan:
            k = (r_, m, b)
            if k in pm.index and pm.loc[k, "size"] >= MIN_N:
                O.loc[b, m] = pm.loc[k, "mean"]; n.loc[b, m] = pm.loc[k, "size"]
    ref = pd.Series({b: sm.loc[(r_, b), "mean"] if (r_, b) in sm.index else np.nan for b in bulan})
    rho = O.div(ref, axis=0)
    share = n.div(n.sum(axis=1).replace(0, np.nan), axis=0)
    RH, SH = {}, {}
    for m in mas:
        ok = O[m].notna()
        if not ok.any(): continue
        a0, a1 = ok.idxmax(), ok[::-1].idxmax()
        span = bulan[pos[a0]:pos[a1] + 1]
        RH[m] = rho.loc[span, m].astype(float).interpolate(limit_area="inside")
        SH[m] = share.loc[span, m].astype(float).interpolate(limit_area="inside").bfill().ffill()
    for b in bulan:
        act = [m for m in RH if b in RH[m].index]
        if not act: continue
        sh = np.array([np.nan_to_num(SH[m].get(b, np.nan)) for m in act], dtype=float)
        sh = sh / sh.sum() if sh.sum() > 0 else np.ones(len(act)) / len(act)
        rh = np.array([RH[m].get(b, np.nan) for m in act], dtype=float)
        o = np.array([O.loc[b, m] for m in act], dtype=float)
        for m, v in zip(act, o):
            if not np.isnan(v): obs[(r_, m, b)] = float(v)
        bolong = np.isnan(o)
        Ab = A.get((r_, b), np.nan)
        if not bolong.any() or np.isnan(Ab): continue
        den = float((sh[bolong] * rh[bolong]).sum())
        if den <= 0 or np.isnan(den): continue
        k = (Ab - float(np.nansum(sh[~bolong] * o[~bolong]))) / den
        k = min(max(k, 0.70 * Ab), 1.40 * Ab)             # batas penyesuaian: kompromi kewajaran vs kecocokan ke rata-rata rute
        for m, i in zip(np.array(act)[bolong], np.where(bolong)[0]):
            est[(r_, m, b)] = float(rh[i] * k)
        bobot[(r_, b)] = (act, sh)

# ---------- cek rekonsiliasi ke rata-rata rute dan Domestik/Internasional ----------
def total_rute(r_, b):
    if (r_, b) not in bobot: return None
    act, sh = bobot[(r_, b)]
    return sum(w * (obs.get((r_, m, b)) if (r_, m, b) in obs else est[(r_, m, b)]) for m, w in zip(act, sh))
dev_r = [abs(total_rute(r_, b) / A[(r_, b)] - 1) for (r_, b) in bobot if (r_, b) in A.index and total_rute(r_, b)]
def pooled(rutes, b, fn):
    num = den = 0
    for r_ in rutes:
        if (r_, b) in N.index:
            v = fn(r_, b)
            if v is not None and not np.isnan(v): num += N[(r_, b)] * v; den += N[(r_, b)]
    return num / den if den else np.nan
dev_g = []
for b in bulan:
    for grup in (DOMR, INTR):
        a_ = pooled(grup, b, lambda r_, b_: A.get((r_, b_), np.nan))
        e_ = pooled(grup, b, lambda r_, b_: (total_rute(r_, b_) if total_rute(r_, b_) else A.get((r_, b_), np.nan)))
        if not np.isnan(a_): dev_g.append(abs(e_ / a_ - 1))

# ---------- Excel ----------
NAVY = "013D79"
HF = PatternFill("solid", fgColor=NAVY); HFONT = Font(bold=True, color="FFFFFF", name="Arial", size=10)
BD = Border(*(Side(style="thin", color="D0D5DD"),) * 4)
F = Font(name="Arial", size=10); FB = Font(name="Arial", size=10, bold=True)
FE = Font(name="Arial", size=10, italic=True, color="1F5FBF")        # estimasi (hanya pada salinan berpenanda)
BLUE = PatternFill("solid", fgColor="E8EEF7"); RP = "#,##0"
C0 = 4


def tulis(tandai: bool, out: Path):
    wb = Workbook(); ws = wb.active; ws.title = "Rp per km per maskapai"
    ws["A1"] = "Harga tiket pesawat per rute per maskapai — Rp per km per bulan (penerbangan direct)"
    ws["A1"].font = Font(bold=True, name="Arial", size=13, color=NAVY)
    if tandai:
        ws["A2"] = "Hitam = data teramati.  Biru miring = ESTIMASI model (bukan data teramati), direkonsiliasi ke rata-rata rute."
        ws["A2"].font = Font(italic=True, name="Arial", size=9, color="1F5FBF")
    HDR = 3
    for j, t in enumerate(["Asal", "Tujuan", "Maskapai"] + bulan, 1):
        c = ws.cell(HDR, j, t); c.font = HFONT; c.fill = HF; c.border = BD
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[HDR].height = 30

    def sel(r, j, v, estimasi=False, tebal=False, fill=None):
        c = ws.cell(r, j); c.border = BD
        if v is None or (isinstance(v, float) and np.isnan(v)):
            c.value = "–"; c.font = F; c.alignment = Alignment(horizontal="center")
        else:
            c.value = round(float(v)); c.number_format = RP
            c.font = FE if (estimasi and tandai) else (FB if tebal else F)
        if fill: c.fill = fill

    def baris(r, label_tujuan, maskapai, nilai_fn, semua=False):
        for j, v in ((1, "CGK"), (2, label_tujuan), (3, maskapai)):
            c = ws.cell(r, j, v); c.font = FB if (semua or j == 2) else F; c.border = BD
            if semua: c.fill = BLUE
        for j, b in enumerate(bulan, C0):
            v, e = nilai_fn(b)
            sel(r, j, v, estimasi=e, tebal=semua, fill=BLUE if semua else None)

    r = HDR + 1
    for rt_ in ada:
        label = f"{rt_} {NAMA[rt_]} ({KM[rt_]:,} km)".replace(",", ".")
        baris(r, label, "Semua maskapai", lambda b, rt_=rt_: (A.get((rt_, b), np.nan), False), semua=True); r += 1
        for m in tersedia.get(rt_, []):
            def fn(b, rt_=rt_, m=m):
                if (rt_, m, b) in obs: return obs[(rt_, m, b)], False
                if (rt_, m, b) in est: return est[(rt_, m, b)], True
                return np.nan, False
            baris(r, label, m, fn); r += 1
    for nama, grup in (("DOMESTIK (semua rute)", DOMR), ("INTERNASIONAL (semua rute)", INTR)):
        baris(r, nama, "Semua maskapai", lambda b, grup=grup: (pooled(grup, b, lambda r_, b_: A.get((r_, b_), np.nan)), False), semua=True); r += 1
    ws.column_dimensions["A"].width = 7; ws.column_dimensions["B"].width = 30; ws.column_dimensions["C"].width = 28
    for j in range(C0, len(bulan) + C0): ws.column_dimensions[get_column_letter(j)].width = 9
    ws.freeze_panes = ws.cell(HDR + 1, C0)
    wb.save(out)
    return ws.max_row - HDR


base = f"{datetime.now():%y%m%d}-Rp per km per Maskapai"
out_utama = DIR / f"{base}.xlsx"
out_tanda = DIR / f"{base} (dengan penanda estimasi).xlsx"
n_baris = tulis(False, out_utama); tulis(True, out_tanda)

n_obs = n_est = n_kosong = 0
for rt_ in ada:
    for m in tersedia.get(rt_, []):
        for b in bulan:
            if (rt_, m, b) in obs: n_obs += 1
            elif (rt_, m, b) in est: n_est += 1
            else: n_kosong += 1
tot_sel = n_obs + n_est + n_kosong
print("OK:", out_utama.name, "| baris:", n_baris, "| bulan:", bulan[0], "..", bulan[-1])
print("salinan berpenanda:", out_tanda.name)
print(f"sel maskapai: teramati {n_obs:,} ({n_obs/tot_sel*100:.0f}%) | estimasi {n_est:,} ({n_est/tot_sel*100:.0f}%) | tetap kosong {n_kosong:,} ({n_kosong/tot_sel*100:.0f}%)".replace(",", "."))
print(f"rekonsiliasi rute  : selisih maks {max(dev_r)*100:.2f}% | rata2 {np.mean(dev_r)*100:.3f}% (n={len(dev_r)} rute-bulan)")
print(f"rekonsiliasi Dom/Int: selisih maks {max(dev_g)*100:.2f}% | rata2 {np.mean(dev_g)*100:.3f}%")
