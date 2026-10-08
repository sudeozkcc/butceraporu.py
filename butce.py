import datetime
import io
import re
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from openpyxl.utils import get_column_letter
from plotly.subplots import make_subplots

st.set_page_config(page_title="Lojistik Ciro & Bütçe Analizi", page_icon="🚚", layout="wide")

# ===================================================================
# STİL & TEMA DEFINITIONS
# ===================================================================
BLUE, ORANGE, GREY, GREEN, RED = "#1F4E9C", "#E07B00", "#9AA5B1", "#2E7D32", "#C62828"
BLUE_SEQ = ["#1F4E9C", "#2F64B5", "#4479C4", "#5A8ED0", "#74A2DA",
            "#8FB5E2", "#A9C6EA", "#C0D6F0", "#D3E3F6", "#E3EDF9"]

st.markdown(
    """
<style>
div[data-testid="stMetric"]{background:#f7f9fc;border:1px solid #e3e8ef;border-radius:10px;padding:10px 14px;}
div[data-testid="stMetricLabel"], div[data-testid="stMetricLabel"] p{color:#5b6574 !important;font-size:0.8rem !important;}
div[data-testid="stMetricValue"]{color:#1f2937 !important;font-size:1.15rem !important;line-height:1.3 !important;}
div[data-testid="stMetricValue"] > div{overflow:visible !important;text-overflow:clip !important;white-space:normal !important;word-break:break-word;}
div[data-testid="stMetricDelta"]{font-size:0.75rem !important;}
h1{font-size:1.7rem !important;} h2{font-size:1.35rem !important;} h3{font-size:1.1rem !important;}
div[data-testid="stMarkdownContainer"] p, div[data-testid="stMarkdownContainer"] li{font-size:0.92rem;}
.pro-wrap{overflow-x:auto;border-radius:10px;border:1px solid #e3e8ef;box-shadow:0 1px 3px rgba(16,24,40,.06);margin:4px 0 12px 0;}
.pro-tbl{border-collapse:collapse;width:100%;}
.pro-tbl th{padding:9px 12px;color:#fff;font-weight:700;text-align:center;white-space:nowrap;border-right:1px solid rgba(255,255,255,.35);}
.pro-tbl td{padding:6px 12px;border-bottom:1px solid #eef1f6;white-space:nowrap;}
.pro-tbl td:first-child{white-space:normal;}
.hero{background:linear-gradient(100deg,#12306b,#1F4E9C 60%,#2F7BD0);color:#fff;border-radius:14px;padding:18px 24px;margin-bottom:10px;box-shadow:0 2px 8px rgba(16,24,40,.15);}
.hero .t{font-size:1.55rem;font-weight:800;letter-spacing:.2px;}
.hero .s{font-size:0.9rem;opacity:.85;margin-top:2px;}
.chip{display:inline-block;padding:3px 10px;border-radius:999px;color:#fff;font-weight:600;font-size:.8rem;margin-right:6px;}
</style>
""",
    unsafe_allow_html=True,
)

TR_AY = ["", "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran",
         "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]

def _st_ver():
    try:
        return tuple(int(x) for x in st.__version__.split(".")[:2])
    except Exception:
        return (0, 0)

NEW_API = _st_ver() >= (1, 50)

def stretch():
    return {"width": "stretch"} if NEW_API else {"use_container_width": True}

def show(fig):
    try:
        st.plotly_chart(fig, **stretch())
    except Exception:
        st.plotly_chart(fig, use_container_width=True)

def editor(df, **kw):
    try:
        return st.data_editor(df, **stretch(), **kw)
    except Exception:
        return st.data_editor(df, use_container_width=True, **kw)

# ===================================================================
# YARDIMCI BİÇİMLENDİRME VE HESAP FONKSİYONLARI
# ===================================================================
_TR = str.maketrans("İıŞşĞğÜüÖöÇç", "IISSGGUUOOCC")

def norm(x) -> str:
    if x is None or (not isinstance(x, (list, tuple)) and pd.isna(x)):
        return ""
    return str(x).translate(_TR).upper().strip()

def name_key(x) -> str:
    return re.sub(r"\s+", " ", norm(x))

def compact(x) -> str:
    return re.sub(r"[^A-Z0-9]", "", norm(x))

_GENERIC = {"GIDA", "SAN", "TIC", "VE", "AS", "ANONIM", "SIRKETI", "LTD", "STI", "TURKIYE", "SANAYI",
            "TICARET", "FOOD", "FOODS", "GRUP", "HOLDING", "A", "S", "ISLETMELERI", "PAZARLAMA"}

def names_match(target_name, excel_name) -> bool:
    e = compact(excel_name)
    if len(e) < 4:
        return False
    etoks = set(re.findall(r"[A-Z0-9]+", norm(excel_name)))
    for part in str(target_name).split("/"):
        c = compact(part)
        if len(c) >= 4 and (c in e or e in c):
            return True
        toks = [t for t in re.findall(r"[A-Z0-9]+", norm(part)) if len(t) >= 3]
        if len(toks) >= 2 and all(t in e for t in toks):
            return True
        dt = [t for t in re.findall(r"[A-Z0-9]+", norm(part)) if t not in _GENERIC and len(t) >= 2]
        if dt and sum(len(t) for t in dt) >= 3 and all(t in etoks for t in dt):
            return True
    return False

def to_num(v) -> float:
    if isinstance(v, (int, float, np.number)):
        return 0.0 if pd.isna(v) else float(v)
    if v is None:
        return 0.0
    s = str(v).replace("₺", "").replace("TL", "").replace("\xa0", "").replace(" ", "").strip()
    if s in ("", "-", "nan", "NaN"):
        return 0.0
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".") if s.count(",") == 1 else s.replace(",", "")
    elif re.fullmatch(r"-?\d{1,3}(\.\d{3})+", s):
        s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return 0.0

def tr_num(x, dec=0) -> str:
    if x is None or pd.isna(x):
        return "0"
    s = f"{float(x):,.{dec}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_tl(x) -> str:
    return tr_num(x, dec=0) + " ₺"

def fmt_short_tl(x) -> str:
    if abs(x) >= 1e6:
        t = tr_num(x / 1e6, 2)
        t = t[:-3] if t.endswith(",00") else (t[:-1] if t.endswith("0") else t)
        return t + " M ₺"
    if abs(x) >= 1e3:
        return tr_num(x / 1e3, 0) + " B ₺"
    return fmt_tl(x)

def fmt_palet(x) -> str:
    return tr_num(x, dec=0) + " ad."

def fmt_pct(x, dec=1) -> str:
    if x is None or pd.isna(x):
        return "—"
    return tr_num(x, dec) + " %"

def hex_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))

def mix(rgb, a):
    r, g, b = [int(round(255 - (255 - c) * a)) for c in rgb]
    return f"rgb({r},{g},{b})"

def shade(values, rgb, lo=0.06, hi=0.40):
    v = np.nan_to_num(np.array(values, dtype=float))
    m = np.abs(v).max() if len(v) else 0
    if not m:
        return [mix(rgb, lo)] * len(v)
    return [mix(rgb, lo + (hi - lo) * abs(x) / m) for x in v]

def diverge(values, center=0.0):
    v = np.nan_to_num(np.array(values, dtype=float)) - center
    m = np.abs(v).max() if len(v) else 0
    out = []
    for x in v:
        if not m or x == 0:
            out.append("rgb(247,248,250)")
        elif x > 0:
            out.append(mix((46, 160, 67), 0.10 + 0.30 * abs(x) / m))
        else:
            out.append(mix((214, 39, 40), 0.10 + 0.30 * abs(x) / m))
    return out

# ===================================================================
# RENDER & KART BİLEŞENLERİ
# ===================================================================
def render_table(headers, cols, fills, widths=None, font_size=12, font_colors=None, header_fill=BLUE):
    n = len(cols[0]) if cols else 0
    esc = lambda x: str(x).replace("&", "&amp;")
    th = "".join(f'<th style="background:{header_fill};font-size:{font_size + 1}px;">{esc(h)}</th>' for h in headers)
    body = []
    for i in range(n):
        tds = []
        for j, c in enumerate(cols):
            fill = fills[j][i] if isinstance(fills[j], (list, tuple, np.ndarray)) else fills[j]
            fc = font_colors[j][i] if font_colors and isinstance(font_colors[j], (list, tuple, np.ndarray)) else (font_colors[j] if font_colors else "#1a202c")
            al = "left" if j == 0 else "right"
            tds.append(f'<td style="background:{fill};color:{fc};text-align:{al};font-size:{font_size}px;">{esc(c[i])}</td>')
        body.append("<tr>" + "".join(tds) + "</tr>")
    html = (f'<div class="pro-wrap"><table class="pro-tbl"><thead><tr>{th}</tr></thead>'
            f'<tbody>{"".join(body)}</tbody></table></div>')
    st.markdown(html, unsafe_allow_html=True)

def rgba(color, a):
    r, g, b = hex_rgb(color)
    return f"rgba({r},{g},{b},{a})"

def kpis(items):
    cols = st.columns(len(items))
    for col, it in zip(cols, items):
        label, value, color = it[0], it[1], it[2]
        sub = it[3] if len(it) > 3 and it[3] else ""
        sub_color = it[4] if len(it) > 4 else "#5b6574"
        col.markdown(
            f'<div style="background:linear-gradient(135deg,{rgba(color, 0.13)},#ffffff);border-left:6px solid {color};'
            f'border-radius:10px;padding:12px 16px;box-shadow:0 1px 4px rgba(0,0,0,.07);">'
            f'<div style="font-size:0.78rem;color:#5b6574;">{label}</div>'
            f'<div style="font-size:1.25rem;font-weight:700;color:{color};line-height:1.35;word-break:break-word;">{value}</div>'
            f'<div style="font-size:0.78rem;font-weight:600;color:{sub_color};min-height:1em;">{sub}</div></div>',
            unsafe_allow_html=True)

def banner(text, color=BLUE):
    st.markdown(
        f'<div style="background:linear-gradient(90deg,{color},{color}CC);color:white;padding:10px 16px;'
        f'border-radius:10px;font-weight:700;font-size:1.05rem;margin:16px 0 8px 0;">{text}</div>',
        unsafe_allow_html=True)

# ===================================================================
# YENİ GRAFİK BİLEŞENLERİ (NOKTASAL ÇİZGİ & DÖNEMSEL İZLEME)
# ===================================================================
def monthly_line_trend(mon_df):
    """Ocak'tan bulunulan aya kadar trendi noktasal (marker) çizgi grafikle, yeşil/kırmızı gösterir."""
    if mon_df is None or mon_df.empty:
        return None
    piv = mon_df.groupby("Ay").agg({"Ciro": "sum", "Palet": "sum"}).reset_index()
    piv["Ay_Adı"] = piv["Ay"].apply(lambda m: TR_AY[m])
    
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    
    # Ciro Çizgisi
    ciro_diff = piv["Ciro"].diff().fillna(0)
    colors_c = [GREEN if x >= 0 else RED for x in ciro_diff]
    
    fig.add_trace(go.Scatter(
        x=piv["Ay_Adı"], y=piv["Ciro"], name="Ciro (₺)",
        mode="lines+markers+text",
        line=dict(color=BLUE, width=3),
        marker=dict(size=10, color=colors_c, symbol="circle"),
        text=[fmt_short_tl(v) for v in piv["Ciro"]],
        textposition="top center",
        hovertemplate="<b>%{x} Ciro:</b> %{y:,.0f} ₺<extra></extra>"
    ), secondary_y=False)
    
    # Palet Çizgisi
    fig.add_trace(go.Scatter(
        x=piv["Ay_Adı"], y=piv["Palet"], name="Palet (ad)",
        mode="lines+markers",
        line=dict(color=ORANGE, width=2, dash="dash"),
        marker=dict(size=8, color=ORANGE, symbol="diamond"),
        hovertemplate="<b>%{x} Palet:</b> %{y:,.0f} ad.<extra></extra>"
    ), secondary_y=True)

    fig.update_layout(
        template="plotly_white", separators=",.", height=400,
        title=dict(text="<b>Ocak - Güncel Ay Trend Analizi (Noktasal Grafik)</b>", x=0.5),
        legend=dict(orientation="h", y=1.1, x=0.35),
        margin=dict(l=20, r=20, t=60, b=20)
    )
    fig.update_yaxes(title_text="Ciro (₺)", showgrid=True, secondary_y=False)
    fig.update_yaxes(title_text="Palet (ad)", showgrid=False, secondary_y=True)
    return fig

# ===================================================================
# EXCEL PARS SİSTEMİ
# ===================================================================
CIRO_RE = re.compile(r"CIRO|TUTAR|GELIR|SATIS|\bTL\b|₺")
EXCL_RE = re.compile(r"\bORT\b|ORTALAMA|BASI|BASINA|ORAN|\bPAY\b|FARK|BUYUME|%|ARTIS|DEGISIM|HEDEF|BUTCE")
TOTAL_RE = re.compile(r"TOPLAM|GENEL|YILLIK")
MONTH_FULL = {"OCAK": 1, "SUBAT": 2, "MART": 3, "NISAN": 4, "MAYIS": 5, "HAZIRAN": 6,
              "TEMMUZ": 7, "AGUSTOS": 8, "EYLUL": 9, "EKIM": 10, "KASIM": 11, "ARALIK": 12}
MONTH_ABBR = {"OCA": 1, "SUB": 2, "MAR": 3, "NIS": 4, "MAY": 5, "HAZ": 6,
              "TEM": 7, "AGU": 8, "AGO": 8, "EYL": 9, "EKI": 10, "KAS": 11, "ARA": 12}

def cell_month(v):
    if isinstance(v, (pd.Timestamp, datetime.datetime, datetime.date)):
        return v.month
    if isinstance(v, str):
        for tok in re.findall(r"[A-Z]+", norm(v)):
            if tok in MONTH_FULL:
                return MONTH_FULL[tok]
            if tok in MONTH_ABBR:
                return MONTH_ABBR[tok]
    return None

@st.cache_data(show_spinner=False)
def get_sheet_names(file_bytes: bytes):
    return pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl").sheet_names

@st.cache_data(show_spinner=False)
def read_raw(file_bytes: bytes, sheet: str) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet, header=None, engine="openpyxl")

def col_label(raw, j, hdr):
    parts = []
    for v in raw.iloc[: hdr + 1, j]:
        if pd.notna(v) and not isinstance(v, (int, float, np.number)):
            parts.append(str(v).strip())
    return " ".join(dict.fromkeys(parts))

def month_map(raw, hdr, cand_cols):
    n = raw.shape[1]
    rows = [[cell_month(raw.iat[i, j]) for j in range(n)] for i in range(hdr + 1)]
    direct = [None] * n
    for ms in rows:
        for j, m in enumerate(ms):
            if m and direct[j] is None:
                direct[j] = m
    inherited = [None] * n
    if rows:
        best = max(rows, key=lambda ms: sum(1 for m in ms if m))
        if any(best):
            cur = None
            for j, m in enumerate(best):
                if m:
                    cur = m
                inherited[j] = cur
    out = {}
    for j in cand_cols:
        m = direct[j] or inherited[j]
        if m:
            out[j] = m
    return out

def detect_layout(raw: pd.DataFrame) -> dict:
    best_i, best_cnt = None, 0
    for i in range(min(10, len(raw))):
        cnt = 0
        for j in range(raw.shape[1]):
            v = raw.iat[i, j]
            if isinstance(v, str):
                n = norm(v)
                if "PALET" in n or CIRO_RE.search(n):
                    cnt += 1
        if cnt > best_cnt:
            best_i, best_cnt = i, cnt
    hdr = best_i if best_i is not None else 2

    labels = [col_label(raw, j, hdr) for j in range(raw.shape[1])]
    ciro_all, palet_all = [], []
    for j, lab in enumerate(labels):
        n = norm(lab)
        if not n or EXCL_RE.search(n):
            continue
        if "PALET" in n:
            palet_all.append(j)
        elif CIRO_RE.search(n):
            ciro_all.append(j)

    def prefer_total(cols):
        tot = [j for j in cols if TOTAL_RE.search(norm(labels[j]))]
        return tot if tot else cols

    non_total = lambda cols: [j for j in cols if not TOTAL_RE.search(norm(labels[j]))]
    ciro_m = month_map(raw, hdr, non_total(ciro_all))
    palet_m = month_map(raw, hdr, non_total(palet_all))
    ciro_c, palet_c = prefer_total(ciro_all), prefer_total(palet_all)

    start = hdr + 1
    cust = 0
    for j in range(raw.shape[1]):
        if j in ciro_c or j in palet_c:
            continue
        vals = raw.iloc[start:, j].dropna().astype(str)
        if len(vals) > 0 and sum(not re.fullmatch(r"[\d\.,\s-]+", v) for v in vals) / len(vals) > 0.5:
            cust = j
            break
    return dict(hdr=hdr, start=start, cust=cust, ciro=ciro_c, palet=palet_c,
                ciro_m=ciro_m, palet_m=palet_m, labels=labels)

def _clean_name(nm):
    if pd.isna(nm):
        return None
    nm = str(nm).strip()
    if not nm or nm.lower() in ("nan", "none"):
        return None
    return nm

def build_df(raw, cust_col, ciro_cols, palet_cols, start_row):
    rows, excel_total = [], None
    sub = raw.iloc[start_row:]
    for i in range(len(sub)):
        nm = _clean_name(sub.iat[i, cust_col])
        if nm is None:
            continue
        ciro = sum(to_num(sub.iat[i, j]) for j in ciro_cols)
        palet = sum(to_num(sub.iat[i, j]) for j in palet_cols)
        if "TOPLAM" in norm(nm):
            excel_total = (ciro, palet)
            continue
        rows.append((nm, ciro, palet))
    if not rows:
        return pd.DataFrame(columns=["Anahtar", "Müşteri", "Ciro", "Palet"]), excel_total
    df = pd.DataFrame(rows, columns=["Müşteri", "Ciro", "Palet"])
    df["Anahtar"] = df["Müşteri"].map(name_key)
    g = df.groupby("Anahtar", as_index=False, sort=False).agg(
        Müşteri=("Müşteri", "first"), Ciro=("Ciro", "sum"), Palet=("Palet", "sum"))
    g = g[(g["Ciro"] != 0) | (g["Palet"] != 0)].reset_index(drop=True)
    return g, excel_total

def build_monthly(raw, cust_col, ciro_m, palet_m, start_row):
    cols = ["Anahtar", "Müşteri", "Ay", "Ciro", "Palet"]
    if not ciro_m and not palet_m:
        return pd.DataFrame(columns=cols)
    recs = []
    sub = raw.iloc[start_row:]
    for i in range(len(sub)):
        nm = _clean_name(sub.iat[i, cust_col])
        if nm is None or "TOPLAM" in norm(nm):
            continue
        acc = {}
        for j, m in ciro_m.items():
            acc.setdefault(m, [0.0, 0.0])[0] += to_num(sub.iat[i, j])
        for j, m in palet_m.items():
            acc.setdefault(m, [0.0, 0.0])[1] += to_num(sub.iat[i, j])
        for m, (c, p) in acc.items():
            recs.append((nm, m, c, p))
    if not recs:
        return pd.DataFrame(columns=cols)
    df = pd.DataFrame(recs, columns=["Müşteri", "Ay", "Ciro", "Palet"])
    df["Anahtar"] = df["Müşteri"].map(name_key)
    g = df.groupby(["Anahtar", "Ay"], as_index=False).agg(
        Müşteri=("Müşteri", "first"), Ciro=("Ciro", "sum"), Palet=("Palet", "sum"))
    return g[cols]

def load_year(year: int):
    empty = pd.DataFrame(columns=["Anahtar", "Müşteri", "Ciro", "Palet"])
    empty_m = pd.DataFrame(columns=["Anahtar", "Müşteri", "Ay", "Ciro", "Palet"])
    up = st.sidebar.file_uploader(f"{year} Excel Raporunu Yükleyin (.xlsx)", type=["xlsx"], key=f"u{year}")
    if up is None:
        return empty, empty_m
    fb = up.getvalue()
    sheet = st.sidebar.selectbox(f"{year} Sayfa Seçin:", get_sheet_names(fb), key=f"s{year}")
    try:
        raw = read_raw(fb, sheet)
    except Exception as e:
        st.sidebar.error(f"{year} dosyası okunamadı: {e}")
        return empty, empty_m

    lay = detect_layout(raw)
    uid = f"{year}_{up.name}_{sheet}"
    opts = list(range(raw.shape[1]))
    fmt = lambda j: f"{get_column_letter(j + 1)} · {lay['labels'][j] or '(başlıksız)'}"

    with st.sidebar.expander(f"⚙️ {year} sütun eşleştirme", expanded=False):
        cust = st.selectbox("Müşteri sütunu", opts, index=lay["cust"], format_func=fmt, key=f"c_{uid}")
        ciro_cols = st.multiselect("Ciro sütun(lar)ı", opts, default=lay["ciro"], format_func=fmt, key=f"ci_{uid}")
        palet_cols = st.multiselect("Palet sütun(lar)ı", opts, default=lay["palet"], format_func=fmt, key=f"pa_{uid}")
        start = st.number_input("Veri başlangıç satırı", min_value=1, max_value=max(len(raw), 1), value=lay["start"] + 1, key=f"st_{uid}") - 1

    df, xl_tot = build_df(raw, cust, ciro_cols, palet_cols, int(start))
    if df.empty:
        return empty, empty_m

    mon = build_monthly(raw, cust, lay["ciro_m"], lay["palet_m"], int(start))
    return df, mon

# ===================================================================
# HEDEF VERİLERİ (MOD 3 İÇİN)
# ===================================================================
SEGMENTS = [
    ("1. SOĞUK – ET / TAVUK / ŞARKÜTERİ", [
        ("BANVİT", 2000), ("ŞENPİLİÇ", 2000), ("BEYPİ / BEYPİLİÇ", 1500), ("LEZİTA", 1500),
        ("NAMET GIDA", 1500), ("PINAR ET", 1200), ("POLONEZ", 1000), ("AYTAÇ", 1000)]),
    ("2. SOĞUK – SÜT / PEYNİR / SÜT ÜRÜNLERİ", [
        ("SÜTAŞ", 2000), ("PINAR SÜT", 1500), ("AK GIDA / İÇİM", 1500), ("EKER GIDA", 1200)]),
    ("3. DONUK / DONDURULMUŞ GIDA", [
        ("KEREVİTAŞ / SUPERFRESH", 2500), ("ÖZGÖRKEY GIDA / FEAST", 2000), ("PEK FOOD", 2000)]),
    ("4. ZİNCİR RESTORAN / FOOD SERVICE", [
        ("TAB GIDA", 3000), ("TAVUK DÜNYASI", 2000), ("DÜRÜMLE", 1500), ("BIGCHEFS", 1500)])
]

# ===================================================================
# ARAYÜZ LOHİSTİK UYGULAMASI
# ===================================================================
st.markdown('<div class="hero"><div class="t">🚚 Lojistik Ciro, Palet &amp; Bütçe Sunumu</div>'
            '<div class="s">Gerçekleşen performans · Yıllık karşılaştırma · Büyüme & Düşüş Analizleri</div></div>', unsafe_allow_html=True)

st.sidebar.header("📁 Excel Dosya Yükleme")
df_2025, mon_2025 = load_year(2025)
df_2026, mon_2026 = load_year(2026)

MODES = ["📊 2026'da Ne Yaptık?", "⚔️ Karşılaştırmalı Dönem Analizi", "🚀 2027 Hedefleri"]
mode = st.sidebar.radio("📌 Çalışma Modunu Seçin:", MODES)

# -------------------------------------------------------------------
# MOD 1: 2026 PERFORMANSI
# -------------------------------------------------------------------
if mode == MODES[0]:
    st.header("📊 2026 Yılı Performans Özeti")
    if df_2026.empty:
        st.info("👈 Sol menüden 2026 Excel dosyasını yükleyin.")
        st.stop()

    tot_ciro, tot_palet = df_2026["Ciro"].sum(), df_2026["Palet"].sum()
    avg_ptl = tot_ciro / tot_palet if tot_palet > 0 else 0.0

    kpis([("Toplam Ciro", fmt_tl(tot_ciro), BLUE),
          ("Toplam Palet", fmt_palet(tot_palet), ORANGE),
          ("Palet Başı Gelir", fmt_tl(avg_ptl), GREEN),
          ("Müşteri Sayısı", f"{len(df_2026)} firma", GREY)])

    st.markdown("---")
    st.subheader("📈 Aylık Seyir ve Noktasal Trend (Ocak - Güncel Ay)")
    tr_fig = monthly_line_trend(mon_2026)
    if tr_fig:
        show(tr_fig)
    else:
        st.caption("Aylık detay verisi bulunamadı.")

    st.subheader("📋 Genel Müşteri Sıralama Tablosu")
    t = df_2026.sort_values("Ciro", ascending=False).reset_index(drop=True)
    t["Palet Başı ₺"] = t["Ciro"].div(t["Palet"].where(t["Palet"] > 0)).fillna(0.0)
    
    st.dataframe(
        t[["Müşteri", "Ciro", "Palet", "Palet Başı ₺"]].style.format({
            "Ciro": lambda v: fmt_tl(v),
            "Palet": lambda v: fmt_palet(v),
            "Palet Başı ₺": lambda v: fmt_tl(v)
        }),
        use_container_width=True
    )

# -------------------------------------------------------------------
# MOD 2: KARŞILAŞTIRMALI DÖNEM ANALİZİ (GÜNCELLENEN KISIM)
# -------------------------------------------------------------------
elif mode == MODES[1]:
    st.header("⚔️ Karşılaştırmalı Dönem Analizi")
    if df_2025.empty or df_2026.empty:
        st.warning("⚠️ Karşılaştırma yapabilmek için sol menüden hem **2025** hem de **2026** dosyalarını yükleyiniz.")
        st.stop()

    # Dönem Seçim Filtresi
    st.subheader("⚙️ Karşılaştırılacak Dönem / Ay Seçimi")
    c_sel1, c_sel2 = st.columns(2)
    
    m_list = [("Genel (Ocak - Güncel Ay)", "GENEL")] + [(f"{TR_AY[i]} Ayı", i) for i in range(1, 13)]
    
    with c_sel1:
        selected_period = st.selectbox("Analiz Tipi / Dönem Seçin:", m_list, format_func=lambda x: x[0])
    
    sel_val = selected_period[1]

    # Veriyi Döneme Göre Filtreleme
    if sel_val == "GENEL":
        f_df25, f_df26 = df_2025.copy(), df_2026.copy()
        period_label = "Tüm Dönem (Genel)"
    else:
        f_df25 = mon_2025[mon_2025["Ay"] == sel_val] if not mon_2025.empty else pd.DataFrame()
        f_df26 = mon_2026[mon_2026["Ay"] == sel_val] if not mon_2026.empty else pd.DataFrame()
        period_label = f"{TR_AY[sel_val]} Ayı"

    # Merge İşlemleri (İki Dönemi Yan Yana Getirme)
    a = f_df25[["Anahtar", "Müşteri", "Ciro", "Palet"]].rename(columns={"Müşteri": "M25", "Ciro": "Ciro_2025", "Palet": "Palet_2025"}) if not f_df25.empty else pd.DataFrame(columns=["Anahtar", "M25", "Ciro_2025", "Palet_2025"])
    b = f_df26[["Anahtar", "Müşteri", "Ciro", "Palet"]].rename(columns={"Müşteri": "M26", "Ciro": "Ciro_2026", "Palet": "Palet_2026"}) if not f_df26.empty else pd.DataFrame(columns=["Anahtar", "M26", "Ciro_2026", "Palet_2026"])
    
    cmp = pd.merge(b, a, on="Anahtar", how="outer")
    cmp["Müşteri"] = cmp["M26"].fillna(cmp["M25"])
    for col in ["Ciro_2025", "Ciro_2026", "Palet_2025", "Palet_2026"]:
        cmp[col] = cmp[col].fillna(0.0)

    # Fark Hesapları
    cmp["Ciro Farkı"] = cmp["Ciro_2026"] - cmp["Ciro_2025"]
    cmp["Palet Farkı"] = cmp["Palet_2026"] - cmp["Palet_2025"]
    
    def calc_status(row):
        if row["Ciro_2025"] == 0 and row["Ciro_2026"] > 0:
            return "Yeni Müşteri 🟢"
        elif row["Ciro_2026"] == 0 and row["Ciro_2025"] > 0:
            return "Kayıp / Terk 🔴"
        elif row["Ciro Farkı"] > 0:
            return "Artış Var 🟢"
        elif row["Ciro Farkı"] < 0:
            return "Düşüş Var 🔴"
        return "Değişim Yok ⚪"

    cmp["Durum"] = cmp.apply(calc_status, axis=1)

    # ÖZET KPİ BANTLARI
    t25, t26 = cmp["Ciro_2025"].sum(), cmp["Ciro_2026"].sum()
    p25, p26 = cmp["Palet_2025"].sum(), cmp["Palet_2026"].sum()
    c_diff = t26 - t25
    c_pct = (c_diff / t25 * 100) if t25 > 0 else 0

    st.markdown(f"### 📍 Seçilen Dönem: **{period_label}**")
    
    kpis([
        ("2025 Dönem Ciro", fmt_tl(t25), GREY),
        ("2026 Dönem Ciro", fmt_tl(t26), BLUE, f"{'+' if c_diff>=0 else ''}{fmt_tl(c_diff)} (%{tr_num(c_pct, 1)})", GREEN if c_diff>=0 else RED),
        ("2025 Dönem Palet", fmt_palet(p25), GREY),
        ("2026 Dönem Palet", fmt_palet(p26), ORANGE, f"{'+' if p26-p25>=0 else ''}{fmt_palet(p26-p25)}", GREEN if p26-p25>=0 else RED)
    ])

    st.markdown("---")

    # MÜŞTERİ BAZLI YORUM VE TESPİT PANELİ (Belirgin Düşüşler)
    st.subheader("🤖 Müşteri Bazlı Özel Yorum ve Tespitler")
    
    drop_customers = cmp[cmp["Ciro Farkı"] < 0].sort_values("Ciro Farkı", ascending=True)
    rise_customers = cmp[cmp["Ciro Farkı"] > 0].sort_values("Ciro Farkı", ascending=False)
    
    notes = []
    if not drop_customers.empty:
        for _, r in drop_customers.head(4).iterrows():
            pct = (abs(r["Ciro Farkı"]) / r["Ciro_2025"] * 100) if r["Ciro_2025"] > 0 else 0
            notes.append(f"* 🔴 **{r['Müşteri']}**: {period_label} döneminde cirosu **{fmt_tl(r['Ciro_2025'])}** seviyesinden **{fmt_tl(r['Ciro_2026'])}** seviyesine düştü. (Net Kayıp: **{fmt_tl(abs(r['Ciro Farkı']))}**, %{tr_num(pct, 1)} gerileme).")
    
    if not rise_customers.empty:
        for _, r in rise_customers.head(3).iterrows():
            pct = (r["Ciro Farkı"] / r["Ciro_2025"] * 100) if r["Ciro_2025"] > 0 else 0
            notes.append(f"* 🟢 **{r['Müşteri']}**: {period_label} döneminde cirosunu **{fmt_tl(r['Ciro_2025'])}** seviyesinden **{fmt_tl(r['Ciro_2026'])}** seviyesine çıkardı. (Net Artış: **{fmt_tl(r['Ciro Farkı'])}**).")

    if notes:
        st.info("\n".join(notes))
    else:
        st.caption("Dönemsel karşılaştırmada belirgin bir değişim tespit edilemedi.")

    # TABLO GÖSTERİMİ
    st.subheader("📋 Yan Yana Karşılaştırmalı Müşteri Tablosu")
    st.caption("Tablo sütun başlıklarına tıklayarak sıralama yapabilirsiniz. Tüm rakamlar okunabilir formattadır.")

    # Formatlama
    disp_df = cmp[["Müşteri", "Ciro_2025", "Ciro_2026", "Ciro Farkı", "Palet_2025", "Palet_2026", "Palet Farkı", "Durum"]].copy()
    disp_df = disp_df.sort_values("Ciro_2026", ascending=False).reset_index(drop=True)

    st.dataframe(
        disp_df.style.format({
            "Ciro_2025": lambda v: fmt_tl(v),
            "Ciro_2026": lambda v: fmt_tl(v),
            "Ciro Farkı": lambda v: f"{'+' if v>0 else ''}{fmt_tl(v)}",
            "Palet_2025": lambda v: fmt_palet(v),
            "Palet_2026": lambda v: fmt_palet(v),
            "Palet Farkı": lambda v: f"{'+' if v>0 else ''}{fmt_palet(v)}",
        }),
        use_container_width=True
    )

# -------------------------------------------------------------------
# MOD 3: 2027 HEDEFLERİ
# -------------------------------------------------------------------
elif mode == MODES[2]:
    st.header("🚀 2027 Yılı Hedef ve Bütçe Planlama")
    st.caption("Mevcut veriler ve hedef müşteri potansiyeline göre 2027 bütçelemesi.")

    unit_price = st.number_input("2027 Tahmini Palet Başı Gelir (₺):", min_value=500, max_value=20000, value=2500, step=100)
    
    rows = []
    for title, list_items in SEGMENTS:
        for nm, p in list_items:
            rows.append({"Firma": nm, "Segment": title, "Hedef Palet": p, "Tahmini Ciro": p * unit_price})
    
    df_tgt = pd.DataFrame(rows)
    
    tot_p = df_tgt["Hedef Palet"].sum()
    tot_c = df_tgt["Tahmini Ciro"].sum()

    kpis([
        ("Hedeflenen Toplam Ciro", fmt_tl(tot_c), BLUE),
        ("Hedeflenen Toplam Palet", fmt_palet(tot_p), ORANGE),
        ("Ortalama Palet Başı Gelir", fmt_tl(unit_price), GREEN)
    ])

    st.markdown("---")
    st.subheader("📋 2027 Potansiyel Hedef Müşteri Listesi")
    
    st.dataframe(
        df_tgt.style.format({
            "Hedef Palet": lambda v: fmt_palet(v),
            "Tahmini Ciro": lambda v: fmt_tl(v)
        }),
        use_container_width=True
    )
