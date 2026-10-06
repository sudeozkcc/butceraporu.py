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
# STİL & RENK PALETLERİ
# ===================================================================
BLUE, ORANGE, GREY = "#1F4E9C", "#E07B00", "#9AA5B1"
COLOR_2025 = "#FF8C00"  # Canlı Turuncu
COLOR_2026 = "#1E88E5"  # Canlı Mavi

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
# YARDIMCI FONKSİYONLAR
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

def names_match(target_name, excel_name) -> bool:
    e = compact(excel_name)
    if len(e) < 3:
        return False
    for part in str(target_name).split("/"):
        c = compact(part)
        if len(c) >= 3 and (c in e or e in c):
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
    if x >= 1e6:
        t = tr_num(x / 1e6, 2)
        t = t[:-3] if t.endswith(",00") else (t[:-1] if t.endswith("0") else t)
        return t + " M ₺"
    if x >= 1e3:
        return tr_num(x / 1e3, 0) + " B ₺"
    return fmt_tl(x)

def fmt_palet(x) -> str:
    return tr_num(x, dec=0)

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

def shade(values, rgb, lo=0.03, hi=0.22):
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
            out.append(mix((46, 160, 67), 0.08 + 0.20 * abs(x) / m))
        else:
            out.append(mix((214, 39, 40), 0.08 + 0.20 * abs(x) / m))
    return out

# ===================================================================
# GRAFİK & TABLO RENDERING
# ===================================================================
def render_table(headers, cols, fills, widths=None, font_size=12):
    n = len(cols[0]) if cols else 0
    fig = go.Figure(go.Table(
        columnwidth=widths,
        header=dict(values=[f"<b>{h}</b>" for h in headers], fill_color=BLUE,
                    font=dict(color="white", size=13), align="center", height=36, line_color="white"),
        cells=dict(values=cols, fill_color=fills, align=["left"] + ["right"] * (len(cols) - 1),
                   font=dict(size=font_size, color="#1a202c"), height=28, line_color="white"),
    ))
    fig.update_layout(margin=dict(l=0, r=0, t=6, b=6), height=min(70 + 31 * n, 1100))
    show(fig)

def ciro_palet_chart(top, avg_ptl=0.0):
    d = top.sort_values("Ciro", ascending=False)
    names = d["Müşteri"].tolist()
    n = len(names)
    fig = make_subplots(rows=1, cols=3, shared_yaxes=True, horizontal_spacing=0.015,
                        column_widths=[0.45, 0.30, 0.25],
                        subplot_titles=("<b>TOPLAM CİRO (₺)</b>", "<b>PALET (adet)</b>", "<b>PALET BAŞI GELİR (₺)</b>"))
    specs = [(1, "Ciro", BLUE, lambda v: f"<b>{tr_num(v)} ₺</b>"),
             (2, "Palet", ORANGE, lambda v: f"<b>{fmt_palet(v)}</b>")]
    for col, key, color, fmt in specs:
        vals = d[key].tolist()
        mx = max(vals) if vals and max(vals) > 0 else 1.0
        tx, ty, fx, fy = [], [], [], []
        for nm, v in zip(names, vals):
            tx += [0, mx, None]
            ty += [nm, nm, None]
            fx += [0, v, None]
            fy += [nm, nm, None]
        fig.add_trace(go.Scatter(x=tx, y=ty, mode="lines", line=dict(color="#e3e8ef", width=10),
                                 hoverinfo="skip", showlegend=False), row=1, col=col)
        fig.add_trace(go.Scatter(x=fx, y=fy, mode="lines", line=dict(color=color, width=10),
                                 hoverinfo="skip", showlegend=False), row=1, col=col)
        hover = [f"<b>{r['Müşteri']}</b><br>Ciro: {tr_num(r['Ciro'])} ₺<br>Palet: {fmt_palet(r['Palet'])}"
                 f"<br>Palet başı: {tr_num(r['Palet Başı TL'])} ₺" for _, r in d.iterrows()]
        fig.add_trace(go.Scatter(x=vals, y=names, mode="markers", showlegend=False,
                                 marker=dict(size=19, color=color, line=dict(color="white", width=3)),
                                 hovertext=hover, hoverinfo="text"), row=1, col=col)
        fig.add_trace(go.Scatter(x=[mx] * n, y=names, mode="text", showlegend=False, hoverinfo="skip",
                                 text=[fmt(v) for v in vals], textposition="middle right",
                                 textfont=dict(size=14, color=color)), row=1, col=col)
        fig.update_xaxes(range=[0, mx * 1.7], showticklabels=False, showgrid=False, zeroline=False, row=1, col=col)
    ptl = d["Palet Başı TL"].tolist()
    dots = ["#2E7D32" if (v >= avg_ptl) else "#C62828" for v in ptl]
    fig.add_trace(go.Scatter(x=[0] * n, y=names, mode="markers+text", showlegend=False,
                             marker=dict(size=11, color=dots), textposition="middle right",
                             text=[f"  <b>{tr_num(v)} ₺</b>" for v in ptl], textfont=dict(size=14, color="#1f2937"),
                             hoverinfo="skip"), row=1, col=3)
    fig.update_xaxes(range=[-0.05, 1], showticklabels=False, showgrid=False, zeroline=False, row=1, col=3)
    fig.update_yaxes(type="category", categoryorder="array", categoryarray=names,
                     autorange="reversed", automargin=True, showgrid=False)
    fig.update_annotations(font_size=12)
    fig.update_layout(template="plotly_white", separators=",.", height=120 + 40 * n,
                      margin=dict(l=10, r=10, t=50, b=10))
    return fig

def pie_fig(df, col, title, top_n=10):
    d = df[["Müşteri", col]].sort_values(col, ascending=False)
    d = d[d[col] > 0]
    head = d.head(top_n)
    rest = d.iloc[top_n:][col].sum()
    labels, values = head["Müşteri"].tolist(), head[col].tolist()
    colors = BLUE_SEQ[: len(labels)]
    if rest > 0:
        labels.append("Diğer")
        values.append(rest)
        colors.append(GREY)
    fig = go.Figure(go.Pie(
        labels=labels, values=values, hole=0.5, sort=False,
        marker=dict(colors=colors, line=dict(color="white", width=2)),
        textinfo="percent", textposition="inside",
        hovertemplate="<b>%{label}</b><br>%{value:,.0f}<br>%{percent}<extra></extra>",
    ))
    fig.update_layout(
        template="plotly_white", separators=",.", height=460,
        title=dict(text=f"<b>{title}</b>", x=0.5),
        annotations=[dict(text=f"<b>{tr_num(sum(values))}</b>", x=0.5, y=0.5, showarrow=False, font=dict(size=14))],
        legend=dict(orientation="h", y=-0.12, font=dict(size=10)),
        margin=dict(l=10, r=10, t=60, b=10),
    )
    return fig

def dumbbell_fig(d, c25, c26, title, unit):
    """
    Daha renkli, canlı ve genişletilmiş (yazıların kesilmediği) Nokta/Dumbbell Grafiği.
    """
    names = d["Müşteri"].tolist()
    fmt = (lambda v: tr_num(v) + " ₺") if unit == "₺" else fmt_palet
    
    lx, ly = [], []
    line_colors = []
    
    for n, a, b in zip(names, d[c25], d[c26]):
        lx += [a, b, None]
        ly += [n, n, None]
        # Artışa veya düşüşe göre çizgi rengi belirleme
        if b >= a:
            line_colors.append("#4CAF50") # Yeşil
        else:
            line_colors.append("#EF5350") # Kırmızı

    # Maksimum değer tespiti ve genişletilmiş eksen aralığı (Metin kesilmesini önler)
    xmax = max(float(d[[c25, c26]].max().max()), 1.0)
    
    fig = go.Figure()

    # Aradaki bağlantı çizgileri (Renkli)
    for i, (n, a, b) in enumerate(zip(names, d[c25], d[c26])):
        color = "#2E7D32" if b >= a else "#C62828"
        fig.add_trace(go.Scatter(
            x=[a, b], y=[n, n], mode="lines",
            line=dict(color=color, width=3.5),
            hoverinfo="skip", showlegend=False
        ))

    # 2025 Noktaları (Canlı Turuncu)
    fig.add_trace(go.Scatter(
        x=d[c25], y=names, mode="markers", name="2025",
        marker=dict(size=13, color=COLOR_2025, line=dict(color="white", width=2)),
        hovertemplate="<b>%{y}</b><br>2025: %{x:,.0f}<extra></extra>"
    ))

    # 2026 Noktaları (Canlı Mavi)
    fig.add_trace(go.Scatter(
        x=d[c26], y=names, mode="markers", name="2026",
        marker=dict(size=13, color=COLOR_2026, line=dict(color="white", width=2)),
        hovertemplate="<b>%{y}</b><br>2026: %{x:,.0f}<extra></extra>"
    ))

    # Sağ Taraf Değer Metinleri (Metin taşmasını önlemek için eksen aralığı xmax * 1.55 yapıldı)
    fig.add_trace(go.Scatter(
        x=d[[c25, c26]].max(axis=1), y=names, mode="text", showlegend=False, hoverinfo="skip",
        text=[f" <b>2025:</b> {fmt(a)} → <b>2026:</b> {fmt(b)}" for a, b in zip(d[c25], d[c26])],
        textposition="middle right", textfont=dict(size=12, color="#1e293b")))

    fig.update_layout(
        template="plotly_white", separators=",.", height=130 + 38 * len(names),
        xaxis=dict(range=[0, xmax * 1.55], showticklabels=False, showgrid=False, zeroline=False),
        yaxis=dict(type="category", categoryorder="array", categoryarray=names, autorange="reversed",
                   automargin=True, showgrid=False),
        legend=dict(orientation="h", y=1.03, x=0.5, xanchor="center", font=dict(size=12, color="#1f2937")),
        margin=dict(l=10, r=80, t=40, b=10)) # Sağ marjin genişletildi
    return fig

def monthly_fig(info):
    tbm, metric = info["tbm"], info["metric"]
    fv = fmt_palet if metric == "Palet" else fmt_tl
    fig = go.Figure(go.Bar(x=[TR_AY[m] for m in tbm.index], y=tbm.values, marker_color=BLUE,
                           text=[fv(v) for v in tbm.values], textposition="outside"))
    fig.update_layout(template="plotly_white", separators=",.", height=380,
                      title=dict(text=f"<b>Aylık Toplam {metric}</b>", x=0.5),
                      yaxis=dict(visible=False, range=[0, float(tbm.max()) * 1.2]),
                      margin=dict(l=10, r=10, t=60, b=10))
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
        start = st.number_input("Veri başlangıç satırı", min_value=1, max_value=max(len(raw), 1),
                                value=lay["start"] + 1, key=f"st_{uid}") - 1

    df, xl_tot = build_df(raw, cust, ciro_cols, palet_cols, int(start))
    if df.empty:
        st.sidebar.error(f"{year}: veri okunamadı.")
        return empty, empty_m

    mon = build_monthly(raw, cust, lay["ciro_m"], lay["palet_m"], int(start))
    st.sidebar.caption(
        f"✅ {year}: {len(df)} müşteri · Ciro {fmt_tl(df['Ciro'].sum())} · Palet {fmt_palet(df['Palet'].sum())}"
    )
    return df, mon
