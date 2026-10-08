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
# STİL (sade: mavi ana renk, turuncu vurgu, yeşil/kırmızı sadece artış/düşüş)
# ===================================================================
BLUE, ORANGE, GREY = "#1F4E9C", "#E07B00", "#9AA5B1"
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
.pro-tbl td.wrap{white-space:normal;min-width:240px;}
.pro-tbl.compact th{padding:5px 8px;}
.pro-tbl.compact td{padding:3px 8px;}
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

_GENERIC = {"GIDA", "SAN", "TIC", "VE", "AS", "ANONIM", "SIRKETI", "LTD", "STI", "TURKIYE", "SANAYI",
            "TICARET", "FOOD", "FOODS", "GRUP", "HOLDING", "A", "S", "ISLETMELERI", "PAZARLAMA"}

def names_match(target_name, excel_name) -> bool:
    """Hedef listedeki firma adı ile Excel'deki müşteri adını (boşluk/noktalama/sıra farkı olmadan) eşleştirir.
    Son kural: firmanın ayırt edici sözcükleri (GIDA, FOOD, A.Ş. gibi genel sözcükler hariç) Excel adında
    tam sözcük olarak geçiyorsa eşleşir (ör. 'PEK FOOD' ↔ 'PEK GIDA SAN. A.Ş.')."""
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

def growth_candidates(df25, df26, exclude, k=5):
    """2025'e göre cirosu artmış, 2026'da ilk 25 içinde yer alan mevcut müşterilerin anahtarları (artış tutarına göre)."""
    if df25 is None or df25.empty or df26.empty:
        return []
    m = df26.merge(df25[["Anahtar", "Ciro"]], on="Anahtar", suffixes=("", "_25"))
    m = m[(m["Ciro_25"] > 0) & (m["Ciro"] > m["Ciro_25"])]
    top = set(df26.sort_values("Ciro", ascending=False).head(25)["Anahtar"])
    m = m[m["Anahtar"].isin(top)].assign(Fark=lambda x: x["Ciro"] - x["Ciro_25"]).sort_values("Fark", ascending=False)
    return [x for x in m["Anahtar"] if x not in exclude][:k]

def to_num(v) -> float:
    if isinstance(v, (int, float, np.number)):
        return 0.0 if pd.isna(v) else float(v)
    if v is None:
        return 0.0
    s = str(v).replace("₺", "").replace("TL", "").replace("\xa0", "").replace(" ", "").strip()
    if s in ("", "-", "nan", "NaN"):
        return 0.0
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
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
# GRAFİK & TABLO
# ===================================================================
def _auto_widths(headers, cols):
    out = []
    for h, c in zip(headers, cols):
        longest = max([len(re.sub(r"<[^>]+>", "", str(x))) for x in c] + [len(str(h)) * 0.9])
        out.append(max(longest, 4))
    return out

def render_table(headers, cols, fills, widths=None, font_size=12, font_colors=None, header_fill=BLUE, left_cols=(0,), compact=False):
    """HTML tablo: hücreler tek satırda kalır, uzun rakamlar (ör. 84.123.456 ₺) asla kesilmez; geniş tablo kaydırılır."""
    n = len(cols[0]) if cols else 0
    esc = lambda x: str(x).replace("&", "&amp;")
    th = "".join(f'<th style="background:{header_fill};font-size:{font_size + 1}px;">{esc(h)}</th>' for h in headers)
    body = []
    for i in range(n):
        tds = []
        for j, c in enumerate(cols):
            fill = fills[j][i] if isinstance(fills[j], (list, tuple, np.ndarray)) else fills[j]
            if font_colors is None:
                fc = "#1a202c"
            elif isinstance(font_colors[j], (list, tuple, np.ndarray)):
                fc = font_colors[j][i]
            else:
                fc = font_colors[j]
            al = "left" if j in left_cols else "right"
            cls = ' class="wrap"' if (j in left_cols and j != 0) else ""
            tds.append(f'<td{cls} style="background:{fill};color:{fc};text-align:{al};font-size:{font_size}px;">{esc(c[i])}</td>')
        body.append("<tr>" + "".join(tds) + "</tr>")
    html = (f'<div class="pro-wrap"><table class="pro-tbl{" compact" if compact else ""}"><thead><tr>{th}</tr></thead>'
            f'<tbody>{"".join(body)}</tbody></table></div>')
    st.markdown(html, unsafe_allow_html=True)

# ----- renkli kartlar ve başlık bantları -----
def rgba(color, a):
    r, g, b = hex_rgb(color)
    return f"rgba({r},{g},{b},{a})"

def kpis(items):
    """items: (etiket, değer, renk, alt_metin, alt_renk). HTML kart: uzun rakamlar kesilmez."""
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

def shade_rows(values, colors, lo=0.06, hi=0.45):
    """Her satırı kendi grubunun rengiyle, değer büyüklüğüne göre boyar."""
    v = np.nan_to_num(np.array(values, dtype=float))
    m = np.abs(v).max() if len(v) else 0
    return [mix(hex_rgb(c), (lo + (hi - lo) * abs(x) / m) if m else lo) for x, c in zip(v, colors)]

def ciro_palet_chart(top, avg_ptl=0.0):
    """Slider tarzı grafik: uzun gri hat = en büyük değer, mavi/turuncu dolgu = müşterinin değeri.
    Sağda toplam ciro, palet ve doğrudan palet başı gelir yazılıdır."""
    d = top.sort_values("Ciro", ascending=False)
    names = d["Müşteri"].tolist()
    n = len(names)
    fig = make_subplots(rows=1, cols=3, shared_yaxes=True, horizontal_spacing=0.015,
                        column_widths=[0.52, 0.25, 0.23],
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
                                 text=["     " + fmt(v) for v in vals], textposition="middle right",
                                 cliponaxis=False, textfont=dict(size=13, color=color)), row=1, col=col)
        fig.update_xaxes(range=[0, mx * (2.5 if col == 1 else 2.1)], showticklabels=False, showgrid=False,
                         zeroline=False, row=1, col=col)
    ptl = d["Palet Başı TL"].tolist()
    dots = ["#2E7D32" if (v >= avg_ptl) else "#C62828" for v in ptl]
    fig.add_trace(go.Scatter(x=[0] * n, y=names, mode="markers+text", showlegend=False,
                             marker=dict(size=11, color=dots), textposition="middle right",
                             text=[f"  <b>{tr_num(v)} ₺</b>" for v in ptl], textfont=dict(size=13, color="#1f2937"),
                             cliponaxis=False, hoverinfo="skip"), row=1, col=3)
    fig.update_xaxes(range=[-0.05, 1], showticklabels=False, showgrid=False, zeroline=False, row=1, col=3)
    fig.update_yaxes(type="category", categoryorder="array", categoryarray=names,
                     autorange="reversed", automargin=True, showgrid=False)
    fig.update_annotations(font_size=12)
    fig.update_layout(template="plotly_white", separators=",.", height=120 + 40 * n,
                      margin=dict(l=10, r=40, t=50, b=10))
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
        annotations=[dict(text=f"<b>{tr_num(sum(values))}</b>", x=0.5, y=0.5, showarrow=False,
                          font=dict(size=14 if len(tr_num(sum(values))) <= 11 else 11))],
        legend=dict(orientation="h", y=-0.12, font=dict(size=10)),
        margin=dict(l=10, r=10, t=60, b=10),
    )
    return fig

def change_chart(d):
    """2025 → 2026: gri çubuk 2025, renkli çubuk 2026 (aynı ölçek, sıfırdan başlar → tam orantılı).
    Sağ panel: ciro değişim yüzdesi (yeşil artış / kırmızı düşüş, ±%150'de kırpılır; etiket gerçek değeri gösterir)."""
    names = d["Müşteri"].tolist()
    n = len(names)
    UP, DOWN, NEW = "#2E7D32", "#C62828", "#2A9D8F"
    CAP = 150.0
    fig = make_subplots(rows=1, cols=3, shared_yaxes=True, horizontal_spacing=0.02,
                        column_widths=[0.38, 0.30, 0.32],
                        subplot_titles=("<b>CİRO (₺)</b>", "<b>PALET (adet)</b>", "<b>CİRO DEĞİŞİMİ</b>"))
    specs = [(1, "Ciro_2025", "Ciro_2026", BLUE, lambda v: fmt_short_tl(v), "Ciro"),
             (2, "Palet_2025", "Palet_2026", ORANGE, lambda v: fmt_palet(v), "Palet")]
    for col, c25, c26, color, fmt, lab in specs:
        a, b = d[c25].tolist(), d[c26].tolist()
        mx = max(a + b + [1.0])
        fig.add_trace(go.Bar(y=names, x=a, orientation="h", name="2025", marker_color="#C9D1DB",
                             text=[fmt(v) if v > 0 else "" for v in a], textposition="outside", cliponaxis=False,
                             textfont=dict(size=11, color="#6b7280"),
                             hovertext=[f"<b>{nm}</b><br>2025 {lab}: {tr_num(v)}" for nm, v in zip(names, a)],
                             hoverinfo="text", showlegend=False), row=1, col=col)
        fig.add_trace(go.Bar(y=names, x=b, orientation="h", name="2026", marker_color=color,
                             text=[f"<b>{fmt(v)}</b>" if v > 0 else "" for v in b], textposition="outside", cliponaxis=False,
                             textfont=dict(size=12, color=color),
                             hovertext=[f"<b>{nm}</b><br>2026 {lab}: {tr_num(v)}" for nm, v in zip(names, b)],
                             hoverinfo="text", showlegend=False), row=1, col=col)
        fig.update_xaxes(range=[0, mx * 1.55], showticklabels=False, showgrid=False, zeroline=False, row=1, col=col)

    def pct_val(x, y):
        return None if x <= 0 else (y - x) / x * 100

    xs, cols_, txt = [], [], []
    for _, r in d.iterrows():
        v = pct_val(r["Ciro_2025"], r["Ciro_2026"])
        pv = pct_val(r["Palet_2025"], r["Palet_2026"])
        ptxt = "Yeni" if pv is None and r["Palet_2026"] > 0 else ("—" if pv is None else f"{'+' if pv >= 0 else '−'}{tr_num(abs(pv), 0)} %")
        if v is None:
            xs.append(CAP if r["Ciro_2026"] > 0 else 0)
            cols_.append(NEW)
            txt.append(f"<b>Yeni</b> · palet {ptxt}")
        else:
            xs.append(max(-CAP, min(CAP, v)))
            cols_.append(UP if v >= 0 else DOWN)
            txt.append(f"<b>{'+' if v >= 0 else '−'}{tr_num(abs(v), 0)} %</b> · palet {ptxt}")
    fig.add_trace(go.Bar(y=names, x=xs, orientation="h", marker_color=cols_, text=txt, textposition="outside",
                         cliponaxis=False, textfont=dict(size=12, color="#1f2937"), hoverinfo="skip",
                         showlegend=False), row=1, col=3)
    fig.update_xaxes(range=[-CAP * 1.1, CAP * 2.6], showticklabels=False, showgrid=False, zeroline=True,
                     zerolinecolor="#9AA5B1", row=1, col=3)
    fig.update_yaxes(type="category", categoryorder="array", categoryarray=names,
                     autorange="reversed", automargin=True, showgrid=False)
    fig.update_annotations(font_size=12)
    fig.update_layout(template="plotly_white", separators=",.", barmode="group", bargap=0.35, bargroupgap=0.05,
                      height=120 + 46 * n, margin=dict(l=10, r=10, t=50, b=10))
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
    """Ay başlıklarını sütunlara eşler (birleştirilmiş hücreler için sağa doğru taşır)."""
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
            idx = [j for j, m in enumerate(best) if m]
            gaps = [b - a for a, b in zip(idx, idx[1:])]
            stride = max(int(np.median(gaps)), 1) if gaps else 3   # bir ayın kapladığı sütun sayısı
            cur, since = None, 0
            for j, m in enumerate(best):
                if m:
                    cur, since = m, 0
                elif cur is not None:
                    since += 1
                # Son ayın sağındaki sütunlar (toplam, tahmin, geçen yıl vb.) son aya yazılmasın
                inherited[j] = cur if (cur is not None and since < stride) else None
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
        ciro_cols = st.multiselect("Ciro sütun(lar)ı (seçilenler toplanır)", opts,
                                   default=lay["ciro"], format_func=fmt, key=f"ci_{uid}")
        palet_cols = st.multiselect("Palet sütun(lar)ı (seçilenler toplanır)", opts,
                                    default=lay["palet"], format_func=fmt, key=f"pa_{uid}")
        start = st.number_input("Veri başlangıç satırı (Excel satır no)", min_value=1,
                                max_value=max(len(raw), 1), value=lay["start"] + 1, key=f"st_{uid}") - 1

    df, xl_tot = build_df(raw, cust, ciro_cols, palet_cols, int(start))
    if df.empty:
        st.sidebar.error(f"{year}: veri okunamadı. '⚙️ sütun eşleştirme' bölümünden sütunları seçin.")
        return empty, empty_m

    mon = build_monthly(raw, cust, lay["ciro_m"], lay["palet_m"], int(start))

    # Henüz gelmemiş aylar (ör. 2026'da Kasım/Aralık) veri içerse bile hesaba katılmaz
    today = datetime.date.today()
    if not mon.empty and year >= today.year:
        last_ok = today.month if year == today.year else 0
        fut = mon[mon["Ay"] > last_ok]
        fut = fut[(fut["Ciro"] != 0) | (fut["Palet"] != 0)]
        if not fut.empty:
            sub_f = fut.groupby("Anahtar")[["Ciro", "Palet"]].sum()
            df = df.merge(sub_f, left_on="Anahtar", right_index=True, how="left", suffixes=("", "_f"))
            df["Ciro"] = (df["Ciro"] - df["Ciro_f"].fillna(0.0)).clip(lower=0.0)
            df["Palet"] = (df["Palet"] - df["Palet_f"].fillna(0.0)).clip(lower=0.0)
            df = df.drop(columns=["Ciro_f", "Palet_f"])
            df = df[(df["Ciro"] != 0) | (df["Palet"] != 0)].reset_index(drop=True)
            ay_txt = ", ".join(TR_AY[m] for m in sorted(fut["Ay"].unique()))
            st.sidebar.caption(f"🕒 {year}: henüz gelmemiş ay(lar) ({ay_txt}) hesaptan çıkarıldı "
                               f"({fmt_tl(fut['Ciro'].sum())} · {fmt_palet(fut['Palet'].sum())} palet).")
        mon = mon[mon["Ay"] <= last_ok].reset_index(drop=True)
        if df.empty:
            return empty, empty_m
    st.sidebar.caption(f"✅ {year}: {len(df)} müşteri · Ciro {fmt_tl(df['Ciro'].sum())} · Palet {fmt_palet(df['Palet'].sum())}")
    if mon.empty:
        st.sidebar.caption(f"ℹ️ {year}: aylık sütun bulunamadı (aylık ivme yorumları üretilemez).")
    else:
        ms = sorted(mon["Ay"].unique())
        st.sidebar.caption(f"📅 {year}: aylık veri {TR_AY[ms[0]]}–{TR_AY[ms[-1]]}")
    if len(ciro_cols) > 1:
        st.sidebar.warning(f"{year}: {len(ciro_cols)} ciro sütunu toplanıyor. Biri 'toplam' sütunuysa çift sayım olur; eşleştirmeyi kontrol edin.")
    if xl_tot and xl_tot[0] > 0 and abs(df["Ciro"].sum() - xl_tot[0]) / xl_tot[0] > 0.01:
        st.sidebar.warning(f"{year}: Excel'deki TOPLAM satırı {fmt_tl(xl_tot[0])}, okunan toplam {fmt_tl(df['Ciro'].sum())}. Sütun eşleştirmeyi kontrol edin.")
    return df, mon

# ===================================================================
# AYLIK İVME YORUMLARI
# ===================================================================
def momentum_notes(mon):
    """Aylık veriden: geç katılıp hızlı büyüyenler, ivmelenenler, gerileyenler ve genel trend."""
    if mon is None or mon.empty:
        return None
    metric = "Palet" if mon["Palet"].sum() > 0 else "Ciro"
    fv = fmt_palet if metric == "Palet" else fmt_tl
    piv = mon.pivot_table(index="Anahtar", columns="Ay", values=metric, aggfunc="sum", fill_value=0.0)
    names = mon.groupby("Anahtar")["Müşteri"].first()
    tbm = piv.sum()
    act = [m for m in piv.columns if tbm[m] > 0]
    if len(act) < 2:
        return None
    first_m, last_m = min(act), max(act)
    months = list(range(first_m, last_m + 1))
    piv = piv.reindex(columns=months, fill_value=0.0)
    tbm = piv.sum()
    tot = piv.sum(axis=1).sort_values(ascending=False)

    new, up, down = [], [], []
    for rank, k in enumerate(tot.index[:30], 1):
        s = piv.loc[k]
        nz = [m for m in months if s[m] > 0]
        if not nz:
            continue
        start = nz[0]
        span = last_m - start + 1
        total_k = float(s.sum())
        avg = total_k / span
        if start > first_m and span <= 8:
            new.append((k, total_k, start, span, avg, rank))
        elif start == first_m and len(months) >= 6:
            head, tail = float(s[months[:3]].mean()), float(s[months[-3:]].mean())
            if head > 0:
                ch = (tail - head) / head * 100
                if ch >= 40:
                    up.append((k, total_k, head, tail, ch, rank))
                elif ch <= -40:
                    down.append((k, total_k, head, tail, ch, rank))

    reach = (lambda t: f"{fmt_palet(t)} palete") if metric == "Palet" else (lambda t: f"{fmt_tl(t)} ciroya")
    rising, falling, riser_keys = [], [], []
    new_keys, up_keys = [], []
    for k, total_k, start, span, avg, rank in sorted(new, key=lambda x: -x[1])[:5]:
        span_txt = "ilk ayında" if span == 1 else f"yalnızca {span} ayda"
        rising.append(f"* **{names[k]}**, {TR_AY[start]} ayında portföye katılmasına rağmen {span_txt} "
                      f"{reach(total_k)} ulaştı (aylık ortalama {fv(avg)}). "
                      f"Yıl içinde hızlı ivme yakalayan isimlerden.")
        riser_keys.append(k)
        new_keys.append(k)
    for k, total_k, head, tail, ch, rank in sorted(up, key=lambda x: -x[4])[:5]:
        rising.append(f"* **{names[k]}**, yıl boyunca ivmelendi: ilk 3 aylık ortalaması {fv(head)} iken "
                      f"son 3 aylık ortalaması {fv(tail)} oldu (%{tr_num(ch, 0)} artış).")
        riser_keys.append(k)
        up_keys.append(k)
    for k, total_k, head, tail, ch, rank in sorted(down, key=lambda x: x[4])[:5]:
        falling.append(f"* **{names[k]}**: ilk 3 aylık ortalaması {fv(head)} iken son 3 aylık ortalaması "
                       f"{fv(tail)}'e geriledi (%{tr_num(abs(ch), 0)} düşüş). Yakından takip edilmeli.")

    general = []
    bm = tbm.idxmax()
    general.append(f"* **En Yoğun Ay:** {TR_AY[bm]} ayı, toplam {fv(tbm[bm])} {metric.lower()} ile yılın zirvesi oldu.")
    if len(months) >= 2 and tbm[months[-2]] > 0:
        ch = (tbm[last_m] - tbm[months[-2]]) / tbm[months[-2]] * 100
        general.append(f"* **Son Ay Trendi:** {TR_AY[last_m]} ayı, {TR_AY[months[-2]]} ayına göre "
                       f"%{tr_num(abs(ch), 1)} {'artış' if ch >= 0 else 'düşüş'} gösterdi "
                       f"(ay tamamlanmadıysa kısmi veri olabilir).")
    return dict(metric=metric, piv=piv, names=names, tbm=tbm, months=months, general=general,
                rising=rising, falling=falling, riser_keys=riser_keys,
                new_keys=new_keys, up_keys=up_keys)

def dormant_list(mon, min_silent=2):
    """Son `min_silent` aydır (veri içindeki son aya göre) hiç palet/ciro vermeyen müşteriler."""
    if mon is None or mon.empty:
        return {}, "Palet", None
    metric = "Palet" if mon["Palet"].sum() > 0 else "Ciro"
    piv = mon.pivot_table(index="Anahtar", columns="Ay", values=metric, aggfunc="sum", fill_value=0.0)
    names = mon.groupby("Anahtar")["Müşteri"].first()
    tbm = piv.sum()
    act = [m for m in piv.columns if tbm[m] > 0]
    if len(act) < 3:
        return {}, metric, None
    last_m = max(act)
    out = {}
    for k, row in piv.iterrows():
        nz = [m for m in piv.columns if row[m] > 0 and m <= last_m]
        if not nz:
            continue
        la = max(nz)
        if last_m - la >= min_silent:
            out[k] = dict(name=names[k], total=float(row.sum()), first=min(nz), last=la)
    return out, metric, last_m

def render_momentum(info):
    if info is None:
        st.caption("ℹ️ Excel'de aylık sütun bulunamadığı için aylık ivme yorumları üretilemedi. "
                   "Sol menüde 'sütun eşleştirme' bölümünü ve başlıklardaki ay adlarını kontrol edin.")
        return
    if info["rising"]:
        st.subheader("🚀 Yıl İçi Hızlı İvme Yakalayanlar")
        st.success("\n".join(info["rising"]))
    if info["falling"]:
        st.subheader("⚠️ Gerileyen Müşteriler")
        st.warning("\n".join(info["falling"]))
    st.subheader("📅 Aylık Seyir")
    st.info("\n".join(info["general"]))
    show(monthly_fig(info))

# ===================================================================
# HEDEF MÜŞTERİ VERİSİ
# ===================================================================
SEGMENTS = [
    ("1. SOĞUK – ET / TAVUK / ŞARKÜTERİ", [
        ("BANVİT", 2000), ("ŞENPİLİÇ", 2000), ("BEYPİ / BEYPİLİÇ", 1500), ("LEZİTA", 1500),
        ("NAMET GIDA", 1500), ("PINAR ET", 1200), ("POLONEZ", 1000), ("AYTAÇ", 1000),
        ("APİKOĞLU", 800), ("COŞKUN ET", 800), ("GEDİK TAVUK", 800), ("HASTAVUK", 800),
        ("KESKİNOĞLU", 600), ("CUMHURİYET SUCUKLARI", 500), ("BAŞYAZICI", 500),
        ("ŞAHİN SUCUKLARI", 500), ("SULTAN ET", 500), ("İKBAL GIDA", 500)]),
    ("2. SOĞUK – SÜT / PEYNİR / SÜT ÜRÜNLERİ", [
        ("SÜTAŞ", 2000), ("PINAR SÜT", 1500), ("AK GIDA / İÇİM", 1500), ("EKER GIDA", 1200),
        ("YÖRÜKOĞLU", 1000), ("MURATBEY", 1000), ("TAT GIDA", 1000), ("TEKSÜT", 800),
        ("YÖRSAN / MATLI GRUP", 800), ("SEK SÜT / SEKSÜT", 800), ("CEBECİ SÜT", 500),
        ("TARAKLI SÜT", 500), ("GÜNEŞOĞLU SÜT", 500), ("TEZCANLAR SÜT", 400),
        ("ÜNAL PEYNİRCİLİK", 400), ("HAN EGE SÜT", 300), ("KUZUCU SÜT", 300), ("BİRNİCİ SÜT", 300)]),
    ("3. DONUK / DONDURULMUŞ GIDA", [
        ("KEREVİTAŞ / SUPERFRESH", 2500), ("ÖZGÖRKEY GIDA / FEAST", 2000), ("PEK FOOD", 2000),
        ("PENGUEN GIDA", 1500), ("BONDUELLE TÜRKİYE", 1200), ("DARDANEL", 1000), ("BENMARK GIDA", 800),
        ("KEKSBAKERY", 600), ("AKTAŞLAR LEZZET GRUBU", 600), ("PARDO FOOD", 600), ("MAPEK", 500),
        ("GUSTO GIDA", 500), ("FİNE FOOD", 500), ("FRİGOPAK", 400), ("MANTİYE GIDA", 300),
        ("LEZZET MANTI", 300)]),
    ("4. ZİNCİR RESTORAN / FOOD SERVICE", [
        ("TAB GIDA", 3000), ("BIGCHEFS", 1500),
        ("GÜNAYDIN KÖFTE & DÖNER", 1000), ("KÖFTECİ RAMİZ", 800), ("KASAP DÖNER", 800),
        ("BURSA KEBAP EVİ", 600), ("KAHVE DÜNYASI", 600),
        ("ARABICA COFFEE", 400), ("DAVID PEOPLE COFFEE", 300), ("BODRUM MANTI", 300), ("SUSHICO", 300),
        ("KAYSERİ MUTFAĞI", 300), ("GREEN SALADS", 300), ("REALLY FRIED CHICKEN", 200),
        ("BIG BAKER", 200), ("ÇÖPS", 200)]),
    ("5. DİĞER – YENİ EKLENEN", [
        ("İSTANBUL SÜTLÜ TATLI", 500), ("KAYA ÇİFTLİĞİ", 500), ("THİO", 400), ("YURT KONSERVE", 400),
        ("UNIFO GIDA", 400), ("YELKEN GIDA", 400), ("PEYSAN", 400)]),
]

# Listeye mutlaka girecek firmalar (en büyük 10 sınırına takılmaz)
MUST_ADD = {"FİNE FOOD", "İSTANBUL SÜTLÜ TATLI", "KAYA ÇİFTLİĞİ", "THİO", "YURT KONSERVE",
            "UNIFO GIDA", "YELKEN GIDA", "PEYSAN"}

SEG_SHORT = ["Et/Tavuk", "Süt", "Donuk", "Restoran", "Diğer"]
TEAL = "#2A9D8F"

# ===================================================================
# ARAYÜZ
# ===================================================================
st.markdown('<div class="hero"><div class="t">🚚 Lojistik Ciro, Palet &amp; Bütçe Sunumu</div>'
            '<div class="s">Gerçekleşen performans · Yıllık karşılaştırma · 2027 hedefleri</div></div>', unsafe_allow_html=True)

st.sidebar.header("📁 Excel Dosya Yükleme")
df_2025, mon_2025 = load_year(2025)
df_2026, mon_2026 = load_year(2026)

MODES = ["📊 2026'da Ne Yaptık?", "⚔️ 2025'e Göre Ne Değişti?", "🚀 2027'de Ne Yapacağız?"]
mode = st.sidebar.radio("📌 Çalışma Modunu Seçin:", MODES)

# ===================================================================
# MOD 1
# ===================================================================
if mode == MODES[0]:
    st.header("📊 2026'da Ne Yaptık? (Ciro ve Performans Özeti)")
    if df_2026.empty:
        st.info("👈 Lütfen sol menüden 2026 yılı Excel dosyanızı yükleyin.")
        st.stop()

    base = df_2026.sort_values("Ciro", ascending=False).reset_index(drop=True)
    base["Palet Başı TL"] = base["Ciro"].div(base["Palet"].where(base["Palet"] > 0)).fillna(0.0)
    tot_ciro, tot_palet = base["Ciro"].sum(), base["Palet"].sum()
    avg_ptl = tot_ciro / tot_palet if tot_palet > 0 else 0.0
    top20 = base.head(20).copy()
    top20_share = top20["Ciro"].sum() / tot_ciro * 100 if tot_ciro else 0.0

    if tot_palet == 0:
        st.warning("⚠️ Palet toplamı 0 görünüyor. Sol menüde '⚙️ 2026 sütun eşleştirme' bölümünden palet sütununu seçin.")

    kpis([("Toplam Ciro", fmt_tl(tot_ciro), BLUE), ("Toplam Palet", fmt_palet(tot_palet), ORANGE),
          ("Palet Başı Ort. Gelir", fmt_tl(avg_ptl), TEAL), ("Müşteri Sayısı", fmt_palet(len(base)), "#8E6BBF"),
          ("Top 20'nin Ciro Payı", fmt_pct(top20_share), "#455A64")])

    st.markdown("---")
    st.subheader("📊 Müşteri Bazında Ciro ve Palet (Top 20)")
    st.caption(f"Her satır bir müşteri; gri hat en büyük müşteriyi, dolu kısım müşterinin payını gösterir. "
               f"Palet başı gelirde 🟢 ortalamanın ({fmt_tl(avg_ptl)}) üstü, 🔴 altı.")
    show(ciro_palet_chart(top20, avg_ptl))

    st.subheader("🥧 Müşteri Payları (İlk 10 + Diğer)")
    p1, p2 = st.columns(2)
    with p1:
        show(pie_fig(base, "Ciro", "Ciro Dağılımı (₺)"))
    with p2:
        show(pie_fig(base, "Palet", "Palet Dağılımı (adet)"))

    st.subheader("📋 Ciro & Palet Tablosu (Top 20)")
    t = top20.copy()
    t["Pay"] = t["Ciro"] / tot_ciro * 100 if tot_ciro else 0.0
    n_t = len(t)
    zebra = ["rgb(255,255,255)" if i % 2 == 0 else "rgb(246,248,251)" for i in range(n_t)]
    dark = ["#1a202c"] * n_t
    render_table(
        ["#", "Müşteri Unvanı", "Ciro", "Ciro Payı", "Palet", "Palet Başı TL"],
        [[str(i + 1) for i in range(n_t)], t["Müşteri"].tolist(),
         [fmt_tl(v) for v in t["Ciro"]], [fmt_pct(v) for v in t["Pay"]],
         [fmt_palet(v) for v in t["Palet"]], [fmt_tl(v) for v in t["Palet Başı TL"]]],
        [zebra] * 6,
        font_colors=[dark] * 6,
        header_fill="#374151",
    )

    st.subheader("💡 2026 Performans Yorumlarımız")
    notes = [f"* **Genel Verimlilik:** 2026 genelinde palet başına ortalama **{fmt_tl(avg_ptl)}** ciro elde edilmiştir.",
             f"* **Yoğunlaşma:** İlk 20 müşteri toplam cironun **{fmt_pct(top20_share)}** kadarını oluşturuyor."]
    valid = top20[top20["Palet"] > 0]
    if not valid.empty:
        star = valid.sort_values("Palet Başı TL", ascending=False).iloc[0]
        notes.append(f"* **En Verimli Müşteri:** **{star['Müşteri']}**, palet başına **{fmt_tl(star['Palet Başı TL'])}** ile Top 20 içinde en yüksek verimliliği yakalamıştır.")
        low = valid[valid["Palet Başı TL"] < avg_ptl * 0.85]
        if not low.empty:
            nm = ", ".join(low.sort_values("Palet Başı TL").head(3)["Müşteri"])
            notes.append(f"* **Dikkat Edilecekler:** {nm} palet başı geliri ortalamanın belirgin altında; fiyatlama/hacim gözden geçirilebilir.")
    st.info("\n".join(notes))

    render_momentum(momentum_notes(mon_2026))

# ===================================================================
# MOD 2
# ===================================================================
elif mode == MODES[1]:
    st.header("⚔️ 2025'e Göre Ne Değişti? (Karşılaştırma Analizi)")
    if df_2025.empty:
        st.warning("⚠️ **Elinizde 2025 yılı verisi bulunmuyor.** Karşılaştırma için lütfen sol menüden **2025 Excel Raporunu** da yükleyin.")
        st.stop()
    if df_2026.empty:
        st.warning("⚠️ Lütfen karşılaştırma için **2026 Excel Raporunu** da yükleyin.")
        st.stop()

    a = df_2025[["Anahtar", "Müşteri", "Ciro", "Palet"]].rename(columns={"Müşteri": "M25", "Ciro": "Ciro_2025", "Palet": "Palet_2025"})
    b = df_2026[["Anahtar", "Müşteri", "Ciro", "Palet"]].rename(columns={"Müşteri": "M26", "Ciro": "Ciro_2026", "Palet": "Palet_2026"})
    cmp = b.merge(a, on="Anahtar", how="outer")
    cmp["Müşteri"] = cmp["M26"].fillna(cmp["M25"])
    for c in ["Ciro_2025", "Ciro_2026", "Palet_2025", "Palet_2026"]:
        cmp[c] = cmp[c].fillna(0.0)
    cmp["Fark"] = cmp["Ciro_2026"] - cmp["Ciro_2025"]
    cmp["Palet Fark"] = cmp["Palet_2026"] - cmp["Palet_2025"]
    cmp["Değişim %"] = cmp["Fark"] / cmp["Ciro_2025"].where(cmp["Ciro_2025"] > 0) * 100
    cmp = cmp.sort_values("Ciro_2026", ascending=False).reset_index(drop=True)
    top = cmp.head(20).copy()

    t25, t26 = df_2025["Ciro"].sum(), df_2026["Ciro"].sum()
    p25, p26 = df_2025["Palet"].sum(), df_2026["Palet"].sum()
    d_c = (t26 - t25) / t25 * 100 if t25 > 0 else None
    d_p = (p26 - p25) / p25 * 100 if p25 > 0 else None
    sgn = lambda x: f"{'+' if x >= 0 else ''}{tr_num(x, 1)} %"

    dc = lambda x: ("#2E7D32" if x >= 0 else "#C62828")
    kpis([("2025 Toplam Ciro", fmt_tl(t25), GREY),
          ("2026 Toplam Ciro", fmt_tl(t26), BLUE, sgn(d_c) if d_c is not None else "", dc(d_c or 0)),
          ("2025 Toplam Palet", fmt_palet(p25), GREY),
          ("2026 Toplam Palet", fmt_palet(p26), ORANGE, sgn(d_p) if d_p is not None else "", dc(d_p or 0))])

    st.markdown("---")
    st.subheader("📊 2025 → 2026 Değişim (Top 20)")
    st.caption("Her müşteri için gri çubuk 2025'i, renkli çubuk 2026'yı gösterir; tüm çubuklar sıfırdan ve aynı ölçekte başlar. Sağdaki panel ciro değişimidir (yeşil artış, kırmızı düşüş, turkuaz yeni müşteri; ±%150 üzeri görsel olarak kırpılır, etiket gerçek oranı verir).")
    show(change_chart(top))

    st.subheader("🥧 Müşteri Payları (Ciro)")
    q1, q2 = st.columns(2)
    with q1:
        show(pie_fig(cmp, "Ciro_2025", "2025 Ciro Dağılımı"))
    with q2:
        show(pie_fig(cmp, "Ciro_2026", "2026 Ciro Dağılımı"))

    st.subheader("📋 Top 20 Müşteri Değişim Tablosu")
    st.caption("🟩 artış, 🟥 düşüş. 'Yeni' = 2025'te kaydı olmayan müşteri.")
    render_table(
        ["Müşteri Unvanı", "2025 Ciro", "2026 Ciro", "Fark (₺)", "Değişim", "2025 Palet", "2026 Palet", "Palet Fark"],
        [top["Müşteri"].tolist(), [fmt_tl(v) for v in top["Ciro_2025"]], [fmt_tl(v) for v in top["Ciro_2026"]],
         [fmt_tl(v) for v in top["Fark"]],
         [("Yeni" if r["Ciro_2025"] == 0 else fmt_pct(r["Değişim %"])) for _, r in top.iterrows()],
         [fmt_palet(v) for v in top["Palet_2025"]], [fmt_palet(v) for v in top["Palet_2026"]],
         [fmt_palet(v) for v in top["Palet Fark"]]],
        [["rgb(247,248,250)"] * len(top), shade(top["Ciro_2025"], hex_rgb(GREY)), shade(top["Ciro_2026"], hex_rgb(BLUE)),
         diverge(top["Fark"]), diverge(top["Değişim %"].fillna(100.0)),
         shade(top["Palet_2025"], hex_rgb(GREY)), shade(top["Palet_2026"], hex_rgb(ORANGE)),
         diverge(top["Palet Fark"])],
        widths=[3, 2, 2, 2, 1.3, 1.3, 1.3, 1.3],
    )

    st.subheader("💡 Yorumlarımız")
    notes = []
    if d_c is not None:
        notes.append(f"* **Genel Tablo:** Toplam ciro 2025'te {fmt_tl(t25)} iken 2026'da {fmt_tl(t26)} oldu "
                     f"(%{tr_num(abs(d_c), 1)} {'artış' if d_c >= 0 else 'düşüş'}).")
    both = cmp[(cmp["Ciro_2025"] > 0) & (cmp["Ciro_2026"] > 0)]
    for _, r in both.sort_values("Fark", ascending=False).head(3).iterrows():
        if r["Fark"] > 0:
            notes.append(f"* **{r['Müşteri']}** ciroyu 2025'te {fmt_tl(r['Ciro_2025'])} seviyesinden 2026'da {fmt_tl(r['Ciro_2026'])} seviyesine taşıdı (%{tr_num(r['Değişim %'], 0)} artış).")
    for _, r in both.sort_values("Fark").head(3).iterrows():
        if r["Fark"] < 0:
            notes.append(f"* **{r['Müşteri']}** 2025'te {fmt_tl(r['Ciro_2025'])} iken 2026'da {fmt_tl(r['Ciro_2026'])} seviyesine geriledi (%{tr_num(abs(r['Değişim %']), 0)} düşüş).")
    newc = top[(top["Ciro_2025"] == 0) & (top["Ciro_2026"] > 0)].head(5)
    for _, r in newc.iterrows():
        notes.append(f"* **{r['Müşteri']}** 2025'te kaydı yokken 2026'da portföye girerek {fmt_tl(r['Ciro_2026'])} ciro ve {fmt_palet(r['Palet_2026'])} palete ulaştı.")
    st.info("\n".join(notes) if notes else "Yorum üretmek için yeterli veri yok.")

    render_momentum(momentum_notes(mon_2026))

# ===================================================================
# MOD 3
# ===================================================================
elif mode == MODES[2]:
    st.header("🚀 2027'de Ne Yapacağız?")
    st.caption("2027 hedefi = 2026 yıl sonu ciro beklentisi + mevcut portföyün büyümesi + hedef firmalardan gelecek ek hacim.")

    TOP_N, UNIT, MULT_R, MULT_G = 10, 2500, 2.5, 1.3

    cc1, cc2 = st.columns(2)
    with cc1:
        yk = st.number_input("2026 yıl sonu ciro beklentisi (milyon ₺)", min_value=0.0, value=450.0, step=5.0, key="yk26") * 1e6
    with cc2:
        gr = st.slider("Mevcut portföy büyüme varsayımı (fiyat + hacim, %)", 0, 60, 20, key="gr27")
    if not df_2026.empty:
        c26x = float(df_2026["Ciro"].sum())
        st.caption(f"Excel'de şu ana kadar okunan 2026 cirosu: {fmt_short_tl(c26x)}.")

    info = momentum_notes(mon_2026) if not df_2026.empty else None
    dorm, d_metric, d_last = dormant_list(mon_2026) if not df_2026.empty else ({}, "Palet", None)
    d_unit = "palet" if d_metric == "Palet" else "ciro"
    kn = dict(zip(df_2026["Anahtar"], df_2026["Müşteri"])) if not df_2026.empty else {}
    new_keys = list(info["new_keys"]) if info else []
    grow_keys = [k for k in (info["up_keys"] if info else []) if k not in new_keys]
    grow_keys += growth_candidates(df_2025, df_2026, set(new_keys) | set(grow_keys))
    rising_names = [kn[k] for k in new_keys if k in kn]
    growing_names = [kn[k] for k in grow_keys if k in kn and kn[k] not in rising_names][:3]
    if df_2026.empty:
        st.info("👈 Hızlı ivme, büyüme ve palet yorumları için 2026 Excel dosyasını yükleyin. Şimdilik yalnızca hedef firmalar hesaplanıyor.")

    rise_rows = df_2026[df_2026["Müşteri"].isin(rising_names)] if rising_names else df_2026.iloc[0:0]
    grow_rows = df_2026[df_2026["Müşteri"].isin(growing_names)] if growing_names else df_2026.iloc[0:0]
    picked = rise_rows["Müşteri"].tolist() + grow_rows["Müşteri"].tolist()

    tgt = pd.DataFrame([(n, p, SEG_SHORT[i]) for i, (_, rs) in enumerate(SEGMENTS) for n, p in rs],
                       columns=["Firma", "P", "Seg"]).sort_values("P", ascending=False, kind="stable")
    tgt = tgt[[not any(names_match(f, r) for r in picked) for f in tgt["Firma"]]]
    must = tgt["Firma"].isin(MUST_ADD)
    tgt = pd.concat([tgt[~must].head(TOP_N), tgt[must]]).sort_values("P", ascending=False, kind="stable")

    def make_note(key, p26, tur):
        if key and key in dorm:
            i_ = dorm[key]
            verb = "palet vermedi" if d_metric == "Palet" else "ciro yapmadı"
            return (f"Son {d_unit} {TR_AY[i_['last']]} ayında; {TR_AY[min(i_['last'] + 1, 12)]} ayından beri hiç {verb}. "
                    f"2027'de tekrar denenebilir.")
        if tur == "Hızlı ivme":
            return "Yıl içinde hızlı ivme yakaladı."
        if tur == "Büyüyen mevcut müşteri":
            return "Büyüme trendinde; hacim artırılacak."
        if p26 > 0:
            return "Mevcut müşteri; hacim artırılacak."
        return "2026'da çalışılmadı; yeni hedef."

    rows = []
    for _, r in tgt.iterrows():
        act = None
        if not df_2026.empty:
            m = df_2026[df_2026["Müşteri"].map(lambda x: names_match(r["Firma"], x))]
            if not m.empty:
                act = m.sort_values("Ciro", ascending=False).iloc[0]
        p26 = float(act["Palet"]) if act is not None else 0.0
        target = max(float(r["P"]), float(round(p26 * MULT_G))) if p26 > 0 else float(r["P"])
        key = act["Anahtar"] if act is not None else ""
        rows.append({"Firma": r["Firma"], "Tür": f"Hedef müşteri ({r['Seg']})", "2026 Palet": p26,
                     "2026 Ciro": float(act["Ciro"]) if act is not None else 0.0,
                     "Palet Başı ₺": float(UNIT), "2027 Hedef Palet": target,
                     "Excel Eşleşmesi": act["Müşteri"] if act is not None else "—",
                     "Anahtar": key, "Not": make_note(key, p26, "Hedef müşteri")})
    for df_pick, tur, mult in ((rise_rows, "Hızlı ivme", MULT_R), (grow_rows, "Büyüyen mevcut müşteri", MULT_G)):
        for _, r in df_pick.iterrows():
            ptl = r["Ciro"] / r["Palet"] if r["Palet"] > 0 and r["Ciro"] > 0 else float(UNIT)
            rows.append({"Firma": r["Müşteri"], "Tür": tur, "2026 Palet": float(r["Palet"]),
                         "2026 Ciro": float(r["Ciro"]), "Palet Başı ₺": round(float(ptl)),
                         "2027 Hedef Palet": float(round(r["Palet"] * mult)), "Excel Eşleşmesi": r["Müşteri"],
                         "Anahtar": r["Anahtar"], "Not": make_note(r["Anahtar"], float(r["Palet"]), tur)})
    if not rows:
        st.warning("Hesaplanacak firma yok.")
        st.stop()

    base_df = pd.DataFrame(rows).sort_values("2027 Hedef Palet", ascending=False).reset_index(drop=True)
    with st.expander("✏️ Rakamları düzenle (2027 hedef palet ve palet başı gelir)", expanded=False):
        edited = editor(
            base_df, hide_index=True,
            disabled=["Firma", "Tür", "2026 Palet", "2026 Ciro", "Excel Eşleşmesi", "Anahtar", "Not"],
            column_config={
                "Anahtar": None, "Not": None,
                "Firma": st.column_config.TextColumn("Firma", width="large"),
                "Tür": st.column_config.TextColumn("Tür", width="medium"),
                "2026 Palet": st.column_config.NumberColumn("2026 Gerçekleşen Palet", format="%.0f"),
                "2026 Ciro": st.column_config.NumberColumn("2026 Gerçekleşen Ciro (₺)", format="%.0f"),
                "Palet Başı ₺": st.column_config.NumberColumn("Palet Başı Gelir (₺)", format="%.0f", min_value=0.0, step=50.0),
                "2027 Hedef Palet": st.column_config.NumberColumn("2027 Hedef Palet", format="%.0f", min_value=0.0, step=50.0),
                "Excel Eşleşmesi": st.column_config.TextColumn("Excel'deki Ad", width="large"),
            })
    d = edited.copy()
    d["2027 Hedef Ciro"] = d["2027 Hedef Palet"] * d["Palet Başı ₺"]
    d["Ek Ciro"] = (d["2027 Hedef Palet"] - d["2026 Palet"]).clip(lower=0) * d["Palet Başı ₺"]
    n = len(d)
    tot_p, tot_c = d["2027 Hedef Palet"].sum(), d["2027 Hedef Ciro"].sum()
    p26_all = d["2026 Palet"].sum()
    base27 = yk * (1 + gr / 100)
    ek = float(d["Ek Ciro"].sum())
    hedef27 = base27 + ek
    growth = (hedef27 / yk - 1) * 100 if yk > 0 else None

    banner("🎯 2027 İÇİN ŞUNLARI HEDEFLİYORUZ", BLUE)
    kpis([("2026 Yıl Sonu Beklentisi", fmt_short_tl(yk), GREY),
          ("2027 Hedef Ciro", fmt_short_tl(hedef27), BLUE,
           (f"+{tr_num(growth, 0)} %" if growth is not None else ""), "#2E7D32"),
          (f"Mevcut Portföy (+%{gr})", fmt_short_tl(base27), TEAL),
          ("Hedef Firmalardan Ek Ciro", fmt_short_tl(ek), ORANGE, f"{n} firma", "#5b6574")])

    # --- Yan yana iki küçük tablo (tek slayta sığacak) ---
    DROP_TOK = {"ANONIM", "SIRKETI", "AS", "LTD", "STI", "LIMITED", "SAN", "TIC", "VE", "SANAYI",
                "TICARET", "PAZARLAMA", "TURKIYE", "ISLETMELERI"}

    def short_name(x):
        toks = []
        for t in str(x).split():
            c = re.sub(r"[^A-Z0-9]", "", norm(t))
            if c and c not in DROP_TOK:
                toks.append(t)
        return " ".join(toks[:3]) if toks else str(x)

    d["Gösterim"] = [f if str(t).startswith("Hedef") else short_name(f) for f, t in zip(d["Firma"], d["Tür"])]
    d["Mevcut"] = (d["2026 Palet"] > 0) | (d["2026 Ciro"] > 0)
    tot_fill = "rgb(226,233,244)"

    def mini_table(sub, mevcut, color):
        m_ = len(sub)
        if m_ == 0:
            st.caption("Bu grupta firma yok." if not df_2026.empty or not mevcut else "2026 Excel dosyasını yükleyin.")
            return
        z = ["rgb(255,255,255)" if i % 2 == 0 else "rgb(246,248,251)" for i in range(m_)] + [tot_fill]
        nm = [f"{x}{' 🔁' if k in dorm else ''}" for x, k in zip(sub["Gösterim"], sub["Anahtar"])]
        cols = [nm + ["<b>TOPLAM</b>"]]
        heads = ["Firma"]
        if mevcut:
            heads.append("2026 Palet")
            cols.append([fmt_palet(v) for v in sub["2026 Palet"]] + [f"<b>{fmt_palet(sub['2026 Palet'].sum())}</b>"])
        heads += ["2027 Hedef Palet", "2027 Hedef Ciro"]
        cols.append([fmt_palet(v) for v in sub["2027 Hedef Palet"]] + [f"<b>{fmt_palet(sub['2027 Hedef Palet'].sum())}</b>"])
        cols.append([fmt_short_tl(v) for v in sub["2027 Hedef Ciro"]] + [f"<b>{fmt_short_tl(sub['2027 Hedef Ciro'].sum())}</b>"])
        render_table(heads, cols, [z] * len(cols), font_size=11, header_fill=color, compact=True)

    tl, tr_ = st.columns(2)
    with tl:
        yeni = d[~d["Mevcut"]]
        banner(f"🎯 Hedef Müşteriler ({len(yeni)} firma)", ORANGE)
        mini_table(yeni, False, ORANGE)
    with tr_:
        mev = d[d["Mevcut"]]
        banner(f"✅ Mevcut Müşteriler ({len(mev)} firma)", BLUE)
        mini_table(mev, True, BLUE)
    st.caption("🔁 = son aylarda palet vermiyor, 2027'de tekrar denenebilir. Hedef ciro = hedef palet × palet başı gelir.")

    # --- Yorumlar ---
    st.subheader("💡 Yorumlarımız")
    notes = [f"* **Hedef:** 2026 yıl sonu **{fmt_short_tl(yk)}** beklentisine mevcut portföyde %{gr} büyüme eklendiğinde taban "
             f"**{fmt_short_tl(base27)}** oluyor; hedef firmalardan **{fmt_short_tl(ek)}** ek ciro ile 2027 hedefi "
             f"**{fmt_short_tl(hedef27)}**" + (f" (%{tr_num(growth, 0)} artış)." if growth is not None else ".")]
    if not df_2026.empty and yk > c26x:
        notes.append(f"* **Yıl Sonu Yolu:** Excel'deki ciro {fmt_short_tl(c26x)}; {fmt_short_tl(yk)} beklentisine ulaşmak için "
                     f"yıl sonuna kadar yaklaşık **{fmt_short_tl(yk - c26x)}** daha gerekiyor.")
    in_tbl = set(d["Anahtar"].astype(str))
    tbl_dorm = [r_ for r_ in d["Firma"][d["Anahtar"].isin(list(dorm.keys()))]] if dorm else []
    if tbl_dorm:
        notes.append(f"* **Tekrar Denenebilir:** {', '.join(tbl_dorm)} listemizde; son aylarda {d_unit} vermiyorlar, 2027'de yeniden denenmeli.")
    st.info("\n".join(notes))

    if dorm and not df_2026.empty:
        top_keys = set(df_2026.sort_values("Ciro", ascending=False).head(40)["Anahtar"])
        cand = [(k, v) for k, v in dorm.items() if k in top_keys and k not in in_tbl]
        cand = sorted(cand, key=lambda x: -x[1]["total"])[:6]
        if cand:
            st.subheader("🔁 2027'de Tekrar Denenebilecek Müşteriler")
            lines = []
            for k, v in cand:
                tot_txt = f"{fmt_palet(v['total'])} palet" if d_metric == "Palet" else fmt_tl(v["total"])
                verb = "palet vermedi" if d_metric == "Palet" else "ciro yapmadı"
                lines.append(f"* **{v['name']}**: 2026'da {tot_txt} ({TR_AY[v['first']]}–{TR_AY[v['last']]} arası); "
                             f"{TR_AY[min(v['last'] + 1, 12)]} ayından beri hiç {verb}. 2027'de tekrar denenebilir.")
            st.warning("\n".join(lines))
