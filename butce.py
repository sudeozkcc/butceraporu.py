import io
import re
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Lojistik Ciro & Bütçe Analizi", page_icon="🚚", layout="wide")

# ===================================================================
# STİL
# ===================================================================
st.markdown(
    """
<style>
div[data-testid="stMetric"]{
    background:linear-gradient(135deg,#eaf2ff 0%,#fff3e6 100%);
    border:1px solid #dde5f0;border-radius:14px;padding:14px 18px;
    box-shadow:0 2px 8px rgba(31,78,156,.08);
}
div[data-testid="stMetricLabel"], div[data-testid="stMetricLabel"] p{color:#4a5568 !important;}
div[data-testid="stMetricValue"]{color:#1f4e9c !important;}
</style>
""",
    unsafe_allow_html=True,
)

PALETTE = [
    "#1F77B4", "#FF7F0E", "#2CA02C", "#D62728", "#9467BD", "#8C564B",
    "#E377C2", "#17BECF", "#BCBD22", "#636EFA", "#EF553B", "#00CC96",
    "#AB63FA", "#FFA15A", "#19D3F3", "#FF6692", "#B6E880", "#FF97FF",
    "#FECB52", "#7F7F7F", "#3366CC", "#DC3912", "#109618", "#990099",
]
GREY = "#BDBDBD"

# Streamlit sürümüne göre "tam genişlik" parametresi
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

# ===================================================================
# BİÇİMLENDİRME & YARDIMCI FONKSİYONLAR
# ===================================================================
_TR = str.maketrans("İıŞşĞğÜüÖöÇç", "IISSGGUUOOCC")

def norm(x) -> str:
    if x is None or (not isinstance(x, (list, tuple)) and pd.isna(x)):
        return ""
    return str(x).translate(_TR).upper().strip()

def name_key(x) -> str:
    """Yıllar arası eşleştirme için müşteri adı anahtarı (büyük/küçük harf, boşluk, Türkçe karakter farkı yok sayılır)."""
    return re.sub(r"\s+", " ", norm(x))

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
    return tr_num(x, dec=2) + " ₺"

def fmt_palet(x) -> str:
    return tr_num(x, dec=0)

def fmt_pct(x, dec=1) -> str:
    if x is None or pd.isna(x):
        return "—"
    return tr_num(x, dec) + " %"

def short(s, n=22):
    s = str(s)
    return s if len(s) <= n else s[: n - 1] + "…"

# ----- renk yardımcıları (tablo hücreleri opak renk kullanır: koyu/açık temada okunur) -----
def hex_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))

def mix(rgb, a):
    r, g, b = [int(round(255 - (255 - c) * a)) for c in rgb]
    return f"rgb({r},{g},{b})"

def shade(values, rgb, lo=0.10, hi=0.65):
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
            out.append("rgb(245,247,251)")
        elif x > 0:
            out.append(mix((46, 160, 67), 0.15 + 0.45 * abs(x) / m))
        else:
            out.append(mix((214, 39, 40), 0.15 + 0.45 * abs(x) / m))
    return out

def build_cmap(names):
    cmap = {n: PALETTE[i % len(PALETTE)] for i, n in enumerate(names)}
    cmap["Diğer"] = GREY
    return cmap

# ===================================================================
# GRAFİK & TABLO FONKSİYONLARI
# ===================================================================
def render_table(headers, cols, fills, widths=None):
    n = len(cols[0]) if cols else 0
    fig = go.Figure(go.Table(
        columnwidth=widths,
        header=dict(
            values=[f"<b>{h}</b>" for h in headers],
            fill_color="#1f4e9c", font=dict(color="white", size=13),
            align="center", height=36, line_color="white",
        ),
        cells=dict(
            values=cols, fill_color=fills,
            align=["left"] + ["right"] * (len(cols) - 1),
            font=dict(size=12, color="#1a202c"), height=30, line_color="white",
        ),
    ))
    fig.update_layout(margin=dict(l=0, r=0, t=6, b=6), height=min(70 + 31 * n, 1000))
    show(fig)

def scatter_fig(top, cmap, avg_ptl):
    """Noktasal grafik: X=Palet, Y=Ciro, nokta büyüklüğü=Palet başı TL."""
    ptl = top["Palet Başı TL"].values
    lo, hi = (ptl.min(), ptl.max()) if len(ptl) else (0, 0)
    fig = go.Figure()
    xmax = max(top["Palet"].max() * 1.1, 1)
    fig.add_trace(go.Scatter(
        x=[0, xmax], y=[0, avg_ptl * xmax], mode="lines",
        name="Ortalama verimlilik çizgisi",
        line=dict(color="#8d99ae", dash="dash", width=2), hoverinfo="skip",
    ))
    for _, r in top.iterrows():
        size = 14 + 30 * ((r["Palet Başı TL"] - lo) / (hi - lo)) if hi > lo else 26
        fig.add_trace(go.Scatter(
            x=[r["Palet"]], y=[r["Ciro"]], mode="markers+text", name=r["Müşteri"],
            text=[short(r["Müşteri"], 16)], textposition="top center", textfont=dict(size=10),
            marker=dict(size=size, color=cmap.get(r["Müşteri"], GREY), opacity=0.9,
                        line=dict(color="white", width=2)),
            hovertext=(f"<b>{r['Müşteri']}</b><br>Ciro: {fmt_tl(r['Ciro'])}"
                       f"<br>Palet: {fmt_palet(r['Palet'])}<br>Palet başı: {fmt_tl(r['Palet Başı TL'])}"),
            hoverinfo="text",
        ))
    fig.update_layout(
        template="plotly_white", separators=",.", height=600,
        xaxis=dict(title="Palet Adedi", tickformat=",d", gridcolor="#eef1f6"),
        yaxis=dict(title="Ciro (₺)", tickformat=",.0f", gridcolor="#eef1f6"),
        legend=dict(orientation="v", font=dict(size=11)),
        margin=dict(l=20, r=20, t=30, b=20),
    )
    return fig

def pie_fig(df, col, title, cmap, top_n=10):
    d = df[["Müşteri", col]].sort_values(col, ascending=False)
    d = d[d[col] > 0]
    head = d.head(top_n)
    rest = d.iloc[top_n:][col].sum()
    labels = head["Müşteri"].tolist()
    values = head[col].tolist()
    if rest > 0:
        labels.append("Diğer")
        values.append(rest)
    colors = [cmap.get(l, GREY) for l in labels]
    total = float(sum(values))
    fig = go.Figure(go.Pie(
        labels=labels, values=values, hole=0.5, sort=False,
        marker=dict(colors=colors, line=dict(color="white", width=2)),
        textinfo="percent", textposition="inside", insidetextorientation="radial",
        hovertemplate="<b>%{label}</b><br>%{value:,.0f}<br>%{percent}<extra></extra>",
    ))
    fig.update_layout(
        template="plotly_white", separators=",.", height=480,
        title=dict(text=f"<b>{title}</b>", x=0.5),
        annotations=[dict(text=f"<b>{tr_num(total)}</b>", x=0.5, y=0.5, showarrow=False, font=dict(size=14))],
        legend=dict(orientation="h", y=-0.1, font=dict(size=10)),
        margin=dict(l=10, r=10, t=60, b=10),
    )
    return fig

def dumbbell_fig(d, c25, c26, title, unit):
    names = d["Müşteri"].tolist()
    lx, ly = [], []
    for n, a, b in zip(names, d[c25], d[c26]):
        lx += [a, b, None]
        ly += [n, n, None]
    fmt = fmt_tl if unit == "₺" else fmt_palet
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=lx, y=ly, mode="lines", line=dict(color="#c9ced6", width=3),
                             hoverinfo="skip", showlegend=False))
    fig.add_trace(go.Scatter(
        x=d[c25], y=names, mode="markers", name="2025",
        marker=dict(size=14, color="#FF7F0E", line=dict(color="white", width=2)),
        hovertext=[f"<b>{n}</b><br>2025: {fmt(v)}" for n, v in zip(names, d[c25])], hoverinfo="text"))
    fig.add_trace(go.Scatter(
        x=d[c26], y=names, mode="markers", name="2026",
        marker=dict(size=14, color="#2CA02C", line=dict(color="white", width=2)),
        hovertext=[f"<b>{n}</b><br>2026: {fmt(v)}" for n, v in zip(names, d[c26])], hoverinfo="text"))
    fig.update_layout(
        template="plotly_white", separators=",.", title=dict(text=f"<b>{title}</b>", x=0.5),
        height=max(420, 60 + 30 * len(names)),
        xaxis=dict(title=f"{title} ({unit})", tickformat=",.0f", gridcolor="#eef1f6"),
        yaxis=dict(autorange="reversed", gridcolor="#f3f5f9"),
        legend=dict(orientation="h", y=1.08, x=0.5, xanchor="center"),
        margin=dict(l=20, r=20, t=70, b=30),
    )
    return fig

# ===================================================================
# EXCEL PARS SİSTEMİ
# ===================================================================
CIRO_RE = re.compile(r"CIRO|TUTAR|GELIR|SATIS|\bTL\b|₺")
# Türetilmiş / toplanmaması gereken sütunlar (ör. "Palet Başı Ciro", "Ort.", "% Pay", "Bütçe", "Hedef")
EXCL_RE = re.compile(r"\bORT\b|ORTALAMA|BASI|BASINA|ORAN|\bPAY\b|FARK|BUYUME|%|ARTIS|DEGISIM|HEDEF|BUTCE")
TOTAL_RE = re.compile(r"TOPLAM|GENEL|YILLIK")

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

def detect_layout(raw: pd.DataFrame) -> dict:
    """Başlık satırını, Ciro/Palet sütunlarını ve müşteri sütununu tahmin eder (kullanıcı değiştirebilir)."""
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
    ciro_c, palet_c = [], []
    for j, lab in enumerate(labels):
        n = norm(lab)
        if not n or EXCL_RE.search(n):
            continue
        if "PALET" in n:
            palet_c.append(j)
        elif CIRO_RE.search(n):
            ciro_c.append(j)

    # Aynı anda hem "Toplam" hem aylık sütunlar varsa sadece toplam sütununu al (çift sayımı önler)
    def prefer_total(cols):
        tot = [j for j in cols if TOTAL_RE.search(norm(labels[j]))]
        return tot if tot else cols

    ciro_c, palet_c = prefer_total(ciro_c), prefer_total(palet_c)

    start = hdr + 1
    cust = 0
    for j in range(raw.shape[1]):
        if j in ciro_c or j in palet_c:
            continue
        vals = raw.iloc[start:, j].dropna().astype(str)
        if len(vals) > 0 and sum(not re.fullmatch(r"[\d\.,\s-]+", v) for v in vals) / len(vals) > 0.5:
            cust = j
            break
    return dict(hdr=hdr, start=start, cust=cust, ciro=ciro_c, palet=palet_c, labels=labels)

def build_df(raw, cust_col, ciro_cols, palet_cols, start_row):
    rows, excel_total = [], None
    sub = raw.iloc[start_row:]
    for i in range(len(sub)):
        nm = sub.iat[i, cust_col]
        if pd.isna(nm):
            continue
        nm = str(nm).strip()
        if not nm or nm.lower() in ("nan", "none"):
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

def load_year(year: int) -> pd.DataFrame:
    empty = pd.DataFrame(columns=["Anahtar", "Müşteri", "Ciro", "Palet"])
    up = st.sidebar.file_uploader(f"{year} Excel Raporunu Yükleyin (.xlsx)", type=["xlsx"], key=f"u{year}")
    if up is None:
        return empty
    fb = up.getvalue()
    sheet = st.sidebar.selectbox(f"{year} Sayfa Seçin:", get_sheet_names(fb), key=f"s{year}")
    try:
        raw = read_raw(fb, sheet)
    except Exception as e:
        st.sidebar.error(f"{year} dosyası okunamadı: {e}")
        return empty

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
        return empty

    st.sidebar.caption(f"✅ {year}: {len(df)} müşteri · Ciro {fmt_tl(df['Ciro'].sum())} · Palet {fmt_palet(df['Palet'].sum())}")
    if len(ciro_cols) > 1:
        st.sidebar.warning(f"{year}: {len(ciro_cols)} ciro sütunu toplanıyor. Biri 'toplam' sütunuysa çift sayım olur; eşleştirmeyi kontrol edin.")
    if xl_tot and xl_tot[0] > 0 and abs(df["Ciro"].sum() - xl_tot[0]) / xl_tot[0] > 0.01:
        st.sidebar.warning(f"{year}: Excel'deki TOPLAM satırı {fmt_tl(xl_tot[0])}, okunan toplam {fmt_tl(df['Ciro'].sum())}. Sütun eşleştirmeyi kontrol edin.")
    return df

# ===================================================================
# ARAYÜZ
# ===================================================================
st.title("🚚 Lojistik Ciro, Palet & Bütçe Sunum Portalı")

st.sidebar.header("📁 Excel Dosya Yükleme")
df_2025 = load_year(2025)
df_2026 = load_year(2026)

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
    cmap = build_cmap(base["Müşteri"].tolist())

    if tot_palet == 0:
        st.warning("⚠️ Palet toplamı 0 görünüyor. Sol menüde '⚙️ 2026 sütun eşleştirme' bölümünden palet sütununu seçin.")

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Toplam Ciro (tüm müşteriler)", fmt_tl(tot_ciro))
    c2.metric("Toplam Palet", fmt_palet(tot_palet))
    c3.metric("Palet Başı Ort. Gelir", fmt_tl(avg_ptl))
    c4.metric("Müşteri Sayısı", fmt_palet(len(base)))
    c5.metric("Top 20'nin Ciro Payı", fmt_pct(top20_share))

    st.markdown("---")
    st.subheader("🔵 Noktasal Grafik: Ciro × Palet (Top 20)")
    st.caption("Nokta büyüklüğü = palet başı gelir. Kesikli çizginin **üstündeki** müşteriler ortalamadan verimli, **altındakiler** daha az verimlidir.")
    show(scatter_fig(top20, cmap, avg_ptl))

    st.subheader("🥧 Pasta Grafikleri: Müşteri Payları (İlk 10 + Diğer)")
    p1, p2 = st.columns(2)
    with p1:
        show(pie_fig(base, "Ciro", "Ciro Dağılımı (₺)", cmap))
    with p2:
        show(pie_fig(base, "Palet", "Palet Dağılımı (adet)", cmap))

    st.subheader("📋 Ciro & Palet Tablosu (Top 20)")
    st.caption("Mavi = ciro, turuncu = palet büyüklüğü. Palet başı gelir: 🟩 ortalamanın üstünde, 🟥 altında.")
    t = top20.copy()
    t["Pay"] = t["Ciro"] / tot_ciro * 100 if tot_ciro else 0.0
    render_table(
        ["#", "Müşteri Unvanı", "Ciro", "Ciro Payı", "Palet", "Palet Başı TL"],
        [
            [str(i + 1) for i in range(len(t))],
            t["Müşteri"].tolist(),
            [fmt_tl(v) for v in t["Ciro"]],
            [fmt_pct(v) for v in t["Pay"]],
            [fmt_palet(v) for v in t["Palet"]],
            [fmt_tl(v) for v in t["Palet Başı TL"]],
        ],
        [
            ["rgb(245,247,251)"] * len(t),
            [mix(hex_rgb(cmap[n]), 0.30) for n in t["Müşteri"]],
            shade(t["Ciro"], (31, 119, 180)),
            shade(t["Pay"], (31, 119, 180)),
            shade(t["Palet"], (255, 127, 14)),
            diverge(t["Palet Başı TL"], center=avg_ptl),
        ],
        widths=[0.5, 3, 2, 1.3, 1.3, 2],
    )

    st.subheader("💡 2026 Performans Yorumlarımız ve Öne Çıkanlar")
    notes = [f"* **Genel Verimlilik:** 2026 genelinde palet başına ortalama **{fmt_tl(avg_ptl)}** ciro elde edilmiştir."]
    notes.append(f"* **Yoğunlaşma:** İlk 20 müşteri toplam cironun **{fmt_pct(top20_share)}** kadarını oluşturuyor.")
    if len(top20) and top20["Palet"].gt(0).any():
        valid = top20[top20["Palet"] > 0]
        star = valid.sort_values("Palet Başı TL", ascending=False).iloc[0]
        notes.append(f"* **En Verimli Müşteri:** **{star['Müşteri']}**, palet başına **{fmt_tl(star['Palet Başı TL'])}** ile Top 20 içinde en yüksek verimliliği yakalamıştır.")
        low = valid[valid["Palet Başı TL"] < avg_ptl * 0.85]
        if not low.empty:
            names = ", ".join(low.sort_values("Palet Başı TL").head(3)["Müşteri"])
            notes.append(f"* **Dikkat Edilecekler:** {names} palet başı geliri ortalamanın belirgin altında; fiyatlama/hacim gözden geçirilebilir.")
    st.info("\n".join(notes))

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

    a = df_2025[["Anahtar", "Müşteri", "Ciro", "Palet"]].rename(
        columns={"Müşteri": "M25", "Ciro": "Ciro_2025", "Palet": "Palet_2025"})
    b = df_2026[["Anahtar", "Müşteri", "Ciro", "Palet"]].rename(
        columns={"Müşteri": "M26", "Ciro": "Ciro_2026", "Palet": "Palet_2026"})
    cmp = b.merge(a, on="Anahtar", how="outer")
    cmp["Müşteri"] = cmp["M26"].fillna(cmp["M25"])
    for c in ["Ciro_2025", "Ciro_2026", "Palet_2025", "Palet_2026"]:
        cmp[c] = cmp[c].fillna(0.0)
    cmp["Fark"] = cmp["Ciro_2026"] - cmp["Ciro_2025"]
    cmp["Palet Fark"] = cmp["Palet_2026"] - cmp["Palet_2025"]
    cmp["Değişim %"] = np.where(cmp["Ciro_2025"] > 0, cmp["Fark"] / cmp["Ciro_2025"].where(cmp["Ciro_2025"] > 0) * 100, np.nan)

    cmp = cmp.sort_values("Ciro_2026", ascending=False).reset_index(drop=True)
    top = cmp.head(20).copy()
    cmap = build_cmap(cmp["Müşteri"].tolist())

    t25, t26 = df_2025["Ciro"].sum(), df_2026["Ciro"].sum()
    p25, p26 = df_2025["Palet"].sum(), df_2026["Palet"].sum()
    d_c = (t26 - t25) / t25 * 100 if t25 > 0 else None
    d_p = (p26 - p25) / p25 * 100 if p25 > 0 else None

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("2025 Toplam Ciro", fmt_tl(t25))
    m2.metric("2026 Toplam Ciro", fmt_tl(t26), delta=(f"{'+' if d_c >= 0 else ''}{tr_num(d_c, 1)} %" if d_c is not None else None))
    m3.metric("2025 Toplam Palet", fmt_palet(p25))
    m4.metric("2026 Toplam Palet", fmt_palet(p26), delta=(f"{'+' if d_p >= 0 else ''}{tr_num(d_p, 1)} %" if d_p is not None else None))

    st.markdown("---")
    st.subheader("🔵 Noktasal Değişim: 2025 → 2026 (Top 20)")
    tab1, tab2 = st.tabs(["💰 Ciro", "📦 Palet"])
    with tab1:
        show(dumbbell_fig(top, "Ciro_2025", "Ciro_2026", "Ciro", "₺"))
    with tab2:
        show(dumbbell_fig(top, "Palet_2025", "Palet_2026", "Palet", "adet"))

    new_stars = top[(top["Ciro_2025"] == 0) & (top["Ciro_2026"] > 0)]
    if not new_stars.empty:
        st.success(f"🚀 **2026'da Portföye Katılan Müşteriler (Top 20 içinde):** {', '.join(new_stars['Müşteri'].tolist())} — 2025'te kaydı yok, 2026'da yüksek ciroyla listeye girdi.")
    lost = cmp[(cmp["Ciro_2026"] == 0) & (cmp["Ciro_2025"] > 0)].sort_values("Ciro_2025", ascending=False).head(5)
    if not lost.empty:
        st.warning("⚠️ **2026'da ciro görünmeyen 2025 müşterileri (ilk 5):** " + ", ".join(lost["Müşteri"].tolist()))

    st.subheader("🥧 Müşteri Payları: 2025 vs 2026 (Ciro)")
    q1, q2 = st.columns(2)
    with q1:
        show(pie_fig(cmp, "Ciro_2025", "2025 Ciro Dağılımı", cmap))
    with q2:
        show(pie_fig(cmp, "Ciro_2026", "2026 Ciro Dağılımı", cmap))

    st.subheader("📋 Top 20 Müşteri Değişim Tablosu")
    st.caption("🟩 artış, 🟥 düşüş. 'Yeni' = 2025'te kaydı olmayan müşteri.")
    render_table(
        ["Müşteri Unvanı", "2025 Ciro", "2026 Ciro", "Fark (₺)", "Değişim", "2025 Palet", "2026 Palet", "Palet Fark"],
        [
            top["Müşteri"].tolist(),
            [fmt_tl(v) for v in top["Ciro_2025"]],
            [fmt_tl(v) for v in top["Ciro_2026"]],
            [fmt_tl(v) for v in top["Fark"]],
            [("Yeni" if r["Ciro_2025"] == 0 else fmt_pct(r["Değişim %"])) for _, r in top.iterrows()],
            [fmt_palet(v) for v in top["Palet_2025"]],
            [fmt_palet(v) for v in top["Palet_2026"]],
            [fmt_palet(v) for v in top["Palet Fark"]],
        ],
        [
            [mix(hex_rgb(cmap[n]), 0.30) for n in top["Müşteri"]],
            shade(top["Ciro_2025"], (255, 127, 14)),
            shade(top["Ciro_2026"], (44, 160, 44)),
            diverge(top["Fark"]),
            diverge(top["Değişim %"].fillna(100.0)),
            shade(top["Palet_2025"], (255, 127, 14)),
            shade(top["Palet_2026"], (44, 160, 44)),
            diverge(top["Palet Fark"]),
        ],
        widths=[3, 2, 2, 2, 1.3, 1.3, 1.3, 1.3],
    )

# ===================================================================
# MOD 3
# ===================================================================
elif mode == MODES[2]:
    st.header("🚀 2027'de Ne Yapacağız?")
    if df_2026.empty:
        st.info("👈 Hedefleri belirlemek için sol menüden 2026 Excel dosyanızı yükleyin.")
        st.stop()

    st.write("2027 için **en önemli 10 müşterimizden** beklentilerimizi simüle ediyoruz. Büyüme oranlarını tablodan değiştirebilirsiniz.")

    top10 = df_2026.sort_values("Ciro", ascending=False).head(10).reset_index(drop=True)
    top10["2027 Hedef Büyüme (%)"] = 25.0

    editor_kwargs = dict(
        hide_index=True,
        disabled=["Müşteri", "Ciro", "Palet"],
        column_config={
            "Müşteri": st.column_config.TextColumn("Hedef Müşterilerimiz (Top 10)", width="large"),
            "Ciro": st.column_config.NumberColumn("2026 Mevcut Ciro (₺)", format="%.0f"),
            "Palet": st.column_config.NumberColumn("2026 Mevcut Palet", format="%.0f"),
            "2027 Hedef Büyüme (%)": st.column_config.NumberColumn("2027 Beklenen Büyüme (%)", format="%.1f", min_value=-100.0, step=1.0),
        },
    )
    try:
        edited = st.data_editor(top10[["Müşteri", "Ciro", "Palet", "2027 Hedef Büyüme (%)"]], **stretch(), **editor_kwargs)
    except Exception:
        edited = st.data_editor(top10[["Müşteri", "Ciro", "Palet", "2027 Hedef Büyüme (%)"]], use_container_width=True, **editor_kwargs)

    edited = edited.copy()
    g = 1 + edited["2027 Hedef Büyüme (%)"] / 100
    edited["Ciro 2027"] = edited["Ciro"] * g
    edited["Palet 2027"] = edited["Palet"] * g

    t_c26, t_p26 = edited["Ciro"].sum(), edited["Palet"].sum()
    t_c27, t_p27 = edited["Ciro 2027"].sum(), edited["Palet 2027"].sum()
    avg27 = t_c27 / t_p27 if t_p27 > 0 else 0.0
    growth = (t_c27 - t_c26) / t_c26 * 100 if t_c26 else 0.0

    st.markdown("---")
    st.subheader("🎯 2027 İÇİN ŞUNLARI HEDEFLİYORUZ (Top 10 müşteri)")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("2027 Hedeflenen Ciro", fmt_tl(t_c27), delta=f"{'+' if growth >= 0 else ''}{tr_num(growth, 1)} %")
    k2.metric("2027 Hedeflenen Palet", fmt_palet(t_p27))
    k3.metric("2027 Palet Başı Ort. Gelir", fmt_tl(avg27))
    k4.metric("2026 Top 10 Ciro", fmt_tl(t_c26))

    cmap = build_cmap(edited["Müşteri"].tolist())
    g1, g2 = st.columns([3, 2])
    with g1:
        fig = go.Figure()
        fig.add_trace(go.Bar(x=edited["Müşteri"], y=edited["Ciro"], name="2026 Ciro", marker_color="#4C78A8",
                             hovertemplate="<b>%{x}</b><br>2026: %{y:,.0f} ₺<extra></extra>"))
        fig.add_trace(go.Bar(x=edited["Müşteri"], y=edited["Ciro 2027"], name="2027 Hedef", marker_color="#54A24B",
                             hovertemplate="<b>%{x}</b><br>2027: %{y:,.0f} ₺<extra></extra>"))
        fig.update_layout(template="plotly_white", separators=",.", barmode="group", height=480,
                          title=dict(text="<b>2026 Gerçekleşen vs 2027 Hedef Ciro</b>", x=0.5),
                          xaxis=dict(tickangle=-35), yaxis=dict(title="Ciro (₺)", tickformat=",.0f", gridcolor="#eef1f6"),
                          legend=dict(orientation="h", y=1.1, x=0.5, xanchor="center"),
                          margin=dict(l=20, r=20, t=70, b=20))
        show(fig)
    with g2:
        show(pie_fig(edited.rename(columns={"Ciro 2027": "Ciro27"}), "Ciro27", "2027 Hedef Ciro Payları", cmap, top_n=10))

    render_table(
        ["Müşteri", "2026 Ciro", "Büyüme", "2027 Hedef Ciro", "2026 Palet", "2027 Hedef Palet"],
        [
            edited["Müşteri"].tolist(),
            [fmt_tl(v) for v in edited["Ciro"]],
            [fmt_pct(v) for v in edited["2027 Hedef Büyüme (%)"]],
            [fmt_tl(v) for v in edited["Ciro 2027"]],
            [fmt_palet(v) for v in edited["Palet"]],
            [fmt_palet(v) for v in edited["Palet 2027"]],
        ],
        [
            [mix(hex_rgb(cmap[n]), 0.30) for n in edited["Müşteri"]],
            shade(edited["Ciro"], (76, 120, 168)),
            diverge(edited["2027 Hedef Büyüme (%)"]),
            shade(edited["Ciro 2027"], (84, 162, 75)),
            shade(edited["Palet"], (255, 127, 14)),
            shade(edited["Palet 2027"], (255, 127, 14)),
        ],
        widths=[3, 2, 1.2, 2, 1.3, 1.5],
    )

    st.info(f"📌 **2027 Bütçe Beklentilerimiz:** 2026'da ana ciro omurgasını oluşturan en önemli 10 müşterimizle 2027'de toplam **{fmt_tl(t_c27)}** ciro ve **{fmt_palet(t_p27)}** palet hacmine ulaşmayı hedefliyoruz.")
