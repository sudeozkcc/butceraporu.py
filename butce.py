import io
import re

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Ciro & Bütçeleme Analiz Portalı", page_icon="📊", layout="wide")

# ===================================================================
# YARDIMCI FONKSİYONLAR & BİÇİMLENDİRME
# ===================================================================
AYLAR = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran",
         "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]

_TR = str.maketrans("İıŞşĞğÜüÖöÇç", "IISSGGUUOOCC")


def norm(x) -> str:
    """Türkçe karakterleri sadeleştirip BÜYÜK harfe çevirir."""
    if x is None or (not isinstance(x, (list, tuple)) and pd.isna(x)):
        return ""
    return str(x).translate(_TR).upper().strip()


AY_KEYS = [norm(a) for a in AYLAR]


def find_month(x, strict=False):
    if isinstance(x, pd.Timestamp):
        return x.month
    n = norm(x)
    if not n:
        return None
    for i, k in enumerate(AY_KEYS):
        if (n.startswith(k) or n == k) if strict else (k in n):
            return i + 1
    return None


def find_metric(n: str):
    if "PALET" in n:
        return "Palet"
    if re.search(r"CIRO|TUTAR|\bTL\b|TRY|GELIR|SATIS", n):
        return "Ciro"
    return None


def has_info(row) -> int:
    c = 0
    for v in row:
        n = norm(v)
        if n and (find_metric(n) or "HEDEF" in n or "GERCEK" in n or re.search(r"20\d\d", n)):
            c += 1
    return c


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
    return "₺" + tr_num(x, dec=2)


def pct(a, b):
    return (b - a) / a * 100 if a and a != 0 else None


def find_sheet(names, *groups) -> int:
    for g in groups:
        for i, n in enumerate(names):
            nn = norm(n)
            if all(k in nn for k in g):
                return i
    return 0


def to_excel(sheets: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for name, d in sheets.items():
            d.to_excel(w, sheet_name=name[:31], index=False)
    return buf.getvalue()


# ===================================================================
# EXCEL OKUMA & PARS ETME
# ===================================================================
@st.cache_data(show_spinner=False)
def get_sheet_names(file_bytes: bytes):
    return pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl").sheet_names


@st.cache_data(show_spinner=False)
def get_raw(file_bytes: bytes, sheet: str) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet, header=None, engine="openpyxl")


def _row_label_parts(hdr: pd.DataFrame, j: int):
    cells = [v for v in hdr.iloc[:, j] if pd.notna(v)]
    return cells, " ".join(str(v) for v in cells)


def _parse_cols_as_months(raw: pd.DataFrame, default_year: int):
    month_row = None
    for r in range(min(15, len(raw))):
        if sum(1 for v in raw.iloc[r] if find_month(v)) >= 3:
            month_row = r
            break
    if month_row is None:
        return None

    rows = [month_row]
    for r in (month_row - 1, month_row - 2):
        if r >= 0 and has_info(raw.iloc[r]) >= 2:
            rows.insert(0, r)
        else:
            break
    r_below = month_row + 1
    if r_below < len(raw) and has_info(raw.iloc[r_below]) >= 2:
        rows.append(r_below)
    start = max(rows) + 1

    hdr = raw.iloc[rows].ffill(axis=1)
    kept = {}
    for j in range(raw.shape[1]):
        cells, label = _row_label_parts(hdr, j)
        n = norm(label)
        month = next((m for m in (find_month(c) for c in cells) if m), None)
        metric = find_metric(n)
        if not month or not metric or "TOPLAM" in n:
            continue
        ym = re.search(r"(20\d{2})", label)
        year = int(ym.group(1)) if ym else default_year
        tur = "Hedef" if ("HEDEF" in n or "BUTCE" in n) else "Gerçekleşen"
        kept[j] = (month, metric, year, tur)
    if not kept:
        return None

    data = raw.iloc[start:].reset_index(drop=True)
    cust_idx = 0
    for c in range(raw.shape[1]):
        if c in kept:
            continue
        col = data.iloc[:, c].dropna()
        if len(col) == 0:
            continue
        strs = sum(isinstance(v, str) and not re.fullmatch(r"[\d.,\s₺%-]+", v) for v in col)
        if strs / len(col) >= 0.5:
            cust_idx = c
            break

    cust = data.iloc[:, cust_idx].astype(object)
    cust_s = cust.map(lambda v: "" if pd.isna(v) else str(v).strip())
    ok = (cust_s != "") & ~cust_s.map(lambda s: "TOPLAM" in norm(s))

    frames = []
    for c, (m, met, yr, tur) in kept.items():
        frames.append(pd.DataFrame({
            "Müşteri": cust_s, "Yıl": yr, "Tür": tur, "Ay": m, "Metrik": met,
            "Değer": data.iloc[:, c].map(to_num)})[ok])
    return pd.concat(frames) if frames else None


def _parse_rows_as_months(raw: pd.DataFrame, default_year: int):
    mcol, idx = None, []
    for c in range(min(4, raw.shape[1])):
        idx = [r for r in range(len(raw)) if find_month(raw.iat[r, c], strict=True)]
        if len(idx) >= 3:
            mcol = c
            break
    if mcol is None:
        return None
    first = idx[0]
    hdr = raw.iloc[max(0, first - 4):first].ffill(axis=1)
    kept = {}
    for j in range(raw.shape[1]):
        if j == mcol:
            continue
        _, label = _row_label_parts(hdr, j)
        n = norm(label)
        metric = find_metric(n)
        if not metric or "TOPLAM" in n:
            continue
        ym = re.search(r"(20\d{2})", label)
        kept[j] = (metric, int(ym.group(1)) if ym else default_year,
                   "Hedef" if ("HEDEF" in n or "BUTCE" in n) else "Gerçekleşen")
    if not kept:
        return None
    rec, seen = [], set()
    for r in idx:
        m = find_month(raw.iat[r, mcol], strict=True)
        if m in seen:
            continue
        seen.add(m)
        for j, (met, yr, tur) in kept.items():
            rec.append({"Müşteri": "Toplam", "Yıl": yr, "Tür": tur, "Ay": m,
                        "Metrik": met, "Değer": to_num(raw.iat[r, j])})
    return pd.DataFrame(rec)


@st.cache_data(show_spinner=False)
def parse_sheet(file_bytes: bytes, sheet: str, default_year: int) -> pd.DataFrame:
    raw = get_raw(file_bytes, sheet)
    long = _parse_cols_as_months(raw, default_year)
    if long is None or long.empty:
        long = _parse_rows_as_months(raw, default_year)
    cols = ["Müşteri", "Yıl", "Tür", "Ay", "Palet", "Ciro"]
    if long is None or long.empty:
        return pd.DataFrame(columns=cols)
    piv = long.pivot_table(index=["Müşteri", "Yıl", "Tür", "Ay"], columns="Metrik",
                           values="Değer", aggfunc="sum", fill_value=0).reset_index()
    for m in ("Palet", "Ciro"):
        if m not in piv:
            piv[m] = 0.0
    return piv[cols]


def monthly_series(df: pd.DataFrame, year: int, tur: str, metric: str):
    s = df[(df["Yıl"] == year) & (df["Tür"] == tur)].groupby("Ay")[metric].sum()
    s = s.reindex(range(1, 13), fill_value=0.0)
    return s if s.abs().sum() > 0 else None


# ===================================================================
# OTURUM HAFIZASI
# ===================================================================
ss = st.session_state
ss.setdefault("file_bytes", None)
ss.setdefault("file_name", None)
ss.setdefault("uploader_key", 0)

POT_COLS = ["Firma Adı", "Durum", "Olasılık (%)", "2027 Tahmini Palet", "2027 Tahmini Ciro (₺)"]
DURUMLAR = ["Soğuk Temas", "Görüşülüyor", "Teklif Verildi", "Sözleşme Aşamasında"]


def empty_pot():
    return pd.DataFrame({"Firma Adı": pd.Series(dtype="object"), "Durum": pd.Series(dtype="object"),
                         "Olasılık (%)": pd.Series(dtype="float"),
                         "2027 Tahmini Palet": pd.Series(dtype="float"),
                         "2027 Tahmini Ciro (₺)": pd.Series(dtype="float")})


if "pot" not in ss:
    ss["pot"] = {c: empty_pot() for c in ["Soğuk Zincir", "Donuk Gıda", "Et & Et Ürünleri",
                                         "Süt & Süt Ürünleri", "Kuru Gıda & Lojistik"]}


def pot_totals():
    palet = ciro = 0.0
    for d in ss["pot"].values():
        if len(d):
            w = d["Olasılık (%)"].fillna(0) / 100
            palet += float((d["2027 Tahmini Palet"].fillna(0) * w).sum())
            ciro += float((d["2027 Tahmini Ciro (₺)"].fillna(0) * w).sum())
    return palet, ciro


# ===================================================================
# YAN MENÜ & DOSYA YÜKLEME
# ===================================================================
st.title("📊 Bütçeleme, Ciro Analizi ve Hedefleme Portalı")
st.sidebar.header("📁 Veri Yükleme & Modlar")

up = st.sidebar.file_uploader("Excel Raporunu Yükleyin (.xlsx)", type=["xlsx"],
                              key=f"excel_{ss['uploader_key']}")
if up is not None:
    ss["file_bytes"] = up.getvalue()
    ss["file_name"] = up.name

if ss["file_bytes"] is not None and st.sidebar.button("🗑️ Dosyayı Kaldır / Yenile"):
    ss["file_bytes"] = None
    ss["file_name"] = None
    ss["uploader_key"] += 1
    st.rerun()

mode = st.sidebar.radio("Çalışma Modunu Seçin:", [
    "📈 Aylık ve Genel Ciro Analizi",
    "🎯 2025 - 2026 Hedef Karşılaştırması",
    "🚀 2027 Hedef Oluşturucu",
    "🔍 Potansiyel Müşteriler & Kategori Hedefleri"])
default_year = st.sidebar.number_input("Başlıkta yıl yazmıyorsa varsayılan yıl", 2020, 2035, 2026)

if ss["file_bytes"] is None and mode != "🔍 Potansiyel Müşteriler & Kategori Hedefleri":
    st.info("👈 Başlamak için sol menüden Excel ciro raporunuzu yükleyin.")
    st.stop()

FB = ss["file_bytes"]
sheet_names = []
if FB is not None:
    st.success(f"📌 **Aktif Dosya:** `{ss['file_name']}`")
    try:
        sheet_names = get_sheet_names(FB)
    except Exception as e:
        st.error(f"Excel dosyası okunamadı: {e}")
        st.stop()


def load_parsed(label, default_idx):
    sheet = st.selectbox(label, sheet_names, index=default_idx)
    df = parse_sheet(FB, sheet, int(default_year))
    with st.expander("🔎 Sayfa önizleme / algılanan yapı"):
        st.caption(f"Algılanan kayıt: {len(df)} satır | Yıllar: {sorted(df['Yıl'].unique().tolist())} | "
                   f"Türler: {df['Tür'].unique().tolist()}")
        st.dataframe(get_raw(FB, sheet).head(15).astype(str), use_container_width=True)
    if df.empty:
        st.warning("Bu sayfada ay / palet / ciro başlıkları algılanamadı.")
        st.stop()
    return df


# ===================================================================
# MOD 1 — AYLIK VE GENEL CİRO ANALİZİ
# ===================================================================
if mode.startswith("📈"):
    st.header("📈 Ciro & Palet Analizi")
    df = load_parsed("Ciro Analizi Sayfası:", find_sheet(sheet_names, ("PALET", "CIRO"), ("2026",)))
    data = df[df["Tür"] == "Gerçekleşen"]
    if data.empty:
        st.info("Sayfada 'Gerçekleşen' veri bulunamadı, tüm veriler kullanılıyor.")
        data = df

    years = sorted(data["Yıl"].unique().tolist())
    c0, c1 = st.columns([1, 2])
    year = c0.selectbox("Yıl Seçin:", years, index=len(years) - 1)

    tot = data[data["Yıl"] == year].groupby("Ay")[["Palet", "Ciro"]].sum().reindex(range(1, 13), fill_value=0)

    # En son veri olan ayı bulma
    last_active_month = max([i + 1 for i in range(12) if tot["Ciro"].iloc[i] > 0 or tot["Palet"].iloc[i] > 0], default=1)

    tab_aylik, tab_genel = st.tabs(["🔄 İki Ay Karşılaştırma", "🌐 Genel Tablo (Ocak - Gelinen Ay)"])

    # ---------------------------------------------------------------
    # SEKMESİ 1: İKİ AY KARŞILAŞTIRMASI (Örn: Eylül vs Ekim)
    # ---------------------------------------------------------------
    with tab_aylik:
        pairs = [f"{AYLAR[i]} → {AYLAR[i + 1]}" for i in range(11)]
        default_pair_idx = min(last_active_month - 2, 10) if last_active_month >= 2 else 0
        pair = st.selectbox("Karşılaştırılacak İki Dönem Seçin:", pairs, index=default_pair_idx)
        pi = pairs.index(pair)
        pm, cm = pi + 1, pi + 2

        # Önceki Ay ve Şu Anki Ay için Müşteri Bazlı Toplamlar
        a = data[(data["Yıl"] == year) & (data["Ay"] == pm)].groupby("Müşteri")[["Palet", "Ciro"]].sum().add_suffix(f" ({AYLAR[pm - 1]})")
        b = data[(data["Yıl"] == year) & (data["Ay"] == cm)].groupby("Müşteri")[["Palet", "Ciro"]].sum().add_suffix(f" ({AYLAR[cm - 1]})")

        cmp_df = a.join(b, how="outer").fillna(0.0).reset_index()

        col_c_prev = f"Ciro ({AYLAR[pm - 1]})"
        col_c_curr = f"Ciro ({AYLAR[cm - 1]})"
        col_p_prev = f"Palet ({AYLAR[pm - 1]})"
        col_p_curr = f"Palet ({AYLAR[cm - 1]})"

        cmp_df["Ciro Farkı (₺)"] = cmp_df[col_c_curr] - cmp_df[col_c_prev]
        cmp_df["Ciro Değişim (%)"] = cmp_df.apply(lambda r: pct(r[col_c_prev], r[col_c_curr]), axis=1)
        cmp_df["Palet Farkı"] = cmp_df[col_p_curr] - cmp_df[col_p_prev]
        cmp_df["Palet Değişim (%)"] = cmp_df.apply(lambda r: pct(r[col_p_prev], r[col_p_curr]), axis=1)

        # Durum belirleme (Artış, Düşüş, Yeni, Kayıp)
        def calc_status(r):
            if r[col_c_prev] == 0 and r[col_c_curr] > 0:
                return "🟢 Yeni Müşteri"
            if r[col_c_prev] > 0 and r[col_c_curr] == 0:
                return "🔴 Kaybedilen Müşteri"
            if r["Ciro Farkı (₺)"] > 0:
                return "🟢 Artış"
            if r["Ciro Farkı (₺)"] < 0:
                return "🔴 Düşüş"
            return "⚪ Değişmedi"

        cmp_df["Durum"] = cmp_df.apply(calc_status, axis=1)

        tot_c_prev = tot["Ciro"][pm]
        tot_c_curr = tot["Ciro"][cm]
        tot_p_prev = tot["Palet"][pm]
        tot_p_curr = tot["Palet"][cm]

        st.subheader(f"📊 {AYLAR[pm - 1]} vs {AYLAR[cm - 1]} {year} Karşılaştırma Özeti")

        k1, k2, k3, k4 = st.columns(4)
        k1.metric(f"Ciro ({AYLAR[pm - 1]})", fmt_tl(tot_c_prev))
        k2.metric(f"Ciro ({AYLAR[cm - 1]})", fmt_tl(tot_c_curr),
                  f"%{pct(tot_c_prev, tot_c_curr):.1f}" if pct(tot_c_prev, tot_c_curr) is not None else None)
        k3.metric(f"Palet ({AYLAR[pm - 1]})", tr_num(tot_p_prev))
        k4.metric(f"Palet ({AYLAR[cm - 1]})", tr_num(tot_p_curr),
                  f"%{pct(tot_p_prev, tot_p_curr):.1f}" if pct(tot_p_prev, tot_p_curr) is not None else None)

        st.markdown("---")
        st.write(f"### 📋 Müşteri Bazlı Detaylı Dön
