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
st.title("📊 Bütçeleme, Ciro Analizi ve Akıllı Trend Portalı")
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
    st.header("📈 Ciro & Palet Analizi (Akıllı Müşteri Trendleri)")
    df = load_parsed("Ciro Analizi Sayfası:", find_sheet(sheet_names, ("PALET", "CIRO"), ("2026",)))
    data = df[df["Tür"] == "Gerçekleşen"]
    if data.empty:
        st.info("Sayfada 'Gerçekleşen' veri bulunamadı, tüm veriler kullanılıyor.")
        data = df

    years = sorted(data["Yıl"].unique().tolist())
    c0, c1 = st.columns([1, 2])
    year = c0.selectbox("Yıl Seçin:", years, index=len(years) - 1)

    tot = data[data["Yıl"] == year].groupby("Ay")[["Palet", "Ciro"]].sum().reindex(range(1, 13), fill_value=0)
    last_active_month = max([i + 1 for i in range(12) if tot["Ciro"].iloc[i] > 0 or tot["Palet"].iloc[i] > 0], default=1)

    # Karşılaştırma Seçenekleri
    pairs = ["🌐 Genel (Yıl Başı - Gelinen Ay)"] + [f"{AYLAR[i]} vs {AYLAR[i + 1]}" for i in range(11)]
    default_pair_idx = min(last_active_month - 1, 11) if last_active_month >= 2 else 0

    selected_period = st.selectbox("📅 Karşılaştırılacak Dönem Veya Analiz Modu Seçin:", pairs, index=default_pair_idx)

    # ---------------------------------------------------------------
    # DURUM A: "🌐 Genel (Yıl Başı - Gelinen Ay)" SEÇİLİRSE
    # ---------------------------------------------------------------
    if selected_period.startswith("🌐"):
        st.subheader(f"🌐 {year} Yılı Genel Tablo (Ocak → {AYLAR[last_active_month - 1]})")

        ytd_data = data[(data["Yıl"] == year) & (data["Ay"] <= last_active_month)]

        # 1. HER AYI YAN YANA SÜTUN OLARAK GÖSTEREN MATRIX
        st.write(f"### 📋 Müşteri Bazında Aylık Ciro Matrix'i (Ocak - {AYLAR[last_active_month - 1]})")

        piv_ciro = ytd_data.pivot_table(index="Müşteri", columns="Ay", values="Ciro", aggfunc="sum", fill_value=0.0)
        piv_ciro.columns = [f"{AYLAR[m - 1]} Ciro" for m in piv_ciro.columns]
        piv_ciro["Toplam Ciro (₺)"] = piv_ciro.sum(axis=1)
        piv_ciro = piv_ciro.sort_values("Toplam Ciro (₺)", ascending=False).reset_index()

        cfg_ciro = {"Müşteri": st.column_config.TextColumn("Müşteri", width="medium")}
        for col in piv_ciro.columns:
            if col != "Müşteri":
                cfg_ciro[col] = st.column_config.NumberColumn(col, format="%.2f ₺")

        st.dataframe(piv_ciro, use_container_width=True, hide_index=True, column_config=cfg_ciro)

        st.write(f"### 📦 Müşteri Bazında Aylık Palet Matrix'i (Ocak - {AYLAR[last_active_month - 1]})")
        piv_palet = ytd_data.pivot_table(index="Müşteri", columns="Ay", values="Palet", aggfunc="sum", fill_value=0.0)
        piv_palet.columns = [f"{AYLAR[m - 1]} Palet" for m in piv_palet.columns]
        piv_palet["Toplam Palet"] = piv_palet.sum(axis=1)
        piv_palet = piv_palet.sort_values("Toplam Palet", ascending=False).reset_index()

        cfg_palet = {"Müşteri": st.column_config.TextColumn("Müşteri", width="medium")}
        for col in piv_palet.columns:
            if col != "Müşteri":
                cfg_palet[col] = st.column_config.NumberColumn(col, format="%.0f")

        st.dataframe(piv_palet, use_container_width=True, hide_index=True, column_config=cfg_palet)

        # 2. GENEL TREND GRAFİĞİ
        st.markdown("---")
        st.write("### 📉 Genel Aylık Ciro ve Palet Trendi (Noktasal & Çizgili)")
        aylik_toplamlar = ytd_data.groupby("Ay")[["Ciro", "Palet"]].sum().reindex(range(1, last_active_month + 1), fill_value=0).reset_index()
        aylik_toplamlar["Ay Adı"] = aylik_toplamlar["Ay"].map(lambda x: AYLAR[x - 1])

        fig_g = go.Figure()
        fig_g.add_trace(go.Scatter(x=aylik_toplamlar["Ay Adı"], y=aylik_toplamlar["Ciro"],
                                   mode="lines+markers+text", name="Ciro (₺)",
                                   line=dict(color="#2CA02C", width=3),
                                   marker=dict(size=10),
                                   text=[fmt_tl(v) for v in aylik_toplamlar["Ciro"]],
                                   textposition="top center"))

        fig_g.update_layout(title="Yıl Başı Gelinen Aya Kadar Aylık Ciro Gelişimi", yaxis_title="Ciro (₺)")
        st.plotly_chart(fig_g, use_container_width=True)

        # 3. YAZILIMSAL AKILLI YORUMLAR (GENEL MÜŞTERİ YORUMU)
        st.subheader("💡 Sistem Yorumu & Trend Analizi")

        top_cust = piv_ciro.iloc[0]["Müşteri"] if len(piv_ciro) else "-"
        top_val = piv_ciro.iloc[0]["Toplam Ciro (₺)"] if len(piv_ciro) else 0

        # Israrla Düşen Müşteriler (Son 3 ay düşüşte olanlar)
        persistent_drop = []
        if last_active_month >= 3:
            m1_col, m2_col, m3_col = AYLAR[last_active_month - 3], AYLAR[last_active_month - 2], AYLAR[last_active_month - 1]
            p_raw = ytd_data.pivot_table(index="Müşteri", columns="Ay", values="Ciro", aggfunc="sum", fill_value=0)
            for cust_name, r in p_raw.iterrows():
                if r.get(last_active_month - 2, 0) > 0:
                    if r.get(last_active_month - 2, 0) < r.get(last_active_month - 3, 0) and r.get(last_active_month - 1, 0) < r.get(last_active_month - 2, 0):
                        persistent_drop.append(cust_name)

        st.info(f"""
        * **En Yüksek Katkı Veren Müşteri:** Yıl başından itibaren **{top_cust}** firması toplam **{fmt_tl(top_val)}** ciro ile ilk sırada yer almaktadır.
        * **Trend Gidişatı:** {year} yılında **{AYLAR[last_active_month - 1]}** ayına kadar toplam **{fmt_tl(aylik_toplamlar['Ciro'].sum())}** ciro ve **{tr_num(aylik_toplamlar['Palet'].sum())}** palet hacmine ulaşıldı.
        """ + (f"\n* **🚨 Israrla Gerileyen Müşteriler:** Son 3 aydır cirosu sürekli düşüş eğiliminde olan müşteriler: **{', '.join(persistent_drop)}**. Bu firmalar ile iletişime geçilmesi önerilir." if persistent_drop else "\n* **Müşteri Stabilizasyonu:** Son 3 ayda sürekli gerileyen belirgin bir müşteri grubu tespit edilmemiştir."))

    # ---------------------------------------------------------------
    # DURUM B: İKİ AY KARŞILAŞTIRMASI SEÇİLİRSE (Örn: Kasım vs Aralık)
    # ---------------------------------------------------------------
    else:
        pi = pairs.index(selected_period) - 1
        pm, cm = pi + 1, pi + 2

        st.subheader(f"📊 {AYLAR[pm - 1]} vs {AYLAR[cm - 1]} {year} Detaylı Karşılaştırması")

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

        tot_c_prev = tot["Ciro"][pm]
        tot_c_curr = tot["Ciro"][cm]
        tot_p_prev = tot["Palet"][pm]
        tot_p_curr = tot["Palet"][cm]

        # Üst Özet Kartları
        k1, k2, k3, k4 = st.columns(4)
        k1.metric(f"Ciro ({AYLAR[pm - 1]})", fmt_tl(tot_c_prev))
        k2.metric(f"Ciro ({AYLAR[cm - 1]})", fmt_tl(tot_c_curr),
                  f"%{pct(tot_c_prev, tot_c_curr):.1f}" if pct(tot_c_prev, tot_c_curr) is not None else None)
        k3.metric(f"Palet ({AYLAR[pm - 1]})", tr_num(tot_p_prev))
        k4.metric(f"Palet ({AYLAR[cm - 1]})", tr_num(tot_p_curr),
                  f"%{pct(tot_p_prev, tot_p_curr):.1f}" if pct(tot_p_prev, tot_p_curr) is not None else None)

        st.markdown("---")

        # 1. AKILLI MÜŞTERİ YORUMU VE BELİRGİN DÜŞÜŞ TESPİTİ
        st.write("### 💡 Otomatik Müşteri Trendi ve Analiz Yorumu")

        # Ciddi Düşüş Gösteren Müşteriler (%25+ düşüş veya 10.000 TL üzeri azalma)
        significant_drops = cmp_df[
            (cmp_df[col_c_prev] >= 5000) &
            ((cmp_df["Ciro Değişim (%)"] <= -25) | (cmp_df["Ciro Farkı (₺)"] <= -10000))
        ].sort_values("Ciro Farkı (₺)")

        # Ciddi Artış Gösterenler
        significant_gains = cmp_df[
            (cmp_df["Ciro Farkı (₺)"] >= 10000) | (cmp_df["Ciro Değişim (%)"] >= 30)
        ].sort_values("Ciro Farkı (₺)", ascending=False)

        if not significant_drops.empty:
            st.warning(f"⚠️ **{AYLAR[pm - 1]} → {AYLAR[cm - 1]} Döneminde Belirgin Düşüş Gösteren Müşteriler:**")
            for _, r in significant_drops.head(5).iterrows():
                st.write(f"- 🔴 **{r['Müşteri']}**: {AYLAR[pm - 1]} ayında **{fmt_tl(r[col_c_prev])}** ciro yaparken, {AYLAR[cm - 1]} ayında **{fmt_tl(r[col_c_curr])}** seviyesine düştü. (Fark: **{fmt_tl(r['Ciro Farkı (₺)'])}**, Değişim: **%{r['Ciro Değişim (%)']:.1f}**)")
        else:
            st.success(f"✅ {AYLAR[pm - 1]} ayından {AYLAR[cm - 1]} ayına geçişte kritik düzeyde yüksek düşüş gösteren müşteri bulunmuyor.")

        st.markdown("---")
        st.write(f"### 📋 Müşteri Bazlı Detaylı Tablo ({AYLAR[pm - 1]} vs {AYLAR[cm - 1]})")

        st.dataframe(
            cmp_df.sort_values(col_c_curr, ascending=False),
            use_container_width=True,
            hide_index=True,
            column_config={
                "Müşteri": st.column_config.TextColumn("Müşteri", width="medium"),
                col_c_prev: st.column_config.NumberColumn(f"{AYLAR[pm - 1]} Ciro", format="%.2f ₺"),
                col_c_curr: st.column_config.NumberColumn(f"{AYLAR[cm - 1]} Ciro", format="%.2f ₺"),
                "Ciro Farkı (₺)": st.column_config.NumberColumn("Ciro Farkı", format="%.2f ₺"),
                "Ciro Değişim (%)": st.column_config.NumberColumn("Ciro Değişim", format="%.1f %%"),
                col_p_prev: st.column_config.NumberColumn(f"{AYLAR[pm - 1]} Palet", format="%.0f"),
                col_p_curr: st.column_config.NumberColumn(f"{AYLAR[cm - 1]} Palet", format="%.0f"),
                "Palet Farkı": st.column_config.NumberColumn("Palet Farkı", format="%.0f"),
                "Palet Değişim (%)": st.column_config.NumberColumn("Palet Değişim", format="%.1f %%")
            }
        )

        st.markdown("---")
        # 2. İKİ DÖNEM GÖRSEL GRAFİĞİ
        st.write(f"### 📉 En Yüksek Ciroya Sahip 10 Müşterinin {AYLAR[pm - 1]} & {AYLAR[cm - 1]} Noktasal Karşılaştırması")
        top10 = cmp_df.nlargest(10, col_c_curr)

        fig_dot = go.Figure()
        fig_dot.add_trace(go.Scatter(
            x=top10["Müşteri"], y=top10[col_c_prev],
            mode="markers+lines", name=AYLAR[pm - 1],
            marker=dict(size=12, color="#FFA15A")
        ))
        fig_dot.add_trace(go.Scatter(
            x=top10["Müşteri"], y=top10[col_c_curr],
            mode="markers+lines", name=AYLAR[cm - 1],
            marker=dict(size=12, color="#00CC96")
        ))
        fig_dot.update_layout(title=f"Müşteri Ciro Değişimi ({AYLAR[pm - 1]} vs {AYLAR[cm - 1]})", yaxis_title="Ciro (₺)")
        st.plotly_chart(fig_dot, use_container_width=True)

# ===================================================================
# MOD 2 — 2025 / 2026 HEDEF KARŞILAŞTIRMASI
# ===================================================================
elif mode.startswith("🎯"):
    st.header("🎯 Yıllar Arası Karşılaştırma: Gerçekleşen & Hedef")
    df = load_parsed("Hedef Sayfası:", find_sheet(sheet_names, ("HEDEF",), ("2026",)))
    years = sorted(df["Yıl"].unique().tolist())
    if len(years) < 2:
        st.warning(f"Sayfada tek yıl algılandı: {years}. Başlıklarda yıl (2025, 2026) yazmalı.")
    c1, c2 = st.columns(2)
    y_base = c1.selectbox("Baz Yıl:", years, index=max(0, len(years) - 2))
    y_cmp = c2.selectbox("Karşılaştırılan Yıl:", years, index=len(years) - 1)

    for met in ("Ciro", "Palet"):
        st.subheader(f"{'💰' if met == 'Ciro' else '📦'} {met} Karşılaştırması")
        series = {
            f"{y_base} Gerçekleşen": monthly_series(df, y_base, "Gerçekleşen", met),
            f"{y_base} Hedef": monthly_series(df, y_base, "Hedef", met),
            f"{y_cmp} Hedef": monthly_series(df, y_cmp, "Hedef", met),
            f"{y_cmp} Gerçekleşen": monthly_series(df, y_cmp, "Gerçekleşen", met),
        }
        series = {k: v for k, v in series.items() if v is not None}
        if not series:
            st.info("Veri bulunamadı.")
            continue
        table = pd.DataFrame(series)
        table.insert(0, "Ay", AYLAR)

        fig = go.Figure([go.Bar(name=k, x=AYLAR, y=v.values) for k, v in series.items()])
        fig.update_layout(barmode="group", title=f"Aylık {met}")
        st.plotly_chart(fig, use_container_width=True)

        ba, ca, ch = (series.get(f"{y_base} Gerçekleşen"), series.get(f"{y_cmp} Gerçekleşen"),
                      series.get(f"{y_cmp} Hedef"))
        if ba is not None and ca is not None:
            last = max(ca[ca > 0].index, default=0)
            ytd_b, ytd_c = ba.loc[:last].sum(), ca.loc[:last].sum()
            k = st.columns(3)
            k[0].metric(f"{y_base} ({AYLAR[last - 1] if last else '-'}'a kadar)", fmt_tl(ytd_b) if met == "Ciro" else tr_num(ytd_b))
            k[1].metric(f"{y_cmp} ({AYLAR[last - 1] if last else '-'}'a kadar)", fmt_tl(ytd_c) if met == "Ciro" else tr_num(ytd_c),
                        f"%{pct(ytd_b, ytd_c):.1f}" if pct(ytd_b, ytd_c) is None else None)
            if ch is not None:
                ch_sum = ch.loc[:last].sum()
                ca_sum = ca.loc[:last].sum()
                ratio = (ca_sum / ch_sum * 100) if ch_sum > 0 else 0
                k[2].metric(f"{y_cmp} Hedef Gerçekleşme", f"%{ratio:.1f}" if ch_sum > 0 else "-")

        cfg = {"Ay": st.column_config.TextColumn("Ay")}
        for col_name in table.columns:
            if col_name != "Ay":
                cfg[col_name] = st.column_config.NumberColumn(
                    col_name,
                    format="%.2f ₺" if met == "Ciro" else "%.0f"
                )
        st.dataframe(table, use_container_width=True, hide_index=True, column_config=cfg)

# ===================================================================
# MOD 3 — 2027 HEDEF OLUŞTURUCU
# ===================================================================
elif mode.startswith("🚀"):
    st.header("🚀 2026'dan Hareketle 2027 Hedef Simülasyonu")
    df = load_parsed("Baz alınacak sayfa:", find_sheet(sheet_names, ("HEDEF",), ("PALET", "CIRO")))
    c1, c2, c3 = st.columns(3)
    y_base = c1.selectbox("Baz Yıl:", sorted(df["Yıl"].unique().tolist()),
                          index=len(df["Yıl"].unique()) - 1)
    turler = df[df["Yıl"] == y_base]["Tür"].unique().tolist()
    tur = c2.selectbox("Baz Veri Türü:", turler)
    fill = c3.checkbox("Verisi olmayan ayları ortalamayla doldur (yıllıklandır)", value=True)
    base = df[(df["Yıl"] == y_base) & (df["Tür"] == tur)]

    s1, s2 = st.columns(2)
    g_ciro = s1.slider("Genel Ciro Büyüme Oranı (%)", 0, 150, 25, 5)
    g_palet = s2.slider("Genel Palet Büyüme Oranı (%)", 0, 150, 15, 5)

    def month_fill(s):
        s = s.reindex(range(1, 13), fill_value=0.0)
        if fill and (s > 0).any():
            s = s.where(s > 0, s[s > 0].mean())
        return s

    raw_m = {m: base.groupby("Ay")[m].sum().reindex(range(1, 13), fill_value=0.0) for m in ("Ciro", "Palet")}
    fill_m = {m: month_fill(raw_m[m]) for m in raw_m}
    factor = {m: (fill_m[m].sum() / raw_m[m].sum() if raw_m[m].sum() > 0 else 1.0) for m in raw_m}

    st.subheader("👥 Müşteri Bazlı Hedefler")
    cust = base.groupby("Müşteri")[["Ciro", "Palet"]].sum().sort_values("Ciro", ascending=False).reset_index()
    cust["Ciro"] *= factor["Ciro"]
    cust["Palet"] *= factor["Palet"]

    ed_in = pd.DataFrame({"Müşteri": cust["Müşteri"], f"{y_base} Ciro (₺)": cust["Ciro"].round(2),
                          f"{y_base} Palet": cust["Palet"].round(0),
                          "Ciro Büyüme (%)": float(g_ciro), "Palet Büyüme (%)": float(g_palet)})

    edited = st.data_editor(ed_in, hide_index=True, use_container_width=True,
                            disabled=["Müşteri", f"{y_base} Ciro (₺)", f"{y_base} Palet"],
                            key=f"cust_ed_{y_base}_{tur}_{g_ciro}_{g_palet}_{fill}",
                            column_config={
                                f"{y_base} Ciro (₺)": st.column_config.NumberColumn(format="%.2f ₺"),
                                f"{y_base} Palet": st.column_config.NumberColumn(format="%.0f"),
                                "Ciro Büyüme (%)": st.column_config.NumberColumn(format="%.1f %%"),
                                "Palet Büyüme (%)": st.column_config.NumberColumn(format="%.1f %%")
                            })

    res = edited.copy()
    res["2027 Hedef Ciro (₺)"] = res[f"{y_base} Ciro (₺)"] * (1 + res["Ciro Büyüme (%)"] / 100)
    res["2027 Hedef Palet"] = res[f"{y_base} Palet"] * (1 + res["Palet Büyüme (%)"] / 100)

    pp, pc = pot_totals()
    inc_pot = st.checkbox("Potansiyel müşterilerin olasılık ağırlıklı katkısını ekle", value=True)
    t_ciro = res["2027 Hedef Ciro (₺)"].sum() + (pc if inc_pot else 0)
    t_palet = res["2027 Hedef Palet"].sum() + (pp if inc_pot else 0)
    b_ciro, b_palet = edited[f"{y_base} Ciro (₺)"].sum(), edited[f"{y_base} Palet"].sum()

    k = st.columns(4)
    k[0].metric(f"{y_base} Ciro (taban)", fmt_tl(b_ciro))
    k[1].metric("2027 Hedef Ciro", fmt_tl(t_ciro), f"%{pct(b_ciro, t_ciro):.1f}" if pct(b_ciro, t_ciro) is not None else None)
    k[2].metric(f"{y_base} Palet (taban)", tr_num(b_palet))
    k[3].metric("2027 Hedef Palet", tr_num(t_palet), f"%{pct(b_palet, t_palet):.1f}" if pct(b_palet, t_palet) is not None else None)

    def dist(met, total):
        s = fill_m[met]
        share = s / s.sum() if s.sum() > 0 else pd.Series(1 / 12, index=s.index)
        return share * total

    monthly = pd.DataFrame({
        "Ay": AYLAR,
        f"{y_base} Ciro (₺)": fill_m["Ciro"].values, "2027 Hedef Ciro (₺)": dist("Ciro", t_ciro).values,
        f"{y_base} Palet": fill_m["Palet"].values, "2027 Hedef Palet": dist("Palet", t_palet).values})

    fig = go.Figure([go.Bar(name=str(y_base), x=AYLAR, y=monthly[f"{y_base} Ciro (₺)"]),
                     go.Bar(name="2027 Hedef", x=AYLAR, y=monthly["2027 Hedef Ciro (₺)"])])
    fig.update_layout(barmode="group", title="Aylık Ciro: Taban vs 2027 Hedef")
    st.plotly_chart(fig, use_container_width=True)

    st.dataframe(
        monthly,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Ay": st.column_config.TextColumn("Ay"),
            f"{y_base} Ciro (₺)": st.column_config.NumberColumn(format="%.2f ₺"),
            "2027 Hedef Ciro (₺)": st.column_config.NumberColumn(format="%.2f ₺"),
            f"{y_base} Palet": st.column_config.NumberColumn(format="%.0f"),
            "2027 Hedef Palet": st.column_config.NumberColumn(format="%.0f")
        }
    )

    pot_rows = pd.concat([d.assign(Kategori=k_) for k_, d in ss["pot"].items() if len(d)]) \
        if any(len(d) for d in ss["pot"].values()) else pd.DataFrame()
    st.download_button("⬇️️ 2027 Hedeflerini Excel'e Aktar",
                       to_excel({"Aylık Hedef": monthly, "Müşteri Hedefleri": res,
                                 "Potansiyel": pot_rows if len(pot_rows) else pd.DataFrame({"Bilgi": ["Yok"]})}),
                       file_name="2027_hedefler.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ===================================================================
# MOD 4 — POTANSİYEL MÜŞTERİLER
# ===================================================================
else:
    st.header("🔍 Potansiyel Müşteriler & Kategori Bazlı 2027 Hedefleri")
    st.caption("Veriler oturum boyunca saklanır. Sayfayı kapatmadan önce Excel olarak indirin.")

    with st.expander("➕ Kategori ekle / 📤 Önceki listeyi geri yükle"):
        n1, n2 = st.columns(2)
        new_cat = n1.text_input("Yeni kategori adı")
        if n1.button("Kategori Ekle") and new_cat and new_cat not in ss["pot"]:
            ss["pot"][new_cat] = empty_pot()
            st.rerun()
        back = n2.file_uploader("İndirdiğiniz potansiyel listesi (.xlsx)", type=["xlsx"], key="pot_up")
        if back is not None and n2.button("Geri Yükle"):
            loaded = pd.read_excel(back, sheet_name=None, engine="openpyxl")
            for name, d in loaded.items():
                if set(POT_COLS).issubset(d.columns):
                    ss["pot"][name] = d[POT_COLS].copy()
            st.rerun()

    tabs = st.tabs(list(ss["pot"].keys()))
    for tab, cat in zip(tabs, list(ss["pot"].keys())):
        with tab:
            ed = st.data_editor(
                ss["pot"][cat], num_rows="dynamic", use_container_width=True, key=f"pot_{cat}",
                column_config={
                    "Durum": st.column_config.SelectboxColumn("Durum", options=DURUMLAR),
                    "Olasılık (%)": st.column_config.NumberColumn(min_value=0, max_value=100, step=5, format="%.0f %%"),
                    "2027 Tahmini Palet": st.column_config.NumberColumn(min_value=0, step=100, format="%.0f"),
                    "2027 Tahmini Ciro (₺)": st.column_config.NumberColumn(min_value=0, step=100000, format="%.2f ₺")})
            ss["pot"][cat] = ed
            w = ed["Olasılık (%)"].fillna(0) / 100
            c1, c2 = st.columns(2)
            c1.metric("Toplam Tahmini Ciro", fmt_tl(ed["2027 Tahmini Ciro (₺)"].fillna(0).sum()))
            c2.metric("Olasılık Ağırlıklı Ciro", fmt_tl((ed["2027 Tahmini Ciro (₺)"].fillna(0) * w).sum()))

    st.subheader("📊 Tüm Kategoriler Özeti")
    summ = pd.DataFrame([{
        "Kategori": c,
        "Firma Sayısı": len(d),
        "Tahmini Palet": d["2027 Tahmini Palet"].fillna(0).sum(),
        "Tahmini Ciro (₺)": d["2027 Tahmini Ciro (₺)"].fillna(0).sum(),
        "Ağırlıklı Ciro (₺)": (d["2027 Tahmini Ciro (₺)"].fillna(0) * d["Olasılık (%)"].fillna(0) / 100).sum()}
        for c, d in ss["pot"].items()])

    st.dataframe(
        summ,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Kategori": st.column_config.TextColumn("Kategori"),
            "Firma Sayısı": st.column_config.NumberColumn("Firma Sayısı", format="%d"),
            "Tahmini Palet": st.column_config.NumberColumn("Tahmini Palet", format="%.0f"),
            "Tahmini Ciro (₺)": st.column_config.NumberColumn("Tahmini Ciro", format="%.2f ₺"),
            "Ağırlıklı Ciro (₺)": st.column_config.NumberColumn("Ağırlıklı Ciro", format="%.2f ₺")
        }
    )

    if summ["Tahmini Ciro (₺)"].sum() > 0:
        fig = go.Figure([go.Bar(name="Tahmini", x=summ["Kategori"], y=summ["Tahmini Ciro (₺)"]),
                         go.Bar(name="Olasılık Ağırlıklı", x=summ["Kategori"], y=summ["Ağırlıklı Ciro (₺)"])])
        fig.update_layout(barmode="group", title="Kategori Bazlı 2027 Potansiyel Ciro")
        st.plotly_chart(fig, use_container_width=True)

    st.download_button("⬇️ Potansiyel Listesini Excel'e Aktar",
                       to_excel({c: d if len(d) else empty_pot() for c, d in ss["pot"].items()}),
                       file_name="potansiyel_musteriler.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
