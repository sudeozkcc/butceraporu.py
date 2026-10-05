import io
import re
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Lojistik Ciro & Bütçe Analizi", page_icon="🚚", layout="wide")

# ===================================================================
# YARDIMCI BİÇİMLENDİRME FONKSİYONLARI
# ===================================================================
_TR = str.maketrans("İıŞşĞğÜüÖöÇç", "IISSGGUUOOCC")

def norm(x) -> str:
    if x is None or (not isinstance(x, (list, tuple)) and pd.isna(x)):
        return ""
    return str(x).translate(_TR).upper().strip()

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

# ===================================================================
# EXCEL VERİ OKUMA SİSTEMİ
# ===================================================================
@st.cache_data(show_spinner=False)
def get_sheet_names(file_bytes: bytes):
    return pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl").sheet_names

@st.cache_data(show_spinner=False)
def parse_logistic_excel(file_bytes: bytes, sheet: str, default_year: int) -> pd.DataFrame:
    raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet, header=None, engine="openpyxl")
    
    metric_cols = {}
    for j in range(raw.shape[1]):
        col_text = " ".join([str(v) for v in raw.iloc[:10, j] if pd.notna(v)])
        n = norm(col_text)
        
        met = None
        if "PALET" in n:
            met = "Palet"
        elif re.search(r"CIRO|TUTAR|\bTL\b|TRY|GELIR|SATIS", n):
            met = "Ciro"
            
        if met:
            ym = re.search(r"\b(202[0-9])\b", col_text)
            yr = int(ym.group(1)) if ym else default_year
            metric_cols[j] = (met, yr)
            
    if not metric_cols:
        return pd.DataFrame()
        
    cust_col = 0
    for j in range(raw.shape[1]):
        if j in metric_cols:
            continue
        vals = raw.iloc[3:, j].dropna().astype(str)
        if len(vals) > 0 and sum(not v.replace('.', '').isdigit() for v in vals) / len(vals) > 0.5:
            cust_col = j
            break

    records = []
    for r in range(3, len(raw)):
        c_name = str(raw.iat[r, cust_col]).strip() if pd.notna(raw.iat[r, cust_col]) else ""
        if not c_name or "TOPLAM" in norm(c_name) or c_name.lower() in ("nan", "none"):
            continue
            
        for j, (met, yr) in metric_cols.items():
            val = to_num(raw.iat[r, j])
            records.append({"Müşteri": c_name, "Yıl": yr, "Metrik": met, "Değer": val})
            
    if not records:
        return pd.DataFrame()
        
    piv = pd.DataFrame(records).pivot_table(
        index=["Müşteri", "Yıl"], columns="Metrik", values="Değer", aggfunc="sum", fill_value=0.0
    ).reset_index()
    
    if "Ciro" not in piv: piv["Ciro"] = 0.0
    if "Palet" not in piv: piv["Palet"] = 0.0
    return piv

# ===================================================================
# STREAMLIT ARAYÜZ
# ===================================================================
st.title("🚚 Lojistik Ciro, Palet & Bütçe Sunum Portalı")

up = st.sidebar.file_uploader("Excel Raporunu Yükleyin (.xlsx)", type=["xlsx"])
default_yr = st.sidebar.number_input("Varsayılan Yıl (Başlıkta Yoksa)", 2020, 2030, 2025)

if up is None:
    st.info("👈 Lütfen sol menüden Excel dosyanızı yükleyin.")
    st.stop()

FB = up.getvalue()
sheet_names = get_sheet_names(FB)
sheet = st.sidebar.selectbox("Sayfa Seçin:", sheet_names)

df_all = parse_logistic_excel(FB, sheet, int(default_yr))

if df_all.empty:
    st.error("Excel dosyasında Ciro/Palet verileri ayrıştırılamadı. Sütun başlıklarını kontrol edin.")
    st.stop()

mode = st.sidebar.radio("📌 Çalışma Modunu Seçin:", [
    "📊 2025 Yılı Ciro Analizi",
    "📊 2026 Yılı Ciro Analizi",
    "⚔️ 2025 ve 2026 Karşılaştırması",
    "🚀 2027 Hedeflerimiz (Top 10 Müşteri Beklentisi)"
])

# ===================================================================
# TEKİL YIL ANALİZ FONKSİYONU
# ===================================================================
def render_single_year(target_year: int):
    st.header(f"📊 {target_year} YILI CİRO ANALİZİ")
    
    df_yr = df_all[df_all["Yıl"] == target_year].copy()
    if df_yr.empty:
        st.warning(f"{target_year} yılına ait veri bulunamadı.")
        return
        
    top20 = df_yr.sort_values("Ciro", ascending=False).head(20).copy()
    top20["Palet Başı TL"] = top20.apply(lambda r: (r["Ciro"] / r["Palet"]) if r["Palet"] > 0 else 0.0, axis=1)
    
    # 1. Noktasal Grafik (Ciro ve Palet)
    st.subheader(f"📈 {target_year} Yılı Noktasal Ciro ve Palet Grafiği")
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=top20["Müşteri"], y=top20["Ciro"],
        mode='markers+lines', name="Ciro (₺)",
        marker=dict(size=10, color='#1f77b4'),
        yaxis="y1"
    ))
    fig.add_trace(go.Scatter(
        x=top20["Müşteri"], y=top20["Palet"],
        mode='markers+lines', name="Palet Adedi",
        marker=dict(size=10, color='#ff7f0e'),
        yaxis="y2"
    ))
    
    fig.update_layout(
        xaxis=dict(title="Müşteriler", tickangle=-45),
        yaxis=dict(title="Ciro (₺)", side="left"),
        yaxis2=dict(title="Palet Adedi", side="right", overlaying="y", showgrid=False),
        legend=dict(x=0.8, y=1.1, orientation="h"),
        margin=dict(l=20, r=20, t=30, b=100)
    )
    st.plotly_chart(fig, use_container_width=True)
    
    # 2. Palet Başı Ortalama Tutar Kartları
    tot_ciro = top20["Ciro"].sum()
    tot_palet = top20["Palet"].sum()
    avg_ptl = (tot_ciro / tot_palet) if tot_palet > 0 else 0.0
    
    c1, c2, c3 = st.columns(3)
    c1.metric("Palet Başı Ort. Gelir", fmt_tl(avg_ptl))
    c2.metric(f"{target_year} Top 20 Toplam Ciro", fmt_tl(tot_ciro))
    c3.metric(f"{target_year} Top 20 Toplam Palet", fmt_palet(tot_palet))
    
    st.markdown("---")
    
    # 3. Top 20 Müşteri Tablosu
    st.subheader(f"🏆 {target_year} Yılı Top 20 Müşterimiz")
    st.dataframe(
        top20,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Müşteri": st.column_config.TextColumn("Müşteri Unvanı", width="large"),
            "Ciro": st.column_config.NumberColumn("Ciro", format="%.2f ₺"),
            "Palet": st.column_config.NumberColumn("Palet", format="%.0f"),
            "Palet Başı TL": st.column_config.NumberColumn("Palet Başı Ort. TL", format="%.2f ₺")
        }
    )

# ===================================================================
# MOD SEÇİMLERİ
# ===================================================================
if mode == "📊 2025 Yılı Ciro Analizi":
    render_single_year(2025)

elif mode == "📊 2026 Yılı Ciro Analizi":
    render_single_year(2026)

elif mode == "⚔️ 2025 ve 2026 Karşılaştırması":
    st.header("⚔️ 2025 vs 2026 KARŞILAŞTIRMA ANALİZİ")
    
    df_25 = df_all[df_all["Yıl"] == 2025].groupby("Müşteri")[["Ciro", "Palet"]].sum().add_suffix("_2025")
    df_26 = df_all[df_all["Yıl"] == 2026].groupby("Müşteri")[["Ciro", "Palet"]].sum().add_suffix("_2026")
    
    cmp = df_26.join(df_25, how="outer").fillna(0.0).reset_index()
    top20_cmp = cmp.sort_values("Ciro_2026", ascending=False).head(20)
    
    st.subheader("📈 Noktasal Ciro Karşılaştırması (2025 vs 2026)")
    fig_cmp = go.Figure()
    fig_cmp.add_trace(go.Scatter(x=top20_cmp["Müşteri"], y=top20_cmp["Ciro_2025"], mode='markers+lines', name="2025 Ciro", marker=dict(size=10, color='orange')))
    fig_cmp.add_trace(go.Scatter(x=top20_cmp["Müşteri"], y=top20_cmp["Ciro_2026"], mode='markers+lines', name="2026 Ciro", marker=dict(size=10, color='green')))
    fig_cmp.update_layout(xaxis=dict(tickangle=-45), yaxis_title="Ciro (₺)")
    st.plotly_chart(fig_cmp, use_container_width=True)
    
    top20_cmp["Ciro Farkı (₺)"] = top20_cmp["Ciro_2026"] - top20_cmp["Ciro_2025"]
    
    st.subheader("📋 Top 20 Müşteri Performans Değişimi")
    st.dataframe(
        top20_cmp,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Müşteri": st.column_config.TextColumn("Müşteri Unvanı", width="large"),
            "Ciro_2025": st.column_config.NumberColumn("2025 Ciro", format="%.2f ₺"),
            "Ciro_2026": st.column_config.NumberColumn("2026 Ciro", format="%.2f ₺"),
            "Ciro Farkı (₺)": st.column_config.NumberColumn("Fark (₺)", format="%.2f ₺"),
            "Palet_2025": st.column_config.NumberColumn("2025 Palet", format="%.0f"),
            "Palet_2026": st.column_config.NumberColumn("2026 Palet", format="%.0f")
        }
    )

elif mode == "🚀 2027 Hedeflerimiz (Top 10 Müşteri Beklentisi)":
    st.header("🚀 2027 BÜTÇE HEDEFLERİ (TOP 10 MÜŞTERİ)")
    
    last_yr = 2026 if 2026 in df_all["Yıl"].values else df_all["Yıl"].max()
    base_df = df_all[df_all["Yıl"] == last_yr].sort_values("Ciro", ascending=False).head(10).copy()
    
    if base_df.empty:
        st.warning("Hedef belirlemek için geçmiş yıl verisi bulunamadı.")
        st.stop()
        
    base_df["2027 Hedef Büyüme (%)"] = 20.0
    
    edited_10 = st.data_editor(
        base_df[["Müşteri", "Ciro", "Palet", "2027 Hedef Büyüme (%)"]],
        use_container_width=True,
        hide_index=True,
        disabled=["Müşteri", "Ciro", "Palet"],
        column_config={
            "Müşteri": st.column_config.TextColumn("En Önemli 10 Müşteri", width="large"),
            "Ciro": st.column_config.NumberColumn(f"Mevcut Ciro ({last_yr})", format="%.2f ₺"),
            "Palet": st.column_config.NumberColumn(f"Mevcut Palet ({last_yr})", format="%.0f"),
            "2027 Hedef Büyüme (%)": st.column_config.NumberColumn("Beklenen Büyüme (%)", format="%.1f %%")
        }
    )
    
    edited_10["2027 Beklenen Ciro (₺)"] = edited_10["Ciro"] * (1 + edited_10["2027 Hedef Büyüme (%)"] / 100)
    edited_10["2027 Beklenen Palet"] = edited_10["Palet"] * (1 + edited_10["2027 Hedef Büyüme (%)"] / 100)
    
    t_ciro_27 = edited_10["2027 Beklenen Ciro (₺)"].sum()
    t_palet_27 = edited_10["2027 Beklenen Palet"].sum()
    avg_ptl_27 = (t_ciro_27 / t_palet_27) if t_palet_27 > 0 else 0.0
    
    st.markdown("---")
    st.subheader("🎯 Top 10 Müşteri 2027 Konsolide Hedefleri")
    
    k1, k2, k3 = st.columns(3)
    k1.metric("2027 Hedeflenen Ciro", fmt_tl(t_ciro_27))
    k2.metric("2027 Hedeflenen Palet", fmt_palet(t_palet_27))
    k3.metric("2027 Palet Başı Ort. Gelir", fmt_tl(avg_ptl_27))
