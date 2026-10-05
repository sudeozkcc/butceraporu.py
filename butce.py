import io
import re
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Lojistik Ciro & Bütçe Analizi", page_icon="🚚", layout="wide")

# ===================================================================
# BİÇİMLENDİRME & YARDIMCI FONKSİYONLAR
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
# EXCEL PARS SİSTEMİ
# ===================================================================
@st.cache_data(show_spinner=False)
def get_sheet_names(file_bytes: bytes):
    return pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl").sheet_names

@st.cache_data(show_spinner=False)
def parse_logistic_excel(file_bytes: bytes, sheet: str) -> pd.DataFrame:
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
            metric_cols[j] = met
            
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
            
        for j, met in metric_cols.items():
            val = to_num(raw.iat[r, j])
            records.append({"Müşteri": c_name, "Metrik": met, "Değer": val})
            
    if not records:
        return pd.DataFrame()
        
    piv = pd.DataFrame(records).pivot_table(
        index=["Müşteri"], columns="Metrik", values="Değer", aggfunc="sum", fill_value=0.0
    ).reset_index()
    
    if "Ciro" not in piv: piv["Ciro"] = 0.0
    if "Palet" not in piv: piv["Palet"] = 0.0
    return piv

# ===================================================================
# STREAMLIT ARAYÜZ (ÇİFT DOSYA YÜKLEME)
# ===================================================================
st.title("🚚 Lojistik Ciro, Palet & Bütçe Sunum Portalı")

st.sidebar.header("📁 Excel Dosya Yükleme")
up_2025 = st.sidebar.file_uploader("2025 Excel Raporunu Yükleyin (.xlsx)", type=["xlsx"], key="u25")
up_2026 = st.sidebar.file_uploader("2026 Excel Raporunu Yükleyin (.xlsx)", type=["xlsx"], key="u26")

df_2025 = pd.DataFrame()
df_2026 = pd.DataFrame()

# 2025 İşleme
if up_2025 is not None:
    sheets_25 = get_sheet_names(up_2025.getvalue())
    s_25 = st.sidebar.selectbox("2025 Sayfa Seçin:", sheets_25, key="s25")
    df_2025 = parse_logistic_excel(up_2025.getvalue(), s_25)

# 2026 İşleme
if up_2026 is not None:
    sheets_26 = get_sheet_names(up_2026.getvalue())
    s_26 = st.sidebar.selectbox("2026 Sayfa Seçin:", sheets_26, key="s26")
    df_2026 = parse_logistic_excel(up_2026.getvalue(), s_26)

mode = st.sidebar.radio("📌 Çalışma Modunu Seçin:", [
    "📊 2026'da Ne Yaptık?",
    "⚔️ 2025'e Göre Ne Değişti?",
    "🚀 2027'de Ne Yapacağız?"
])

# ===================================================================
# MOD 1: 2026'DA NE YAPTIK?
# ===================================================================
if mode == "📊 2026'da Ne Yaptık?":
    st.header("📊 2026'da Ne Yaptık? (Ciro ve Performans Özeti)")
    
    if df_2026.empty:
        st.info("👈 Lütfen sol menüden 2026 yılı Excel dosyanızı yükleyin.")
        st.stop()
        
    top20 = df_2026.sort_values("Ciro", ascending=False).head(20).copy()
    top20["Palet Başı TL"] = top20.apply(lambda r: (r["Ciro"] / r["Palet"]) if r["Palet"] > 0 else 0.0, axis=1)
    
    # Grafik
    st.subheader("📈 Müşteri Bazlı Ciro ve Palet Dağılımı")
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
    
    # Metrik Kartları
    tot_ciro = top20["Ciro"].sum()
    tot_palet = top20["Palet"].sum()
    avg_ptl = (tot_ciro / tot_palet) if tot_palet > 0 else 0.0
    
    c1, c2, c3 = st.columns(3)
    c1.metric("Palet Başı Ort. Gelir", fmt_tl(avg_ptl))
    c2.metric("2026 Top 20 Toplam Ciro", fmt_tl(tot_ciro))
    c3.metric("2026 Top 20 Toplam Palet", fmt_palet(tot_palet))
    
    st.markdown("---")
    
    # Yorumlar
    st.subheader("💡 2026 Performans Yorumlarımız ve Öne Çıkanlar")
    high_value_custs = top20[top20["Palet Başı TL"] > (avg_ptl * 1.15)]
    
    notes = [f"* **Genel Verimlilik:** 2026 genelinde palet başına ortalama **{fmt_tl(avg_ptl)}** ciro elde edilmiştir."]
    if not high_value_custs.empty:
        star_cust = high_value_custs.iloc[0]["Müşteri"]
        star_val = high_value_custs.iloc[0]["Palet Başı TL"]
        notes.append(f"* **Hızlı İvme Gösteren Müşteri:** **{star_cust}** palet başına **{fmt_tl(star_val)}** ile genel ortalamanın üzerinde verimlilik yakalamış, yıl içi hızlı ivmesiyle dikkat çekmiştir.")
    
    st.info("\n".join(notes))
    
    disp_top20 = pd.DataFrame({
        "Müşteri Unvanı": top20["Müşteri"],
        "Ciro": top20["Ciro"].apply(fmt_tl),
        "Palet": top20["Palet"].apply(fmt_palet),
        "Palet Başı Ort. TL": top20["Palet Başı TL"].apply(fmt_tl)
    })
    st.dataframe(disp_top20, use_container_width=True, hide_index=True)

# ===================================================================
# MOD 2: 2025'E GÖRE NE DEĞİŞTİ?
# ===================================================================
elif mode == "⚔️ 2025'e Göre Ne Değişti?":
    st.header("⚔️ 2025'e Göre Ne Değişti? (Karşılaştırma Analizi)")
    
    # KONTROL: 2025 verisi var mı?
    if df_2025.empty:
        st.warning("⚠️ **Elinizde 2025 yılı verisi bulunmuyor.** 2025'e göre ne yaptığımızı karşılaştırabilmek için lütfen sol menüden **2025 Excel Raporunu** da yükleyin.")
        st.stop()
        
    if df_2026.empty:
        st.warning("⚠️ Lütfen karşılaştırma için **2026 Excel Raporunu** da yükleyin.")
        st.stop()

    df_25_m = df_2025.set_index("Müşteri").add_suffix("_2025")
    df_26_m = df_2026.set_index("Müşteri").add_suffix("_2026")
    
    cmp = df_26_m.join(df_25_m, how="outer").fillna(0.0).reset_index()
    top20_cmp = cmp.sort_values("Ciro_2026", ascending=False).head(20).copy()
    
    st.subheader("📈 Noktasal Ciro Değişimi (2025 vs 2026)")
    fig_cmp = go.Figure()
    fig_cmp.add_trace(go.Scatter(x=top20_cmp["Müşteri"], y=top20_cmp["Ciro_2025"], mode='markers+lines', name="2025 Ciro", marker=dict(size=10, color='orange')))
    fig_cmp.add_trace(go.Scatter(x=top20_cmp["Müşteri"], y=top20_cmp["Ciro_2026"], mode='markers+lines', name="2026 Ciro", marker=dict(size=10, color='green')))
    fig_cmp.update_layout(xaxis=dict(tickangle=-45), yaxis_title="Ciro (₺)")
    st.plotly_chart(fig_cmp, use_container_width=True)
    
    top20_cmp["Ciro Farkı (₺)"] = top20_cmp["Ciro_2026"] - top20_cmp["Ciro_2025"]
    
    new_stars = top20_cmp[(top20_cmp["Ciro_2025"] == 0) & (top20_cmp["Ciro_2026"] > 0)]
    if not new_stars.empty:
        st.success(f"🚀 **2026'da Portföye Katılan Hızlı İvmeli Müşteriler:** {', '.join(new_stars['Müşteri'].tolist())} 2025'te olmayıp 2026'da portföye katılarak yüksek ivme yakalamıştır.")
        
    st.subheader("📋 Top 20 Müşteri Değişim Tablosu")
    disp_cmp = pd.DataFrame({
        "Müşteri Unvanı": top20_cmp["Müşteri"],
        "2025 Ciro": top20_cmp["Ciro_2025"].apply(fmt_tl),
        "2026 Ciro": top20_cmp["Ciro_2026"].apply(fmt_tl),
        "Fark (₺)": top20_cmp["Ciro Farkı (₺)"].apply(fmt_tl),
        "2025 Palet": top20_cmp["Palet_2025"].apply(fmt_palet),
        "2026 Palet": top20_cmp["Palet_2026"].apply(fmt_palet)
    })
    st.dataframe(disp_cmp, use_container_width=True, hide_index=True)

# ===================================================================
# MOD 3: 2027'DE NE YAPACAĞIZ?
# ===================================================================
elif mode == "🚀 2027'de Ne Yapacağız?":
    st.header("🚀 2027'de Ne Yapacağız?")
    
    if df_2026.empty:
        st.info("👈 Hedefleri belirlemek için sol menüden 2026 Excel dosyanızı yükleyin.")
        st.stop()
        
    st.write("2027 yılı için **en önemli 10 müşterimizden ve yıl içi yüksek ivme yakalayan isimlerden** beklentilerimizi simüle ediyoruz.")
    
    top10 = df_2026.sort_values("Ciro", ascending=False).head(10).copy()
    top10["2027 Hedef Büyüme (%)"] = 25.0
    
    edited_10 = st.data_editor(
        top10[["Müşteri", "Ciro", "Palet", "2027 Hedef Büyüme (%)"]],
        use_container_width=True,
        hide_index=True,
        disabled=["Müşteri", "Ciro", "Palet"],
        column_config={
            "Müşteri": st.column_config.TextColumn("Hedef Müşterilerimiz (Top 10)", width="large"),
            "Ciro": st.column_config.NumberColumn("2026 Mevcut Ciro", format="%.2f ₺"),
            "Palet": st.column_config.NumberColumn("2026 Mevcut Palet", format="%.0f"),
            "2027 Hedef Büyüme (%)": st.column_config.NumberColumn("2027 Beklenen Büyüme (%)", format="%.1f %%")
        }
    )
    
    edited_10["2027 Beklenen Ciro (₺)"] = edited_10["Ciro"] * (1 + edited_10["2027 Hedef Büyüme (%)"] / 100)
    edited_10["2027 Beklenen Palet"] = edited_10["Palet"] * (1 + edited_10["2027 Hedef Büyüme (%)"] / 100)
    
    t_ciro_27 = edited_10["2027 Beklenen Ciro (₺)"].sum()
    t_palet_27 = edited_10["2027 Beklenen Palet"].sum()
    avg_ptl_27 = (t_ciro_27 / t_palet_27) if t_palet_27 > 0 else 0.0
    
    st.markdown("---")
    st.subheader("🎯 2027 İÇİN ŞUNLARI HEDEFLİYORUZ:")
    
    k1, k2, k3 = st.columns(3)
    k1.metric("2027 Hedeflenen Ciro", fmt_tl(t_ciro_27))
    k2.metric("2027 Hedeflenen Palet", fmt_palet(t_palet_27))
    k3.metric("2027 Palet Başı Ort. Gelir", fmt_tl(avg_ptl_27))
    
    st.info(f"📌 **2027 Bütçe Beklentilerimiz:** 2026 yılında ana ciro omurgasını oluşturan en önemli 10 müşterimiz ve yıl içerisinde yakalanan ivmeyle birlikte 2027'de toplam **{fmt_tl(t_ciro_27)}** ciro ve **{fmt_palet(t_palet_27)}** palet hacmine ulaşmayı hedefliyoruz.")
