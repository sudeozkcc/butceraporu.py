import io
import re
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Bütçe ve Ciro Raporu Portalı", page_icon="📊", layout="wide")

# ===================================================================
# Biçimlendirme ve Yardımcı Fonksiyonlar
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

def pct(a, b):
    return (b - a) / a * 100 if a and a != 0 else None

def to_excel(sheets: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for name, d in sheets.items():
            d.to_excel(w, sheet_name=name[:31], index=False)
    return buf.getvalue()

# ===================================================================
# Excel Parsing (Genel Veri Okuma)
# ===================================================================
@st.cache_data(show_spinner=False)
def get_sheet_names(file_bytes: bytes):
    return pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl").sheet_names

@st.cache_data(show_spinner=False)
def get_raw(file_bytes: bytes, sheet: str) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet, header=None, engine="openpyxl")

def find_metric(n: str):
    if "PALET" in n:
        return "Palet"
    if re.search(r"CIRO|TUTAR|\bTL\b|TRY|GELIR|SATIS", n):
        return "Ciro"
    return None

@st.cache_data(show_spinner=False)
def parse_general_sheet(file_bytes: bytes, sheet: str, default_year: int) -> pd.DataFrame:
    raw = get_raw(file_bytes, sheet)
    
    # Sütunlarda Ciro/Palet/Yıl bulma
    metric_cols = {}
    for j in range(raw.shape[1]):
        col_text = " ".join([str(v) for v in raw.iloc[:10, j] if pd.notna(v)])
        n = norm(col_text)
        met = find_metric(n)
        if met:
            ym = re.search(r"(20\d{2})", col_text)
            yr = int(ym.group(1)) if ym else default_year
            metric_cols[j] = (met, yr)
            
    if not metric_cols:
        return pd.DataFrame()
        
    # Müşteri sütununu tespit etme
    cust_col = 0
    for j in range(raw.shape[1]):
        if j in metric_cols:
            continue
        vals = raw.iloc[5:, j].dropna().astype(str)
        if len(vals) > 0 and sum(not v.replace('.', '').isdigit() for v in vals) / len(vals) > 0.5:
            cust_col = j
            break

    records = []
    for r in range(5, len(raw)):
        c_name = str(raw.iat[r, cust_col]).strip() if pd.notna(raw.iat[r, cust_col]) else ""
        if not c_name or "TOPLAM" in norm(c_name) or c_name.lower() in ("nan", "none"):
            continue
            
        for j, (met, yr) in metric_cols.items():
            val = to_num(raw.iat[r, j])
            records.append({"Müşteri": c_name, "Yıl": yr, "Metrik": met, "Değer": val})
            
    if not records:
        return pd.DataFrame()
        
    df_long = pd.DataFrame(records)
    piv = df_long.pivot_table(index=["Müşteri", "Yıl"], columns="Metrik", values="Değer", aggfunc="sum", fill_value=0.0).reset_index()
    if "Ciro" not in piv: piv["Ciro"] = 0.0
    if "Palet" not in piv: piv["Palet"] = 0.0
    return piv

# ===================================================================
# Streamlit Arayüz
# ===================================================================
ss = st.session_state
ss.setdefault("file_bytes", None)
ss.setdefault("file_name", None)

st.title("📊 Genel Ciro, Bütçe Raporlama & Hedef Sunum Portalı")
st.sidebar.header("📁 Rapor Yükleme")

up = st.sidebar.file_uploader("Excel Dosyasını Yükleyin (.xlsx)", type=["xlsx"])
if up is not None:
    ss["file_bytes"] = up.getvalue()
    ss["file_name"] = up.name

default_year = st.sidebar.number_input("Varsayılan Yıl", 2020, 2035, 2026)

if ss["file_bytes"] is None:
    st.info("👈 Başlamak için sol taraf menüden 2025/2026 verilerinizi içeren Excel dosyasını yükleyin.")
    st.stop()

FB = ss["file_bytes"]
sheet_names = get_sheet_names(FB)
sheet = st.sidebar.selectbox("Analiz Edilecek Sayfa:", sheet_names)

df_all = parse_general_sheet(FB, sheet, int(default_year))

if df_all.empty:
    st.error("Seçilen sayfada uygun Ciro/Palet verisi ayrıştırılamadı.")
    st.stop()

# ===================================================================
# GENEL RAKAMLAR VE İLK 20 MÜŞTERİ ANALİZİ
# ===================================================================
st.header("📈 Genel Ciro & Palet Performans Özeti")

years = sorted(df_all["Yıl"].unique())
sel_year = st.selectbox("🎯 Analiz Edilecek Yılı Seçin:", years, index=len(years)-1)

df_year = df_all[df_all["Yıl"] == sel_year].copy()

# En yüksek cirolu 20 Müşteriyi Filtreleme
top20_df = df_year.sort_values("Ciro", ascending=False).head(20).copy()
top20_df["Palet Başına TL"] = top20_df.apply(lambda r: (r["Ciro"] / r["Palet"]) if r["Palet"] > 0 else 0.0, axis=1)

# Genel Metrik Kartları
genel_ciro = top20_df["Ciro"].sum()
genel_palet = top20_df["Palet"].sum()
ort_palet_tl = (genel_ciro / genel_palet) if genel_palet > 0 else 0.0

c1, c2, c3 = st.columns(3)
c1.metric(f"{sel_year} Top 20 Toplam Ciro", fmt_tl(genel_ciro))
c2.metric(f"{sel_year} Top 20 Toplam Palet", fmt_palet(genel_palet))
c3.metric("Ortalama Palet Başı Ciro (₺/Palet)", fmt_tl(ort_palet_tl))

st.markdown("---")
st.subheader(f"📋 {sel_year} Yılı En Yüksek Cirolu 20 Müşteri Tablosu")

st.dataframe(
    top20_df,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Müşteri": st.column_config.TextColumn("Müşteri Adı", width="large"),
        "Yıl": st.column_config.NumberColumn("Yıl", format="%d"),
        "Ciro": st.column_config.NumberColumn("Toplam Ciro (₺)", format="%.2f ₺"),
        "Palet": st.column_config.NumberColumn("Toplam Palet", format="%.0f"),
        "Palet Başına TL": st.column_config.NumberColumn("Palet Başı Ort. TL", format="%.2f ₺")
    }
)

# ===================================================================
# PASTA GRAFİKLER & PAZAR PAYI
# ===================================================================
st.markdown("---")
st.subheader("🥧 İlk 20 Müşteri Ciro ve Palet Dağılımı (Pasta Grafik)")

g_col1, g_col2 = st.columns(2)

fig_ciro = px.pie(top20_df, values="Ciro", names="Müşteri", title=f"{sel_year} Ciro Dağılımı (Top 20)",
                  hole=0.3, color_discrete_sequence=px.colors.sequential.RdBu)
fig_ciro.update_traces(textposition='inside', textinfo='percent+label')
g_col1.plotly_chart(fig_ciro, use_container_width=True)

fig_palet = px.pie(top20_df, values="Palet", names="Müşteri", title=f"{sel_year} Palet Dağılımı (Top 20)",
                   hole=0.3, color_discrete_sequence=px.colors.sequential.Blugrn)
fig_palet.update_traces(textposition='inside', textinfo='percent+label')
g_col2.plotly_chart(fig_palet, use_container_width=True)

# Yorum Alanı
st.subheader("💡 Bütçe Raporu ve Sunum Yorumu")
top1 = top20_df.iloc[0]
top1_share = (top1["Ciro"] / genel_ciro) * 100 if genel_ciro > 0 else 0

st.info(f"""
* **Müşteri Konsantrasyonu:** En yüksek ciroyu getiren **{top1['Müşteri']}**, ilk 20 müşteri cirosunun tek başına **%{top1_share:.1f}** kadarını oluşturmaktadır (**{fmt_tl(top1['Ciro'])}**).
* **Verimlilik Analizi (Palet Başına Ciro):** Toplamda **{fmt_palet(genel_palet)}** palet sevkiyatı yapılmış olup, palet başına elde edilen ortalama gelir **{fmt_tl(ort_palet_tl)}** olarak gerçekleşmiştir.
* **Karlılık & Verim Odağı:** Palet başı getirisi ortalamanın altında kalan müşterilerde birim fiyat güncellemesi veya operasyonel verimlilik artışı hedeflenebilir.
""")

# ===================================================================
# YILLIK KARŞILAŞTIRMA & 2027 HEDEF OLUŞTURUCU
# ===================================================================
st.markdown("---")
st.header("🎯 Yıllık Karşılaştırma & Gelecek Yıl Bütçe Hedef Oluşturucu")

if len(years) >= 2:
    y_old, y_new = years[0], years[1]
    df_old = df_all[df_all["Yıl"] == y_old].groupby("Müşteri")[["Ciro", "Palet"]].sum().add_suffix(f"_{y_old}")
    df_new = df_all[df_all["Yıl"] == y_new].groupby("Müşteri")[["Ciro", "Palet"]].sum().add_suffix(f"_{y_new}")
    
    cmp_df = df_new.join(df_old, how="left").fillna(0.0).reset_index()
    cmp_df = cmp_df.sort_values(f"Ciro_{y_new}", ascending=False).head(20) # Sadece top 20
    
    cmp_df["Ciro Değişim (%)"] = cmp_df.apply(lambda r: pct(r[f"Ciro_{y_old}"], r[f"Ciro_{y_new}"]), axis=1)
    
    st.subheader(f"🔄 {y_old} vs {y_new} İlk 20 Müşteri Karşılaştırması")
    st.dataframe(cmp_df, use_container_width=True, hide_index=True)

# Hedef Verme Ekranı (Top 20 Müşteriye Hedef Simülasyonu)
st.subheader(f"🚀 Top 20 Müşteri İçin Bütçe Hedefleri Belirleme")

g_grow = st.slider("Genel Hedef Büyüme Oranı (%)", 0, 100, 20, 5)

hed_df = top20_df[["Müşteri", "Ciro", "Palet"]].copy()
hed_df["Büyüme Hedefi (%)"] = float(g_grow)

edited_target = st.data_editor(
    hed_df,
    hide_index=True,
    use_container_width=True,
    disabled=["Müşteri", "Ciro", "Palet"],
    column_config={
        "Ciro": st.column_config.NumberColumn(f"Mevcut Ciro ({sel_year})", format="%.2f ₺"),
        "Palet": st.column_config.NumberColumn(f"Mevcut Palet ({sel_year})", format="%.0f"),
        "Büyüme Hedefi (%)": st.column_config.NumberColumn("Hedef Büyüme (%)", format="%.1f %%")
    }
)

edited_target["Yeni Hedef Ciro (₺)"] = edited_target["Ciro"] * (1 + edited_target["Büyüme Hedefi (%)"] / 100)
edited_target["Yeni Hedef Palet"] = edited_target["Palet"] * (1 + edited_target["Büyüme Hedefi (%)"] / 100)

t_ciro = edited_target["Yeni Hedef Ciro (₺)"].sum()
t_palet = edited_target["Yeni Hedef Palet"].sum()

st.success(f"🎯 **Bütçelenen Toplam Yeni Ciro:** {fmt_tl(t_ciro)} | **Toplam Hedef Palet:** {fmt_palet(t_palet)}")

st.download_button(
    "⬇️ Bütçe Raporunu Excel Olarak İndir",
    to_excel({"Top 20 Müşteri Genel": top20_df, "Gelecek Yıl Bütçe Hedefleri": edited_target}),
    file_name="Butce_Raporu_Top20.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
