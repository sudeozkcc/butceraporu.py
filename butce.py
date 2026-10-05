import io
import re
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Lojistik Bütçe & Ciro Portal", page_icon="🚚", layout="wide")

# ===================================================================
# YARDIMCI FONKSİYONLAR & BİÇİMLENDİRME
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
# EXCEL OKUMA & YAPI AYRIŞTIRMA
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
# OTURUM HAFIZASI (POTANSİYEL MÜŞTERİLER)
# ===================================================================
ss = st.session_state
ss.setdefault("file_bytes", None)
ss.setdefault("file_name", None)

POT_CATEGORIES = ["Süt & Süt Ürünleri", "Donuk Gıda", "Et & Et Ürünleri", "Kuru Gıda & Lojistik", "Soğuk Zincir"]
if "pot_data" not in ss:
    ss["pot_data"] = {cat: pd.DataFrame(columns=["Firma Adı", "Olasılık (%)", "2027 Hedef Palet", "2027 Hedef Ciro (₺)"]) for cat in POT_CATEGORIES}

# ===================================================================
# YAN MENÜ & DOSYA YÜKLEME
# ===================================================================
st.title("🚚 Lojistik Ciro, Palet Verimliliği & Bütçe Portalı")
st.sidebar.header("📁 Veri Yükleme")

up = st.sidebar.file_uploader("Excel Raporunu Yükleyin (.xlsx)", type=["xlsx"])
if up is not None:
    ss["file_bytes"] = up.getvalue()
    ss["file_name"] = up.name

default_year = st.sidebar.number_input("Varsayılan Yıl (Başlıkta yazmıyorsa)", 2020, 2035, 2026)

mode = st.sidebar.radio("📌 Çalışma Modunu Seçin:", [
    "📊 Yıllık Bazda Durum (2025 veya 2026 Özeti)",
    "⚔️ Yıllar Arası Karşılaştırma (2025 vs 2026)",
    "🚀 2027 Yılı Bütçesi ve Hedefler",
    "🎯 Kategorik Potansiyel Müşteri Yönetimi"
])

if ss["file_bytes"] is None and mode != "🎯 Kategorik Potansiyel Müşteri Yönetimi":
    st.info("👈 Analize başlamak için tekil yıl veya karşılaştırmalı Excel dosyanızı sol menüden yükleyin.")
    st.stop()

df_all = pd.DataFrame()
if ss["file_bytes"] is not None:
    FB = ss["file_bytes"]
    sheet_names = get_sheet_names(FB)
    sheet = st.sidebar.selectbox("Analiz Edilecek Sayfa:", sheet_names)
    df_all = parse_general_sheet(FB, sheet, int(default_year))

# ===================================================================
# MOD 1: YILLIK BAZDA DURUM (2025 veya 2026 ODAKLI)
# ===================================================================
if mode.startswith("📊"):
    st.header("📊 Yıllık Lojistik Operasyon ve Ciro Analizi")
    
    available_years = sorted(df_all["Yıl"].unique())
    sel_year = st.selectbox("İncelemek İstediğiniz Yılı Seçin:", available_years, index=len(available_years)-1)
    
    df_yr = df_all[df_all["Yıl"] == sel_year].copy()
    top20 = df_yr.sort_values("Ciro", ascending=False).head(20).copy()
    top20["Palet Başı TL"] = top20.apply(lambda r: (r["Ciro"] / r["Palet"]) if r["Palet"] > 0 else 0.0, axis=1)
    
    t_ciro = top20["Ciro"].sum()
    t_palet = top20["Palet"].sum()
    avg_ptl = (t_ciro / t_palet) if t_palet > 0 else 0.0
    
    # Metrik Kartları
    m1, m2, m3 = st.columns(3)
    m1.metric(f"{sel_year} Top 20 Toplam Ciro", fmt_tl(t_ciro))
    m2.metric(f"{sel_year} Top 20 Toplam Palet", fmt_palet(t_palet))
    m3.metric("Palet Başına Ort. Gelir", fmt_tl(avg_ptl))
    
    st.markdown("---")
    st.subheader(f"📋 {sel_year} Yılı En Yüksek Cirolu İlk 20 Müşteri")
    st.dataframe(
        top20,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Müşteri": st.column_config.TextColumn("Müşteri Adı", width="large"),
            "Ciro": st.column_config.NumberColumn("Toplam Ciro (₺)", format="%.2f ₺"),
            "Palet": st.column_config.NumberColumn("Toplam Palet", format="%.0f"),
            "Palet Başı TL": st.column_config.NumberColumn("Palet Başı Fiyat (₺/Palet)", format="%.2f ₺")
        }
    )
    
    # Grafik & Yorumlar
    st.subheader("🥧 Ciro ve Palet Dağılımları (Top 20)")
    g1, g2 = st.columns(2)
    
    fig_c = px.pie(top20, values="Ciro", names="Müşteri", title="Ciro Payları (%)", hole=0.3)
    fig_c.update_traces(textposition='inside', textinfo='percent+label')
    g1.plotly_chart(fig_c, use_container_width=True)
    
    fig_p = px.pie(top20, values="Palet", names="Müşteri", title="Palet Hacim Payları (%)", hole=0.3)
    fig_p.update_traces(textposition='inside', textinfo='percent+label')
    g2.plotly_chart(fig_p, use_container_width=True)
    
    top_c = top20.iloc[0] if len(top20) else None
    if top_c is not None:
        st.info(f"""
        💡 **Stratejik Sunum Notları ({sel_year}):**
        * En büyük lojistik hacmini veren **{top_c['Müşteri']}**, toplam Top 20 cirosunun **%{ (top_c['Ciro']/t_ciro)*100:.1f}** kadarını oluşturmuştur.
        * Firmamız {sel_year} yılında ortalama palet başına **{fmt_tl(avg_ptl)}** taşıma geliri elde etmiştir.
        * Palet başı getirisi bu ortalamanın altında kalan müşteriler ile yeni dönemde birim fiyat revizyonu görüşülebilir.
        """)

# ===================================================================
# MOD 2: YILLAR ARASI KARŞILAŞTIRMA (2025 vs 2026)
# ===================================================================
elif mode.startswith("⚔️"):
    st.header("⚔️ Yıllar Arası Ciro & Palet Karşılaştırması")
    
    years = sorted(df_all["Yıl"].unique())
    if len(years) < 2:
        st.warning("⚠️ Karşılaştırma yapabilmek için dosyanızda en az 2 farklı yılın (örn: 2025 ve 2026) verisi bulunmalıdır.")
        st.stop()
        
    c1, c2 = st.columns(2)
    y_base = c1.selectbox("Baz Yıl Seçin:", years, index=0)
    y_cmp = c2.selectbox("Karşılaştırılacak Yıl Seçin:", years, index=len(years)-1)
    
    df_b = df_all[df_all["Yıl"] == y_base].groupby("Müşteri")[["Ciro", "Palet"]].sum().add_suffix(f"_{y_base}")
    df_c = df_all[df_all["Yıl"] == y_cmp].groupby("Müşteri")[["Ciro", "Palet"]].sum().add_suffix(f"_{y_cmp}")
    
    cmp = df_c.join(df_b, how="outer").fillna(0.0).reset_index()
    cmp["Ciro Farkı (₺)"] = cmp[f"Ciro_{y_cmp}"] - cmp[f"Ciro_{y_base}"]
    cmp["Ciro Değişim (%)"] = cmp.apply(lambda r: pct(r[f"Ciro_{y_base}"], r[f"Ciro_{y_cmp}"]), axis=1)
    
    top20_cmp = cmp.sort_values(f"Ciro_{y_cmp}", ascending=False).head(20)
    
    # Noktasal Karşılaştırma Grafiği
    st.subheader(f"📈 Top 20 Müşteri Ciro Değişimi ({y_base} vs {y_cmp})")
    fig_dot = go.Figure()
    fig_dot.add_trace(go.Scatter(x=top20_cmp["Müşteri"], y=top20_cmp[f"Ciro_{y_base}"], mode='markers+lines', name=str(y_base), marker=dict(size=10, color='orange')))
    fig_dot.add_trace(go.Scatter(x=top20_cmp["Müşteri"], y=top20_cmp[f"Ciro_{y_cmp}"], mode='markers+lines', name=str(y_cmp), marker=dict(size=10, color='green')))
    fig_dot.update_layout(title="Müşteri Ciro Karşılaştırma Noktaları", yaxis_title="Ciro (₺)")
    st.plotly_chart(fig_dot, use_container_width=True)
    
    # Kırmızı/Yeşil Renklendirilmiş Değişim Tablosu
    st.subheader("📋 Müşteri Bazlı Performans Değişimi (Top 20)")
    
    def highlight_change(val):
        if pd.isna(val): return ''
        color = '#d4edda' if val > 0 else ('#f8d7da' if val < 0 else '')
        text_color = 'green' if val > 0 else ('red' if val < 0 else 'black')
        return f'background-color: {color}; color: {text_color}; font-weight: bold;'

    styled_df = top20_cmp.style.applymap(highlight_change, subset=["Ciro Farkı (₺)", "Ciro Değişim (%)"])\
        .format({
            f"Ciro_{y_base}": "{:,.2f} ₺", f"Ciro_{y_cmp}": "{:,.2f} ₺", "Ciro Farkı (₺)": "{:,.2f} ₺",
            f"Palet_{y_base}": "{:,.0f}", f"Palet_{y_cmp}": "{:,.0f}", "Ciro Değişim (%)": "{:.1f} %"
        })
        
    st.dataframe(styled_df, use_container_width=True, hide_index=True)

# ===================================================================
# MOD 3: 2027 YILI BÜTÇESİ VE HEDEFLER
# ===================================================================
elif mode.startswith("🚀"):
    st.header("🚀 2027 Lojistik Bütçesi ve Büyüme Simülasyonu")
    
    years = sorted(df_all["Yıl"].unique())
    base_yr = st.selectbox("Hedefler Hangi Yılın Üzerine Büyüme Olarak Hesaplansın?", years, index=len(years)-1)
    
    df_b = df_all[df_all["Yıl"] == base_yr].sort_values("Ciro", ascending=False).head(20).copy()
    
    g_ratio = st.slider("Genel Lojistik Büyüme Hedefi (%)", 0, 100, 25, 5)
    
    df_b["Büyüme (%)"] = float(g_ratio)
    
    edited = st.data_editor(
        df_b[["Müşteri", "Ciro", "Palet", "Büyüme (%)"]],
        use_container_width=True,
        hide_index=True,
        disabled=["Müşteri", "Ciro", "Palet"],
        column_config={
            "Ciro": st.column_config.NumberColumn(f"Mevcut Ciro ({base_yr})", format="%.2f ₺"),
            "Palet": st.column_config.NumberColumn(f"Mevcut Palet ({base_yr})", format="%.0f"),
            "Büyüme (%)": st.column_config.NumberColumn("2027 Hedef Büyüme (%)", format="%.1f %%")
        }
    )
    
    edited["2027 Hedef Ciro (₺)"] = edited["Ciro"] * (1 + edited["Büyüme (%)"] / 100)
    edited["2027 Hedef Palet"] = edited["Palet"] * (1 + edited["Büyüme (%)"] / 100)
    
    t_c = edited["2027 Hedef Ciro (₺)"].sum()
    t_p = edited["2027 Hedef Palet"].sum()
    
    st.success(f"🎯 **2027 Yılı Toplam Bütçelenen Ciro:** {fmt_tl(t_c)} | **Hedef Sevk Paleti:** {fmt_palet(t_p)}")
    
    st.download_button(
        "⬇️ 2027 Bütçe Raporunu Excel Olarak İndir",
        to_excel({"2027 Müşteri Hedefleri": edited}),
        file_name="2027_Lojistik_Butcesi.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ===================================================================
# MOD 4: KATEGORİK POTANSİYEL MÜŞTERİ YÖNETİMİ
# ===================================================================
else:
    st.header("🎯 Kategorize Potansiyel Müşteri ve Sektörel Hedefler")
    st.caption("Süt Grubu, Donuk Gıda, Et Grubu gibi lojistik kategorilerinize özel yeni fırsatları yönetin.")
    
    tabs = st.tabs(POT_CATEGORIES)
    
    for i, cat in enumerate(POT_CATEGORIES):
        with tabs[i]:
            st.subheader(f"📦 {cat} Sektörü Potansiyel Müşterileri")
            
            df_cat = st.data_editor(
                ss["pot_data"][cat],
                num_rows="dynamic",
                use_container_width=True,
                key=f"editor_{cat}",
                column_config={
                    "Firma Adı": st.column_config.TextColumn("Firma Adı"),
                    "Olasılık (%)": st.column_config.NumberColumn("Kazanma Olasılığı (%)", format="%.0f %%", min_value=0, max_value=100),
                    "2027 Hedef Palet": st.column_config.NumberColumn("Tahmini Palet Hacmi", format="%.0f"),
                    "2027 Hedef Ciro (₺)": st.column_config.NumberColumn("Tahmini Ciro (₺)", format="%.2f ₺")
                }
            )
            ss["pot_data"][cat] = df_cat
            
            # Kategori içi hesaplama
            w = df_cat["Olasılık (%)"].fillna(0) / 100
            weighted_ciro = (df_cat["2027 Hedef Ciro (₺)"].fillna(0) * w).sum()
            st.metric(f"{cat} - Olasılık Ağırlıklı Tahmini Ciro", fmt_tl(weighted_ciro))

    st.markdown("---")
    st.download_button(
        "⬇️ Tüm Potansiyel Müşteri Listesini Excel Olarak İndir",
        to_excel({cat: df if len(df) else pd.DataFrame(columns=["Firma Adı", "Olasılık (%)", "2027 Hedef Palet", "2027 Hedef Ciro (₺)"]) for cat, df in ss["pot_data"].items()}),
        file_name="Potansiyel_Lojistik_Musterileri.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
