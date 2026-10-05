import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import tempfile
import os

# Sayfa Yapılandırması
st.set_page_config(
    page_title="Ciro & Bütçeleme Analiz Portalı",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Bütçeleme, Ciro Analizi ve Hedefleme Portalı")

# -------------------------------------------------------------------
# 1. SESSION STATE (OTURUM DURUMU) KURULUMU
# -------------------------------------------------------------------
if "file_bytes" not in st.session_state:
    st.session_state["file_bytes"] = None
if "file_name" not in st.session_state:
    st.session_state["file_name"] = None

# Yan Menü: Dosya Yükleme & Modlar
st.sidebar.header("📁 Veri Yükleme & Modlar")

uploaded_file = st.sidebar.file_uploader(
    "Excel Raporunu Yükleyin (.xlsx)", 
    type=["xlsx", "xls"],
    key="excel_uploader"
)

# Yeni bir dosya yüklendiğinde oturuma kaydet
if uploaded_file is not None:
    st.session_state["file_bytes"] = uploaded_file.getvalue()
    st.session_state["file_name"] = uploaded_file.name

# Yüklü dosyayı kaldırma butonu
if st.session_state["file_bytes"] is not None:
    if st.sidebar.button("🗑️ Dosyayı Kaldır / Yenile"):
        st.session_state["file_bytes"] = None
        st.session_state["file_name"] = None
        st.rerun()

mode = st.sidebar.radio(
    "Çalışma Modunu Seçin:",
    [
        "📈 2026 Ciro Analizi (Aylık Karşılaştırma)",
        "🎯 2025 - 2026 Hedef Karşılaştırması",
        "🚀 2027 Hedef Oluşturucu (Simülasyon)",
        "🔍 Potansiyel Müşteriler & Kategori Hedefleri"
    ]
)

# -------------------------------------------------------------------
# 2. DOSYA KONTROLÜ VE GEÇİCİ DOSYA OLUŞTURMA (HATA ÖNLEYİCİ)
# -------------------------------------------------------------------
if st.session_state["file_bytes"] is None:
    st.info("👈 Lütfen analizlere başlamak için sol taraftaki menüden Excel ciro raporu dosyanızı yükleyin.")
    st.stop()
else:
    st.success(f"📌 **Aktif Kullanılan Dosya:** `{st.session_state['file_name']}`")

# Dosyayı diske geçici dosya olarak kaydet ve oradan oku (en güvenli yöntem)
try:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
        tmp_file.write(st.session_state["file_bytes"])
        tmp_file_path = tmp_file.name

    xls = pd.ExcelFile(tmp_file_path, engine="openpyxl")
    sheet_names = xls.sheet_names
except Exception as e:
    st.error(f"Excel dosyası okunamadı. Lütfen dosyanın bozuk olmadığını ve geçerli bir .xlsx dosyası olduğunu kontrol edin.\nHata: {e}")
    st.stop()

# -------------------------------------------------------------------
# MOD 1: 2026 Ciro Analizi (Aylık Karşılaştırma)
# -------------------------------------------------------------------
if mode == "📈 2026 Ciro Analizi (Aylık Karşılaştırma)":
    st.header("📈 2026 Ciro & Palet Analizi (Aylar Arası Karşılaştırma)")
    
    ciro_sheet = st.selectbox("Ciro Analizi Sayfasını Seçin:", sheet_names, index=0)
    df_ciro = pd.read_excel(tmp_file_path, sheet_name=ciro_sheet, engine="openpyxl")
    
    aylar = ["OCAK", "ŞUBAT", "MART", "NİSAN", "MAYIS", "HAZİRAN", "TEMMUZ", "AĞUSTOS", "EYLÜL", "EKİM", "KASIM", "ARALIK"]
    
    col1, col2 = st.columns(2)
    with col1:
        ay_1 = st.selectbox("1. Ayı Seçin (Önceki Ay):", aylar, index=2) # Mart
    with col2:
        ay_2 = st.selectbox("2. Ayı Seçin (Sonraki Ay):", aylar, index=3) # Nisan
        
    st.subheader(f"🔄 {ay_1} vs {ay_2} Dönem Performansı")
    
    palet_col_1 = [c for c in df_ciro.columns if ay_1 in str(c).upper() and "PALET" in str(c).upper()]
    tutar_col_1 = [c for c in df_ciro.columns if ay_1 in str(c).upper() and ("TUTAR" in str(c).upper() or "CİRO" in str(c).upper())]
    
    palet_col_2 = [c for c in df_ciro.columns if ay_2 in str(c).upper() and "PALET" in str(c).upper()]
    tutar_col_2 = [c for c in df_ciro.columns if ay_2 in str(c).upper() and ("TUTAR" in str(c).upper() or "CİRO" in str(c).upper())]

    if palet_col_1 and tutar_col_1 and palet_col_2 and tutar_col_2:
        p1, t1 = palet_col_1[0], tutar_col_1[0]
        p2, t2 = palet_col_2[0], tutar_col_2[0]
        
        musteri_col = df_ciro.columns[0]
        df_sub = df_ciro[[musteri_col, p1, t1, p2, t2]].dropna(subset=[musteri_col]).copy()
        
        for col in [p1, t1, p2, t2]:
            df_sub[col] = pd.to_numeric(df_sub[col].astype(str).str.replace('₺','').str.replace('.','').str.replace(',','.'), errors='coerce').fillna(0)
        
        df_sub['Ciro_Farki'] = df_sub[t2] - df_sub[t1]
        df_sub['Palet_Farki'] = df_sub[p2] - df_sub[p1]
        
        g_col1, g_col2 = st.columns(2)
        with g_col1:
            top_10_t2 = df_sub.nlargest(8, t2)
            fig_pie_ciro = px.pie(top_10_t2, values=t2, names=musteri_col, title=f"{ay_2} Ayı En Yüksek Ciro Eden Müşteriler")
            st.plotly_chart(fig_pie_ciro, use_container_width=True)
            
        with g_col2:
            top_10_p2 = df_sub.nlargest(8, p2)
            fig_pie_palet = px.pie(top_10_p2, values=p2, names=musteri_col, title=f"{ay_2} Ayı En Yüksek Palet Sayısına Ulaşan Müşteriler")
            st.plotly_chart(fig_pie_palet, use_container_width=True)
            
        st.markdown("---")
        st.subheader(f"📊 {ay_1} - {ay_2} Değişim Analizi")
        
        c_inc, c_dec = st.columns(2)
        with c_inc:
            st.write("🟢 **Cirosu En Çok Artan İlk 5 Müşteri**")
            st.dataframe(df_sub.nlargest(5, 'Ciro_Farki')[[musteri_col, t1, t2, 'Ciro_Farki']].style.format({t1: "₺{:,.2f}", t2: "₺{:,.2f}", 'Ciro_Farki': "₺{:,.2f}"}))
        with c_dec:
            st.write("🔴 **Cirosu En Çok Düşen İlk 5 Müşteri**")
            st.dataframe(df_sub.nsmallest(5, 'Ciro_Farki')[[musteri_col, t1, t2, 'Ciro_Farki']].style.format({t1: "₺{:,.2f}", t2: "₺{:,.2f}", 'Ciro_Farki': "₺{:,.2f}"}))

    else:
        st.warning("Seçilen aylara ait sütunlar sayfada bulunamadı.")

# -------------------------------------------------------------------
# MOD 2: 2025 - 2026 Hedef Karşılaştırması
# -------------------------------------------------------------------
elif mode == "🎯 2025 - 2026 Hedef Karşılaştırması":
    st.header("🎯 2025 Gerçekleşen vs 2026 Hedef & Gerçekleşen")
    
    hedef_sheet = st.selectbox("Hedef Sayfasını Seçin:", sheet_names, index=min(1, len(sheet_names)-1))
    df_hedef = pd.read_excel(tmp_file_path, sheet_name=hedef_sheet, engine="openpyxl")
    
    st.write("📋 **Hedef Tablosu Önizleme:**")
    st.dataframe(df_hedef.head(10))

# -------------------------------------------------------------------
# MOD 3: 2027 Hedef Oluşturucu (Simülasyon)
# -------------------------------------------------------------------
elif mode == "🚀 2027 Hedef Oluşturucu (Simülasyon)":
    st.header("🚀 2026 Verilerinden Hareketle 2027 Hedef Simülasyonu")
    
    artisa_orani = st.slider("2027 Yılı Tahmini Büyüme Oranı (%)", min_value=0, max_value=100, value=25, step=5)
    palet_artis_orani = st.slider("2027 Yılı Palet Kapasite Artış Oranı (%)", min_value=0, max_value=50, value=15, step=5)
    
    col_h1, col_h2, col_h3 = st.columns(3)
    col_h1.metric(label="2026 Toplam Hedef Ciro", value="₺710,000,000")
    col_h2.metric(label="2027 Yeni Hedef Ciro", value=f"₺{710000000 * (1 + artisa_orani/100):,.2f}", delta=f"%{artisa_orani} Büyüme")
    col_h3.metric(label="2027 Hedef Palet Sayısı", value=f"{int(187203 * (1 + palet_artis_orani/100)):,} Palet", delta=f"%{palet_artis_orani} Kapasite Artışı")

# -------------------------------------------------------------------
# MOD 4: Potansiyel Müşteriler & Kategori Hedefleri
# -------------------------------------------------------------------
elif mode == "🔍 Potansiyel Müşteriler & Kategori Hedefleri":
    st.header("🔍 Kategorize Potansiyel Müşteriler & 2027 Bütçeleme")
    
    categories = ["Soğuk Zincir", "Donmuş Gıda", "Et & Et Ürünleri", "Süt & Süt Ürünleri", "Kuru Gıda & Lojistik"]
    cat_select = st.selectbox("Sektör / Kategori Seçin:", categories)
    
    sample_data = pd.DataFrame({
        "Firma Adı": ["Firma A", "Firma B", "Firma C"],
        "Mevcut Durum": ["Görüşülüyor", "Teklif Verildi", "Soğuk Temas"],
        "2027 Tahmini Palet": [500, 1200, 800],
        "2027 Tahmini Ciro (₺)": [1500000, 3600000, 2400000]
    })
    
    edited_df = st.data_editor(sample_data, num_rows="dynamic")
    toplam_potansiyel = edited_df["2027 Tahmini Ciro (₺)"].sum()
    st.info(f"💰 **{cat_select}** Kategorisinden Beklenen Toplam 2027 Ek Ciro Katkısı: **₺{toplam_potansiyel:,.2f}**")

# İşlem bitince geçici dosyayı sil (temizlik)
if os.path.exists(tmp_file_path):
    os.remove(tmp_file_path)
