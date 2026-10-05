import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np

# Sayfa Yapılandırması
st.set_page_config(
    page_title="Ciro & Bütçeleme Analiz Portalı",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Bütçeleme, Ciro Analizi ve Hedefleme Portalı")

# Yan Menü: Dosya Yükleme ve Mod Seçimi
st.sidebar.header("📁 Veri Yükleme & Modlar")
uploaded_file = st.sidebar.file_content_uploader if hasattr(st.sidebar, 'file_content_uploader') else st.sidebar.file_uploader("Excel Raporunu Yükleyin (.xlsx)", type=["xlsx"])

mode = st.sidebar.radio(
    "Çalışma Modunu Seçin:",
    [
        "📈 2026 Ciro Analizi (Aylık Karşılaştırma)",
        "🎯 2025 - 2026 Hedef Karşılaştırması",
        "🚀 2027 Hedef Oluşturucu (Simülasyon)",
        "🔍 Potansiyel Müşteriler & Kategori Hedefleri"
    ]
)

if uploaded_file is not None:
    # Sayfa isimlerini al
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names
    
    # -------------------------------------------------------------------
    # MOD 1: 2026 Ciro Analizi (Aylık Karşılaştırma)
    # -------------------------------------------------------------------
    if mode == "📈 2026 Ciro Analizi (Aylık Karşılaştırma)":
        st.header("📈 2026 Ciro & Palet Analizi (Aylar Arası Karşılaştırma)")
        
        # Sayfa Seçimi
        ciro_sheet = st.selectbox("Ciro Analizi Sayfasını Seçin:", sheet_names, index=0 if "Panthera" in sheet_names[0] else 0)
        df_ciro = pd.read_excel(uploaded_file, sheet_name=ciro_sheet)
        
        # Aylar Listesi
        aylar = ["OCAK", "ŞUBAT", "MART", "NİSAN", "MAYIS", "HAZİRAN", "TEMMUZ", "AĞUSTOS", "EYLÜL", "EKİM", "KASIM", "ARALIK"]
        
        col1, col2 = st.columns(2)
        with col1:
            ay_1 = st.selectbox("1. Ayı Seçin (Önceki Ay):", aylar, index=2) # Mart
        with col2:
            ay_2 = st.selectbox("2. Ayı Seçin (Sonraki Ay):", aylar, index=3) # Nisan
            
        st.subheader(f"🔄 {ay_1} vs {ay_2} Dönem Performansı")
        
        # Sütun Adlarını Dinamik Yakalama
        palet_col_1 = [c for c in df_ciro.columns if ay_1 in str(c).upper() and "PALET" in str(c).upper()]
        tutar_col_1 = [c for c in df_ciro.columns if ay_1 in str(c).upper() and ("TUTAR" in str(c).upper() or "CİRO" in str(c).upper())]
        
        palet_col_2 = [c for c in df_ciro.columns if ay_2 in str(c).upper() and "PALET" in str(c).upper()]
        tutar_col_2 = [c for c in df_ciro.columns if ay_2 in str(c).upper() and ("TUTAR" in str(c).upper() or "CİRO" in str(c).upper())]

        if palet_col_1 and tutar_col_1 and palet_col_2 and tutar_col_2:
            p1, t1 = palet_col_1[0], tutar_col_1[0]
            p2, t2 = palet_col_2[0], tutar_col_2[0]
            
            # Temizlik & Sayısallaştırma
            musteri_col = df_ciro.columns[0]
            df_sub = df_ciro[[musteri_col, p1, t1, p2, t2]].dropna(subset=[musteri_col]).copy()
            for col in [p1, t1, p2, t2]:
                df_sub[col] = pd.to_numeric(df_sub[col].astype(str).str.replace('₺','').str.replace('.','').str.replace(',','.'), errors='coerce').fillna(0)
            
            # Farklar
            df_sub['Ciro_Farki'] = df_sub[t2] - df_sub[t1]
            df_sub['Palet_Farki'] = df_sub[p2] - df_sub[p1]
            
            # Pasta Grafikler
            g_col1, g_col2 = st.columns(2)
            with g_col1:
                top_10_t2 = df_sub.nlargest(8, t2)
                fig_pie_ciro = px.pie(top_10_t2, values=t2, names=musteri_col, title=f"{ay_2} Ayı En Yüksek Ciro Eden Müşteriler (Pasta Grafik)")
                st.plotly_chart(fig_pie_ciro, use_container_width=True)
                
            with g_col2:
                top_10_p2 = df_sub.nlargest(8, p2)
                fig_pie_palet = px.pie(top_10_p2, values=p2, names=musteri_col, title=f"{ay_2} Ayı En Yüksek Palet Sayısına Ulaşan Müşteriler")
                st.plotly_chart(fig_pie_palet, use_container_width=True)
                
            # Ciro ve Paleti Artan / Azalanlar
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
            st.warning("Seçilen aylara ait sütunlar sayfada tam olarak tespit edilemedi. Lütfen sütun başlıklarını kontrol edin.")

    # -------------------------------------------------------------------
    # MOD 2: 2025 - 2026 Hedef Karşılaştırması
    # -------------------------------------------------------------------
    elif mode == "🎯 2025 - 2026 Hedef Karşılaştırması":
        st.header("🎯 2025 Gerçekleşen vs 2026 Hedef & Gerçekleşen")
        
        hedef_sheet = st.selectbox("Hedef Sayfasını Seçin:", sheet_names, index=1 if len(sheet_names)>1 else 0)
        df_hedef = pd.read_excel(uploaded_file, sheet_name=hedef_sheet)
        
        st.write("📋 **Hedef Tablosu Önizleme:**")
        st.dataframe(df_hedef.head(10))
        
        # Örnek Aylar ve Metrik Karşılaştırma Grafiği
        aylar = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]
        
        # Örnek görselleştirme (Gelişmiş Çizgi/Sütun Grafiği)
        st.subheader("📊 Aylık Ciro Gelişimi ve Hedef Sapmaları")
        
        # Tablodan Satır/Sütun eşleştirme örneği
        fig = go.Figure()
        # İllüstratif Veri Bağlantısı
        fig.add_trace(go.Bar(x=aylar, y=[34.9, 32.5, 29.1, 32.0, 38.1, 34.1, 45.8, 40.9, 41.4, 42.4, 32.1, 35.0], name="2025 Gerçekleşen (Milyon ₺)"))
        fig.add_trace(go.Scatter(x=aylar, y=[22.7, 21.0, 23.6, 47.1, 60.6, 62.6, 61.6, 58.6, 56.6, 64.6, 62.6, 67.6], name="2026 Hedef (Milyon ₺)", line=dict(color='orange', width=3)))
        fig.add_trace(go.Scatter(x=aylar, y=[22.7, 21.0, 23.6, 47.1, 30.1, 42.1, 48.1, 41.8, None, None, None, None], name="2026 Gerçekleşen (Milyon ₺)", line=dict(color='green', width=3, dash='dot')))
        
        fig.update_layout(title="2025 vs 2026 Dönemsel Performans", xaxis_title="Aylar", yaxis_title="Tutar (Milyon ₺)", barmode='group')
        st.plotly_chart(fig, use_container_width=True)

    # -------------------------------------------------------------------
    # MOD 3: 2027 Hedef Oluşturucu (Simülasyon)
    # -------------------------------------------------------------------
    elif mode == "🚀 2027 Hedef Oluşturucu (Simülasyon)":
        st.header("🚀 2026 Verilerinden Hareketle 2027 Hedef Simülasyonu")
        
        st.markdown("2026 yılı gerçekleşen verilerinin üzerine **yüzdesel büyüme** veya **enflasyon/hedef katsayısı** ekleyerek 2027 hedeflerini otomatik hesaplayın.")
        
        artisa_orani = st.slider("2027 Yılı Tahmini Büyüme Oranı (%)", min_value=0, max_value=100, value=25, step=5)
        palet_artis_orani = st.slider("2027 Yılı Palet Kapasite Artış Oranı (%)", min_value=0, max_value=50, value=15, step=5)
        
        # Hesaplama alanı
        st.subheader("💡 2027 Yılı Projeksiyonu")
        
        col_h1, col_h2, col_h3 = st.columns(3)
        col_h1.metric(label="2026 Toplam Hedef Ciro", value="₺710,000,000")
        col_h2.metric(label="2027 Yeni Hedef Ciro", value=f"₺{710000000 * (1 + artisa_orani/100):,.2f}", delta=f"%{artisa_orani} Büyüme")
        col_h3.metric(label="2027 Hedef Palet Sayısı", value=f"{int(187203 * (1 + palet_artis_orani/100)):,} Palet", delta=f"%{palet_artis_orani} Kapasite Artışı")

    # -------------------------------------------------------------------
    # MOD 4: Potansiyel Müşteriler & Kategori Hedefleri
    # -------------------------------------------------------------------
    elif mode == "🔍 Potansiyel Müşteriler & Kategori Hedefleri":
        st.header("🔍 Kategorize Potansiyel Müşteriler & 2027 Bütçeleme")
        
        st.markdown("Potansiyel müşterilerinizi sektör bazlı kategorize edin ve 2027 hedeflerine tahmini rakamlar ekleyin.")
        
        # Örnek Kategori Tablosu Yapısı
        categories = ["Soğuk Zincir", "Donmuş Gıda", "Et & Et Ürünleri", "Süt & Süt Ürünleri", "Kuru Gıda & Lojistik"]
        
        cat_select = st.selectbox("Sektör / Kategori Seçin:", categories)
        
        # Dinamik Tablo Ekranı / Düzenleyici
        st.subheader(f"📋 {cat_select} Kategorisi Potansiyel Müşteri Listesi")
        
        sample_data = pd.DataFrame({
            "Firma Adı": ["Firma A", "Firma B", "Firma C"],
            "Mevcut Durum": ["Görüşülüyor", "Teklif Verildi", "Soğuk Temas"],
            "2027 Tahmini Palet": [500, 1200, 800],
            "2027 Tahmini Ciro (₺)": [1500000, 3600000, 2400000]
        })
        
        edited_df = st.data_editor(sample_data, num_rows="dynamic")
        
        toplam_potansiyel = edited_df["2027 Tahmini Ciro (₺)"].sum()
        st.info(f"💰 **{cat_select}** Kategorisinden Beklenen Toplam 2027 Ek Ciro Katkısı: **₺{toplam_potansiyel:,.2f}**")

else:
    st.info("👈 Lütfen sol taraftaki menüden Excel ciro raporu dosyanızı yükleyin.")
