import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ==========================================
# 1. TÜRKÇE SAYI VE PARA FORMATLAMA
# ==========================================
def format_tl(val):
    """Örn: 23343457.78 -> 23.343.457,78 TL"""
    if pd.isna(val) or val is None:
        return "0,00 TL"
    return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") + " TL"

def format_palet(val):
    """Örn: 1568.0 -> 1.568 Palet"""
    if pd.isna(val) or val is None:
        return "0 Palet"
    return f"{int(round(val)):,}".replace(",", ".") + " Palet"

def format_oran(val):
    """Örn: 12.5 -> %12,50"""
    if pd.isna(val) or val is None:
        return "%0,00"
    return f"%{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# ==========================================
# 2. VERİ TEMİZLEME VE HAZIRLAMA (GENEL TOPLAM)
# ==========================================
def hazirla_genel_veri(df_2025, df_2026):
    """
    Ay detaylarını kaldırır, doğrudan genel toplam üzerinden 2025 ve 2026 verilerini birleştirir.
    Sütun isimlerinin 'MUSTERI', 'CIRO', 'PALET' olduğunu varsayar.
    """
    # 2025 Gruplama
    g25 = df_2025.groupby("MUSTERI", as_index=False).agg({
        "CIRO": "sum",
        "PALET": "sum"
    }).rename(columns={"CIRO": "CIRO_2025", "PALET": "PALET_2025"})

    # 2026 Gruplama
    g26 = df_2026.groupby("MUSTERI", as_index=False).agg({
        "CIRO": "sum",
        "PALET": "sum"
    }).rename(columns={"CIRO": "CIRO_2026", "PALET": "PALET_2026"})

    # Birleştirme (Outer Join)
    df_merged = pd.merge(g25, g26, on="MUSTERI", how="outer").fillna(0)

    # Değişim Miktarları
    df_merged["CIRO_DEGISIM"] = df_merged["CIRO_2026"] - df_merged["CIRO_2025"]
    df_merged["PALET_DEGISIM"] = df_merged["PALET_2026"] - df_merged["PALET_2025"]

    return df_merged


# ==========================================
# 3. STREAMLIT BÜTÇE VE SUNUM RAPORU PANELİ
# ==========================================
def render_butce_raporu(df_2025, df_2026):
    st.title("📊 Bütçe & Müşteri Analiz Raporu (2025 - 2026)")
    st.markdown("---")

    # Veriyi İşle
    df_genel = hazirla_genel_veri(df_2025, df_2026)

    # Toplam Metrikler
    tot_ciro_25 = df_genel["CIRO_2025"].sum()
    tot_ciro_26 = df_genel["CIRO_2026"].sum()
    tot_palet_25 = df_genel["PALET_2025"].sum()
    tot_palet_26 = df_genel["PALET_2026"].sum()

    palet_basi_25 = tot_ciro_25 / tot_palet_25 if tot_palet_25 > 0 else 0
    palet_basi_26 = tot_ciro_26 / tot_palet_26 if tot_palet_26 > 0 else 0

    # ------------------------------------------
    # BÖLÜM 1: GENEL PERFORMANS KARTLARI
    # ------------------------------------------
    st.subheader("📈 2025 - 2026 Genel Toplam Özeti")
    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric("2025 Toplam Ciro", format_tl(tot_ciro_25))
        st.metric("2026 Toplam Ciro", format_tl(tot_ciro_26), 
                  delta=format_tl(tot_ciro_26 - tot_ciro_25))

    with col2:
        st.metric("2025 Toplam Palet", format_palet(tot_palet_25))
        st.metric("2026 Toplam Palet", format_palet(tot_palet_26), 
                  delta=format_palet(tot_palet_26 - tot_palet_25))

    with col3:
        st.metric("2025 Ort. Palet Başı Ciro", format_tl(palet_basi_25))
        st.metric("2026 Ort. Palet Başı Ciro", format_tl(palet_basi_26), 
                  delta=format_tl(palet_basi_26 - palet_basi_25))

    st.markdown("---")

    # ------------------------------------------
    # BÖLÜM 2: EN YÜKSEK CİROLU İLK 20 MÜŞTERİ & HEDEF
    # ------------------------------------------
    st.subheader("🎯 Top 20 Müşteri Analizi ve Bütçe Hedefleri")
    
    # En yüksek ciroyu getiren ilk 20 müşteri (2026 veya Toplam üzerinden)
    df_genel["TOPLAM_CIRO"] = df_genel["CIRO_2025"] + df_genel["CIRO_2026"]
    top_20_df = df_genel.sort_values(by="TOPLAM_CIRO", ascending=False).head(20).copy()

    # Müşteri bazında Palet Başına TL
    top_20_df["TL_PALET_2025"] = np.where(top_20_df["PALET_2025"] > 0, top_20_df["CIRO_2025"] / top_20_df["PALET_2025"], 0)
    top_20_df["TL_PALET_2026"] = np.where(top_20_df["PALET_2026"] > 0, top_20_df["CIRO_2026"] / top_20_df["PALET_2026"], 0)

    # Hedef Bütçe Oranı (Kullanıcı dinamik hedef belirleyebilir)
    hedef_artis_orani = st.slider("2027 Bütçe Hedef Büyüme Oranı (%)", min_value=0, max_value=100, value=15)
    top_20_df["2027_HEDEF_CIRO"] = top_20_df["CIRO_2026"] * (1 + hedef_artis_orani / 100)

    # ------------------------------------------
    # BÖLÜM 3: PASTA GRAFİKLER (CİRO VE PALET DAĞILIMI)
    # ------------------------------------------
    col_chart1, col_chart2 = st.columns(2)

    with col_chart1:
        fig_ciro = px.pie(
            top_20_df, 
            values="CIRO_2026", 
            names="MUSTERI", 
            title="2026 Top 20 Müşteri Ciro Payı",
            hole=0.3
        )
        st.plotly_chart(fig_ciro, use_container_width=True)

    with col_chart2:
        fig_palet = px.pie(
            top_20_df, 
            values="PALET_2026", 
            names="MUSTERI", 
            title="2026 Top 20 Müşteri Palet Hacmi Payı",
            hole=0.3
        )
        st.plotly_chart(fig_palet, use_container_width=True)

    # Top 20 Tablosu
    tablo_display = top_20_df[["MUSTERI", "CIRO_2025", "CIRO_2026", "PALET_2025", "PALET_2026", "TL_PALET_2026", "2027_HEDEF_CIRO"]].copy()
    tablo_display["CIRO_2025"] = tablo_display["CIRO_2025"].apply(format_tl)
    tablo_display["CIRO_2026"] = tablo_display["CIRO_2026"].apply(format_tl)
    tablo_display["PALET_2025"] = tablo_display["PALET_2025"].apply(format_palet)
    tablo_display["PALET_2026"] = tablo_display["PALET_2026"].apply(format_palet)
    tablo_display["TL_PALET_2026"] = tablo_display["TL_PALET_2026"].apply(format_tl)
    tablo_display["2027_HEDEF_CIRO"] = tablo_display["2027_HEDEF_CIRO"].apply(format_tl)

    st.dataframe(tablo_display, use_container_width=True)

    st.markdown("---")

    # ------------------------------------------
    # BÖLÜM 4: CİROSU EN ÇOK ARTAN & DÜŞEN TOP 5 MÜŞTERİ
    # ------------------------------------------
    st.subheader("🔥 Genel Ciro Değişim Değerlendirmesi")
    
    col_inc, col_dec = st.columns(2)

    artan_5 = df_genel.sort_values(by="CIRO_DEGISIM", ascending=False).head(5)
    dusan_5 = df_genel.sort_values(by="CIRO_DEGISIM", ascending=True).head(5)

    with col_inc:
        st.success("🟢 Cirosu En Çok Artan 5 Müşteri")
        for _, row in artan_5.iterrows():
            st.write(f"**{row['MUSTERI']}**: +{format_tl(row['CIRO_DEGISIM'])} (2026: {format_tl(row['CIRO_2026'])})")

    with col_dec:
        st.error("🔴 Cirosu En Çok Düşen 5 Müşteri")
        for _, row in dusan_5.iterrows():
            st.write(f"**{row['MUSTERI']}**: {format_tl(row['CIRO_DEGISIM'])} (2026: {format_tl(row['CIRO_2026'])})")

    st.markdown("---")

    # ------------------------------------------
    # BÖLÜM 5: OTOMATİK SUNUM YORUMU (BÜTÇE NOTLARI)
    # ------------------------------------------
    st.subheader("📝 Sunum İçin Otomatik Analiz ve Yönetici Yorumu")
    
    top20_ciro_payi = (top_20_df["CIRO_2026"].sum() / tot_ciro_26 * 100) if tot_ciro_26 > 0 else 0
    palet_verimlilik_degisim = ((palet_basi_26 - palet_basi_25) / palet_basi_25 * 100) if palet_basi_25 > 0 else 0

    st.info(f"""
    **📌 Sunum Slayt Notları:**
    1. **Müşteri Konsantrasyonu:** Odaklandığımız ilk 20 müşteri, 2026 yılı toplam cironun **%{top20_ciro_payi:.1f}** kadarlık kısmını oluşturmaktadır.
    2. **Verimlilik Analizi (Palet Başı TL):** 2025 yılında palet başına ortalama ciro **{format_tl(palet_basi_25)}** iken, 2026 yılında bu değer **{format_tl(palet_basi_26)}** seviyesine ulaşmıştır (Değişim: **%{palet_verimlilik_degisim:.1f}**).
    3. **Bütçe Stratejisi:** En çok ciro artışı sağlanan **{artan_5.iloc[0]['MUSTERI']}** tarafındaki büyüme ivmesi korunmalı; cirosu en fazla düşen **{dusan_5.iloc[0]['MUSTERI']}** için özel aksiyon planı oluşturulmalıdır.
    """)
