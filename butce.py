import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# ==========================================
# STREAMLIT SAYFA AYARLARI
# ==========================================
st.set_page_config(page_title="Bütçe Raporu & Sunum", layout="wide")


# ==========================================
# 1. TÜRKÇE SAYI VE PARA FORMATLAMA
# ==========================================
def format_tl(val):
    if pd.isna(val) or val is None:
        return "0,00 TL"
    return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") + " TL"

def format_palet(val):
    if pd.isna(val) or val is None:
        return "0 Palet"
    return f"{int(round(val)):,}".replace(",", ".") + " Palet"


# ==========================================
# 2. VERİ TEMİZLEME VE HAZIRLAMA (GENEL TOPLAM)
# ==========================================
def hazirla_genel_veri(df_2025, df_2026, col_musteri, col_ciro, col_palet):
    # 2025 Gruplama
    g25 = df_2025.groupby(col_musteri, as_index=False).agg({
        col_ciro: "sum",
        col_palet: "sum"
    }).rename(columns={col_musteri: "MUSTERI", col_ciro: "CIRO_2025", col_palet: "PALET_2025"})

    # 2026 Gruplama
    g26 = df_2026.groupby(col_musteri, as_index=False).agg({
        col_ciro: "sum",
        col_palet: "sum"
    }).rename(columns={col_musteri: "MUSTERI", col_ciro: "CIRO_2026", col_palet: "PALET_2026"})

    # Birleştirme (Genel Toplamlar Üzerinden)
    df_merged = pd.merge(g25, g26, on="MUSTERI", how="outer").fillna(0)

    # Değişim Miktarları
    df_merged["CIRO_DEGISIM"] = df_merged["CIRO_2026"] - df_merged["CIRO_2025"]
    df_merged["PALET_DEGISIM"] = df_merged["PALET_2026"] - df_merged["PALET_2025"]

    return df_merged


# ==========================================
# 3. ANA UYGULAMA VE DOSYA YÜKLEME EKRANI
# ==========================================
st.title("📊 Bütçe & Müşteri Analiz Raporu (2025 - 2026)")
st.markdown("---")

# Sol Panel - Dosya Yükleme
st.sidebar.header("📁 Veri Yükleme")
file_2025 = st.sidebar.file_uploader("2025 Veri Dosyasını Yükleyin (Excel/CSV)", type=["xlsx", "csv"], key="2025")
file_2026 = st.sidebar.file_uploader("2026 Veri Dosyasını Yükleyin (Excel/CSV)", type=["xlsx", "csv"], key="2026")

if file_2025 and file_2026:
    # Verileri Oku
    try:
        df_2025 = pd.read_excel(file_2025) if file_2025.name.endswith(".xlsx") else pd.read_csv(file_2025)
        df_2026 = pd.read_excel(file_2026) if file_2026.name.endswith(".xlsx") else pd.read_csv(file_2026)
    except Exception as e:
        st.error(f"Dosya okunurken bir hata oluştu: {e}")
        st.stop()

    # Sütun Eşleştirme (Yanlış sütun adı kaynaklı kilitlenmeyi önler)
    st.sidebar.subheader("⚙️ Sütun Eşleştirme")
    cols = list(df_2026.columns)
    
    col_musteri = st.sidebar.selectbox("Müşteri Adı Sütunu", cols, index=0)
    col_ciro = st.sidebar.selectbox("Ciro Sütunu", cols, index=1 if len(cols)>1 else 0)
    col_palet = st.sidebar.selectbox("Palet Sütunu", cols, index=2 if len(cols)>2 else 0)

    # Veriyi Genel Bazda İşle
    df_genel = hazirla_genel_veri(df_2025, df_2026, col_musteri, col_ciro, col_palet)

    # Toplamlar
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
    c1, c2, c3 = st.columns(3)

    with c1:
        st.metric("2025 Toplam Ciro", format_tl(tot_ciro_25))
        st.metric("2026 Toplam Ciro", format_tl(tot_ciro_26), delta=format_tl(tot_ciro_26 - tot_ciro_25))

    with c2:
        st.metric("2025 Toplam Palet", format_palet(tot_palet_25))
        st.metric("2026 Toplam Palet", format_palet(tot_palet_26), delta=format_palet(tot_palet_26 - tot_palet_25))

    with c3:
        st.metric("2025 Ort. Palet Başı Ciro", format_tl(palet_basi_25))
        st.metric("2026 Ort. Palet Başı Ciro", format_tl(palet_basi_26), delta=format_tl(palet_basi_26 - palet_basi_25))

    st.markdown("---")

    # ------------------------------------------
    # BÖLÜM 2: TOP 20 MÜŞTERİ & PASTA GRAFİKLER
    # ------------------------------------------
    st.subheader("🎯 En Yüksek Ciro Getiren 20 Müşteri Analizi")
    
    df_genel["TOPLAM_CIRO"] = df_genel["CIRO_2025"] + df_genel["CIRO_2026"]
    top_20 = df_genel.sort_values(by="TOPLAM_CIRO", ascending=False).head(20).copy()

    top_20["TL_PALET_2026"] = np.where(top_20["PALET_2026"] > 0, top_20["CIRO_2026"] / top_20["PALET_2026"], 0)

    # Hedef Slider
    hedef_orani = st.slider("2027 Bütçe Hedef Büyüme Oranı (%)", 0, 100, 15)
    top_20["2027_HEDEF_CIRO"] = top_20["CIRO_2026"] * (1 + hedef_orani / 100)

    # Pasta Grafikler
    col_g1, col_g2 = st.columns(2)
    with col_g1:
        fig_ciro = px.pie(top_20, values="CIRO_2026", names="MUSTERI", title="2026 Top 20 Müşteri Ciro Payı", hole=0.3)
        st.plotly_chart(fig_ciro, use_container_width=True)

    with col_g2:
        fig_palet = px.pie(top_20, values="PALET_2026", names="MUSTERI", title="2026 Top 20 Müşteri Palet Payı", hole=0.3)
        st.plotly_chart(fig_palet, use_container_width=True)

    # Tablo Gösterimi
    tablo = top_20[["MUSTERI", "CIRO_2025", "CIRO_2026", "PALET_2025", "PALET_2026", "TL_PALET_2026", "2027_HEDEF_CIRO"]].copy()
    tablo["CIRO_2025"] = tablo["CIRO_2025"].apply(format_tl)
    tablo["CIRO_2026"] = tablo["CIRO_2026"].apply(format_tl)
    tablo["PALET_2025"] = tablo["PALET_2025"].apply(format_palet)
    tablo["PALET_2026"] = tablo["PALET_2026"].apply(format_palet)
    tablo["TL_PALET_2026"] = tablo["TL_PALET_2026"].apply(format_tl)
    tablo["2027_HEDEF_CIRO"] = tablo["2027_HEDEF_CIRO"].apply(format_tl)

    st.dataframe(tablo, use_container_width=True)

    st.markdown("---")

    # ------------------------------------------
    # BÖLÜM 3: ARTAN VE DÜŞEN MÜŞTERİLER
    # ------------------------------------------
    st.subheader("🔥 Genel Ciro Değişim Değerlendirmesi")
    col_inc, col_dec = st.columns(2)

    artan_5 = df_genel.sort_values(by="CIRO_DEGISIM", ascending=False).head(5)
    dusan_5 = df_genel.sort_values(by="CIRO_DEGISIM", ascending=True).head(5)

    with col_inc:
        st.success("🟢 Cirosu En Çok Artan 5 Müşteri")
        for _, r in artan_5.iterrows():
            st.write(f"**{r['MUSTERI']}**: +{format_tl(r['CIRO_DEGISIM'])} *(2026 Toplam: {format_tl(r['CIRO_2026'])})*")

    with col_dec:
        st.error("🔴 Cirosu En Çok Düşen 5 Müşteri")
        for _, r in dusan_5.iterrows():
            st.write(f"**{r['MUSTERI']}**: {format_tl(r['CIRO_DEGISIM'])} *(2026 Toplam: {format_tl(r['CIRO_2026'])})*")

    st.markdown("---")

    # ------------------------------------------
    # BÖLÜM 4: SUNUM İÇİN YÖNETİCİ NOTLARI
    # ------------------------------------------
    st.subheader("📝 Sunum İçin Otomatik Analiz Notları")
    top20_pay = (top_20["CIRO_2026"].sum() / tot_ciro_26 * 100) if tot_ciro_26 > 0 else 0
    palet_degisim = ((palet_basi_26 - palet_basi_25) / palet_basi_25 * 100) if palet_basi_25 > 0 else 0

    st.info(f"""
    * **Müşteri Yoğunlaşması:** İlk 20 müşteri, 2026 yılı toplam cironun **%{top20_pay:.1f}** kısmını oluşturmaktadır.
    * **Palet Başı Verimlilik:** Palet başına ortalama ciro 2025'te **{format_tl(palet_basi_25)}** iken 2026'da **{format_tl(palet_basi_26)}** olmuştur (Değişim: **%{palet_degisim:.1f}**).
    * **Aksiyon Notu:** En fazla artış sağlanan **{artan_5.iloc[0]['MUSTERI']}** ivmesi sürdürülmeli, en fazla düşüş yaşayan **{dusan_5.iloc[0]['MUSTERI']}** müşterisi ile tekrar görüşülmelidir.
    """)

else:
    st.warning("👈 Lütfen analizin başlayabilmesi için sol taraftaki menüden 2025 ve 2026 Excel/CSV dosyalarınızı yükleyin.")
