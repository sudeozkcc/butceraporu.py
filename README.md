# Bütçeleme, Ciro Analizi ve Hedefleme Portalı

Streamlit uygulaması. Excel ciro raporunu yükleyin; 4 modda analiz edin:

1. **Aylık Ciro Analizi** – Ocak→Şubat, Şubat→Mart … Kasım→Aralık; pasta grafikleri, müşteri bazlı ay karşılaştırması, artan/düşen müşteriler.
2. **Yıllar Arası Hedef Karşılaştırması** – 2025 / 2026 gerçekleşen ve hedef (palet & ciro).
3. **2027 Hedef Oluşturucu** – Büyüme oranları ve müşteri bazlı düzenlenebilir hedefler, Excel çıktısı.
4. **Potansiyel Müşteriler** – Kategori bazlı (soğuk, donuk, et, süt …) firma listesi, olasılık ağırlıklı 2027 tahmini.

## Çalıştırma
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Excel formatı
Başlıklarda ay adı (OCAK, Şubat …) ile PALET / CİRO (veya TUTAR) kelimeleri geçmelidir.
Hedef sayfasında yıl (2025, 2026) ve HEDEF kelimesi ayırıcı olarak kullanılır.
Başlıklar birleştirilmiş hücre, üst başlık satırı veya "TOPLAM" satırı içerebilir.

## Yayınlama (Streamlit Community Cloud)
GitHub'a yükleyin → share.streamlit.io → *New app* → repo seçin → Main file: `app.py`.
