# Autism 2026 — AI Peer Review & Verification Package (`antigravity-review` branch)

Bu branch, otizm spektrum bozukluğu (OSB) olan erkek çocuk ve ergenlerde (`N = 34`, 5–15 yaş) rutin uyarlanmış hareket eğitiminin son 15 dakikasına eklenen 12 haftalık (24 seans) mikro-doz motor beceri antrenmanının etkilerini inceleyen çalışmanın **tam veri, kod, yöntem, bulgu, şekil ve makale taslağı** paketini içerir.

## 1. İnceleme İçin Ana Dosyalar (Öncelikli Okuma Sırası)

1. **Makale Taslakları (`04_manuscript/`):**
   - **Türkçe Ana Taslak (Birincil Atıflı):** `04_manuscript/Q1_Master_Makale_Taslagi_Birincil_Atifli_TR.md` (ve `.docx`)
   - **İngilizce Ana Taslak (Birincil Atıflı):** `04_manuscript/Q1_Master_Manuscript_Primary_Cited_EN.md` (ve `.docx`)
   - **Temiz Kılavuz Kopyaları (Giriş & Tartışma'da yeni kaynak eklemeye hazır):**
     - `04_manuscript/Q1_Master_Makale_Taslagi_Temiz_Kilavuz_TR.md` (ve `.docx`)
     - `04_manuscript/Q1_Master_Manuscript_Clean_Guide_EN.md` (ve `.docx`)
   - **Mevcut Durum Notu:** `ÖZ`, `ABSTRACT`, `2. YÖNTEM / METHODS` ve `3. BULGULAR / RESULTS` (4 APA tablo + 4 yüksek çözünürlüklü şekil) tamamlanmıştır. `1. GİRİŞ / INTRODUCTION`, `4. TARTIŞMA / DISCUSSION` ve `5. PRATİK UYGULAMALAR VE SONUÇ / CONCLUSIONS` bölümleri yeni kaynaklar eklendikten sonra yazılmak üzere **bilinçli olarak adım adım yazım kılavuzu (`[BU BÖLÜM YAZILACAKTIR]`)** formatında bırakılmıştır.

2. **Ham ve İşlenmiş Veriler (`02_data/`):**
   - `02_data/raw_data.csv` — Anonimleştirilmiş başlangıç kohort verisi (`N = 57` tarama kaydı).
   - `02_data/cleaned_data.csv` — Ön-test ve son-testi tamamlayan per-protokol veri seti (`N = 34`; Çocukluk 5–9 yaş `n = 17`, Ergenlik 10–15 yaş `n = 17`), türetilmiş Mirwald PHV ofseti, rölatif kavrama kuvveti, bilateral asimetri ve Sayers güç değişkenleri.

3. **İstatistiksel Analiz, Şekil ve Belge Üretim Kodları (`03_scripts_and_analysis/`):**
   - Veri temizleme & doğrulama: `data_cleaner_and_validator.py`, `generate_clean_dataset.py`, `audit_anomalies.py`, `attrition_bias_check.py`
   - İstatistiksel testler: `run_statistical_tests.py`, `reanalyze_relative_grip.py`, `responder_analysis.py`, `run_regressions.py`
   - Şekil üretimi (300 DPI): `generate_study_design_infographics.py`, `generate_results_figures.py`
   - Makale derleme (TR & EN `.md`/`.docx`): `build_primary_cited_manuscripts_stage5.py`

4. **Tasarım, Müfredat ve Atıf Denetimi:**
   - `01_plan_and_design/Egzersiz_Mufredati_Temel_ve_Genel_Beceriler.md` ve `Q1_Makale_Tasarim_ve_Analiz_Rehberi.md`
   - `TEMEL BECERİLER İÇİN ÇALIŞMALAR FİNAL.pdf` (T1–T38) ve `GENEL BECERİLER İÇİN ÇALIŞMALAR FİNAL.pdf` (G1–G41)
   - `05_citation_audit/Birincil_Kaynak_Atif_Denetim_Raporu_TR_matrix.csv` (ve `.xlsx`, `.docx`): 26 birincil kaynak doğrulama matrisi *(üçüncü taraf yayınevi PDF'leri telif ve depo boyutu nedeniyle hariç tutulmuştur; Mendeley `.ris`/`.bib` dosyaları `04_manuscript/` içindedir)*.

5. **Yayın Şekilleri (`04_manuscript/`):**
   - `Sekil_1_Arastirma_Tasarimi_TR.png` / `Figure_1_Study_Design_EN.png`
   - `Sekil_2_Performans_Degisimleri_TR.png` / `Figure_2_Performance_Adaptations_EN.png`
   - `Sekil_3_SWC_Yanitlayici_Analizi_TR.png` / `Figure_3_SWC_Responder_Analysis_EN.png`
   - `Sekil_4_Regresyon_Iliskileri_TR.png` / `Figure_4_Regression_Relationships_EN.png`

## 2. Review Kontrol Odakları

- `02_data/cleaned_data.csv` ile `04_manuscript/` içindeki tüm tablo ve metin sayılarının (`p`, `t`, `d`, `ηp²`, `%95 GA`, SWC yanıt oranları, OLS regresyon `R² = 0.280`) tutarlılığı
- Yöntem bölümündeki STROBE / CERT raporlama uyumu, 6 aylık ön eğitim dahil etme ölçütü, hareket eğitiminin son 15 dakikasında uygulanan mikro-doz protokol ve ölçüm protokollerinin (Jamar, Mirwald PHV, Sayers keşifsel kestirimi, 2 denemeli Pro-Agility uyarlaması) eleştirel değerlendirmesi
- İstatistiksel model seçimi (bağımlı t-testi, SWC eşiği, eşdoğrusallık altında OLS + iki değişkenli korelasyon yorumlaması) ve hakem gözüyle geliştirilebilecek noktalar
