# Q1 Düzeyinde Makale Tasarım ve Analiz Rehberi
**Proje:** Otizm Spektrum Bozukluğu (OSB) Tanılı Bireylerde Ek Beceri Eğitiminin Algısal ve Motorik Etkileri: Retrospektif Kohort Çalışması
**Örneklem:** N = 57 Erkek (En az 6 ay hareket eğitimi + 6 ay özel eğitim geçmişi)

---

## 1. Çalışmanın Metodolojik Dayanakları (Q1 Savunma Hattı)

Tek gruplu (pre-post) retrospektif çalışmaların Q1 düzeyindeki tıp/spor bilimleri dergilerinde (*Autism Research*, *Journal of Autism and Developmental Disorders*, *Research in Developmental Disabilities*, *Frontiers in Physiology*, *Perceptual and Motor Skills*) en sık karşılaştığı hakem eleştirilerini baştan bertaraf edecek stratejik dayanaklar:

### A. "Plato Kırma (Plateau-Breaking)" ve Doğal Kontrol Durumu
* **Sorun:** "Kontrol grubu olmadan gelişmelerin rutin antrenmandan kaynaklanmadığını nasıl bilebiliriz?"
* **Çözüm:** Dahil edilme kriteri olarak belirlenen **en az 6 aylık düzenli hareket eğitimi (haftada 2 gün 60 dk)**, erken dönem nöromusküler alışma ve ilk motor kazanım evresinin aşıldığını, bireylerin stabil bir bazal plato seviyesinde bulunduğunu belgeler (pandemi kısıtlamaları sonrası reel katılım süresini yansıtır).
* Rutin 60 dakikaya eklenen **15 dakikalık hedefli mikrodosaj beceri çalışması**, bu platoyu aşmak için uygulanan bağımsız değişkendir.

### B. Biyolojik Olgunlaşmanın Kontrolü (Mirwald Maturity Offset)
* **Sorun:** "Farklı doğum tarihlerine sahip bireylerde 10-12 haftadaki değişim doğal büyümeden mi kaynaklanıyor?"
* **Çözüm:** Takvim yaşı tek başına yetersizdir. Boy, Kilo, Oturma Boyu ve Bacak Boyu kullanılarak Mirwald vd. (2002) veya Moore vd. (2015) denklemleri ile **Zirve Boy Hızı (PHV - Maturity Offset)** hesaplanır. İstatistiksel modellerde (RM-ANCOVA / LMM) bu değer ortak değişken (kovaryat) olarak atanarak büyüme etkisi matematiksel olarak ayrıştırılır.

### C. Nöromotor Entegrasyon ve Bilateral Asimetri Analizi
* **Sorun:** "Sadece test skorları artmış, bu mekanistik olarak ne anlama geliyor?"
* **Çözüm:** OSB'de korpus kallozum anomalileri ve inter-hemisferik iletişim bozuklukları bilateral kuvvet asimetrisine yol açar. Sağ ve sol el kavrama kuvvetinden türetilecek **Bilateral Asimetri İndeksi (%)** ve Görsel Reaksiyon ile Çeviklik (Pro-Agility) arasındaki transfer analizi, çalışmaya klinik/nörogelişimsel derinlik kazandırır.

---

## 2. Ham Veriden Türetilecek İkincil Biyomekanik Parametreler

1. **Vücut Kitle İndeksi (VKİ) ve VKİ Z-Skoru (BAZ):**
   $$\text{VKİ} = \frac{\text{Ağırlık (kg)}}{\text{Boy (m)}^2}$$
2. **Biyolojik Olgunlaşma (Maturity Offset - Yıl):**
   Mirwald denklemi: Bireyin PHV'den kaç yıl önce veya sonra olduğunu belirler.
3. **Ape İndeksi (Kulaç / Boy Oranı):**
   $$\text{Ape İndeksi} = \frac{\text{Kulaç Uzunluğu (cm)}}{\text{Boy (cm)}}$$
4. **Rölatif Karış Uzunluğu:**
   $$\text{Rölatif Karış} = \frac{\text{Karış (cm)}}{\text{Boy (cm)}}$$
5. **Bilateral Kavrama Kuvveti Asimetrisi (%):**
   $$\text{Asimetri (\%)} = \frac{|\text{Sağ El} - \text{Sol El}|}{\max(\text{Sağ El}, \text{Sol El})} \times 100$$
6. **Rölatif Toplam Kavrama Kuvveti:**
   $$\text{Rölatif Kuvvet} = \frac{\text{Sağ El (kg)} + \text{Sol El (kg)}}{\text{Ağırlık (kg)}}$$
7. **Alt Ekstremite Zirve Anaerobik Güç (Sayers Formülü - Watt):**
   $$\text{Peak Power (W)} = 60.7 \times \text{Aktif Sıçrama (cm)} + 45.3 \times \text{Ağırlık (kg)} - 2055$$

---

## 3. İstatistiksel Modelleme ve Analiz Aşamaları

### Aşama 1: Tanımlayıcı İstatistikler & Normallik
* Shapiro-Wilk testi ve Q-Q grafikleri.
* Ortalama ± Standart Sapma (parametrik) veya Medyan (Çeyrekler Arası Genişlik - IQR).

### Aşama 2: Zaman Ana Etkisi (Zaman x Olgunlaşma)
* **Tekrarlı Ölçümlerde Kovaryans Analizi (RM-ANCOVA)** veya **Doğrusal Karma Modeller (Linear Mixed Models - LMM)**:
  * Sabit Etki: Zaman (Ön Test vs. Son Test)
  * Kovaryatlar: Takvim Yaşı, Biyolojik Olgunlaşma (PHV Offset), Başlangıç Antropometrisi (VKİ)
  * Etki Büyüklüğü: Kısmi Eta Kare ($\eta_p^2$) ve Cohen's $d_{av}$.

### Aşama 3: Bireysel Yanıtlayıcı Analizi (Responder / Non-Responder)
* **En Küçük Değerli Değişim (Smallest Worthwhile Change - SWC):**
  $$\text{SWC} = 0.2 \times \text{SD}_{\text{ön-test}}$$
* Her bireyin farkı ($\Delta = \text{Son} - \text{Ön}$) SWC ile kıyaslanır:
  * $\Delta > \text{SWC}$: Olumlu Yanıt Veren (Positive Responder)
  * $-\text{SWC} \le \Delta \le \text{SWC}$: Stabil / Nötr (Non-Responder)
  * $\Delta < -\text{SWC}$: Olumsuz Yanıt (Negative Responder)

### Aşama 4: Regresyon ve Belirleyicilik Modelleri
* **Model 1 (Üst Ekstremite):** El kavrama kuvveti artışının ve asimetri azalmasının yordayıcıları (Bağımsız: Karış, Kulaç, VKİ, Biyolojik Yaş).
* **Model 2 (Algısal-Motor ve Çeviklik):** Pro-Agility çeviklik süresindeki değişimin yordayıcıları (Bağımsız: $\Delta$ Görsel Reaksiyon, $\Delta$ Aktif Sıçrama Gücü, Bacak Boyu, Başlangıç Yaşı).

---

## 4. Q1 Düzeyinde Makale Bölüm Taslağı ve Yazım Yol Haritası

```
├── Adım 1: Veri Tabanı ve İstatistiksel Analiz Script'lerinin Kurulması
├── Adım 2: Kesinleştirilmiş Başlık, Hipotezler ve Yapılandırılmış Özet (Abstract)
├── Adım 3: Giriş (Introduction) - Nörogelişimsel Zemin, Plato Hipotezi ve Rasyonel
├── Adım 4: Yöntem (Methods) - Katılımcılar, STROBE Standartları, Test Protokolleri, İstatistik
├── Adım 5: Bulgular (Results) - APA 7.0 Formatında Tablolar ve Grafik Tasarımları
├── Adım 6: Tartışma (Discussion) - Fizyolojik, Nöromotor ve Fonksiyonel Mekanizmalar
└── Adım 7: Kısıtlılıklar, Pratik Uygulama Önerileri ve Sonuç (Limitations & Conclusion)
```
