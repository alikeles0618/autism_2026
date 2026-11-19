import csv

with open('/Users/alik/Desktop/anti_projeler/articles_all/autism_2026_testtekrar/02_data/raw_data.csv', 'r', encoding='utf-8') as f:
    reader = csv.DictReader(f)
    rows = list(reader)

print(f"Toplam Satır Sayısı: {len(rows)}")

# Sütun isimleri
print("Sütunlar:", reader.fieldnames)

anomalies = []

for idx, r in enumerate(rows, 1):
    bid = r['BİREY ID']
    no = r['NO']
    grp = r['grup']
    birth = int(r['DOĞUM T.'])
    age_2021 = 2021 - birth
    
    # Boy
    h_pre = float(r['BOY_PRE'])
    h_post = float(r['BOY_POST'])
    dh = h_post - h_pre
    
    # Kilo
    w_pre = float(r['KİLO_PRE'])
    w_post = float(r['KİLO_POST'])
    dw = w_post - w_pre
    
    # Kulaç
    k_pre = float(r['KULAÇ_PRE'])
    k_post = float(r['KULAÇ_POST'])
    
    # Karış
    kar_pre = float(r['KARIŞ_PRE'])
    kar_post = float(r['KARIŞ_POST'])
    
    # Oturma Boyu
    ot_pre = float(r['OTURMA_BOYU_PRE'])
    ot_post = float(r['OTURMA_BOYU_POST'])
    
    # Görsel Total
    g_pre = float(r['GORSEL_TOTAL_PRE'])
    g_post = float(r['GORSEL_TOTAL_POST'])
    
    # Aktif Sıçrama
    s_pre = float(r['AKTIF_SIC_PRE'])
    s_post = float(r['AKTIF_SIC_POST'])
    
    # Esneklik
    es_pre = float(r['ESNEKLIK_PRE'])
    es_post = float(r['ESNEKLIK_POST'])
    
    # Sağ El
    sag_pre = float(r['SAG_EL_PRE'])
    sag_post = float(r['SAG_EL_POST'])
    
    # Sol El
    sol_pre = float(r['SOL_EL_PRE'])
    sol_post = float(r['SOL_EL_POST'])
    
    # Çeviklik
    cev_pre = float(r['CEVIKLIK_PRE'])
    cev_post = float(r['CEVIKLIK_POST'])
    
    # Kontroller:
    # 1. Boy azalması veya 12 haftada aşırı büyüme (> 4 cm)
    if dh < 0:
        anomalies.append((bid, "BOY", f"Son test boyu ön testten küçük: Ön={h_pre}, Son={h_post} (Fark: {dh:.1f} cm)"))
    elif dh > 4.0:
        anomalies.append((bid, "BOY", f"12 haftada aşırı boy uzaması: Ön={h_pre}, Son={h_post} (Fark: +{dh:.1f} cm)"))
        
    # 2. Oturma boyu kontrolleri
    if ot_pre >= h_pre or ot_post >= h_post:
        anomalies.append((bid, "OTURMA BOYU", f"Oturma boyu ayakta boydan büyük/eşit: Oturma Ön={ot_pre}, Boy Ön={h_pre}"))
    if abs(ot_post - ot_pre) > 5.0:
        anomalies.append((bid, "OTURMA BOYU", f"12 haftada oturma boyunda aşırı değişim: Ön={ot_pre}, Son={ot_post} (Fark: {ot_post-ot_pre:.1f} cm)"))
    # Oran kontrolü: oturma boyu / boy genelde %48 - %60 arasındadır
    ratio_pre = ot_pre / h_pre
    ratio_post = ot_post / h_post
    if ratio_pre < 0.40 or ratio_pre > 0.65 or ratio_post < 0.40 or ratio_post > 0.65:
        anomalies.append((bid, "OTURMA BOYU ORANI", f"Oturma boyu / Boy oranı anormal: Ön Oran=%{ratio_pre*100:.1f}, Son Oran=%{ratio_post*100:.1f}"))

    # 3. Görsel Reaksiyon süresi (genelde 10 - 70 arası, 549 bariz yazım hatası)
    if g_pre > 100 or g_post > 100:
        anomalies.append((bid, "GÖRSEL TOTAL", f"Aşırı yüksek değer (muhtemelen virgül/nokta hatası): Ön={g_pre}, Son={g_post}"))
        
    # 4. Kulaç vs Boy (Ape Index)
    ape_pre = k_pre / h_pre
    ape_post = k_post / h_post
    if ape_pre < 0.85 or ape_pre > 1.15 or ape_post < 0.85 or ape_post > 1.15:
        anomalies.append((bid, "KULAÇ / BOY", f"Ape indeksi uç değerde: Ön={ape_pre:.2f}, Son={ape_post:.2f} (Kulaç={k_pre}, Boy={h_pre})"))
        
    # 5. Karış boyutu
    if kar_pre < 10 or kar_pre > 30 or kar_post < 10 or kar_post > 30:
        anomalies.append((bid, "KARIŞ", f"Karış uzunluğu fizyolojik dışı: Ön={kar_pre}, Son={kar_post}"))

    # 6. Kilo aşırı değişim (> 6 kg değişim 12 haftada çocuklarda)
    if abs(dw) > 7.0:
        anomalies.append((bid, "KİLO", f"12 haftada aşırı kilo değişimi: Ön={w_pre}, Son={w_post} (Fark: {dw:.2f} kg)"))

    # 7. El kavrama kuvveti (aşırı değişim veya uç değerler)
    if abs(sag_post - sag_pre) > 10.0:
        anomalies.append((bid, "SAĞ EL KAVRAMA", f"Kavrama kuvvetinde aşırı sıçrama: Ön={sag_pre}, Son={sag_post}"))
    if abs(sol_post - sol_pre) > 10.0:
        anomalies.append((bid, "SOL EL KAVRAMA", f"Kavrama kuvvetinde aşırı sıçrama: Ön={sol_pre}, Son={sol_post}"))

    # 8. Çeviklik
    if abs(cev_post - cev_pre) > 6.0:
        anomalies.append((bid, "ÇEVİKLİK", f"Çeviklik süresinde aşırı değişim: Ön={cev_pre}, Son={cev_post}"))

print(f"\nToplam Tespit Edilen Olası Anomali/Hata Sayısı: {len(anomalies)}")
for a in anomalies:
    print(f"Birey ID {a[0]} | {a[1]}: {a[2]}")
