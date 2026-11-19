import csv
from collections import Counter

with open('/Users/alik/Desktop/anti_projeler/articles_all/autism_2026_testtekrar/02_data/raw_data.csv', 'r', encoding='utf-8') as f:
    reader = csv.DictReader(f)
    rows = list(reader)

print(f"Toplam Satır: {len(rows)}")

groups = Counter(r['grup'] for r in rows)
print("Grup Dağılımı:")
for g, count in groups.items():
    print(f"  {g}: {count}")

# Check unique BİREY ID
b_ids = [r['BİREY ID'] for r in rows]
print(f"Tekil Birey ID Sayısı: {len(set(b_ids))}")

# Check if pre-test and post-test values are identical for some participants!
# Look at ID 4: BOY 125, 125, KİLO 24.8, 24.8, KULAÇ 121, 121, KARIŞ 15, 15, OTURMA 62.5, 62.5, GORSEL 38.6, 38.6, AKTIF_SIC 14.1, 14.1, ESNEKLIK 6, 6, SAG_EL 6.8, 6.8, SOL_EL 7.4, 7.4, CEVIKLIK 17.18, 17.18
# They are 100% IDENTICAL pre and post!
identical_count = 0
for r in rows:
    cols = ['BOY', 'KİLO', 'KULAÇ', 'KARIŞ', 'OTURMA_BOYU', 'GORSEL_TOTAL', 'AKTIF_SIC', 'ESNEKLIK', 'SAG_EL', 'SOL_EL', 'CEVIKLIK']
    all_same = True
    for c in cols:
        if r[c+'_PRE'] != r[c+'_POST']:
            all_same = False
            break
    if all_same:
        identical_count += 1
        print(f"TÜM DEĞERLERİ BİREBİR AYNI OLAN (Değişim Sıfır / Belki son test yapılmadı?): Birey ID {r['BİREY ID']}, Grup: {r['grup']}")

print(f"\nToplam Değişimi Olmayan (Tüm sütunları Ön=Son olan) Birey Sayısı: {identical_count}")
print(f"Değişim gösteren birey sayısı: {len(rows) - identical_count}")
