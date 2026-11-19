import csv

with open('/Users/alik/Desktop/anti_projeler/articles_all/autism_2026_testtekrar/02_data/raw_data.csv', 'r', encoding='utf-8') as f:
    rows = list(csv.DictReader(f))

cols = ['BOY', 'KİLO', 'KULAÇ', 'KARIŞ', 'OTURMA_BOYU', 'GORSEL_TOTAL', 'AKTIF_SIC', 'ESNEKLIK', 'SAG_EL', 'SOL_EL', 'CEVIKLIK']

he_rows = [r for r in rows if r['grup'] == 'Hareket Eğitimi']
oe_rows = [r for r in rows if r['grup'] == 'Özel Eğitim']

print(f"Hareket Eğitimi Toplam: {len(he_rows)}")
print(f"Özel Eğitim Toplam: {len(oe_rows)}")

def analyze_subset(subset, name):
    print(f"\n--- {name} (N={len(subset)}) ---")
    identical = 0
    changed = 0
    for r in subset:
        is_same = all(r[c+'_PRE'] == r[c+'_POST'] for c in cols)
        if is_same:
            identical += 1
        else:
            changed += 1
    print(f"Ön = Son (Birebir Aynı Olan): {identical}")
    print(f"Ön != Son (Farklı Olan): {changed}")

analyze_subset(he_rows, "Hareket Eğitimi")
analyze_subset(oe_rows, "Özel Eğitim")

# Let's inspect some rows where pre == post
print("\nHareket Eğitimi İçinde Ön=Son olan Bireylerin ID'leri ve NO'ları:")
for r in he_rows:
    if all(r[c+'_PRE'] == r[c+'_POST'] for c in cols):
        print(f"  NO={r['NO']}, ID={r['BİREY ID']}, Doğum={r['DOĞUM T.']}")

print("\nÖzel Eğitim İçinde Ön=Son olan Bireylerin ID'leri ve NO'ları:")
for r in oe_rows:
    if all(r[c+'_PRE'] == r[c+'_POST'] for c in cols):
        print(f"  NO={r['NO']}, ID={r['BİREY ID']}, Doğum={r['DOĞUM T.']}")
