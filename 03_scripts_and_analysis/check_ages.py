import csv

with open('/Users/alik/Desktop/anti_projeler/articles_all/autism_2026_testtekrar/02_data/raw_data.csv', 'r', encoding='utf-8') as f:
    rows = list(csv.DictReader(f))

years = [int(r['DOĞUM T.']) for r in rows]
ages = [2021 - y for y in years]

print(f"Doğum Yılları: Min={min(years)}, Max={max(years)}")
print(f"2021 Yaşları: Min={min(ages)}, Max={max(ages)}, Ortalama={sum(ages)/len(ages):.1f}")

from collections import Counter
print("Yaş Dağılımı (2021 bazlı):")
for age, count in sorted(Counter(ages).items()):
    print(f"  {age} Yaş (Doğum {2021-age}): {count} birey")
