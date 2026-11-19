import csv
import numpy as np

with open('/Users/alik/Desktop/anti_projeler/articles_all/autism_2026_testtekrar/02_data/cleaned_data.csv', 'r', encoding='utf-8') as f:
    data = list(csv.DictReader(f))

keys = ['DELTA_GORSEL', 'DELTA_CEVIKLIK', 'DELTA_SAG_EL', 'DELTA_SOL_EL', 'DELTA_AKTIF_SIC', 'DELTA_ESNEKLIK', 'DELTA_SAYERS_POWER', 'DELTA_ASYMMETRY', 'YAS_2021', 'PHV_OFFSET_PRE', 'VKI_PRE', 'KARIS_PRE', 'KULAC_PRE']
labels = ['ΔGörsel', 'ΔÇeviklik', 'ΔSağEl', 'ΔSolEl', 'ΔSıçrama', 'ΔEsneklik', 'ΔGüç', 'ΔAsimetri', 'Yaş', 'PHV', 'VKİ', 'Karış', 'Kulaç']

mat = []
for k in keys:
    mat.append([float(r[k]) for r in data])

mat = np.array(mat)
corr = np.corrcoef(mat)

print(f"{'Değişken':<12} | " + " | ".join([f"{l:>9}" for l in labels]))
print("-" * 150)
for i, l in enumerate(labels):
    row_str = f"{l:<12} | " + " | ".join([f"{corr[i, j]:>9.3f}" for j in range(len(labels))])
    print(row_str)
