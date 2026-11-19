import csv
import numpy as np

with open('/Users/alik/Desktop/anti_projeler/articles_all/autism_2026_testtekrar/02_data/raw_data.csv', 'r', encoding='utf-8') as f:
    rows = [r for r in csv.DictReader(f) if r['grup'] == 'Hareket Eğitimi']

cols = ['BOY', 'KİLO', 'KULAÇ', 'KARIŞ', 'OTURMA_BOYU', 'GORSEL_TOTAL', 'AKTIF_SIC', 'ESNEKLIK', 'SAG_EL', 'SOL_EL', 'CEVIKLIK']

completers = []
dropouts = []

for r in rows:
    is_same = all(r[c+'_PRE'] == r[c+'_POST'] for c in cols)
    if is_same:
        dropouts.append(r)
    else:
        completers.append(r)

print(f"Hareket Eğitimi: Toplam={len(rows)}, Tamamlayan={len(completers)}, Drop-out={len(dropouts)}")

# Compare baseline variables between completers and dropouts
print("\n--- BASELINE (ÖN TEST) KARŞILAŞTIRMASI: Tamamlayanlar vs Drop-out ---")
print(f"{'Değişken':<18} | {'Tamamlayan (n=34)':<22} | {'Drop-out (n=23)':<22} | {'Fark':<8}")
print("-" * 75)

for c in ['DOĞUM T.', 'BOY_PRE', 'KİLO_PRE', 'KULAÇ_PRE', 'KARIŞ_PRE', 'OTURMA_BOYU_PRE', 'GORSEL_TOTAL_PRE', 'AKTIF_SIC_PRE', 'ESNEKLIK_PRE', 'SAG_EL_PRE', 'SOL_EL_PRE', 'CEVIKLIK_PRE']:
    v_comp = [float(r[c]) for r in completers]
    v_drop = [float(r[c]) for r in dropouts]
    
    m_c, s_c = np.mean(v_comp), np.std(v_comp, ddof=1)
    m_d, s_d = np.mean(v_drop), np.std(v_drop, ddof=1)
    
    label = c.replace('_PRE', '')
    if c == 'DOĞUM T.':
        label = 'Yaş (2021)'
        v_comp = [2021 - float(r['DOĞUM T.']) for r in completers]
        v_drop = [2021 - float(r['DOĞUM T.']) for r in dropouts]
        m_c, s_c = np.mean(v_comp), np.std(v_comp, ddof=1)
        m_d, s_d = np.mean(v_drop), np.std(v_drop, ddof=1)
        
    print(f"{label:<18} | {m_c:>6.2f} ± {s_c:>5.2f}          | {m_d:>6.2f} ± {s_d:>5.2f}          | {m_c - m_d:>+6.2f}")
