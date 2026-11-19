import csv
import math
import numpy as np

with open('/Users/alik/Desktop/anti_projeler/articles_all/autism_2026_testtekrar/02_data/cleaned_data.csv', 'r', encoding='utf-8') as f:
    data = list(csv.DictReader(f))

n = len(data)

print(f"--- SWC (SMALLEST WORTHWHILE CHANGE) BİREYSEL YANITLAYICI ANALİZİ (N={n}) ---")

# For agility and visual RT, NEGATIVE delta is IMPROVEMENT (faster)
# For grip, power, jump, flexibility, POSITIVE delta is IMPROVEMENT (stronger/higher)

metrics = [
    ('GORSEL_TOTAL', 'Görsel Reaksiyon Süresi', True),  # True = lower is better
    ('CEVIKLIK', 'Pro-Agility Çeviklik', True),
    ('SAG_EL', 'Sağ El Kavrama Kuvveti', False),       # False = higher is better
    ('SOL_EL', 'Sol El Kavrama Kuvveti', False),
    ('REL_GRIP', 'Rölatif Kavrama Kuvveti', False),
    ('SAYERS_POWER', 'Sayers Zirve Güç', False),
    ('AKTIF_SIC', 'Aktif Sıçrama', False),
    ('ESNEKLIK', 'Esneklik', False),
    ('ASYMMETRY', 'Bilateral Asimetri', True)          # lower asymmetry is better
]

for key, label, lower_is_better in metrics:
    pre_vals = np.array([float(r[key + '_PRE']) for r in data])
    post_vals = np.array([float(r[key + '_POST']) for r in data])
    diffs = post_vals - pre_vals
    
    sd_pre = np.std(pre_vals, ddof=1)
    swc = 0.2 * sd_pre
    
    if lower_is_better:
        # improvement is diff < -swc
        responders = np.sum(diffs < -swc)
        trivial = np.sum((diffs >= -swc) & (diffs <= swc))
        non_responders = np.sum(diffs > swc)
    else:
        # improvement is diff > swc
        responders = np.sum(diffs > swc)
        trivial = np.sum((diffs >= -swc) & (diffs <= swc))
        non_responders = np.sum(diffs < -swc)
        
    pct_resp = (responders / n) * 100
    pct_triv = (trivial / n) * 100
    pct_non = (non_responders / n) * 100
    
    print(f"\n{label} (SWC = {swc:.2f}):")
    print(f"  • Olumlu Yanıt Veren (Gelişen - Responder): {responders:2d} birey (%{pct_resp:5.1f})")
    print(f"  • Değişmeyen / Stabil (Trivial/Stable):       {trivial:2d} birey (%{pct_triv:5.1f})")
    print(f"  • Olumsuz / Gerileyen (Non-Responder):      {non_responders:2d} birey (%{pct_non:5.1f})")
