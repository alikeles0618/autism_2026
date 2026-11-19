import csv

with open('/Users/alik/Desktop/anti_projeler/articles_all/autism_2026_testtekrar/02_data/raw_data.csv', 'r', encoding='utf-8') as f:
    rows = list(csv.DictReader(f))

cols = ['BOY', 'KİLO', 'KULAÇ', 'KARIŞ', 'OTURMA_BOYU', 'GORSEL_TOTAL', 'AKTIF_SIC', 'ESNEKLIK', 'SAG_EL', 'SOL_EL', 'CEVIKLIK']

he_active = []
for r in rows:
    if r['grup'] == 'Hareket Eğitimi':
        is_same = all(r[c+'_PRE'] == r[c+'_POST'] for c in cols)
        if not is_same:
            he_active.append(r)

print(f"Hareket Eğitimi Aktif Katılımcı Sayısı (Ön ve Son testi olan): {len(he_active)}")

# Let's inspect each active participant in detail
for idx, r in enumerate(he_active, 1):
    bid = r['BİREY ID']
    no = r['NO']
    birth = int(r['DOĞUM T.'])
    age_2021 = 2021 - birth
    
    h_pre, h_post = float(r['BOY_PRE']), float(r['BOY_POST'])
    w_pre, w_post = float(r['KİLO_PRE']), float(r['KİLO_POST'])
    k_pre, k_post = float(r['KULAÇ_PRE']), float(r['KULAÇ_POST'])
    kar_pre, kar_post = float(r['KARIŞ_PRE']), float(r['KARIŞ_POST'])
    ot_pre, ot_post = float(r['OTURMA_BOYU_PRE']), float(r['OTURMA_BOYU_POST'])
    g_pre, g_post = float(r['GORSEL_TOTAL_PRE']), float(r['GORSEL_TOTAL_POST'])
    s_pre, s_post = float(r['AKTIF_SIC_PRE']), float(r['AKTIF_SIC_POST'])
    es_pre, es_post = float(r['ESNEKLIK_PRE']), float(r['ESNEKLIK_POST'])
    sag_pre, sag_post = float(r['SAG_EL_PRE']), float(r['SAG_EL_POST'])
    sol_pre, sol_post = float(r['SOL_EL_PRE']), float(r['SOL_EL_POST'])
    cev_pre, cev_post = float(r['CEVIKLIK_PRE']), float(r['CEVIKLIK_POST'])
    
    dh = h_post - h_pre
    dw = w_post - w_pre
    
    # Check any specific large differences in height or other variables
    notes = []
    if dh > 3.0:
        notes.append(f"Boy uzaması: +{dh:.1f} cm ({h_pre}->{h_post})")
    elif dh < 0:
        notes.append(f"Boy kısalması: {dh:.1f} cm ({h_pre}->{h_post})")
    if abs(dw) > 5.0:
        notes.append(f"Kilo değişimi: {dw:+.1f} kg ({w_pre}->{w_post})")
    if abs(ot_post - ot_pre) > 4.0:
        notes.append(f"Oturma boyu farkı: {ot_post-ot_pre:+.1f} cm ({ot_pre}->{ot_post})")
    if g_pre > 100 or g_post > 100:
        notes.append(f"Görsel uç değer: {g_pre}/{g_post}")
        
    print(f"#{idx:02d} | ID:{bid:>2} | Doğum:{birth} (Yaş {age_2021:02d}) | Boy:{h_pre}->{h_post} | Kilo:{w_pre}->{w_post} | Not: {', '.join(notes) if notes else 'OK'}")
