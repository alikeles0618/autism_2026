import csv
import numpy as np

def calculate_mirwald_phv_boys(height_cm, sitting_height_cm, weight_kg, chronological_age):
    leg_length_cm = height_cm - sitting_height_cm
    # Mirwald 2002 Male Equation
    offset = (
        -9.236
        + 0.0002708 * (leg_length_cm * sitting_height_cm)
        - 0.001663 * (chronological_age * leg_length_cm)
        + 0.007216 * (chronological_age * sitting_height_cm)
        + 0.02292 * (weight_kg / height_cm * 100)
    )
    return offset

def calculate_bmi(weight_kg, height_cm):
    height_m = height_cm / 100.0
    return weight_kg / (height_m ** 2)

def calculate_bilateral_asymmetry(right_grip, left_grip):
    max_grip = max(right_grip, left_grip)
    if max_grip == 0:
        return 0.0
    return (abs(right_grip - left_grip) / max_grip) * 100.0

def calculate_sayers_power(jump_height_cm, weight_kg):
    # Sayers et al. (1999) Peak Power (Watts)
    # Peak Power (W) = 60.7 * Jump Height (cm) + 45.3 * Weight (kg) - 2055
    return 60.7 * jump_height_cm + 45.3 * weight_kg - 2055.0

with open('/Users/alik/Desktop/anti_projeler/articles_all/autism_2026_testtekrar/02_data/raw_data.csv', 'r', encoding='utf-8') as f:
    rows = list(csv.DictReader(f))

cols = ['BOY', 'KİLO', 'KULAÇ', 'KARIŞ', 'OTURMA_BOYU', 'GORSEL_TOTAL', 'AKTIF_SIC', 'ESNEKLIK', 'SAG_EL', 'SOL_EL', 'CEVIKLIK']

cleaned_rows = []

for r in rows:
    # 1. Sadece Hareket Eğitimi
    if r['grup'] != 'Hareket Eğitimi':
        continue
    # 2. Drop-out olanları (ön=son olanları) filtrele
    is_same = all(r[c+'_PRE'] == r[c+'_POST'] for c in cols)
    if is_same:
        continue
        
    bid = r['BİREY ID']
    no = r['NO']
    birth = int(r['DOĞUM T.'])
    age_2021 = 2021 - birth
    
    h_pre = float(r['BOY_PRE'])
    h_post = float(r['BOY_POST'])
    w_pre = float(r['KİLO_PRE'])
    w_post = float(r['KİLO_POST'])
    k_pre = float(r['KULAÇ_PRE'])
    k_post = float(r['KULAÇ_POST'])
    kar_pre = float(r['KARIŞ_PRE'])
    kar_post = float(r['KARIŞ_POST'])
    ot_pre = float(r['OTURMA_BOYU_PRE'])
    ot_post = float(r['OTURMA_BOYU_POST'])
    g_pre = float(r['GORSEL_TOTAL_PRE'])
    g_post = float(r['GORSEL_TOTAL_POST'])
    s_pre = float(r['AKTIF_SIC_PRE'])
    s_post = float(r['AKTIF_SIC_POST'])
    es_pre = float(r['ESNEKLIK_PRE'])
    es_post = float(r['ESNEKLIK_POST'])
    sag_pre = float(r['SAG_EL_PRE'])
    sag_post = float(r['SAG_EL_POST'])
    sol_pre = float(r['SOL_EL_PRE'])
    sol_post = float(r['SOL_EL_POST'])
    cev_pre = float(r['CEVIKLIK_PRE'])
    cev_post = float(r['CEVIKLIK_POST'])
    
    # ID 85 boy kontrolü: 113.2'den 123.0 yazılmış, ancak kulaç 110-113, oturma boyu 58-60.
    # Eğer 123 girilmişse koruyabiliriz veya 115.0 olarak düzeltebiliriz. Burada orijinali 123.0 tutalım ama not düşelim veya mantıklı sınırda inceleyelim.
    
    # İkincil parametrelerin türetilmesi:
    bmi_pre = calculate_bmi(w_pre, h_pre)
    bmi_post = calculate_bmi(w_post, h_post)
    
    phv_pre = calculate_mirwald_phv_boys(h_pre, ot_pre, w_pre, age_2021)
    phv_post = calculate_mirwald_phv_boys(h_post, ot_post, w_post, age_2021 + 0.25) # 12 hafta ~ 0.25 yıl
    
    ape_pre = k_pre / h_pre
    ape_post = k_post / h_post
    
    rel_kar_pre = kar_pre / h_pre
    rel_kar_post = kar_post / h_post
    
    asym_pre = calculate_bilateral_asymmetry(sag_pre, sol_pre)
    asym_post = calculate_bilateral_asymmetry(sag_post, sol_post)
    
    rel_grip_pre = (sag_pre + sol_pre) / w_pre
    rel_grip_post = (sag_post + sol_post) / w_post
    
    sayers_pre = calculate_sayers_power(s_pre, w_pre)
    sayers_post = calculate_sayers_power(s_post, w_post)
    
    cleaned_rows.append({
        'NO': no,
        'BIREY_ID': bid,
        'CINSIYET': 'Erkek',
        'DOGUM_YILI': birth,
        'YAS_2021': age_2021,
        'BOY_PRE': h_pre,
        'BOY_POST': h_post,
        'DELTA_BOY': round(h_post - h_pre, 2),
        'KILO_PRE': w_pre,
        'KILO_POST': w_post,
        'DELTA_KILO': round(w_post - w_pre, 2),
        'VKI_PRE': round(bmi_pre, 2),
        'VKI_POST': round(bmi_post, 2),
        'DELTA_VKI': round(bmi_post - bmi_pre, 2),
        'KULAC_PRE': k_pre,
        'KULAC_POST': k_post,
        'DELTA_KULAC': round(k_post - k_pre, 2),
        'APE_INDEX_PRE': round(ape_pre, 3),
        'APE_INDEX_POST': round(ape_post, 3),
        'KARIS_PRE': kar_pre,
        'KARIS_POST': kar_post,
        'DELTA_KARIS': round(kar_post - kar_pre, 2),
        'OTURMA_BOYU_PRE': ot_pre,
        'OTURMA_BOYU_POST': ot_post,
        'DELTA_OTURMA_BOYU': round(ot_post - ot_pre, 2),
        'PHV_OFFSET_PRE': round(phv_pre, 2),
        'PHV_OFFSET_POST': round(phv_post, 2),
        'GORSEL_TOTAL_PRE': g_pre,
        'GORSEL_TOTAL_POST': g_post,
        'DELTA_GORSEL': round(g_post - g_pre, 2),
        'PCT_DELTA_GORSEL': round(((g_post - g_pre) / g_pre) * 100, 2),
        'AKTIF_SIC_PRE': s_pre,
        'AKTIF_SIC_POST': s_post,
        'DELTA_AKTIF_SIC': round(s_post - s_pre, 2),
        'PCT_DELTA_AKTIF_SIC': round(((s_post - s_pre) / s_pre) * 100, 2),
        'SAYERS_POWER_PRE': round(sayers_pre, 1),
        'SAYERS_POWER_POST': round(sayers_post, 1),
        'DELTA_SAYERS_POWER': round(sayers_post - sayers_pre, 1),
        'ESNEKLIK_PRE': es_pre,
        'ESNEKLIK_POST': es_post,
        'DELTA_ESNEKLIK': round(es_post - es_pre, 2),
        'PCT_DELTA_ESNEKLIK': round(((es_post - es_pre) / es_pre) * 100, 2),
        'SAG_EL_PRE': sag_pre,
        'SAG_EL_POST': sag_post,
        'DELTA_SAG_EL': round(sag_post - sag_pre, 2),
        'SOL_EL_PRE': sol_pre,
        'SOL_EL_POST': sol_post,
        'DELTA_SOL_EL': round(sol_post - sol_pre, 2),
        'ASYMMETRY_PRE': round(asym_pre, 2),
        'ASYMMETRY_POST': round(asym_post, 2),
        'DELTA_ASYMMETRY': round(asym_post - asym_pre, 2),
        'REL_GRIP_PRE': round(rel_grip_pre, 3),
        'REL_GRIP_POST': round(rel_grip_post, 3),
        'DELTA_REL_GRIP': round(rel_grip_post - rel_grip_pre, 3),
        'CEVIKLIK_PRE': cev_pre,
        'CEVIKLIK_POST': cev_post,
        'DELTA_CEVIKLIK': round(cev_post - cev_pre, 3),
        'PCT_DELTA_CEVIKLIK': round(((cev_post - cev_pre) / cev_pre) * 100, 2)
    })

fieldnames = list(cleaned_rows[0].keys())

with open('/Users/alik/Desktop/anti_projeler/articles_all/autism_2026_testtekrar/02_data/cleaned_data.csv', 'w', newline='', encoding='utf-8') as f:
    writer = csv.DictWriter(f, fieldnames=fieldnames)
    writer.writeheader()
    writer.writerows(cleaned_rows)

print(f"cleaned_data.csv başarıyla oluşturuldu! Toplam Katılımcı: {len(cleaned_rows)}")
