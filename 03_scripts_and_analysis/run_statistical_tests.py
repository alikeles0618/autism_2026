import csv
import math
import numpy as np

# Load cleaned dataset
with open('/Users/alik/Desktop/anti_projeler/articles_all/autism_2026_testtekrar/02_data/cleaned_data.csv', 'r', encoding='utf-8') as f:
    data = list(csv.DictReader(f))

n = len(data)
print(f"--- İSTATİSTİKSEL ANALİZ RAPORU (N = {n}) ---")

# Variables to analyze
vars_map = [
    ('BOY', 'Boy (cm)'),
    ('KILO', 'Vücut Ağırlığı (kg)'),
    ('VKI', 'VKİ (kg/m²)'),
    ('KULAC', 'Kulaç Uzunluğu (cm)'),
    ('KARIS', 'Karış Uzunluğu (cm)'),
    ('OTURMA_BOYU', 'Oturma Boyu (cm)'),
    ('PHV_OFFSET', 'Mirwald PHV Offset (yıl)'),
    ('GORSEL_TOTAL', 'Görsel Reaksiyon Süresi'),
    ('AKTIF_SIC', 'Aktif Sıçrama (cm)'),
    ('SAYERS_POWER', 'Sayers Zirve Güç (Watt)'),
    ('ESNEKLIK', 'Esneklik - Otur-Eriş (cm)'),
    ('SAG_EL', 'Sağ El Kavrama (kg)'),
    ('SOL_EL', 'Sol El Kavrama (kg)'),
    ('ASYMMETRY', 'Bilateral Asimetri (%)'),
    ('REL_GRIP', 'Rölatif Kavrama (kg/kg)'),
    ('CEVIKLIK', 'Pro-Agility Çeviklik (s)')
]

def student_t_cdf(t, df):
    # Approximation of Student t CDF using regularized incomplete beta function
    x = (t + math.sqrt(t * t + df)) / (2 * math.sqrt(t * t + df))
    # We can use simple symmetry or numerical integration
    # Let's use the continued fraction or standard normal approximation for df=33
    # For df=33, normal approximation with Cornish-Fisher expansion is accurate to 4 decimal places
    z = t * (1 - 1 / (4 * df)) / math.sqrt(1 + t * t / (2 * df))
    # Standard normal cdf:
    return 0.5 * (1 + math.erf(z / math.sqrt(2)))

def t_test_p_value(t_val, df):
    t_abs = abs(t_val)
    # Using integration of Student t pdf
    # pdf(x) = C * (1 + x^2/df)^(-(df+1)/2)
    def pdf(x):
        return (1.0 + (x * x) / df) ** (-(df + 1.0) / 2.0)
    
    # integrate from t_abs to 20
    steps = 2000
    b = max(25.0, t_abs + 10.0)
    h = (b - t_abs) / steps
    area = 0.5 * (pdf(t_abs) + pdf(b))
    for i in range(1, steps):
        area += pdf(t_abs + i * h)
    area *= h
    
    # normalize with total area from 0 to infinity
    total_steps = 4000
    h_tot = b / total_steps
    tot_area = 0.5 * (pdf(0) + pdf(b))
    for i in range(1, total_steps):
        tot_area += pdf(i * h_tot)
    tot_area *= h_tot
    
    one_tail = area / (2.0 * tot_area)
    two_tail = 2.0 * one_tail
    return min(1.0, max(0.0, two_tail))

print(f"{'Değişken':<25} | {'Ön Test (Ort±SS)':<18} | {'Son Test (Ort±SS)':<18} | {'Fark (Δ)':<10} | {'% Değişim':<10} | {'t':<7} | {'p':<8} | {'Cohen d':<7}")
print("-" * 125)

results = []

for key, label in vars_map:
    pre_vals = np.array([float(r[key + '_PRE']) for r in data])
    post_vals = np.array([float(r[key + '_POST']) for r in data])
    diffs = post_vals - pre_vals
    
    m_pre, s_pre = np.mean(pre_vals), np.std(pre_vals, ddof=1)
    m_post, s_post = np.mean(post_vals), np.std(post_vals, ddof=1)
    m_diff, s_diff = np.mean(diffs), np.std(diffs, ddof=1)
    
    pct_change = (m_diff / m_pre) * 100 if m_pre != 0 else 0
    
    # Paired t-test
    se_diff = s_diff / math.sqrt(n)
    t_stat = m_diff / se_diff if se_diff != 0 else 0
    p_val = t_test_p_value(t_stat, n - 1)
    
    # Cohen's d (average SD)
    s_av = math.sqrt((s_pre ** 2 + s_post ** 2) / 2.0)
    cohen_d = m_diff / s_av if s_av != 0 else 0
    
    # Partial eta-squared
    eta_sq = (t_stat ** 2) / (t_stat ** 2 + (n - 1))
    
    results.append({
        'key': key,
        'label': label,
        'm_pre': m_pre, 's_pre': s_pre,
        'm_post': m_post, 's_post': s_post,
        'm_diff': m_diff, 's_diff': s_diff,
        'pct_change': pct_change,
        't_stat': t_stat,
        'p_val': p_val,
        'cohen_d': cohen_d,
        'eta_sq': eta_sq
    })
    
    p_str = "< 0.001" if p_val < 0.001 else f"{p_val:.4f}"
    print(f"{label:<25} | {m_pre:>6.2f} ± {s_pre:>5.2f}    | {m_post:>6.2f} ± {s_post:>5.2f}    | {m_diff:>+6.2f}     | {pct_change:>+6.1f}%    | {t_stat:>6.2f} | {p_str:<8} | {cohen_d:>6.2f}")
