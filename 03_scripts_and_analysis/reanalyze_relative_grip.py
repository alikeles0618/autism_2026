import csv
import math
import numpy as np

with open('/Users/alik/Desktop/anti_projeler/articles_all/autism_2026_testtekrar/02_data/cleaned_data.csv', 'r', encoding='utf-8') as f:
    data = list(csv.DictReader(f))

n = len(data)

# Calculate Total Grip and Relative Total Grip for all 34 participants
rel_tot_pre = []
rel_tot_post = []
tot_pre = []
tot_post = []

for r in data:
    w_pre = float(r['KILO_PRE'])
    w_post = float(r['KILO_POST'])
    sag_pre = float(r['SAG_EL_PRE'])
    sag_post = float(r['SAG_EL_POST'])
    sol_pre = float(r['SOL_EL_PRE'])
    sol_post = float(r['SOL_EL_POST'])
    
    t_pre = sag_pre + sol_pre
    t_post = sag_post + sol_post
    
    rt_pre = t_pre / w_pre
    rt_post = t_post / w_post
    
    tot_pre.append(t_pre)
    tot_post.append(t_post)
    rel_tot_pre.append(rt_pre)
    rel_tot_post.append(rt_post)

rel_tot_pre = np.array(rel_tot_pre)
rel_tot_post = np.array(rel_tot_post)
tot_pre = np.array(tot_pre)
tot_post = np.array(tot_post)

# 1. Total Grip (kg) statistics
m_t_pre, s_t_pre = np.mean(tot_pre), np.std(tot_pre, ddof=1)
m_t_post, s_t_post = np.mean(tot_post), np.std(tot_post, ddof=1)
d_tot = tot_post - tot_pre
m_dt, s_dt = np.mean(d_tot), np.std(d_tot, ddof=1)
t_stat_tot = m_dt / (s_dt / math.sqrt(n))
pct_tot = (m_dt / m_t_pre) * 100
s_av_tot = math.sqrt((s_t_pre**2 + s_t_post**2) / 2.0)
cohen_d_tot = m_dt / s_av_tot

# 2. Relative Total Grip (kg/kg) statistics
m_rt_pre, s_rt_pre = np.mean(rel_tot_pre), np.std(rel_tot_pre, ddof=1)
m_rt_post, s_rt_post = np.mean(rel_tot_post), np.std(rel_tot_post, ddof=1)
d_rtot = rel_tot_post - rel_tot_pre
m_drt, s_drt = np.mean(d_rtot), np.std(d_rtot, ddof=1)
t_stat_rt = m_drt / (s_drt / math.sqrt(n))
pct_rt = (m_drt / m_rt_pre) * 100
s_av_rt = math.sqrt((s_rt_pre**2 + s_rt_post**2) / 2.0)
cohen_d_rt = m_drt / s_av_rt
eta_sq_rt = (t_stat_rt ** 2) / (t_stat_rt ** 2 + (n - 1))

print("=== KAVRAMA KUVVETİ YENİDEN ANALİZ SONUÇLARI (N = 34) ===")
print(f"Toplam Kavrama (kg):       Ön = {m_t_pre:.2f} ± {s_t_pre:.2f} | Son = {m_t_post:.2f} ± {s_t_post:.2f} | Fark = {m_dt:+.2f} ({pct_tot:+.1f}%) | t = {t_stat_tot:.2f}, d = {cohen_d_tot:.2f}")
print(f"Rölatif Toplam (kg/kg):    Ön = {m_rt_pre:.3f} ± {s_rt_pre:.3f} | Son = {m_rt_post:.3f} ± {s_rt_post:.3f} | Fark = {m_drt:+.3f} ({pct_rt:+.1f}%) | t = {t_stat_rt:.2f}, d = {cohen_d_rt:.2f}, ηp² = {eta_sq_rt:.3f}")

# P-value calculation for t_stat_rt
def t_test_p_value(t_val, df):
    t_abs = abs(t_val)
    def pdf(x):
        return (1.0 + (x * x) / df) ** (-(df + 1.0) / 2.0)
    steps = 2000
    b = max(25.0, t_abs + 10.0)
    h = (b - t_abs) / steps
    area = 0.5 * (pdf(t_abs) + pdf(b))
    for i in range(1, steps):
        area += pdf(t_abs + i * h)
    area *= h
    total_steps = 4000
    h_tot = b / total_steps
    tot_area = 0.5 * (pdf(0) + pdf(b))
    for i in range(1, total_steps):
        tot_area += pdf(i * h_tot)
    tot_area *= h_tot
    return min(1.0, max(0.0, 2.0 * (area / (2.0 * tot_area))))

p_val_rt = t_test_p_value(t_stat_rt, n - 1)
p_val_tot = t_test_p_value(t_stat_tot, n - 1)
print(f"p değeri (Toplam Kavrama):       {p_val_tot:.4f}")
print(f"p değeri (Rölatif Toplam):       {p_val_rt:.4f}")

# SWC Responder Analysis for Relative Total Grip
swc_rt = 0.2 * s_rt_pre
resp_rt = np.sum(d_rtot > swc_rt)
triv_rt = np.sum((d_rtot >= -swc_rt) & (d_rtot <= swc_rt))
non_rt = np.sum(d_rtot < -swc_rt)

print(f"\nSWC Yanıtlayıcı Analizi (Rölatif Toplam, SWC = {swc_rt:.3f} kg/kg):")
print(f"  Gelişen (Responder):     {resp_rt:2d} (%{resp_rt/n*100:.1f})")
print(f"  Stabil (Trivial):        {triv_rt:2d} (%{triv_rt/n*100:.1f})")
print(f"  Gerileyen (Non-Resp):    {non_rt:2d} (%{non_rt/n*100:.1f})")

# Multiple Regression for Relative Total Grip Change
Y_reg = d_rtot
X_reg = np.column_stack([
    [float(r['KARIS_PRE']) for r in data],
    [float(r['KULAC_PRE']) for r in data],
    [float(r['VKI_PRE']) for r in data],
    [float(r['PHV_OFFSET_PRE']) for r in data]
])
X_mat = np.column_stack([np.ones(n), X_reg])
XtX_inv = np.linalg.inv(np.dot(X_mat.T, X_mat))
beta = np.dot(np.dot(XtX_inv, X_mat.T), Y_reg)
residuals = Y_reg - np.dot(X_mat, beta)
SS_tot = np.sum((Y_reg - np.mean(Y_reg))**2)
SS_res = np.sum(residuals**2)
R2 = 1.0 - (SS_res / SS_tot)
R2_adj = 1.0 - ((1.0 - R2) * (n - 1) / (n - 4 - 1))
s2 = SS_res / (n - 5)
se_beta = np.sqrt(np.diagonal(s2 * XtX_inv))
t_beta = beta / se_beta

print(f"\nRegresyon Modeli (ΔRölatif Toplam Kavrama): R² = {R2:.4f}, Düzeltilmiş R² = {R2_adj:.4f}")
names = ['Sabit', 'Karış (cm)', 'Kulaç (cm)', 'VKİ', 'PHV Offset']
for name, b, se, t in zip(names, beta, se_beta, t_beta):
    print(f"  {name:<15}: B = {b:>8.4f}, SE = {se:>7.4f}, t = {t:>6.2f}")

# Also compare Drop-out vs Completers for Relative Total Grip at Baseline
with open('/Users/alik/Desktop/anti_projeler/articles_all/autism_2026_testtekrar/02_data/raw_data.csv', 'r', encoding='utf-8') as f:
    raw_rows = [r for r in csv.DictReader(f) if r['grup'] == 'Hareket Eğitimi']

cols = ['BOY', 'KİLO', 'KULAÇ', 'KARIŞ', 'OTURMA_BOYU', 'GORSEL_TOTAL', 'AKTIF_SIC', 'ESNEKLIK', 'SAG_EL', 'SOL_EL', 'CEVIKLIK']
comp_rt_base = []
drop_rt_base = []
for r in raw_rows:
    is_same = all(r[c+'_PRE'] == r[c+'_POST'] for c in cols)
    w = float(r['KİLO_PRE'])
    s = float(r['SAG_EL_PRE']) + float(r['SOL_EL_PRE'])
    val = s / w
    if is_same:
        drop_rt_base.append(val)
    else:
        comp_rt_base.append(val)

m_c, s_c = np.mean(comp_rt_base), np.std(comp_rt_base, ddof=1)
m_d, s_d = np.mean(drop_rt_base), np.std(drop_rt_base, ddof=1)
diff_b = m_c - m_d
t_b = diff_b / math.sqrt((s_c**2 / len(comp_rt_base)) + (s_d**2 / len(drop_rt_base)))
p_b = t_test_p_value(t_b, 55)
print(f"\nBaşlangıç Yıpranma Kontrolü (Rölatif Toplam): Tamamlayan = {m_c:.3f} ± {s_c:.3f} | Drop-out = {m_d:.3f} ± {s_d:.3f} | t = {t_b:.2f}, p = {p_b:.4f}")
