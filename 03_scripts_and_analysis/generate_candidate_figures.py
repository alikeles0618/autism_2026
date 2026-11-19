import os
os.environ['MPLCONFIGDIR'] = '/tmp/matplotlib'

import csv
import math
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
from matplotlib.patches import Patch

matplotlib.rcParams['font.family'] = 'sans-serif'
matplotlib.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Arial', 'Helvetica']
matplotlib.rcParams['axes.edgecolor'] = '#111111'
matplotlib.rcParams['axes.linewidth'] = 1.0

# Paths
csv_path = '/Users/alik/Desktop/anti_projeler/articles_all/autism_2026_testtekrar/02_data/cleaned_data.csv'
artifact_dir = '/Users/alik/.gemini/antigravity/brain/b070dc08-8fc0-46d0-9bc0-8e6ac510db31'
manuscript_dir = '/Users/alik/Desktop/anti_projeler/articles_all/autism_2026_testtekrar/04_manuscript'

with open(csv_path, 'r', encoding='utf-8') as f:
    rows = list(csv.DictReader(f))

n = len(rows)

# Data arrays
yas = np.array([float(r['YAS_2021']) for r in rows])
phv_pre = np.array([float(r['PHV_OFFSET_PRE']) for r in rows])
vki_pre = np.array([float(r['VKI_PRE']) for r in rows])
kulac_pre = np.array([float(r['KULAC_PRE']) for r in rows])
karis_pre = np.array([float(r['KARIS_PRE']) for r in rows])

reac_pre = np.array([float(r['GORSEL_TOTAL_PRE']) for r in rows])
reac_post = np.array([float(r['GORSEL_TOTAL_POST']) for r in rows])
d_reac = reac_post - reac_pre
pct_reac = ((reac_post - reac_pre) / reac_pre) * 100

sag_pre = np.array([float(r['SAG_EL_PRE']) for r in rows])
sag_post = np.array([float(r['SAG_EL_POST']) for r in rows])
sol_pre = np.array([float(r['SOL_EL_PRE']) for r in rows])
sol_post = np.array([float(r['SOL_EL_POST']) for r in rows])

tot_pre = sag_pre + sol_pre
tot_post = sag_post + sol_post
d_tot = tot_post - tot_pre
pct_tot = ((tot_post - tot_pre) / tot_pre) * 100

w_pre = np.array([float(r['KILO_PRE']) for r in rows])
w_post = np.array([float(r['KILO_POST']) for r in rows])

rel_pre = tot_pre / w_pre
rel_post = tot_post / w_post
d_rel = rel_post - rel_pre
pct_rel = ((rel_post - rel_pre) / rel_pre) * 100

asym_pre = np.array([float(r['ASYMMETRY_PRE']) for r in rows])
asym_post = np.array([float(r['ASYMMETRY_POST']) for r in rows])
d_asym = asym_post - asym_pre

power_pre = np.array([float(r['SAYERS_POWER_PRE']) for r in rows])
power_post = np.array([float(r['SAYERS_POWER_POST']) for r in rows])
d_power = power_post - power_pre
pct_power = ((power_post - power_pre) / power_pre) * 100

agil_pre = np.array([float(r['CEVIKLIK_PRE']) for r in rows])
agil_post = np.array([float(r['CEVIKLIK_POST']) for r in rows])
d_agil = agil_post - agil_pre
pct_agil = ((agil_post - agil_pre) / agil_pre) * 100

flex_pre = np.array([float(r['ESNEKLIK_PRE']) for r in rows])
flex_post = np.array([float(r['ESNEKLIK_POST']) for r in rows])
d_flex = flex_post - flex_pre

def save_fig(fig, base_name):
    p1 = os.path.join(manuscript_dir, base_name)
    p2 = os.path.join(artifact_dir, base_name)
    fig.savefig(p1, dpi=300, bbox_inches='tight', facecolor='white')
    fig.savefig(p2, dpi=300, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    print(f"Saved: {base_name}")

# ==========================================
# 1. RADAR / ÖRÜMCEK Z-SKOR PROFİLİ
# ==========================================
def generate_radar(lang='TR'):
    # Standardize to Pre-test (Mean=0, SD=1)
    # Higher should always mean better in Radar:
    # 1. Reaksiyon Hızı: -Z(Reac)
    # 2. Rölatif Kuvvet: +Z(Rel)
    # 3. Toplam Kuvvet: +Z(Tot)
    # 4. Sayers Zirve Güç: +Z(Power)
    # 5. Çeviklik Hızı: -Z(Agil)
    # 6. Esneklik: +Z(Flex)
    # 7. Bilateral Simetri: -Z(Asym)
    
    def z_trans(pre_arr, post_arr, invert=False):
        m, s = np.mean(pre_arr), np.std(pre_arr, ddof=1)
        z1 = (pre_arr - m) / s
        z2 = (post_arr - m) / s
        if invert:
            return -np.mean(z1), -np.mean(z2)
        return np.mean(z1), np.mean(z2)
    
    val_reac_pre, val_reac_post = z_trans(reac_pre, reac_post, invert=True) # speed
    val_rel_pre, val_rel_post = z_trans(rel_pre, rel_post)
    val_tot_pre, val_tot_post = z_trans(tot_pre, tot_post)
    val_pow_pre, val_pow_post = z_trans(power_pre, power_post)
    val_agil_pre, val_agil_post = z_trans(agil_pre, agil_post, invert=True) # speed
    val_flex_pre, val_flex_post = z_trans(flex_pre, flex_post)
    val_sym_pre, val_sym_post = z_trans(asym_pre, asym_post, invert=True) # symmetry
    
    pre_vals = [val_reac_pre, val_rel_pre, val_tot_pre, val_pow_pre, val_agil_pre, val_flex_pre, val_sym_pre]
    post_vals = [val_reac_post, val_rel_post, val_tot_post, val_pow_post, val_agil_post, val_flex_post, val_sym_post]
    
    if lang == 'TR':
        labels = [
            'Görsel Reaksiyon Hızı\n(İşlemleme Hızı)',
            'Rölatif Toplam Kuvvet\n(Kuvvet / Vücut Ağırlığı)',
            'Toplam Mutlak Kuvvet\n(Sağ + Sol jamar)',
            'Sayers Zirve Mekanik Güç\n(Alt Ekstremite Patlayıcı)',
            'Pro-Agility Çeviklik Hızı\n(5-10-5 Koordinasyon)',
            'Otur-Eriş Esneklik\n(Hamstring/Bel)',
            'Bilateral Simetri Düzeyi\n(Ekstremite Dengesi)'
        ]
        title = "Aday Grafik 1: Çok Boyutlu Nöromotor Z-Skoru Profili (n = 34)"
        pre_label = 'Ön Test (Temel Profil = 0)'
        post_label = '12 Hafta Son Test'
        fname = "Aday_1_Radar_Z_Skor_TR.png"
    else:
        labels = [
            'Visual Reaction Speed\n(Processing Speed)',
            'Relative Total Strength\n(Force / Body Mass)',
            'Total Absolute Strength\n(Right + Left Jamar)',
            'Sayers Peak Mechanical Power\n(Lower-Limb Power)',
            'Pro-Agility Shuttle Speed\n(5-10-5 Coordination)',
            'Sit-and-Reach Flexibility\n(Hamstring/Lower Back)',
            'Bilateral Functional Symmetry\n(Limb Balance)'
        ]
        title = "Candidate Figure 1: Multidimensional Neuromotor Z-Score Profile (n = 34)"
        pre_label = 'Pre-Test (Baseline = 0)'
        post_label = '12-Week Post-Test'
        fname = "Candidate_1_Radar_Z_Score_EN.png"

    num_vars = len(labels)
    angles = np.linspace(0, 2 * np.pi, num_vars, endpoint=False).tolist()
    
    # Complete circular polygon
    pre_vals += pre_vals[:1]
    post_vals += post_vals[:1]
    angles += angles[:1]
    
    fig, ax = plt.subplots(figsize=(8.5, 8.5), subplot_kw=dict(polar=True), dpi=300)
    
    # Draw reference circle at 0
    ax.plot(angles, [0]*len(angles), color='#7f8c8d', linestyle=':', linewidth=1.2, label=pre_label)
    
    # Plot Pre-test
    ax.plot(angles, pre_vals, color='#7f8c8d', linewidth=2.0, linestyle='--', marker='o', markersize=6)
    ax.fill(angles, pre_vals, color='#bdc3c7', alpha=0.25)
    
    # Plot Post-test
    ax.plot(angles, post_vals, color='#111111', linewidth=2.5, linestyle='-', marker='s', markersize=7, label=post_label)
    ax.fill(angles, post_vals, color='#2c3e50', alpha=0.35)
    
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    
    ax.set_thetagrids(np.degrees(angles[:-1]), labels, fontsize=10.5, fontweight='bold', color='#111111')
    ax.set_ylim(-0.35, 0.45)
    ax.set_yticks([-0.2, 0.0, 0.2, 0.4])
    ax.set_yticklabels(['-0.2 SD', '0 (Ön)', '+0.2 SD', '+0.4 SD'] if lang=='TR' else ['-0.2 SD', '0 (Pre)', '+0.2 SD', '+0.4 SD'],
                       fontsize=9, color='#555555')
    ax.grid(True, linestyle='--', alpha=0.5, color='#bdc3c7')
    
    ax.set_title(title, fontsize=13, fontweight='bold', pad=25, color='#111111')
    ax.legend(loc='lower right', bbox_to_anchor=(1.25, 0.0), frameon=True, fontsize=11)
    
    save_fig(fig, fname)

# ==========================================
# 2. BİLATERAL ASİMETRİ & KLİNİK EŞİK GRAFİĞİ
# ==========================================
def generate_asymmetry(lang='TR'):
    fig, axes = plt.subplots(1, 2, figsize=(12, 5.5), dpi=300)
    
    # Panel A: Scatter Sol vs Sağ
    ax1 = axes[0]
    
    max_val = max(np.max(sag_post), np.max(sol_post)) + 3.0
    
    # Identity line y = x
    line_x = np.linspace(0, max_val, 100)
    ax1.plot(line_x, line_x, color='#111111', linestyle='-', linewidth=1.5, label='Tam Simetri Hattı (y = x)' if lang=='TR' else 'Identity Line (y = x)')
    
    # ±10% Asymmetry Zone
    ax1.fill_between(line_x, line_x * 0.90, line_x * 1.10, color='#27ae60', alpha=0.15, 
                     label='Normatif Denge Alanı (±%10)' if lang=='TR' else 'Normative Balance Zone (±10%)')
    
    # ±15% Asymmetry Lines
    ax1.plot(line_x, line_x * 1.15, color='#e67e22', linestyle='--', linewidth=1.1)
    ax1.plot(line_x, line_x * 0.85, color='#e67e22', linestyle='--', linewidth=1.1, 
             label='Klinik Asimetri Sınırı (±%15)' if lang=='TR' else 'Clinical Asymmetry Cutoff (±15%)')
    
    # Points: Pre-test (gray) and Post-test (dark slate)
    for i in range(n):
        ax1.annotate("", xy=(sol_post[i], sag_post[i]), xytext=(sol_pre[i], sag_pre[i]),
                     arrowprops=dict(arrowstyle="->", color='#95a5a6', alpha=0.5, lw=0.9))
        
    ax1.scatter(sol_pre, sag_pre, color='#95a5a6', edgecolors='#555555', s=45, alpha=0.7, 
                label='Ön Test' if lang=='TR' else 'Pre-Test', zorder=4)
    ax1.scatter(sol_post, sag_post, color='#2c3e50', edgecolors='#111111', s=55, alpha=0.9, 
                label='Son Test' if lang=='TR' else 'Post-Test', zorder=5)
    
    ax1.set_xlim(0, max_val)
    ax1.set_ylim(0, max_val)
    ax1.set_xlabel('Sol El Kavrama Kuvveti (kg)' if lang=='TR' else 'Left Handgrip Strength (kg)', fontsize=11, fontweight='bold')
    ax1.set_ylabel('Sağ El Kavrama Kuvveti (kg)' if lang=='TR' else 'Right Handgrip Strength (kg)', fontsize=11, fontweight='bold')
    ax1.set_title('A: Sağ ve Sol El Kuvvet Eşleşmesi ve Klinik Eşikler' if lang=='TR' else 'A: Bilateral Grip Pairing & Clinical Thresholds', fontsize=11.5, fontweight='bold')
    ax1.legend(loc='upper left', frameon=True, fontsize=9.5)
    ax1.grid(True, linestyle='--', alpha=0.4, color='#bdc3c7')
    
    # Panel B: Asymmetry Index pre vs post slope
    ax2 = axes[1]
    x_pre = 1.0
    x_post = 2.0
    
    for i in range(n):
        ax2.plot([x_pre, x_post], [asym_pre[i], asym_post[i]], color='#95a5a6', alpha=0.4, lw=1.1)
        ax2.scatter([x_pre, x_post], [asym_pre[i], asym_post[i]], color='#7f8c8d', alpha=0.5, s=20)
        
    m_asym_pre, s_asym_pre = np.mean(asym_pre), np.std(asym_pre, ddof=1)
    m_asym_post, s_asym_post = np.mean(asym_post), np.std(asym_post, ddof=1)
    
    ax2.errorbar([x_pre, x_post], [m_asym_pre, m_asym_post], 
                 yerr=[1.96 * s_asym_pre / np.sqrt(n), 1.96 * s_asym_post / np.sqrt(n)],
                 color='#111111', lw=2.4, capsize=6, capthick=2.0, zorder=5)
    ax2.plot([x_pre, x_post], [m_asym_pre, m_asym_post], color='#111111', lw=2.6, zorder=5)
    ax2.scatter([x_pre, x_post], [m_asym_pre, m_asym_post], color='#111111', s=70, zorder=6)
    
    # Clinical cutoff line at 15%
    ax2.axhline(15.0, color='#e67e22', linestyle='--', linewidth=1.5, label='Klinik Asimetri Eşiği (%15)' if lang=='TR' else 'Clinical Threshold (15%)')
    
    ax2.set_xlim(0.6, 2.4)
    ax2.set_xticks([x_pre, x_post])
    ax2.set_xticklabels(['Ön Test\n(%16.8 ± 11.9)', 'Son Test\n(%17.0 ± 11.7)'] if lang=='TR' else ['Pre-Test\n(16.8 ± 11.9%)', 'Post-Test\n(17.0 ± 11.7%)'], fontsize=10.5, fontweight='bold')
    ax2.set_ylabel('Fonksiyonel Asimetri İndeksi (%)' if lang=='TR' else 'Functional Asymmetry Index (%)', fontsize=11, fontweight='bold')
    ax2.set_title('B: Bireysel Asimetri Değişimleri (p = 0.914, Stabil)' if lang=='TR' else 'B: Individual Asymmetry Changes (p = 0.914, Stable)', fontsize=11.5, fontweight='bold')
    ax2.legend(loc='upper right', frameon=True, fontsize=10)
    ax2.grid(True, linestyle='--', alpha=0.4, color='#bdc3c7')
    
    plt.tight_layout(pad=2.0)
    fname = "Aday_2_Bilateral_Asimetri_TR.png" if lang=='TR' else "Candidate_2_Bilateral_Asymmetry_EN.png"
    save_fig(fig, fname)

# ==========================================
# 3. DEĞİŞİMLER ARASI KORELASYON ISI HARİTASI
# ==========================================
def generate_heatmap(lang='TR'):
    # Delta variables and baseline anthropometrics
    if lang == 'TR':
        var_names = [
            'Δ Görsel Reaksiyon',
            'Δ Rölatif Kavrama',
            'Δ Toplam Kavrama',
            'Δ Sayers Zirve Güç',
            'Δ Pro-Agility',
            'Δ Otur-Eriş',
            'Başlangıç Karış',
            'Başlangıç Kulaç',
            'Başlangıç VKİ',
            'Mirwald PHV Offset'
        ]
        title = "Aday Grafik 3: Performans Değişimleri ve Morfolojik Korelasyon Isı Haritası"
        fname = "Aday_3_Korelasyon_Isi_Haritasi_TR.png"
    else:
        var_names = [
            'Δ Visual Reaction',
            'Δ Relative Grip',
            'Δ Total Grip',
            'Δ Sayers Peak Power',
            'Δ Pro-Agility',
            'Δ Sit-and-Reach',
            'Baseline Hand Span',
            'Baseline Arm Span',
            'Baseline BMI',
            'Mirwald PHV Offset'
        ]
        title = "Candidate Figure 3: Performance Adaptation & Morphological Correlation Heatmap"
        fname = "Candidate_3_Correlation_Heatmap_EN.png"
        
    data_matrix = np.column_stack([
        d_reac, d_rel, d_tot, d_power, d_agil, d_flex, karis_pre, kulac_pre, vki_pre, phv_pre
    ])
    
    num_vars = len(var_names)
    corr_matrix = np.corrcoef(data_matrix, rowvar=False)
    
    fig, ax = plt.subplots(figsize=(9.5, 8.5), dpi=300)
    
    # Mask upper triangle
    mask = np.triu(np.ones_like(corr_matrix, dtype=bool), k=1)
    plot_corr = np.ma.masked_array(corr_matrix, mask=mask)
    
    cmap = matplotlib.colormaps['coolwarm'].copy()
    cax = ax.imshow(plot_corr, cmap=cmap, vmin=-0.6, vmax=0.6)
    
    # Add text numbers
    for i in range(num_vars):
        for j in range(i + 1):
            val = corr_matrix[i, j]
            # p-value check for t = r * sqrt(n-2) / sqrt(1-r^2)
            df = n - 2
            t_val = val * np.sqrt(df) / np.sqrt(max(1e-9, 1.0 - val**2))
            is_sig = abs(t_val) > 2.037 # p < 0.05 for df = 32
            
            sig_mark = "*" if is_sig and i != j else ""
            txt_color = "white" if abs(val) > 0.35 else "#111111"
            ax.text(j, i, f"{val:+.2f}{sig_mark}", ha='center', va='center',
                    fontsize=9, fontweight='bold' if is_sig else 'normal', color=txt_color)
            
    ax.set_xticks(np.arange(num_vars))
    ax.set_yticks(np.arange(num_vars))
    ax.set_xticklabels(var_names, rotation=45, ha='right', fontsize=9.5, fontweight='bold')
    ax.set_yticklabels(var_names, fontsize=9.5, fontweight='bold')
    
    cbar = fig.colorbar(cax, ax=ax, fraction=0.046, pad=0.04)
    cbar.set_label('Pearson Korelasyon Katsayısı (r)' if lang=='TR' else 'Pearson Correlation Coefficient (r)',
                   fontsize=10.5, fontweight='bold')
    
    ax.set_title(title + "\n(* p < 0.05)", fontsize=11.5, fontweight='bold', pad=15)
    plt.tight_layout()
    save_fig(fig, fname)

# ==========================================
# 4. OLGUNLAŞMA VE YAŞ ORMAN GRAFİĞİ (FOREST PLOT)
# ==========================================
def generate_forest_plot(lang='TR'):
    # Split into 2 maturity subgroups by median PHV Offset (-3.5 years)
    # Early Pre-PHV: Offset < -3.5 years (Younger, ~7 yrs, n=17)
    # Late Pre-PHV: Offset >= -3.5 years (Older, ~11 yrs, n=17)
    
    med_phv = -3.5
    idx_early = np.where(phv_pre < med_phv)[0]
    idx_late = np.where(phv_pre >= med_phv)[0]
    
    metrics = [
        ('Görsel Reaksiyon Hızı' if lang=='TR' else 'Visual Reaction Speed', reac_pre, reac_post, True),
        ('Rölatif Toplam Kavrama' if lang=='TR' else 'Relative Total Grip', rel_pre, rel_post, False),
        ('Toplam Mutlak Kavrama' if lang=='TR' else 'Total Absolute Grip', tot_pre, tot_post, False),
        ('Sayers Zirve Mekanik Güç' if lang=='TR' else 'Sayers Peak Power', power_pre, power_post, False),
    ]
    
    # Calculate Cohen's d and 95% CI for Early, Late, and Whole Cohort
    def calc_d_ci(pre_arr, post_arr, idxs=None, invert=False):
        if idxs is not None:
            pre_arr = pre_arr[idxs]
            post_arr = post_arr[idxs]
        n_sub = len(pre_arr)
        diff = (pre_arr - post_arr) if invert else (post_arr - pre_arr)
        s_av = np.sqrt((np.var(pre_arr, ddof=1) + np.var(post_arr, ddof=1)) / 2.0)
        d = np.mean(diff) / s_av
        # Standard error of Cohen's d
        se_d = np.sqrt((2.0 / n_sub) + (d**2 / (2.0 * n_sub)))
        return d, d - 1.96 * se_d, d + 1.96 * se_d
    
    fig, ax = plt.subplots(figsize=(10, 6.0), dpi=300)
    
    y_pos = 0
    y_ticks = []
    y_labels = []
    
    colors = ['#2c3e50', '#27ae60', '#e67e22'] # Whole, Early, Late
    subgroup_labels = [
        'Tüm Kohort (n = 34)' if lang=='TR' else 'Whole Cohort (n = 34)',
        'Erken Pre-PHV (Maturity Offset < -3.5 yıl, n = 17)' if lang=='TR' else 'Early Pre-PHV (Offset < -3.5 yrs, n = 17)',
        'Geç Pre-PHV (Maturity Offset ≥ -3.5 yıl, n = 17)' if lang=='TR' else 'Late Pre-PHV (Offset ≥ -3.5 yrs, n = 17)'
    ]
    
    for m_label, pre_a, post_a, inv in metrics:
        # Group category label
        y_labels.append(m_label)
        y_ticks.append(y_pos + 1.0)
        
        # Subgroups: Whole, Early, Late
        d_whole, ci_l_w, ci_u_w = calc_d_ci(pre_a, post_a, None, inv)
        d_early, ci_l_e, ci_u_e = calc_d_ci(pre_a, post_a, idx_early, inv)
        d_late, ci_l_l, ci_u_l = calc_d_ci(pre_a, post_a, idx_late, inv)
        
        vals = [(d_whole, ci_l_w, ci_u_w), (d_early, ci_l_e, ci_u_e), (d_late, ci_l_l, ci_u_l)]
        
        for idx_sub, (d_val, ci_l, ci_u) in enumerate(vals):
            cur_y = y_pos + (2 - idx_sub) * 0.7
            ax.errorbar(d_val, cur_y, xerr=[[d_val - ci_l], [ci_u - d_val]],
                        fmt='s' if idx_sub==0 else 'o', color=colors[idx_sub],
                        ecolor=colors[idx_sub], elinewidth=1.8, capsize=4, capthick=1.5,
                        markersize=7 if idx_sub==0 else 6, zorder=4)
            # Text annotation
            ax.text(ci_u + 0.04, cur_y, f"d = {d_val:+.2f} [{ci_l:+.2f}, {ci_u:+.2f}]",
                    va='center', fontsize=8.5, color=colors[idx_sub], fontweight='bold' if idx_sub==0 else 'normal')
            
        y_pos += 3.2
        
    ax.axvline(0.0, color='#111111', linestyle='-', linewidth=1.2, zorder=2)
    ax.axvline(0.2, color='#95a5a6', linestyle='--', linewidth=0.9, zorder=1)
    ax.axvline(0.5, color='#95a5a6', linestyle='--', linewidth=0.9, zorder=1)
    
    ax.text(0.2, y_pos - 0.2, 'Küçük Etki (0.2)' if lang=='TR' else 'Small Effect (0.2)', ha='center', fontsize=8.5, color='#7f8c8d')
    ax.text(0.5, y_pos - 0.2, 'Orta Etki (0.5)' if lang=='TR' else 'Medium Effect (0.5)', ha='center', fontsize=8.5, color='#7f8c8d')
    
    ax.set_yticks(y_ticks)
    ax.set_yticklabels(y_labels, fontsize=11, fontweight='bold', color='#111111')
    ax.set_xlabel("Standardize Etki Büyüklüğü (Cohen's d ve %95 Güven Aralığı)" if lang=='TR' else "Standardized Effect Size (Cohen's d & 95% Confidence Interval)",
                  fontsize=10.5, fontweight='bold', labelpad=8)
    
    ax.set_title("Aday Grafik 4: Biyolojik Olgunlaşma Alt Gruplarına Göre Orman Grafiği" if lang=='TR' else "Candidate Figure 4: Subgroup Forest Plot Across Somatic Maturity (PHV Offset)",
                 fontsize=12, fontweight='bold', pad=14)
    
    # Custom legend
    legend_elements = [
        Patch(facecolor=colors[0], label=subgroup_labels[0]),
        Patch(facecolor=colors[1], label=subgroup_labels[1]),
        Patch(facecolor=colors[2], label=subgroup_labels[2])
    ]
    ax.legend(handles=legend_elements, loc='lower right', frameon=True, fontsize=9.5)
    
    ax.set_xlim(-0.4, 1.1)
    ax.grid(True, linestyle='--', alpha=0.35, color='#bdc3c7', axis='x')
    
    plt.tight_layout()
    fname = "Aday_4_Olgunlasma_Forest_Plot_TR.png" if lang=='TR' else "Candidate_4_Maturity_Forest_Plot_EN.png"
    save_fig(fig, fname)

# ==========================================
# 5. KÜMÜLATİF YANIT DAĞILIMI (eCDF)
# ==========================================
def generate_ecdf(lang='TR'):
    # eCDF for % changes
    # Reaction time: % improvement = -pct_reac (so positive is faster!)
    ecdf_reac = -pct_reac
    ecdf_rel = pct_rel
    ecdf_tot = pct_tot
    ecdf_pow = pct_power
    
    fig, ax = plt.subplots(figsize=(9.5, 5.5), dpi=300)
    
    curves = [
        (ecdf_reac, 'Görsel Reaksiyon Hızlanması' if lang=='TR' else 'Visual Reaction Acceleration', '#2c3e50', '-'),
        (ecdf_rel, 'Rölatif Toplam Kavrama Artışı' if lang=='TR' else 'Relative Total Grip Gain', '#27ae60', '--'),
        (ecdf_tot, 'Toplam Mutlak Kavrama Artışı' if lang=='TR' else 'Total Absolute Grip Gain', '#2980b9', '-.'),
        (ecdf_pow, 'Sayers Zirve Güç Artışı' if lang=='TR' else 'Sayers Peak Power Gain', '#e67e22', ':')
    ]
    
    for arr, label, col, ls in curves:
        sorted_x = np.sort(arr)
        y = np.arange(1, len(sorted_x) + 1) / len(sorted_x) * 100.0
        ax.step(sorted_x, y, where='post', color=col, linestyle=ls, linewidth=2.4, label=label)
        
    ax.axvline(0.0, color='#111111', linestyle='-', linewidth=1.2)
    ax.axvline(10.0, color='#7f8c8d', linestyle='--', linewidth=1.0)
    ax.axvline(20.0, color='#7f8c8d', linestyle='--', linewidth=1.0)
    
    ax.text(10.2, 8, '> %10 Gelişim' if lang=='TR' else '> 10% Gain', fontsize=9, color='#555555', fontweight='bold')
    ax.text(20.2, 8, '> %20 Gelişim' if lang=='TR' else '> 20% Gain', fontsize=9, color='#555555', fontweight='bold')
    
    ax.set_xlim(-25, 65)
    ax.set_ylim(0, 105)
    ax.set_xlabel('Yüzdesel Performans Değişimi (% Δ)' if lang=='TR' else 'Percentage Performance Change (% Δ)', fontsize=11, fontweight='bold', labelpad=8)
    ax.set_ylabel('Kümülatif Kohort Yüzdesi (%)' if lang=='TR' else 'Cumulative Cohort Proportion (%)', fontsize=11, fontweight='bold', labelpad=8)
    ax.set_title("Aday Grafik 5: Kümülatif Yanıt Dağılımı (eCDF Eğrileri, n = 34)" if lang=='TR' else "Candidate Figure 5: Empirical Cumulative Distribution Function (eCDF, n = 34)",
                 fontsize=12, fontweight='bold', pad=12)
    
    ax.legend(loc='lower right', frameon=True, fontsize=10)
    ax.grid(True, linestyle='--', alpha=0.35, color='#bdc3c7')
    
    plt.tight_layout()
    fname = "Aday_5_Kumulatif_eCDF_TR.png" if lang=='TR' else "Candidate_5_Cumulative_eCDF_EN.png"
    save_fig(fig, fname)

if __name__ == '__main__':
    generate_radar('TR')
    generate_radar('EN')
    generate_asymmetry('TR')
    generate_asymmetry('EN')
    generate_heatmap('TR')
    generate_heatmap('EN')
    generate_forest_plot('TR')
    generate_forest_plot('EN')
    generate_ecdf('TR')
    generate_ecdf('EN')
    print("All 5 Candidate Figures Generated Successfully!")
