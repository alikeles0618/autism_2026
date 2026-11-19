import os
os.environ['MPLCONFIGDIR'] = '/tmp/matplotlib'

import csv
import math
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
from matplotlib.lines import Line2D
from PIL import Image
from PIL.PngImagePlugin import PngInfo

# =========================================================================
# Q1 EDITORIAL TYPOGRAPHY & COLORBLIND-SAFE OKABE-ITO PALETTE
# =========================================================================
matplotlib.rcParams['font.family'] = 'sans-serif'
matplotlib.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Arial', 'Helvetica']
matplotlib.rcParams['axes.edgecolor'] = '#334155'
matplotlib.rcParams['axes.linewidth'] = 1.1
matplotlib.rcParams['xtick.color'] = '#1E293B'
matplotlib.rcParams['ytick.color'] = '#1E293B'
matplotlib.rcParams['xtick.direction'] = 'out'
matplotlib.rcParams['ytick.direction'] = 'out'
matplotlib.rcParams['xtick.major.size'] = 4.5
matplotlib.rcParams['ytick.major.size'] = 4.5
matplotlib.rcParams['text.color'] = '#0F172A'

C_PRE = '#64748B'       # Slate gray
C_PRE_LIGHT = '#CBD5E1' # Soft slate
C_POS = '#0072B2'       # Okabe-Ito deep blue
C_POS_LIGHT = '#93C5FD' # Soft blue
C_NEG = '#D55E00'       # Okabe-Ito vermilion orange
C_INK = '#0F172A'       # Crisp dark charcoal
C_GRID = '#CBD5E1'      # Grid line gray
C_STAB = '#94A3B8'      # Stable gray

T_CRIT_33 = 2.034515    # Two-tailed Student-t critical value (df = 33, alpha = 0.05)

CSV_PATH = '/Users/alik/Desktop/anti_projeler/articles_all/autism_2026_testtekrar/02_data/cleaned_data.csv'
MANUSCRIPT_DIR = '/Users/alik/Desktop/anti_projeler/articles_all/autism_2026_testtekrar/04_manuscript'


def load_data():
    with open(CSV_PATH, 'r', encoding='utf-8') as f:
        rows = list(csv.DictReader(f))
    n = len(rows)

    def arr(key):
        return np.array([float(r[key]) for r in rows])

    yas = arr('YAS_2021')
    w_pre, w_post = arr('KILO_PRE'), arr('KILO_POST')
    sag_pre, sag_post = arr('SAG_EL_PRE'), arr('SAG_EL_POST')
    sol_pre, sol_post = arr('SOL_EL_PRE'), arr('SOL_EL_POST')
    mean_grip_pre = (sag_pre + sol_pre) / 2.0
    mean_grip_post = (sag_post + sol_post) / 2.0

    return {
        'n': n,
        'is_child': yas <= 9,
        'is_adol': yas >= 10,
        'karis': arr('KARIS_PRE'),
        'kulac': arr('KULAC_PRE'),
        'reac': (arr('GORSEL_TOTAL_PRE'), arr('GORSEL_TOTAL_POST')),
        'rel': (mean_grip_pre / w_pre, mean_grip_post / w_post),
        'power': (arr('SAYERS_POWER_PRE'), arr('SAYERS_POWER_POST')),
        'agil': (arr('CEVIKLIK_PRE'), arr('CEVIKLIK_POST')),
        'sag': (sag_pre, sag_post),
        'mean_grip': (mean_grip_pre, mean_grip_post),
    }


def style_ax(ax, grid_axis='y'):
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    if grid_axis in ('y', 'both'):
        ax.yaxis.grid(True, linestyle='--', alpha=0.45, color=C_GRID, linewidth=0.8)
    if grid_axis in ('x', 'both'):
        ax.xaxis.grid(True, linestyle='--', alpha=0.45, color=C_GRID, linewidth=0.8)
    ax.set_axisbelow(True)


def save_clean(fig, base_name):
    os.makedirs(MANUSCRIPT_DIR, exist_ok=True)
    path = os.path.join(MANUSCRIPT_DIR, base_name)
    fig.savefig(path, dpi=300, bbox_inches='tight', facecolor='white', pad_inches=0.16)
    plt.close(fig)
    im = Image.open(path)
    im.save(path, 'PNG', pnginfo=PngInfo(), dpi=(300, 300))
    print(f'Saved cleanly (300 DPI, metadata stripped): {base_name}')


def kde_1d(vals, grid, bw_scale=1.05):
    sd = np.std(vals, ddof=1)
    iqr = np.percentile(vals, 75) - np.percentile(vals, 25)
    sigma = min(sd, iqr / 1.34) if iqr > 0 else sd
    bw = max(1e-6, bw_scale * 0.9 * max(sigma, sd * 0.35) * (len(vals) ** -0.2))
    diffs = (grid[:, None] - vals[None, :]) / bw
    return np.mean(np.exp(-0.5 * diffs ** 2), axis=1) / (bw * math.sqrt(2 * math.pi))


def paired_stats(pre, post):
    n = len(pre)
    m_pre, m_post = np.mean(pre), np.mean(post)
    ci_pre = T_CRIT_33 * np.std(pre, ddof=1) / math.sqrt(n)
    ci_post = T_CRIT_33 * np.std(post, ddof=1) / math.sqrt(n)
    return m_pre, ci_pre, m_post, ci_post


def swc_classify(pre, post, lower_is_better):
    swc = 0.20 * np.std(pre, ddof=1)
    imp = (pre - post) if lower_is_better else (post - pre)
    status = np.zeros(len(pre), dtype=int)
    status[imp > swc] = 1
    status[imp < -swc] = -1
    return swc, imp, status


# =========================================================================
# FIGURE 2 (OPTION 2A): PAIRED RAINCLOUD + INDIVIDUAL SWC SPAGHETTI PLOT
# =========================================================================
def generate_figure_2(d, lang='TR'):
    fig, axes = plt.subplots(2, 2, figsize=(12.4, 9.6), dpi=300)
    rng = np.random.default_rng(42)

    if lang == 'TR':
        metrics = [
            ('A', 'Görsel Reaksiyon Süresi', 'Süre (sn)', d['reac'][0], d['reac'][1],
             True, 'p = 0.011* · Δ = -13.0% · d = -0.35'),
            ('B', 'Rölatif Kavrama Kuvveti', 'Rölatif Kuvvet (kg/kg)', d['rel'][0], d['rel'][1],
             False, 'p = 0.033* · Δ = +9.0% · d = +0.15'),
            ('C', 'Sayers Zirve Mekanik Güç', 'Güç (Watt)', d['power'][0], d['power'][1],
             False, 'p = 0.076 · Δ = +11.6% · d = +0.05'),
            ('D', 'Pro-Agility 5-10-5 Çeviklik', 'Süre (sn)', d['agil'][0], d['agil'][1],
             True, 'p = 0.413 (ns) · Δ = -3.7% · d = -0.13'),
        ]
        xlabels = ['Ön Test (Hafta 1)', 'Son Test (Hafta 14)']
        leg_labels = [
            'Olumlu SWC Yanıtı',
            'Stabil / Değişmeyen (±1×SWC)',
            'Olumsuz SWC Yanıtı',
            'Grup Ortalaması ± %95 GA',
        ]
        fname = 'Sekil_2_Performans_Degisimleri_TR.png'
    else:
        metrics = [
            ('A', 'Visual Reaction Time', 'Time (s)', d['reac'][0], d['reac'][1],
             True, 'p = 0.011* · Δ = -13.0% · d = -0.35'),
            ('B', 'Relative Handgrip Strength', 'Relative Strength (kg/kg)', d['rel'][0], d['rel'][1],
             False, 'p = 0.033* · Δ = +9.0% · d = +0.15'),
            ('C', 'Sayers Peak Mechanical Power', 'Power (Watts)', d['power'][0], d['power'][1],
             False, 'p = 0.076 · Δ = +11.6% · d = +0.05'),
            ('D', 'Pro-Agility Shuttle (5-10-5)', 'Time (s)', d['agil'][0], d['agil'][1],
             True, 'p = 0.413 (ns) · Δ = -3.7% · d = -0.13'),
        ]
        xlabels = ['Pre-Test (Week 1)', 'Post-Test (Week 14)']
        leg_labels = [
            'Positive SWC Responder',
            'Stable / Trivial (±1×SWC)',
            'Negative SWC Responder',
            'Cohort Mean ± 95% CI',
        ]
        fname = 'Figure_2_Performance_Adaptations_EN.png'

    for ax, (badge, title, ylabel, pre, post, low_better, stat_txt) in zip(axes.flat, metrics):
        style_ax(ax, 'y')
        m_pre, ci_pre, m_post, ci_post = paired_stats(pre, post)
        _, _, status = swc_classify(pre, post, low_better)

        if 'Sayers' in title:
            ax.axhline(0.0, color='#64748B', linestyle=':', linewidth=1.1, zorder=1)

        lo_pre, hi_pre = np.min(pre), np.max(pre)
        grid_pre = np.linspace(lo_pre, hi_pre, 140)
        dens_pre = kde_1d(pre, grid_pre)
        dens_pre = dens_pre / np.max(dens_pre) * 0.30
        ax.fill_betweenx(grid_pre, 0.96 - dens_pre, 0.96, color=C_PRE_LIGHT, alpha=0.55,
                         edgecolor=C_PRE, linewidth=1.0, zorder=2)

        lo_post, hi_post = np.min(post), np.max(post)
        grid_post = np.linspace(lo_post, hi_post, 140)
        dens_post = kde_1d(post, grid_post)
        dens_post = dens_post / np.max(dens_post) * 0.30
        ax.fill_betweenx(grid_post, 2.04, 2.04 + dens_post, color=C_POS_LIGHT, alpha=0.45,
                         edgecolor=C_POS, linewidth=1.0, zorder=2)

        x_left = 1.11 + rng.uniform(-0.02, 0.02, size=len(pre))
        x_right = 1.89 + rng.uniform(-0.02, 0.02, size=len(post))
        for i in range(len(pre)):
            st = status[i]
            if st == 1:
                col, lw, al, zo = C_POS, 1.35, 0.72, 3
            elif st == -1:
                col, lw, al, zo = C_NEG, 1.35, 0.75, 3
            else:
                col, lw, al, zo = C_STAB, 0.85, 0.42, 2
            ax.plot([x_left[i], x_right[i]], [pre[i], post[i]], color=col, linewidth=lw, alpha=al, zorder=zo)
            ax.scatter([x_left[i], x_right[i]], [pre[i], post[i]], color=col, s=16, alpha=min(0.9, al + 0.15),
                       edgecolors='none', zorder=zo + 0.2)

        ax.plot([1.00, 2.00], [m_pre, m_post], color=C_INK, linewidth=2.6, zorder=6)
        ax.errorbar([1.00, 2.00], [m_pre, m_post], yerr=[ci_pre, ci_post], fmt='none',
                    ecolor=C_INK, elinewidth=2.2, capsize=5, capthick=1.8, zorder=7)
        ax.scatter([1.00], [m_pre], color=C_PRE, edgecolor='white', linewidth=1.3, s=85, zorder=8)
        ax.scatter([2.00], [m_post], color=C_POS, edgecolor='white', linewidth=1.3, s=85, zorder=8)

        all_min = min(np.min(pre), np.min(post))
        all_max = max(np.max(pre), np.max(post))
        span = all_max - all_min
        ax.set_ylim(all_min - 0.08 * span, all_max + 0.24 * span)
        ax.set_xlim(0.55, 2.45)

        brk_y = all_max + 0.09 * span
        brk_h = 0.03 * span
        ax.plot([1.00, 1.00, 2.00, 2.00], [brk_y - brk_h, brk_y, brk_y, brk_y - brk_h],
                color=C_INK, linewidth=1.2, zorder=5)
        p_col = C_POS if '*' in stat_txt else '#475569'
        ax.text(1.5, brk_y + 0.02 * span, stat_txt, ha='center', va='bottom',
                fontsize=9.5, fontweight='bold', color=p_col)

        ax.set_xticks([1.0, 2.0])
        ax.set_xticklabels(xlabels, fontsize=10, fontweight='bold')
        ax.set_ylabel(ylabel, fontsize=10, fontweight='bold')
        ax.set_title(f'{badge}   {title}', loc='left', fontsize=11.5, fontweight='bold', pad=10, color=C_INK)

    legend_items = [
        Line2D([0], [0], color=C_POS, lw=2.0, label=leg_labels[0]),
        Line2D([0], [0], color=C_STAB, lw=1.6, label=leg_labels[1]),
        Line2D([0], [0], color=C_NEG, lw=2.0, label=leg_labels[2]),
        Line2D([0], [0], color=C_INK, lw=2.6, marker='o', markerfacecolor=C_POS,
               markeredgecolor='white', markersize=7, label=leg_labels[3]),
    ]
    fig.legend(handles=legend_items, loc='lower center', bbox_to_anchor=(0.5, -0.02),
               ncol=4, frameon=True, facecolor='#F8FAFC', edgecolor='#CBD5E1', fontsize=9.5)
    plt.tight_layout(rect=[0, 0.04, 1, 1], pad=2.0)
    save_clean(fig, fname)


# =========================================================================
# FIGURE 3 (OPTION 3B): DIVERGING CENTERED SWC RESPONDER CHART
# =========================================================================
def generate_figure_3(lang='TR'):
    fig, ax = plt.subplots(figsize=(11.8, 5.1), dpi=300)
    style_ax(ax, 'x')
    ax.spines['left'].set_visible(False)
    ax.tick_params(axis='y', length=0)

    if lang == 'TR':
        rows_spec = [
            ('Sayers Zirve Mekanik Güç\n(SWC = 255.2 W)', 6, 28, 0),
            ('Pro-Agility Çeviklik Süresi\n(SWC = 0.88 sn)', 10, 17, 7),
            ('Rölatif Kavrama Kuvveti\n(SWC = 0.03 kg/kg)', 10, 23, 1),
            ('Görsel Reaksiyon Süresi\n(SWC = 2.91 sn)', 14, 14, 6),
        ]
        xlabel = '◄── Olumsuz Yanıt (%)       |       Stabil Bölge (±1×SWC)       |       Olumlu Yanıt (%) ──►'
        title = 'En Küçük Anlamlı Değişim (SWC = 0.20 × SS_ön) Bireysel Yanıt Dağılımı (N = 34)'
        center_lbl = 'Merkez'
        leg_labels = [
            'Olumsuz Yanıt (< -1 × SWC)',
            'Stabil / Değişmeyen (±1 × SWC)',
            'Olumlu Yanıt (> +1 × SWC)',
        ]
        fmt_pct = lambda p: f'%{p:.1f}'
        fmt_n_pct = lambda n_c, p: f'{n_c} (%{p:.1f})'
        fname = 'Sekil_3_SWC_Yanitlayici_Analizi_TR.png'
    else:
        rows_spec = [
            ('Sayers Peak Mechanical Power\n(SWC = 255.2 W)', 6, 28, 0),
            ('Pro-Agility Shuttle Time\n(SWC = 0.88 s)', 10, 17, 7),
            ('Relative Handgrip Strength\n(SWC = 0.03 kg/kg)', 10, 23, 1),
            ('Visual Reaction Time\n(SWC = 2.91 s)', 14, 14, 6),
        ]
        xlabel = '◄── Negative Response (%)       |       Stable Zone (±1×SWC)       |       Positive Response (%) ──►'
        title = 'Smallest Worthwhile Change (SWC = 0.20 × SD_pre) Individual Responder Distribution (N = 34)'
        center_lbl = 'Center'
        leg_labels = [
            'Negative Responder (< -1 × SWC)',
            'Stable / Trivial (±1 × SWC)',
            'Positive Responder (> +1 × SWC)',
        ]
        fmt_pct = lambda p: f'{p:.1f}%'
        fmt_n_pct = lambda n_c, p: f'{n_c} ({p:.1f}%)'
        fname = 'Figure_3_SWC_Responder_Analysis_EN.png'

    y_pos = np.arange(len(rows_spec))
    bar_h = 0.54

    for idx, (_, n_pos, n_stb, n_neg) in enumerate(rows_spec):
        p_pos = 100.0 * n_pos / 34.0
        p_stb = 100.0 * n_stb / 34.0
        p_neg = 100.0 * n_neg / 34.0

        left_stb = -p_stb / 2.0
        right_stb = p_stb / 2.0

        ax.barh(idx, p_stb, left=left_stb, height=bar_h, color='#CBD5E1',
                edgecolor='white', linewidth=1.2, zorder=3)
        ax.text(0.0, idx, fmt_n_pct(n_stb, p_stb), va='center', ha='center',
                fontsize=9.5, fontweight='bold', color='#1E293B', zorder=5)

        ax.barh(idx, p_pos, left=right_stb, height=bar_h, color=C_POS,
                edgecolor='white', linewidth=1.2, zorder=3)
        if p_pos >= 15.0:
            ax.text(right_stb + p_pos / 2.0, idx, fmt_n_pct(n_pos, p_pos),
                    va='center', ha='center', fontsize=9.5, fontweight='bold', color='white', zorder=5)
        else:
            ax.text(right_stb + p_pos + 1.5, idx, fmt_n_pct(n_pos, p_pos),
                    va='center', ha='left', fontsize=9.5, fontweight='bold', color=C_POS, zorder=5)

        if n_neg > 0:
            ax.barh(idx, p_neg, left=left_stb - p_neg, height=bar_h, color=C_NEG,
                    edgecolor='white', linewidth=1.2, zorder=3)
            if p_neg >= 13.0:
                ax.text(left_stb - p_neg / 2.0, idx, fmt_n_pct(n_neg, p_neg),
                        va='center', ha='center', fontsize=9.5, fontweight='bold', color='white', zorder=5)
            else:
                ax.annotate(fmt_n_pct(n_neg, p_neg),
                            xy=(left_stb - p_neg * 0.5, idx),
                            xytext=(left_stb - p_neg - 4.5, idx),
                            ha='right', va='center', fontsize=9.5, fontweight='bold', color=C_NEG,
                            arrowprops=dict(arrowstyle='-', color=C_NEG, lw=1.2), zorder=6)
        else:
            ax.text(left_stb - 1.6, idx, fmt_n_pct(0, 0.0), va='center', ha='right',
                    fontsize=9.0, fontweight='bold', color='#64748B', zorder=5)

    ax.axvline(0.0, color='#475569', linestyle=':', linewidth=1.0, zorder=4)
    ax.set_yticks(y_pos)
    ax.set_yticklabels([r[0] for r in rows_spec], fontsize=10, fontweight='bold', color=C_INK)
    ax.set_xlim(-64, 72)
    xticks = [-60, -40, -20, 0, 20, 40, 60]
    ax.set_xticks(xticks)
    ax.set_xticklabels([fmt_pct(abs(v)) if v != 0 else center_lbl for v in xticks], fontsize=9.5, fontweight='bold')
    ax.set_xlabel(xlabel, fontsize=10, fontweight='bold', labelpad=10)
    ax.set_title(title, fontsize=11.8, fontweight='bold', pad=14, color=C_INK)

    legend_items = [
        Patch(facecolor=C_NEG, label=leg_labels[0]),
        Patch(facecolor='#CBD5E1', label=leg_labels[1]),
        Patch(facecolor=C_POS, label=leg_labels[2]),
    ]
    ax.legend(handles=legend_items, loc='lower center', bbox_to_anchor=(0.5, -0.30),
              ncol=3, frameon=True, facecolor='#F8FAFC', edgecolor='#CBD5E1', fontsize=9.5)
    plt.tight_layout()
    save_clean(fig, fname)


# =========================================================================
# FIGURE 4 (OPTION 4A): SUBGROUP-ENCODED REGRESSION + MARGINAL DENSITIES
# =========================================================================
def generate_figure_4(d, lang='TR'):
    fig = plt.figure(figsize=(13.2, 5.9), dpi=300)
    outer = fig.add_gridspec(1, 2, wspace=0.28)
    rng = np.random.default_rng(2026)

    karis = d['karis']
    kulac = d['kulac']
    d_sag = d['sag'][1] - d['sag'][0]
    d_mean = d['mean_grip'][1] - d['mean_grip'][0]
    is_c, is_a = d['is_child'], d['is_adol']

    if lang == 'TR':
        specs = [
            ('A', 'El Karış Uzunluğu ve Kuvvet Gelişimi',
             'Başlangıç El Karış Uzunluğu (cm)', 'Δ Sağ/Dominant El Kavrama Kuvveti (kg)',
             karis, d_sag, rng.uniform(-0.14, 0.14, size=len(karis)), 0.4680, 0.2190, 'p = 0.005**'),
            ('B', 'Kulaç Uzunluğu ve Kuvvet Gelişimi',
             'Başlangıç Kulaç Uzunluğu (cm)', 'Δ Ortalama El Kavrama Kuvveti (kg)',
             kulac, d_mean, rng.uniform(-0.45, 0.45, size=len(kulac)), 0.3658, 0.1338, 'p = 0.033*'),
        ]
        leg_labels = [
            'Çocukluk Alt Grubu (5–9 yaş, n = 17)',
            'Ergenlik Alt Grubu (10–15 yaş, n = 17)',
            'OLS Doğrusal Eğim ve %95 GA',
            'Sıfır Değişim Çizgisi (Δ = 0)',
        ]
        fname = 'Sekil_4_Regresyon_Iliskileri_TR.png'
    else:
        specs = [
            ('A', 'Hand Span and Dominant Force Gain',
             'Baseline Hand Span (cm)', 'Δ Dominant Handgrip Strength (kg)',
             karis, d_sag, rng.uniform(-0.14, 0.14, size=len(karis)), 0.4680, 0.2190, 'p = 0.005**'),
            ('B', 'Arm Span and Mean Bilateral Force Gain',
             'Baseline Arm Span (cm)', 'Δ Mean Handgrip Strength (kg)',
             kulac, d_mean, rng.uniform(-0.45, 0.45, size=len(kulac)), 0.3658, 0.1338, 'p = 0.033*'),
        ]
        leg_labels = [
            'Childhood Subgroup (5–9 yrs, n = 17)',
            'Adolescence Subgroup (10–15 yrs, n = 17)',
            'OLS Linear Fit & 95% CI',
            'Zero-Change Line (Δ = 0)',
        ]
        fname = 'Figure_4_Regression_Relationships_EN.png'

    for idx, (badge, title, xl, yl, xr, yr, jit, rv, r2v, ps) in enumerate(specs):
        sub = outer[idx].subgridspec(2, 2, width_ratios=[4.2, 1.0], height_ratios=[1.0, 4.2],
                                     hspace=0.06, wspace=0.06)
        ax_sc = fig.add_subplot(sub[1, 0])
        ax_top = fig.add_subplot(sub[0, 0], sharex=ax_sc)
        ax_right = fig.add_subplot(sub[1, 1], sharey=ax_sc)

        style_ax(ax_sc, 'both')
        n = len(xr)
        slope, intercept = np.polyfit(xr, yr, 1)
        x_line = np.linspace(np.min(xr) - 0.04 * (np.max(xr) - np.min(xr)),
                             np.max(xr) + 0.04 * (np.max(xr) - np.min(xr)), 140)
        y_line = slope * x_line + intercept
        y_pred = slope * xr + intercept
        s_err = math.sqrt(np.sum((yr - y_pred) ** 2) / (n - 2))
        t_32 = 2.036933
        ci = t_32 * s_err * np.sqrt(1.0 / n + ((x_line - np.mean(xr)) ** 2) / np.sum((xr - np.mean(xr)) ** 2))

        ax_sc.axhline(0.0, color='#475569', linestyle='--', linewidth=1.1, zorder=1)
        ax_sc.fill_between(x_line, y_line - ci, y_line + ci, color='#94A3B8', alpha=0.22, zorder=2)
        ax_sc.plot(x_line, y_line, color=C_INK, linewidth=2.2, zorder=3)

        xp = xr + jit
        ax_sc.scatter(xp[is_c], yr[is_c], color=C_POS, marker='o', s=58, alpha=0.88,
                      edgecolors='white', linewidths=0.9, zorder=4)
        ax_sc.scatter(xp[is_a], yr[is_a], color=C_NEG, marker='s', s=54, alpha=0.88,
                      edgecolors='white', linewidths=0.9, zorder=4)

        y_min_d, y_max_d = np.min(yr), np.max(yr)
        y_span = y_max_d - y_min_d
        ax_sc.set_ylim(y_min_d - 0.12 * y_span, y_max_d + 0.22 * y_span)

        ax_sc.text(0.04, 0.96, f'r = {rv:+.2f}, r² = {r2v:.3f}\n{ps}',
                   transform=ax_sc.transAxes, ha='left', va='top', fontsize=9.5, fontweight='bold', color=C_INK,
                   bbox=dict(boxstyle='round,pad=0.3', facecolor='#F8FAFC', edgecolor='#CBD5E1', lw=0.85))
        ax_sc.text(x_line[-1], -0.025 * y_span, 'Δ = 0', ha='right', va='top',
                   fontsize=8.2, color='#475569', style='italic',
                   bbox=dict(boxstyle='square,pad=0.1', facecolor='white', edgecolor='none', alpha=0.8))

        ax_sc.set_xlabel(xl, fontsize=10, fontweight='bold', labelpad=7)
        ax_sc.set_ylabel(yl, fontsize=10, fontweight='bold', labelpad=7)
        ax_top.set_title(f'{badge}   {title}', loc='left', fontsize=11.3, fontweight='bold', pad=8, color=C_INK)

        ax_top.axis('off')
        xg = np.linspace(np.min(xr) - 0.06 * (np.max(xr) - np.min(xr)),
                         np.max(xr) + 0.06 * (np.max(xr) - np.min(xr)), 120)
        for mask, col in [(is_c, C_POS), (is_a, C_NEG)]:
            dens = kde_1d(xr[mask], xg, bw_scale=1.15)
            ax_top.fill_between(xg, 0, dens, color=col, alpha=0.30)
            ax_top.plot(xg, dens, color=col, linewidth=1.3)

        ax_right.axis('off')
        yg = np.linspace(np.min(yr) - 0.1 * y_span, np.max(yr) + 0.1 * y_span, 120)
        for mask, col in [(is_c, C_POS), (is_a, C_NEG)]:
            dens = kde_1d(yr[mask], yg, bw_scale=1.15)
            ax_right.fill_betweenx(yg, 0, dens, color=col, alpha=0.30)
            ax_right.plot(dens, yg, color=col, linewidth=1.3)

    legend_items = [
        Line2D([0], [0], marker='o', color='w', markerfacecolor=C_POS, markersize=8, label=leg_labels[0]),
        Line2D([0], [0], marker='s', color='w', markerfacecolor=C_NEG, markersize=7.5, label=leg_labels[1]),
        Line2D([0], [0], color=C_INK, lw=2.2, label=leg_labels[2]),
        Line2D([0], [0], color='#475569', lw=1.2, ls='--', label=leg_labels[3]),
    ]
    fig.legend(handles=legend_items, loc='lower center', bbox_to_anchor=(0.5, -0.05),
               ncol=4, frameon=True, facecolor='#F8FAFC', edgecolor='#CBD5E1', fontsize=9.2)
    save_clean(fig, fname)


def main():
    d = load_data()
    for lang in ('TR', 'EN'):
        generate_figure_2(d, lang)
        generate_figure_3(lang)
        generate_figure_4(d, lang)
    print('Updated Figures 2, 3, and 4 (TR & EN) successfully generated in 04_manuscript!')


if __name__ == '__main__':
    main()
