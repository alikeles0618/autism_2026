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
# GLOBAL Q1 EDITORIAL STYLING & COLORBLIND-SAFE OKABE-ITO PALETTE
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
C_NEG_LIGHT = '#FDBA74' # Soft orange
C_INK = '#0F172A'       # Crisp dark charcoal
C_GRID = '#CBD5E1'      # Grid line gray
C_STAB = '#94A3B8'      # Stable gray

T_CRIT_33 = 2.034515    # Two-tailed Student-t critical value (df = 33, alpha = 0.05)

CSV_PATH = '/Users/alik/Desktop/anti_projeler/articles_all/autism_2026_testtekrar/02_data/cleaned_data.csv'
OUT_DIR = '/Users/alik/Desktop/anti_projeler/articles_all/autism_2026_testtekrar/06_sekil_tasarim_secenekleri'
CAND_OUT_DIR = os.path.join(OUT_DIR, 'ek_revize_adaylar')

os.makedirs(OUT_DIR, exist_ok=True)
os.makedirs(CAND_OUT_DIR, exist_ok=True)


def load_data():
    with open(CSV_PATH, 'r', encoding='utf-8') as f:
        rows = list(csv.DictReader(f))
    n = len(rows)

    def arr(key):
        return np.array([float(r[key]) for r in rows])

    yas = arr('YAS_2021')
    phv = arr('PHV_OFFSET_PRE')
    vki = arr('VKI_PRE')
    kulac = arr('KULAC_PRE')
    karis = arr('KARIS_PRE')
    w_pre, w_post = arr('KILO_PRE'), arr('KILO_POST')

    reac_pre, reac_post = arr('GORSEL_TOTAL_PRE'), arr('GORSEL_TOTAL_POST')
    sag_pre, sag_post = arr('SAG_EL_PRE'), arr('SAG_EL_POST')
    sol_pre, sol_post = arr('SOL_EL_PRE'), arr('SOL_EL_POST')
    mean_grip_pre = (sag_pre + sol_pre) / 2.0
    mean_grip_post = (sag_post + sol_post) / 2.0
    tot_grip_pre = sag_pre + sol_pre
    tot_grip_post = sag_post + sol_post
    rel_pre = mean_grip_pre / w_pre
    rel_post = mean_grip_post / w_post

    asym_pre, asym_post = arr('ASYMMETRY_PRE'), arr('ASYMMETRY_POST')
    power_pre, power_post = arr('SAYERS_POWER_PRE'), arr('SAYERS_POWER_POST')
    agil_pre, agil_post = arr('CEVIKLIK_PRE'), arr('CEVIKLIK_POST')
    flex_pre, flex_post = arr('ESNEKLIK_PRE'), arr('ESNEKLIK_POST')

    return {
        'n': n, 'yas': yas, 'is_child': yas <= 9, 'is_adol': yas >= 10,
        'phv': phv, 'vki': vki, 'kulac': kulac, 'karis': karis,
        'w_pre': w_pre, 'w_post': w_post,
        'reac': (reac_pre, reac_post),
        'rel': (rel_pre, rel_post),
        'power': (power_pre, power_post),
        'agil': (agil_pre, agil_post),
        'sag': (sag_pre, sag_post),
        'sol': (sol_pre, sol_post),
        'mean_grip': (mean_grip_pre, mean_grip_post),
        'tot_grip': (tot_grip_pre, tot_grip_post),
        'asym': (asym_pre, asym_post),
        'flex': (flex_pre, flex_post),
    }


def style_ax(ax, grid_axis='y'):
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    if grid_axis in ('y', 'both'):
        ax.yaxis.grid(True, linestyle='--', alpha=0.45, color=C_GRID, linewidth=0.8)
    if grid_axis in ('x', 'both'):
        ax.xaxis.grid(True, linestyle='--', alpha=0.45, color=C_GRID, linewidth=0.8)
    ax.set_axisbelow(True)


def save_clean(fig, filename, subdir=None):
    target_dir = subdir if subdir else OUT_DIR
    os.makedirs(target_dir, exist_ok=True)
    path = os.path.join(target_dir, filename)
    fig.savefig(path, dpi=300, bbox_inches='tight', facecolor='white', pad_inches=0.18)
    plt.close(fig)
    im = Image.open(path)
    im.save(path, 'PNG', pnginfo=PngInfo(), dpi=(300, 300))
    print(f'Saved cleanly (300 DPI, zero metadata): {path}')


def kde_1d(vals, grid, bw_scale=1.05):
    sd = np.std(vals, ddof=1)
    iqr = np.percentile(vals, 75) - np.percentile(vals, 25)
    sigma = min(sd, iqr / 1.34) if iqr > 0 else sd
    bw = max(1e-6, bw_scale * 0.9 * max(sigma, sd * 0.35) * (len(vals) ** -0.2))
    diffs = (grid[:, None] - vals[None, :]) / bw
    dens = np.mean(np.exp(-0.5 * diffs ** 2), axis=1) / (bw * math.sqrt(2 * math.pi))
    return dens


def paired_stats(pre, post):
    n = len(pre)
    m_pre, m_post = np.mean(pre), np.mean(post)
    ci_pre = T_CRIT_33 * np.std(pre, ddof=1) / math.sqrt(n)
    ci_post = T_CRIT_33 * np.std(post, ddof=1) / math.sqrt(n)
    diff = post - pre
    m_diff = np.mean(diff)
    se_diff = np.std(diff, ddof=1) / math.sqrt(n)
    ci_diff = T_CRIT_33 * se_diff
    return m_pre, ci_pre, m_post, ci_post, m_diff, ci_diff, diff


def swc_classify(pre, post, lower_is_better):
    swc = 0.20 * np.std(pre, ddof=1)
    imp = (pre - post) if lower_is_better else (post - pre)
    status = np.zeros(len(pre), dtype=int)
    status[imp > swc] = 1
    status[imp < -swc] = -1
    return swc, imp, status


def four_metrics(d):
    return [
        ('A', 'Görsel Reaksiyon Süresi', 'Süre (sn)', d['reac'][0], d['reac'][1],
         True, 'p = 0.011* · Δ = -13.0% · d = -0.35'),
        ('B', 'Rölatif Kavrama Kuvveti', 'Rölatif Kuvvet (kg/kg)', d['rel'][0], d['rel'][1],
         False, 'p = 0.033* · Δ = +9.0% · d = +0.15'),
        ('C', 'Sayers Zirve Mekanik Güç', 'Güç (Watt)', d['power'][0], d['power'][1],
         False, 'p = 0.076 · Δ = +11.6% · d = +0.05'),
        ('D', 'Pro-Agility 5-10-5 Çeviklik', 'Süre (sn)', d['agil'][0], d['agil'][1],
         True, 'p = 0.413 (ns) · Δ = -3.7% · d = -0.13'),
    ]


# =========================================================================
# 2A. PAIRED RAINCLOUD + SPAGHETTI PLOT (SWC RESPONDER COLORED)
# =========================================================================
def plot_2A_raincloud_spaghetti(d):
    fig, axes = plt.subplots(2, 2, figsize=(12.6, 10.0), dpi=300)
    rng = np.random.default_rng(42)
    metrics = four_metrics(d)

    for ax, (badge, title, ylabel, pre, post, low_better, stat_txt) in zip(axes.flat, metrics):
        style_ax(ax, 'y')
        m_pre, ci_pre, m_post, ci_post, _, _, _ = paired_stats(pre, post)
        _, _, status = swc_classify(pre, post, low_better)

        # Zero reference line for Sayers Peak Power (14/34 negative estimates)
        if 'Sayers' in title:
            ax.axhline(0.0, color='#64748B', linestyle=':', linewidth=1.1, zorder=1)

        # Half-violins (Pre left of 1.0, Post right of 2.0)
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

        # Individual paired spaghetti trajectories colored by SWC responder status
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

        # Cohort Mean ± 95% CI summary overlay
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
        ax.set_xticklabels(['Ön Test (Hafta 1)', 'Son Test (Hafta 14)'], fontsize=10, fontweight='bold')
        ax.set_ylabel(ylabel, fontsize=10, fontweight='bold')
        ax.set_title(f'{badge}   {title}', loc='left', fontsize=11.5, fontweight='bold', pad=10, color=C_INK)

    legend_items = [
        Line2D([0], [0], color=C_POS, lw=2.0, label='Olumlu SWC Yanıtı (Responder)'),
        Line2D([0], [0], color=C_STAB, lw=1.6, label='Stabil / Değişmeyen (±1×SWC)'),
        Line2D([0], [0], color=C_NEG, lw=2.0, label='Olumsuz SWC Yanıtı'),
        Line2D([0], [0], color=C_INK, lw=2.6, marker='o', markerfacecolor=C_POS,
               markeredgecolor='white', markersize=7, label='Grup Ortalaması ± %95 GA'),
    ]
    fig.legend(handles=legend_items, loc='lower center', bbox_to_anchor=(0.5, -0.02),
               ncol=4, frameon=True, facecolor='#F8FAFC', edgecolor='#CBD5E1', fontsize=9.5)
    fig.suptitle('Seçenek 2A: Eşleştirilmiş Yağmur Bulutu ve Bireysel SWC Değişim Çizgileri (N = 34)',
                 fontsize=13, fontweight='bold', y=1.01, color=C_INK)
    plt.tight_layout(rect=[0, 0.04, 1, 0.98], pad=2.0)
    save_clean(fig, 'Sekil_2A_Yagmur_Bulutu_Spaghetti_TR.png')


# =========================================================================
# 2B. CUMMING / GARDNER-ALTMAN PAIRED ESTIMATION PLOT
# =========================================================================
def plot_2B_cumming_estimation(d):
    fig = plt.figure(figsize=(13.2, 10.6), dpi=300)
    outer = fig.add_gridspec(2, 2, hspace=0.40, wspace=0.25)
    rng = np.random.default_rng(7)
    metrics = four_metrics(d)

    for idx, (badge, title, ylabel, pre, post, low_better, stat_txt) in enumerate(metrics):
        r, c = divmod(idx, 2)
        inner = outer[r, c].subgridspec(2, 1, height_ratios=[1.45, 1.0], hspace=0.30)
        ax_top = fig.add_subplot(inner[0])
        ax_bot = fig.add_subplot(inner[1])

        style_ax(ax_top, 'y')
        style_ax(ax_bot, 'x')
        ax_bot.spines['left'].set_visible(False)
        ax_bot.set_yticks([])

        m_pre, ci_pre, m_post, ci_post, m_diff, ci_diff, diff = paired_stats(pre, post)
        _, _, status = swc_classify(pre, post, low_better)

        if 'Sayers' in title:
            ax_top.axhline(0.0, color='#64748B', linestyle=':', linewidth=1.0, zorder=1)

        # Top slopegraph
        for i in range(len(pre)):
            col = C_POS if status[i] == 1 else (C_NEG if status[i] == -1 else '#CBD5E1')
            al = 0.60 if status[i] != 0 else 0.55
            lw = 1.2 if status[i] != 0 else 0.9
            ax_top.plot([0, 1], [pre[i], post[i]], color=col, linewidth=lw, alpha=al, zorder=2)

        ax_top.plot([0, 1], [m_pre, m_post], color=C_INK, linewidth=2.6, zorder=5)
        ax_top.scatter([0], [m_pre], color=C_PRE, edgecolor='white', linewidth=1.2, s=75, zorder=6)
        ax_top.scatter([1], [m_post], color=C_POS, edgecolor='white', linewidth=1.2, s=75, zorder=6)
        ax_top.set_xlim(-0.22, 1.22)
        ax_top.set_xticks([0, 1])
        ax_top.set_xticklabels(['Ön Test', 'Son Test'], fontsize=9.5, fontweight='bold')
        ax_top.set_ylabel(ylabel, fontsize=9.5, fontweight='bold')
        ax_top.set_title(f'{badge}   {title}', loc='left', fontsize=11.5, fontweight='bold', pad=8, color=C_INK)

        # Bottom horizontal paired difference estimation strip
        ax_bot.axvline(0.0, color=C_INK, linestyle='--', linewidth=1.3, zorder=3)
        d_lo, d_hi = np.min(diff), np.max(diff)
        d_pad = 0.14 * max(1e-6, d_hi - d_lo)
        grid = np.linspace(d_lo - d_pad, d_hi + d_pad, 180)
        dens = kde_1d(diff, grid, bw_scale=0.95)
        dens_norm = dens / np.max(dens) * 0.42

        is_sig = (m_diff - ci_diff > 0) or (m_diff + ci_diff < 0)
        band_col = C_POS_LIGHT if is_sig else '#CBD5E1'
        edge_col = C_POS if is_sig else '#475569'
        ax_bot.fill_between(grid, 0.22, 0.22 + dens_norm, color=band_col, alpha=0.55, zorder=2)
        ax_bot.plot(grid, 0.22 + dens_norm, color=edge_col, linewidth=1.2, zorder=3)

        # Individual delta points below zero baseline strip
        y_jit = rng.uniform(-0.32, -0.08, size=len(diff))
        dot_cols = [C_POS if s == 1 else (C_NEG if s == -1 else C_STAB) for s in status]
        ax_bot.scatter(diff, y_jit, c=dot_cols, s=22, alpha=0.80, edgecolors='white', linewidths=0.4, zorder=4)

        # Paired Mean Δ ± 95% CI horizontal bar
        ax_bot.errorbar([m_diff], [0.09], xerr=[[ci_diff], [ci_diff]], fmt='o',
                        color=C_INK, ecolor=C_INK, elinewidth=2.6, capsize=4.5, capthick=2.0,
                        markersize=7.5, zorder=6)

        fmt_ci = f'Δ = {m_diff:+.2f} [{m_diff - ci_diff:+.2f}, {m_diff + ci_diff:+.2f}]'
        if 'Rölatif' in title:
            fmt_ci = f'Δ = {m_diff:+.3f} [{m_diff - ci_diff:+.3f}, {m_diff + ci_diff:+.3f}]'
        elif 'Sayers' in title:
            fmt_ci = f'Δ = {m_diff:+.1f} [{m_diff - ci_diff:+.1f}, {m_diff + ci_diff:+.1f}]'

        ax_bot.text(0.5, 0.98, f'{fmt_ci}  |  {stat_txt.split(" · ")[0]}', transform=ax_bot.transAxes,
                    ha='center', va='top', fontsize=8.8, fontweight='bold',
                    color=C_POS if is_sig else '#334155',
                    bbox=dict(boxstyle='round,pad=0.22', facecolor='#F8FAFC', edgecolor='#CBD5E1', lw=0.8))

        ax_bot.set_ylim(-0.42, 1.04)
        ax_bot.set_xlim(min(d_lo - d_pad, -abs(ci_diff) * 0.6), max(d_hi + d_pad, abs(ci_diff) * 0.6))
        unit = ylabel.split('(')[-1].replace(')', '')
        ax_bot.set_xlabel(f'Eşleştirilmiş Değişim (Δ = Son - Ön, {unit})', fontsize=9.2, fontweight='bold')

    legend_items = [
        Line2D([0], [0], color=C_POS, lw=2.0, marker='o', markersize=5, label='Olumlu SWC Yanıtı'),
        Line2D([0], [0], color=C_STAB, lw=1.6, marker='o', markersize=5, label='Stabil (±1×SWC)'),
        Line2D([0], [0], color=C_NEG, lw=2.0, marker='o', markersize=5, label='Olumsuz SWC Yanıtı'),
        Line2D([0], [0], color=C_INK, lw=2.6, marker='o', markersize=7, label='Grup Δ Ortalaması ± %95 GA'),
    ]
    fig.legend(handles=legend_items, loc='lower center', bbox_to_anchor=(0.5, -0.02),
               ncol=4, frameon=True, facecolor='#F8FAFC', edgecolor='#CBD5E1', fontsize=9.5)
    fig.suptitle('Seçenek 2B: Cumming / Gardner-Altman Çiftli Etki Büyüklüğü (Estimation) Grafiği (N = 34)',
                 fontsize=13, fontweight='bold', y=1.01, color=C_INK)
    save_clean(fig, 'Sekil_2B_Cumming_Etki_Buyuklugu_TR.png')


# =========================================================================
# 2C. MINIMALIST DESPINED DOT-INTERVAL + AGE SUBGROUP STRIP
# =========================================================================
def plot_2C_nokta_aralik_yas(d):
    fig, axes = plt.subplots(2, 2, figsize=(12.2, 9.6), dpi=300)
    rng = np.random.default_rng(101)
    is_child = d['is_child']
    is_adol = d['is_adol']
    metrics = four_metrics(d)

    for ax, (badge, title, ylabel, pre, post, _, stat_txt) in zip(axes.flat, metrics):
        style_ax(ax, 'y')
        m_pre, ci_pre, m_post, ci_post, _, _, _ = paired_stats(pre, post)

        if 'Sayers' in title:
            ax.axhline(0.0, color='#64748B', linestyle=':', linewidth=1.1, zorder=1)

        # Childhood (5-9 yrs) slightly left, Adolescence (10-15 yrs) slightly right around each timepoint
        jit_c = rng.uniform(-0.07, 0.07, size=np.sum(is_child))
        jit_a = rng.uniform(-0.07, 0.07, size=np.sum(is_adol))

        ax.scatter(0.84 + jit_c, pre[is_child], color=C_POS, marker='o', s=34, alpha=0.62,
                   edgecolors='white', linewidths=0.6, zorder=3)
        ax.scatter(1.16 + jit_a, pre[is_adol], color=C_NEG, marker='s', s=32, alpha=0.62,
                   edgecolors='white', linewidths=0.6, zorder=3)

        ax.scatter(1.84 + jit_c, post[is_child], color=C_POS, marker='o', s=34, alpha=0.62,
                   edgecolors='white', linewidths=0.6, zorder=3)
        ax.scatter(2.16 + jit_a, post[is_adol], color=C_NEG, marker='s', s=32, alpha=0.62,
                   edgecolors='white', linewidths=0.6, zorder=3)

        # Cohort Mean ± 95% CI centered cleanly
        ax.plot([1.0, 2.0], [m_pre, m_post], color=C_INK, linestyle='--', linewidth=1.8, zorder=4)
        ax.errorbar([1.0, 2.0], [m_pre, m_post], yerr=[ci_pre, ci_post], fmt='o',
                    color=C_INK, markerfacecolor='white', markeredgewidth=2.2, markersize=8.5,
                    ecolor=C_INK, elinewidth=2.5, capsize=6, capthick=2.0, zorder=6)

        all_min = min(np.min(pre), np.min(post))
        all_max = max(np.max(pre), np.max(post))
        span = all_max - all_min
        ax.set_ylim(all_min - 0.08 * span, all_max + 0.25 * span)
        ax.set_xlim(0.55, 2.45)

        brk_y = all_max + 0.10 * span
        brk_h = 0.03 * span
        ax.plot([1.0, 1.0, 2.0, 2.0], [brk_y - brk_h, brk_y, brk_y, brk_y - brk_h],
                color=C_INK, linewidth=1.2, zorder=5)
        ax.text(1.5, brk_y + 0.02 * span, stat_txt, ha='center', va='bottom',
                fontsize=9.5, fontweight='bold', color=C_POS if '*' in stat_txt else '#475569')

        ax.set_xticks([1.0, 2.0])
        ax.set_xticklabels(['Ön Test (Hafta 1)', 'Son Test (Hafta 14)'], fontsize=10, fontweight='bold')
        ax.set_ylabel(ylabel, fontsize=10, fontweight='bold')
        ax.set_title(f'{badge}   {title}', loc='left', fontsize=11.5, fontweight='bold', pad=10, color=C_INK)

    legend_items = [
        Line2D([0], [0], marker='o', color='w', markerfacecolor=C_POS, markersize=7.5, alpha=0.8,
               label='Çocukluk Alt Grubu (5–9 yaş, n = 17)'),
        Line2D([0], [0], marker='s', color='w', markerfacecolor=C_NEG, markersize=7.5, alpha=0.8,
               label='Ergenlik Alt Grubu (10–15 yaş, n = 17)'),
        Line2D([0], [0], color=C_INK, lw=2.3, marker='o', markerfacecolor='white',
               markeredgewidth=2.2, markersize=8, label='Tüm Kohort Ort. ± %95 GA (N = 34)'),
    ]
    fig.legend(handles=legend_items, loc='lower center', bbox_to_anchor=(0.5, -0.02),
               ncol=3, frameon=True, facecolor='#F8FAFC', edgecolor='#CBD5E1', fontsize=9.5)
    fig.suptitle('Seçenek 2C: Minimalist Nokta–Aralık ve Gelişimsel Evre Dağılım Grafiği (N = 34)',
                 fontsize=13, fontweight='bold', y=1.01, color=C_INK)
    plt.tight_layout(rect=[0, 0.04, 1, 0.98], pad=2.0)
    save_clean(fig, 'Sekil_2C_Nokta_Aralik_Yas_Grubu_TR.png')


# =========================================================================
# 3A. INDIVIDUAL WATERFALL PLOT + ±1×SWC SHADED BAND
# =========================================================================
def plot_3A_waterfall_swc(d):
    fig, axes = plt.subplots(2, 2, figsize=(13.2, 9.6), dpi=300)
    specs = [
        ('A', 'Görsel Reaksiyon Süresi', 'sn', d['reac'][0], d['reac'][1], True, '2.91 sn'),
        ('B', 'Rölatif Kavrama Kuvveti', 'kg/kg', d['rel'][0], d['rel'][1], False, '0.03 kg/kg'),
        ('C', 'Pro-Agility 5-10-5 Çeviklik', 'sn', d['agil'][0], d['agil'][1], True, '0.88 sn'),
        ('D', 'Sayers Zirve Mekanik Güç', 'Watt', d['power'][0], d['power'][1], False, '255.2 W'),
    ]

    for ax, (badge, title, _, pre, post, low_better, swc_lbl) in zip(axes.flat, specs):
        style_ax(ax, 'y')
        swc, imp, status = swc_classify(pre, post, low_better)
        norm_vals = imp / swc
        order = np.argsort(-norm_vals)  # descending: biggest benefit on left
        vals_sorted = norm_vals[order]
        st_sorted = status[order]

        x = np.arange(1, len(vals_sorted) + 1)
        colors = [C_POS if s == 1 else (C_NEG if s == -1 else C_STAB) for s in st_sorted]

        # Shaded trivial/equivalence zone [-1, +1] * SWC
        ax.axhspan(-1.0, 1.0, color='#E2E8F0', alpha=0.80, zorder=0)
        ax.axhline(1.0, color=C_POS, linestyle='--', linewidth=1.1, alpha=0.85, zorder=2)
        ax.axhline(-1.0, color=C_NEG, linestyle='--', linewidth=1.1, alpha=0.85, zorder=2)
        ax.axhline(0.0, color='#334155', linestyle='-', linewidth=0.9, zorder=2)

        ax.bar(x, vals_sorted, width=0.76, color=colors, edgecolor='white', linewidth=0.5, zorder=3)

        n_pos = int(np.sum(status == 1))
        n_stb = int(np.sum(status == 0))
        n_neg = int(np.sum(status == -1))
        pct_pos, pct_stb, pct_neg = 100 * n_pos / 34, 100 * n_stb / 34, 100 * n_neg / 34

        y_min = min(-1.8, np.min(vals_sorted) - 0.8)
        y_max = max(2.8, np.max(vals_sorted) + 1.4)
        ax.set_ylim(y_min, y_max)
        ax.set_xlim(0.2, 34.8)

        summary = (f'SWC = {swc_lbl}\n'
                   f'Olumlu: {n_pos}/34 (%{pct_pos:.1f})\n'
                   f'Stabil: {n_stb}/34 (%{pct_stb:.1f})\n'
                   f'Olumsuz: {n_neg}/34 (%{pct_neg:.1f})')
        ax.text(0.98, 0.96, summary, transform=ax.transAxes, ha='right', va='top',
                fontsize=8.6, fontweight='bold', color=C_INK, linespacing=1.35,
                bbox=dict(boxstyle='round,pad=0.35', facecolor='#F8FAFC', edgecolor='#CBD5E1', lw=0.9))

        ax.text(34.3, 1.06, '+1×SWC', ha='right', va='bottom', fontsize=7.8, fontweight='bold', color=C_POS)
        ax.text(1.0, -1.06, '-1×SWC', ha='left', va='top', fontsize=7.8, fontweight='bold', color=C_NEG)

        ax.set_xticks([1, 8, 17, 26, 34])
        ax.set_xticklabels(['1', '8', '17', '26', '34'], fontsize=9)
        ax.set_xlabel('Bireysel Katılımcı Sırası (En Yüksek Gelişim → En Düşük)', fontsize=9.2, fontweight='bold')
        ax.set_ylabel('Değişim Büyüklüğü (Δ / SWC)', fontsize=9.5, fontweight='bold')
        ax.set_title(f'{badge}   {title}', loc='left', fontsize=11.5, fontweight='bold', pad=8, color=C_INK)

    legend_items = [
        Patch(facecolor=C_POS, label='Olumlu Yanıt (> +1 × SWC)'),
        Patch(facecolor=C_STAB, label='Stabil / Gürültü Bandı (±1 × SWC)'),
        Patch(facecolor=C_NEG, label='Olumsuz Yanıt (< -1 × SWC)'),
        Patch(facecolor='#E2E8F0', edgecolor='#94A3B8', label='SWC Eşdeğerlik Koridoru (±1 × SWC)'),
    ]
    fig.legend(handles=legend_items, loc='lower center', bbox_to_anchor=(0.5, -0.02),
               ncol=4, frameon=True, facecolor='#F8FAFC', edgecolor='#CBD5E1', fontsize=9.3)
    fig.suptitle('Seçenek 3A: En Küçük Anlamlı Değişim (SWC = 0.20 × SS_ön) Bireysel Şelale Grafiği (N = 34)',
                 fontsize=13, fontweight='bold', y=1.01, color=C_INK)
    plt.tight_layout(rect=[0, 0.04, 1, 0.98], pad=2.0)
    save_clean(fig, 'Sekil_3A_Bireysel_Selale_SWC_TR.png')


# =========================================================================
# 3B. DIVERGING CENTERED HORIZONTAL STACKED BAR + OKABE-ITO + CALLOUT
# =========================================================================
def plot_3B_diverging_swc(d):
    fig, ax = plt.subplots(figsize=(11.8, 5.2), dpi=300)
    style_ax(ax, 'x')
    ax.spines['left'].set_visible(False)
    ax.tick_params(axis='y', length=0)

    # Top to bottom order matching Tablo 3
    rows_spec = [
        ('Sayers Zirve Mekanik Güç\n(SWC = 255.2 W)', 6, 28, 0),
        ('Pro-Agility Çeviklik Süresi\n(SWC = 0.88 sn)', 10, 17, 7),
        ('Rölatif Kavrama Kuvveti\n(SWC = 0.03 kg/kg)', 10, 23, 1),
        ('Görsel Reaksiyon Süresi\n(SWC = 2.91 sn)', 14, 14, 6),
    ]
    y_pos = np.arange(len(rows_spec))
    bar_h = 0.54

    for idx, (label, n_pos, n_stb, n_neg) in enumerate(rows_spec):
        p_pos = 100.0 * n_pos / 34.0
        p_stb = 100.0 * n_stb / 34.0
        p_neg = 100.0 * n_neg / 34.0

        # Stable centered at x = 0: [-p_stb/2, +p_stb/2]
        left_stb = -p_stb / 2.0
        right_stb = p_stb / 2.0

        # Stable segment
        ax.barh(idx, p_stb, left=left_stb, height=bar_h, color='#CBD5E1',
                edgecolor='white', linewidth=1.2, zorder=3)
        ax.text(0.0, idx, f'{n_stb} (%{p_stb:.1f})', va='center', ha='center',
                fontsize=9.5, fontweight='bold', color='#1E293B', zorder=5)

        # Positive Responder (extending right from +p_stb/2)
        ax.barh(idx, p_pos, left=right_stb, height=bar_h, color=C_POS,
                edgecolor='white', linewidth=1.2, zorder=3)
        if p_pos >= 15.0:
            ax.text(right_stb + p_pos / 2.0, idx, f'{n_pos} (%{p_pos:.1f})',
                    va='center', ha='center', fontsize=9.5, fontweight='bold', color='white', zorder=5)
        else:
            ax.text(right_stb + p_pos + 1.5, idx, f'{n_pos} (%{p_pos:.1f})',
                    va='center', ha='left', fontsize=9.5, fontweight='bold', color=C_POS, zorder=5)

        # Negative Responder (extending left from -p_stb/2)
        if n_neg > 0:
            ax.barh(idx, p_neg, left=left_stb - p_neg, height=bar_h, color=C_NEG,
                    edgecolor='white', linewidth=1.2, zorder=3)
            if p_neg >= 13.0:
                ax.text(left_stb - p_neg / 2.0, idx, f'{n_neg} (%{p_neg:.1f})',
                        va='center', ha='center', fontsize=9.5, fontweight='bold', color='white', zorder=5)
            else:
                # Explicit outside callout for narrow 1 (%2.9) bar
                ax.annotate(f'{n_neg} (%{p_neg:.1f})',
                            xy=(left_stb - p_neg * 0.5, idx),
                            xytext=(left_stb - p_neg - 4.5, idx),
                            ha='right', va='center', fontsize=9.5, fontweight='bold', color=C_NEG,
                            arrowprops=dict(arrowstyle='-', color=C_NEG, lw=1.2), zorder=6)
        else:
            ax.text(left_stb - 1.6, idx, '0 (%0.0)', va='center', ha='right',
                    fontsize=9.0, fontweight='bold', color='#64748B', zorder=5)

    ax.axvline(0.0, color='#475569', linestyle=':', linewidth=1.0, zorder=4)
    ax.set_yticks(y_pos)
    ax.set_yticklabels([r[0] for r in rows_spec], fontsize=10, fontweight='bold', color=C_INK)
    ax.set_xlim(-64, 72)
    xticks = [-60, -40, -20, 0, 20, 40, 60]
    ax.set_xticks(xticks)
    ax.set_xticklabels([f'%{abs(v)}' if v != 0 else 'Merkez' for v in xticks], fontsize=9.5, fontweight='bold')
    ax.set_xlabel('◄── Olumsuz Yanıt (%)       |       Stabil Bölge (±1×SWC)       |       Olumlu Yanıt (%) ──►',
                  fontsize=10, fontweight='bold', labelpad=10)
    ax.set_title('Seçenek 3B: En Küçük Anlamlı Değişim (SWC = 0.20 × SS_ön) Merkezden Ayrışan Yanıt Dağılımı (N = 34)',
                 fontsize=12, fontweight='bold', pad=14, color=C_INK)

    legend_items = [
        Patch(facecolor=C_NEG, label='Olumsuz Yanıt (< -1 × SWC)'),
        Patch(facecolor='#CBD5E1', label='Stabil / Değişmeyen (±1 × SWC)'),
        Patch(facecolor=C_POS, label='Olumlu Yanıt (> +1 × SWC)'),
    ]
    ax.legend(handles=legend_items, loc='lower center', bbox_to_anchor=(0.5, -0.29),
              ncol=3, frameon=True, facecolor='#F8FAFC', edgecolor='#CBD5E1', fontsize=9.5)
    plt.tight_layout()
    save_clean(fig, 'Sekil_3B_Merkezden_Ayrisan_SWC_TR.png')


# =========================================================================
# 3C. CLEVELAND LOLLIPOP / DOT-PROPORTIONAL RESPONDER CHART
# =========================================================================
def plot_3C_cleveland_lollipop(d):
    fig, ax = plt.subplots(figsize=(11.4, 6.2), dpi=300)
    style_ax(ax, 'x')
    ax.tick_params(axis='y', length=0)

    specs = [
        ('Sayers Zirve Mekanik Güç (SWC = 255.2 W)', [(6, 'Olumlu'), (28, 'Stabil'), (0, 'Olumsuz')]),
        ('Pro-Agility Çeviklik Süresi (SWC = 0.88 sn)', [(10, 'Olumlu'), (17, 'Stabil'), (7, 'Olumsuz')]),
        ('Rölatif Kavrama Kuvveti (SWC = 0.03 kg/kg)', [(10, 'Olumlu'), (23, 'Stabil'), (1, 'Olumsuz')]),
        ('Görsel Reaksiyon Süresi (SWC = 2.91 sn)', [(14, 'Olumlu'), (14, 'Stabil'), (6, 'Olumsuz')]),
    ]

    style_map = {
        'Olumlu': (C_POS, 'o', +0.26),
        'Stabil': (C_STAB, 'D', 0.0),
        'Olumsuz': (C_NEG, 's', -0.26),
    }

    y_centers = np.arange(len(specs)) * 1.15
    for idx, (label, groups) in enumerate(specs):
        yc = y_centers[idx]
        if idx % 2 == 0:
            ax.axhspan(yc - 0.48, yc + 0.48, color='#F8FAFC', zorder=0)
        for n_cnt, cat in groups:
            col, marker, offset = style_map[cat]
            pct = 100.0 * n_cnt / 34.0
            y = yc + offset
            ax.plot([0.0, pct], [y, y], color=col, linewidth=2.3, alpha=0.85, zorder=2)
            ax.scatter([pct], [y], color=col, marker=marker, s=75, edgecolors='white', linewidths=0.9, zorder=4)
            ax.text(pct + 1.8, y, f'{n_cnt}/34 (%{pct:.1f})', va='center', ha='left',
                    fontsize=9.0, fontweight='bold', color=col if n_cnt > 0 else '#64748B')

    ax.set_yticks(y_centers)
    ax.set_yticklabels([s[0] for s in specs], fontsize=10, fontweight='bold', color=C_INK)
    ax.set_xlim(0, 102)
    ax.set_ylim(-0.62, y_centers[-1] + 0.62)
    ax.set_xlabel('Kohort Yüzdesi (%, N = 34)', fontsize=10, fontweight='bold', labelpad=8)
    ax.set_title('Seçenek 3C: En Küçük Anlamlı Değişim (SWC = 0.20 × SS_ön) Cleveland Lolipop Dağılımı',
                 fontsize=12, fontweight='bold', pad=14, color=C_INK)

    legend_items = [
        Line2D([0], [0], color=C_POS, lw=2.3, marker='o', markersize=7.5, label='Olumlu Yanıt (> +1 × SWC)'),
        Line2D([0], [0], color=C_STAB, lw=2.3, marker='D', markersize=6.8, label='Stabil / Değişmeyen (±1 × SWC)'),
        Line2D([0], [0], color=C_NEG, lw=2.3, marker='s', markersize=7.2, label='Olumsuz Yanıt (< -1 × SWC)'),
    ]
    ax.legend(handles=legend_items, loc='lower center', bbox_to_anchor=(0.5, -0.22),
              ncol=3, frameon=True, facecolor='#F8FAFC', edgecolor='#CBD5E1', fontsize=9.5)
    plt.tight_layout()
    save_clean(fig, 'Sekil_3C_Cleveland_Nokta_Oran_TR.png')


# =========================================================================
# HELPER: SUBGROUP SCATTER WITH EXACT OLS LINE + 95% CI
# =========================================================================
def draw_regression_core(ax, x_raw, y_raw, is_child, is_adol, x_jitter,
                         xlabel, ylabel, title, badge, r_val, r2_val, p_str, set_title=True):
    style_ax(ax, 'both')
    n = len(x_raw)
    slope, intercept = np.polyfit(x_raw, y_raw, 1)
    x_line = np.linspace(np.min(x_raw) - 0.04 * (np.max(x_raw) - np.min(x_raw)),
                         np.max(x_raw) + 0.04 * (np.max(x_raw) - np.min(x_raw)), 140)
    y_line = slope * x_line + intercept
    y_pred = slope * x_raw + intercept
    s_err = math.sqrt(np.sum((y_raw - y_pred) ** 2) / (n - 2))
    t_32 = 2.036933
    ci = t_32 * s_err * np.sqrt(1.0 / n + ((x_line - np.mean(x_raw)) ** 2) / np.sum((x_raw - np.mean(x_raw)) ** 2))

    ax.axhline(0.0, color='#475569', linestyle='--', linewidth=1.1, zorder=1)
    ax.fill_between(x_line, y_line - ci, y_line + ci, color='#94A3B8', alpha=0.22, zorder=2)
    ax.plot(x_line, y_line, color=C_INK, linewidth=2.2, zorder=3)

    x_plot = x_raw + x_jitter
    ax.scatter(x_plot[is_child], y_raw[is_child], color=C_POS, marker='o', s=58, alpha=0.88,
               edgecolors='white', linewidths=0.9, zorder=4, label='Çocukluk (5–9 yaş, n = 17)')
    ax.scatter(x_plot[is_adol], y_raw[is_adol], color=C_NEG, marker='s', s=54, alpha=0.88,
               edgecolors='white', linewidths=0.9, zorder=4, label='Ergenlik (10–15 yaş, n = 17)')

    y_min_d, y_max_d = np.min(y_raw), np.max(y_raw)
    y_span = y_max_d - y_min_d
    ax.set_ylim(y_min_d - 0.12 * y_span, y_max_d + 0.22 * y_span)

    ax.text(0.04, 0.96, f'r = {r_val:+.2f}, r² = {r2_val:.3f}\n{p_str}',
            transform=ax.transAxes, ha='left', va='top', fontsize=9.5, fontweight='bold', color=C_INK,
            bbox=dict(boxstyle='round,pad=0.3', facecolor='#F8FAFC', edgecolor='#CBD5E1', lw=0.85))
    ax.text(x_line[-1], -0.025 * y_span, 'Δ = 0', ha='right', va='top',
            fontsize=8.2, color='#475569', style='italic',
            bbox=dict(boxstyle='square,pad=0.1', facecolor='white', edgecolor='none', alpha=0.8))

    ax.set_xlabel(xlabel, fontsize=10, fontweight='bold', labelpad=7)
    ax.set_ylabel(ylabel, fontsize=10, fontweight='bold', labelpad=7)
    if set_title:
        ax.set_title(f'{badge}   {title}', loc='left', fontsize=11.3, fontweight='bold', pad=8, color=C_INK)
    return x_line


# =========================================================================
# 4A. SUBGROUP-ENCODED SCATTER + Δ=0 REFERENCE + MARGINAL DENSITY STRIPS
# =========================================================================
def plot_4A_subgroup_marginal(d):
    fig = plt.figure(figsize=(13.2, 6.2), dpi=300)
    outer = fig.add_gridspec(1, 2, wspace=0.28)
    rng = np.random.default_rng(2026)

    karis = d['karis']
    kulac = d['kulac']
    d_sag = d['sag'][1] - d['sag'][0]
    d_mean = d['mean_grip'][1] - d['mean_grip'][0]
    is_c, is_a = d['is_child'], d['is_adol']

    specs = [
        ('A', 'El Karış Uzunluğu ve Kuvvet Gelişimi',
         'Başlangıç El Karış Uzunluğu (cm)', 'Δ Sağ/Dominant El Kavrama Kuvveti (kg)',
         karis, d_sag, rng.uniform(-0.14, 0.14, size=len(karis)), 0.4680, 0.2190, 'p = 0.005**'),
        ('B', 'Kulaç Uzunluğu ve Kuvvet Gelişimi',
         'Başlangıç Kulaç Uzunluğu (cm)', 'Δ Ortalama El Kavrama Kuvveti (kg)',
         kulac, d_mean, rng.uniform(-0.45, 0.45, size=len(kulac)), 0.3658, 0.1338, 'p = 0.033*'),
    ]

    for idx, (badge, title, xl, yl, xr, yr, jit, rv, r2v, ps) in enumerate(specs):
        sub = outer[idx].subgridspec(2, 2, width_ratios=[4.2, 1.0], height_ratios=[1.0, 4.2],
                                     hspace=0.06, wspace=0.06)
        ax_sc = fig.add_subplot(sub[1, 0])
        ax_top = fig.add_subplot(sub[0, 0], sharex=ax_sc)
        ax_right = fig.add_subplot(sub[1, 1], sharey=ax_sc)

        draw_regression_core(ax_sc, xr, yr, is_c, is_a, jit, xl, yl, '', badge, rv, r2v, ps, set_title=False)
        ax_top.set_title(f'{badge}   {title}', loc='left', fontsize=11.3, fontweight='bold', pad=8, color=C_INK)

        # Top marginal KDE
        ax_top.axis('off')
        xg = np.linspace(np.min(xr) - 0.06 * (np.max(xr) - np.min(xr)),
                         np.max(xr) + 0.06 * (np.max(xr) - np.min(xr)), 120)
        for mask, col in [(is_c, C_POS), (is_a, C_NEG)]:
            dens = kde_1d(xr[mask], xg, bw_scale=1.15)
            ax_top.fill_between(xg, 0, dens, color=col, alpha=0.30)
            ax_top.plot(xg, dens, color=col, linewidth=1.3)

        # Right marginal KDE
        ax_right.axis('off')
        yg = np.linspace(np.min(yr) - 0.1 * (np.max(yr) - np.min(yr)),
                         np.max(yr) + 0.1 * (np.max(yr) - np.min(yr)), 120)
        for mask, col in [(is_c, C_POS), (is_a, C_NEG)]:
            dens = kde_1d(yr[mask], yg, bw_scale=1.15)
            ax_right.fill_betweenx(yg, 0, dens, color=col, alpha=0.30)
            ax_right.plot(dens, yg, color=col, linewidth=1.3)

    legend_items = [
        Line2D([0], [0], marker='o', color='w', markerfacecolor=C_POS, markersize=8,
               label='Çocukluk Alt Grubu (5–9 yaş, n = 17)'),
        Line2D([0], [0], marker='s', color='w', markerfacecolor=C_NEG, markersize=7.5,
               label='Ergenlik Alt Grubu (10–15 yaş, n = 17)'),
        Line2D([0], [0], color=C_INK, lw=2.2, label='OLS Doğrusal Eğim ve %95 GA'),
        Line2D([0], [0], color='#475569', lw=1.2, ls='--', label='Sıfır Değişim Çizgisi (Δ = 0)'),
    ]
    fig.legend(handles=legend_items, loc='lower center', bbox_to_anchor=(0.5, -0.04),
               ncol=4, frameon=True, facecolor='#F8FAFC', edgecolor='#CBD5E1', fontsize=9.2)
    fig.suptitle('Seçenek 4A: Gelişimsel Evre Kodlu Antropometrik Regresyon ve Marjinal Dağılımlar (N = 34)',
                 fontsize=12.8, fontweight='bold', y=1.02, color=C_INK)
    save_clean(fig, 'Sekil_4A_Yas_Kodlu_Regresyon_TR.png')


# =========================================================================
# 4B. 3-PANEL MECHANISTIC FIGURE (SCATTER A, B + COLLINEARITY r vs BETA C)
# =========================================================================
def plot_4B_mechanistic_3panel(d):
    fig = plt.figure(figsize=(15.4, 5.6), dpi=300)
    gs = fig.add_gridspec(1, 3, width_ratios=[1.12, 1.12, 1.10], wspace=0.42)
    rng = np.random.default_rng(2026)

    karis, kulac = d['karis'], d['kulac']
    d_sag = d['sag'][1] - d['sag'][0]
    d_mean = d['mean_grip'][1] - d['mean_grip'][0]
    is_c, is_a = d['is_child'], d['is_adol']

    ax1 = fig.add_subplot(gs[0])
    draw_regression_core(ax1, karis, d_sag, is_c, is_a, rng.uniform(-0.14, 0.14, len(karis)),
                         'Başlangıç El Karış Uzunluğu (cm)', 'Δ Sağ/Dominant Kavrama (kg)',
                         'El Karışı – Dominant Kuvvet', 'A', 0.4680, 0.2190, 'p = 0.005**')

    ax2 = fig.add_subplot(gs[1])
    draw_regression_core(ax2, kulac, d_mean, is_c, is_a, rng.uniform(-0.45, 0.45, len(kulac)),
                         'Başlangıç Kulaç Uzunluğu (cm)', 'Δ Ortalama Kavrama (kg)',
                         'Kulaç Uzunluğu – Ortalama Kuvvet', 'B', 0.3658, 0.1338, 'p = 0.033*')

    # Panel C: Bivariate Pearson r vs Multiple OLS Standardized Beta (from Tablo 4)
    ax3 = fig.add_subplot(gs[2])
    style_ax(ax3, 'x')
    preds = [
        'Başlangıç VKİ\n(kg/m²)',
        'Mirwald PHV\nOfseti (yıl)',
        'Başlangıç El\nKarışı (cm)',
        'Başlangıç Kulaç\nUzunluğu (cm)',
    ]
    r_vals = [-0.04, 0.25, 0.36, 0.37]      # bivariate Pearson r vs Δ Mean Grip
    b_vals = [-0.50, -0.04, -0.04, 0.77]    # Tablo 4 standardized β
    r_sigs = ['', '', '*', '*']

    y = np.arange(len(preds))
    h = 0.34
    ax3.axvline(0.0, color=C_INK, linewidth=1.1, zorder=2)
    ax3.barh(y + h / 2, r_vals, height=h, color=C_POS, edgecolor='white', label='İki Değişkenli (Pearson r)')
    ax3.barh(y - h / 2, b_vals, height=h, color=C_NEG, edgecolor='white', alpha=0.88,
             label='Çoklu OLS Modeli (Standart β)')

    for idx, (rv, bv, rs) in enumerate(zip(r_vals, b_vals, r_sigs)):
        if rv >= 0:
            ax3.text(rv + 0.03, y[idx] + h / 2, f'r={rv:+.2f}{rs}', va='center', ha='left',
                     fontsize=8.3, fontweight='bold', color=C_POS)
        else:
            ax3.text(rv - 0.03, y[idx] + h / 2, f'r={rv:+.2f}', va='center', ha='right',
                     fontsize=8.3, fontweight='bold', color=C_POS)
        if bv >= 0:
            ax3.text(bv + 0.03, y[idx] - h / 2, f'β={bv:+.2f}', va='center', ha='left',
                     fontsize=8.3, fontweight='bold', color=C_NEG)
        else:
            ax3.text(bv - 0.03, y[idx] - h / 2, f'β={bv:+.2f}', va='center', ha='right',
                     fontsize=8.3, fontweight='bold', color=C_NEG)

    ax3.set_yticks(y)
    ax3.set_yticklabels(preds, fontsize=8.8, fontweight='bold')
    ax3.set_xlim(-1.05, 1.25)
    ax3.set_ylim(-0.6, 4.15)
    ax3.set_xlabel('İlişki / Standartlaştırılmış Katsayı Büyüklüğü', fontsize=9.3, fontweight='bold', labelpad=7)
    ax3.set_title('C   Eşdoğrusallık: r vs. Çoklu Model β', loc='left', fontsize=11.3, fontweight='bold', pad=8, color=C_INK)
    ax3.text(0.5, 0.96, 'OLS Modeli: R² = 0.280, p = 0.044* · PHV: p = 0.939',
             transform=ax3.transAxes, ha='center', va='top', fontsize=8.1, fontweight='bold', color='#334155',
             bbox=dict(boxstyle='round,pad=0.25', facecolor='#F8FAFC', edgecolor='#CBD5E1', lw=0.8))

    legend_items = [
        Line2D([0], [0], marker='o', color='w', markerfacecolor=C_POS, markersize=7.5, label='Çocukluk (5–9 yaş, n=17)'),
        Line2D([0], [0], marker='s', color='w', markerfacecolor=C_NEG, markersize=7.2, label='Ergenlik (10–15 yaş, n=17)'),
        Line2D([0], [0], color=C_INK, lw=2.2, label='OLS Regresyon (%95 GA)'),
        Patch(facecolor=C_POS, label='İki Değişkenli r (Δ Ort. Kavrama)'),
        Patch(facecolor=C_NEG, alpha=0.88, label='Çoklu OLS Standart β'),
    ]
    fig.legend(handles=legend_items, loc='lower center', bbox_to_anchor=(0.5, -0.06),
               ncol=5, frameon=True, facecolor='#F8FAFC', edgecolor='#CBD5E1', fontsize=8.7)
    fig.suptitle('Seçenek 4B: Morfolojik Yordayıcılar ve Olgunlaşma Eşdoğrusallığı Mekanistik Özeti (N = 34)',
                 fontsize=12.8, fontweight='bold', y=1.02, color=C_INK)
    save_clean(fig, 'Sekil_4B_Uc_Panelli_Mekanistik_Regresyon_TR.png')


# =========================================================================
# BONUS REVISED VERSIONS OF CANDIDATES 1..4 (FIXING OBSERVED LAYOUT BUGS)
# =========================================================================
def plot_revised_candidates(d):
    n = d['n']

    # --- Aday 1R: Radar Z-score without label/legend clipping ---
    def z_mean(pre_arr, post_arr, invert=False):
        m, s = np.mean(pre_arr), np.std(pre_arr, ddof=1)
        z1 = (pre_arr - m) / s
        z2 = (post_arr - m) / s
        return (-np.mean(z1), -np.mean(z2)) if invert else (np.mean(z1), np.mean(z2))

    axes_specs = [
        ('Görsel Reaksiyon\nHızı', d['reac'][0], d['reac'][1], True),
        ('Rölatif Kavrama\nKuvveti', d['rel'][0], d['rel'][1], False),
        ('Toplam Mutlak\nKavrama', d['tot_grip'][0], d['tot_grip'][1], False),
        ('Sayers Zirve\nMekanik Güç', d['power'][0], d['power'][1], False),
        ('Pro-Agility\nÇeviklik Hızı', d['agil'][0], d['agil'][1], True),
        ('Otur-Eriş\nEsneklik', d['flex'][0], d['flex'][1], False),
        ('Bilateral Simetri\nDüzeyi', d['asym'][0], d['asym'][1], True),
    ]
    pre_vals, post_vals, labels = [], [], []
    for lbl, pa, qa, inv in axes_specs:
        v1, v2 = z_mean(pa, qa, inv)
        pre_vals.append(v1)
        post_vals.append(v2)
        labels.append(lbl)

    angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
    pre_vals += pre_vals[:1]
    post_vals += post_vals[:1]
    angles += angles[:1]

    fig, ax = plt.subplots(figsize=(9.2, 9.0), subplot_kw=dict(polar=True), dpi=300)
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    ax.plot(angles, pre_vals, color=C_PRE, linewidth=2.0, linestyle='--', marker='o', markersize=5.5,
            label='Ön Test (Z = 0 Referansı)')
    ax.fill(angles, pre_vals, color=C_PRE_LIGHT, alpha=0.22)
    ax.plot(angles, post_vals, color=C_POS, linewidth=2.6, linestyle='-', marker='s', markersize=6.5,
            label='12 Hafta Son Test')
    ax.fill(angles, post_vals, color=C_POS_LIGHT, alpha=0.30)

    ax.set_thetagrids(np.degrees(angles[:-1]), labels, fontsize=10, fontweight='bold', color=C_INK)
    ax.tick_params(axis='x', pad=22)
    ax.set_ylim(-0.32, 0.45)
    ax.set_yticks([-0.2, 0.0, 0.2, 0.4])
    ax.set_yticklabels(['-0.2 SD', '0', '+0.2 SD', '+0.4 SD'], fontsize=8.5, color='#475569')
    ax.set_rlabel_position(25)
    ax.grid(True, linestyle='--', alpha=0.5, color=C_GRID)
    ax.spines['polar'].set_color('#94A3B8')
    ax.set_title('Revize Aday 1R: Çok Boyutlu Nöromotor Z-Skoru Profili (N = 34)',
                 fontsize=12.5, fontweight='bold', pad=34, color=C_INK)
    ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.10), ncol=2, frameon=True,
              facecolor='#F8FAFC', edgecolor='#CBD5E1', fontsize=10)
    save_clean(fig, 'Aday_1R_Radar_Z_Skor_Revize_TR.png', CAND_OUT_DIR)

    # --- Aday 2R: Despined Bilateral Asymmetry + Okabe-Ito ---
    sag_pre, sag_post = d['sag']
    sol_pre, sol_post = d['sol']
    asym_pre, asym_post = d['asym']
    fig, axes = plt.subplots(1, 2, figsize=(12.4, 5.4), dpi=300)

    ax1, ax2 = axes[0], axes[1]
    style_ax(ax1, 'both')
    style_ax(ax2, 'y')
    max_v = max(np.max(sag_post), np.max(sol_post)) + 3.0
    lx = np.linspace(0, max_v, 100)
    ax1.plot(lx, lx, color=C_INK, lw=1.4, label='Tam Simetri (y = x)')
    ax1.fill_between(lx, lx * 0.90, lx * 1.10, color=C_POS_LIGHT, alpha=0.25, label='Normatif Alan (±%10)')
    ax1.plot(lx, lx * 1.15, color=C_NEG, ls='--', lw=1.1, label='Klinik Eşik (±%15)')
    ax1.plot(lx, lx * 0.85, color=C_NEG, ls='--', lw=1.1)
    for i in range(n):
        ax1.annotate('', xy=(sol_post[i], sag_post[i]), xytext=(sol_pre[i], sag_pre[i]),
                     arrowprops=dict(arrowstyle='->', color='#94A3B8', alpha=0.55, lw=0.9))
    ax1.scatter(sol_pre, sag_pre, color=C_PRE, edgecolor='white', s=44, alpha=0.8, label='Ön Test', zorder=4)
    ax1.scatter(sol_post, sag_post, color=C_POS, edgecolor='white', s=54, alpha=0.92, label='Son Test', zorder=5)
    ax1.set_xlim(0, max_v)
    ax1.set_ylim(0, max_v)
    ax1.set_xlabel('Sol El Kavrama Kuvveti (kg)', fontsize=10, fontweight='bold')
    ax1.set_ylabel('Sağ El Kavrama Kuvveti (kg)', fontsize=10, fontweight='bold')
    ax1.set_title('A   Sağ–Sol El Eşleşmesi ve Klinik Eşikler', loc='left', fontsize=11.3, fontweight='bold', pad=8)
    ax1.legend(loc='upper left', frameon=True, facecolor='#F8FAFC', edgecolor='#CBD5E1', fontsize=8.8)

    for i in range(n):
        ax2.plot([1.0, 2.0], [asym_pre[i], asym_post[i]], color='#CBD5E1', lw=1.0, zorder=2)
        ax2.scatter([1.0, 2.0], [asym_pre[i], asym_post[i]], color=C_STAB, s=20, alpha=0.65, zorder=3)
    m_ap, _, m_apo, _, _, _, _ = paired_stats(asym_pre, asym_post)
    ci_ap = T_CRIT_33 * np.std(asym_pre, ddof=1) / math.sqrt(n)
    ci_apo = T_CRIT_33 * np.std(asym_post, ddof=1) / math.sqrt(n)
    ax2.errorbar([1.0, 2.0], [m_ap, m_apo], yerr=[ci_ap, ci_apo], fmt='o-', color=C_INK,
                 lw=2.5, capsize=6, capthick=2.0, markersize=7.5, zorder=5)
    ax2.axhline(15.0, color=C_NEG, ls='--', lw=1.4, label='Klinik Asimetri Eşiği (%15)')
    ax2.set_xlim(0.6, 2.4)
    ax2.set_xticks([1.0, 2.0])
    ax2.set_xticklabels(['Ön Test\n(%16.8 ± 11.9)', 'Son Test\n(%17.0 ± 11.7)'], fontsize=9.8, fontweight='bold')
    ax2.set_ylabel('Bilateral Asimetri İndeksi (%)', fontsize=10, fontweight='bold')
    ax2.set_title('B   Bireysel Asimetri Seyri (p = 0.914, Stabil)', loc='left', fontsize=11.3, fontweight='bold', pad=8)
    ax2.legend(loc='upper right', frameon=True, facecolor='#F8FAFC', edgecolor='#CBD5E1', fontsize=9.0)
    fig.suptitle('Revize Aday 2R: Bilateral Ekstremite Kuvvet Simetrisi ve Klinik Eşik Analizi (N = 34)',
                 fontsize=12.5, fontweight='bold', y=1.02, color=C_INK)
    plt.tight_layout()
    save_clean(fig, 'Aday_2R_Bilateral_Asimetri_Revize_TR.png', CAND_OUT_DIR)

    # --- Aday 3R: Full [-1.0, +1.0] Scale Correlation Heatmap ---
    d_reac = d['reac'][1] - d['reac'][0]
    d_rel = d['rel'][1] - d['rel'][0]
    d_tot = d['tot_grip'][1] - d['tot_grip'][0]
    d_pow = d['power'][1] - d['power'][0]
    d_agil = d['agil'][1] - d['agil'][0]
    d_flex = d['flex'][1] - d['flex'][0]
    vnames = ['Δ Görsel Reaksiyon', 'Δ Rölatif Kavrama', 'Δ Toplam Kavrama', 'Δ Sayers Güç',
              'Δ Pro-Agility', 'Δ Otur-Eriş', 'Başlangıç Karış', 'Başlangıç Kulaç',
              'Başlangıç VKİ', 'Mirwald PHV Ofseti']
    mat = np.column_stack([d_reac, d_rel, d_tot, d_pow, d_agil, d_flex,
                           d['karis'], d['kulac'], d['vki'], d['phv']])
    corr = np.corrcoef(mat, rowvar=False)
    mask = np.triu(np.ones_like(corr, dtype=bool), k=1)
    plot_c = np.ma.masked_array(corr, mask=mask)

    fig, ax = plt.subplots(figsize=(9.8, 8.6), dpi=300)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    cax = ax.imshow(plot_c, cmap='RdBu_r', vmin=-1.0, vmax=1.0)
    for i in range(len(vnames)):
        for j in range(i + 1):
            val = corr[i, j]
            t_v = val * math.sqrt(n - 2) / math.sqrt(max(1e-9, 1.0 - val ** 2))
            sig = '**' if abs(t_v) > 2.738 and i != j else ('*' if abs(t_v) > 2.037 and i != j else '')
            txt_c = 'white' if abs(val) > 0.52 else C_INK
            ax.text(j, i, f'{val:+.2f}{sig}', ha='center', va='center',
                    fontsize=8.6, fontweight='bold' if sig else 'normal', color=txt_c)
    ax.set_xticks(np.arange(len(vnames)))
    ax.set_yticks(np.arange(len(vnames)))
    ax.set_xticklabels(vnames, rotation=40, ha='right', fontsize=9.2, fontweight='bold')
    ax.set_yticklabels(vnames, fontsize=9.2, fontweight='bold')
    cbar = fig.colorbar(cax, ax=ax, fraction=0.044, pad=0.04)
    cbar.set_label('Pearson Korelasyon Katsayısı (r, Tam Ölçek [-1.0, +1.0])', fontsize=9.8, fontweight='bold')
    ax.set_title('Revize Aday 3R: Performans Değişimleri ve Morfolojik Korelasyon Matrisi (*p < 0.05, **p < 0.01)',
                 fontsize=11.5, fontweight='bold', pad=14, color=C_INK)
    plt.tight_layout()
    save_clean(fig, 'Aday_3R_Korelasyon_Isi_Haritasi_Revize_TR.png', CAND_OUT_DIR)

    # --- Aday 4R: Fixed Subgroup Forest Plot (zero overlap, top legend) ---
    idx_early = np.where(d['phv'] < -3.5)[0]
    idx_late = np.where(d['phv'] >= -3.5)[0]
    f_metrics = [
        ('Görsel Reaksiyon Hızı', d['reac'][0], d['reac'][1], True),
        ('Rölatif Kavrama Kuvveti', d['rel'][0], d['rel'][1], False),
        ('Toplam Mutlak Kavrama', d['tot_grip'][0], d['tot_grip'][1], False),
        ('Sayers Zirve Mekanik Güç', d['power'][0], d['power'][1], False),
    ]

    def calc_d(pa, qa, idxs=None, invert=False):
        if idxs is not None:
            pa, qa = pa[idxs], qa[idxs]
        ns = len(pa)
        diff = (pa - qa) if invert else (qa - pa)
        s_av = math.sqrt((np.var(pa, ddof=1) + np.var(qa, ddof=1)) / 2.0)
        dv = np.mean(diff) / s_av
        se = math.sqrt((2.0 / ns) + (dv ** 2 / (2.0 * ns)))
        return dv, dv - 1.96 * se, dv + 1.96 * se

    fig, ax = plt.subplots(figsize=(10.8, 6.8), dpi=300)
    style_ax(ax, 'x')
    y_pos = 0.0
    y_ticks, y_labels = [], []
    cols = [C_INK, C_POS, C_NEG]
    markers = ['D', 'o', 's']

    for lbl, pa, qa, inv in f_metrics:
        y_labels.append(lbl)
        y_ticks.append(y_pos + 0.75)
        for s_idx, idxs in enumerate([None, idx_early, idx_late]):
            dv, cl, cu = calc_d(pa, qa, idxs, inv)
            cy = y_pos + (2 - s_idx) * 0.75
            ax.errorbar(dv, cy, xerr=[[dv - cl], [cu - dv]], fmt=markers[s_idx],
                        color=cols[s_idx], ecolor=cols[s_idx], elinewidth=1.8, capsize=4, capthick=1.5,
                        markersize=7 if s_idx == 0 else 6, zorder=4)
            ax.text(cu + 0.04, cy, f'd = {dv:+.2f} [{cl:+.2f}, {cu:+.2f}]',
                    va='center', fontsize=8.6, color=cols[s_idx], fontweight='bold' if s_idx == 0 else 'normal')
        y_pos += 3.3

    ax.axvline(0.0, color=C_INK, ls='-', lw=1.2, zorder=2)
    ax.axvline(0.2, color='#94A3B8', ls='--', lw=0.9, zorder=1)
    ax.axvline(0.5, color='#94A3B8', ls='--', lw=0.9, zorder=1)
    ax.set_ylim(-0.6, y_pos + 0.1)
    ax.set_xlim(-0.72, 1.55)
    ax.text(0.2, y_pos - 0.20, 'Küçük (0.2)', ha='center', va='top', fontsize=8.4, color='#475569',
            fontweight='bold', bbox=dict(boxstyle='square,pad=0.15', facecolor='white', edgecolor='none'))
    ax.text(0.5, y_pos - 0.20, 'Orta (0.5)', ha='center', va='top', fontsize=8.4, color='#475569',
            fontweight='bold', bbox=dict(boxstyle='square,pad=0.15', facecolor='white', edgecolor='none'))

    ax.set_yticks(y_ticks)
    ax.set_yticklabels(y_labels, fontsize=10.5, fontweight='bold', color=C_INK)
    ax.set_xlabel("Standartlaştırılmış Etki Büyüklüğü (Cohen's d ve %95 Güven Aralığı)", fontsize=10, fontweight='bold', labelpad=8)

    legend_items = [
        Line2D([0], [0], color=C_INK, lw=1.8, marker='D', markersize=6.5, label='Tüm Kohort (N = 34)'),
        Line2D([0], [0], color=C_POS, lw=1.8, marker='o', markersize=6.5, label='Erken Pre-PHV (Ofset < -3.5 yıl, n = 17)'),
        Line2D([0], [0], color=C_NEG, lw=1.8, marker='s', markersize=6.2, label='Geç Pre-PHV (Ofset ≥ -3.5 yıl, n = 17)'),
    ]
    ax.legend(handles=legend_items, loc='lower center', bbox_to_anchor=(0.5, -0.22), ncol=3,
              frameon=True, facecolor='#F8FAFC', edgecolor='#CBD5E1', fontsize=8.8)
    ax.set_title('Revize Aday 4R: Biyolojik Olgunlaşma Alt Gruplarına Göre Orman Grafiği',
                 fontsize=12.2, fontweight='bold', pad=14, color=C_INK)
    plt.tight_layout()
    save_clean(fig, 'Aday_4R_Olgunlasma_Forest_Plot_Revize_TR.png', CAND_OUT_DIR)


def main():
    d = load_data()
    plot_2A_raincloud_spaghetti(d)
    plot_2B_cumming_estimation(d)
    plot_2C_nokta_aralik_yas(d)
    plot_3A_waterfall_swc(d)
    plot_3B_diverging_swc(d)
    plot_3C_cleveland_lollipop(d)
    plot_4A_subgroup_marginal(d)
    plot_4B_mechanistic_3panel(d)
    plot_revised_candidates(d)
    print('All 12 Q1 visual redesign options generated cleanly!')


if __name__ == '__main__':
    main()
