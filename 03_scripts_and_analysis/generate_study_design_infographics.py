import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch
from PIL import Image
from PIL.PngImagePlugin import PngInfo

os.environ['MPLCONFIGDIR'] = '/tmp/matplotlib'


def save_clean_png(fig, output_path, dpi=300):
    temp_path = output_path + ".tmp.png"
    fig.savefig(temp_path, dpi=dpi, facecolor="#FFFFFF", bbox_inches="tight", pad_inches=0.15)
    plt.close(fig)
    with Image.open(temp_path) as img:
        clean_meta = PngInfo()
        img.save(output_path, format="PNG", pnginfo=clean_meta, dpi=(dpi, dpi))
    if os.path.exists(temp_path):
        os.remove(temp_path)


def create_infographic(lang='TR', output_path='study_design.png'):
    fig = plt.figure(figsize=(14.5, 4.7), dpi=300)
    ax = fig.add_subplot(111)
    ax.set_xlim(0, 100)
    ax.set_ylim(16, 98)
    ax.axis('off')

    fig.patch.set_facecolor("#FFFFFF")
    ax.set_facecolor("#FFFFFF")

    # Academic Q1 Monochrome + Navy/Slate Palette
    c_navy = "#1E3A8A"       # Primary academic navy
    c_slate_hdr = "#334155"  # Secondary dark slate
    c_border = "#94A3B8"     # Crisp slate border
    c_sub_bg = "#F8FAFC"     # Subtle box fill
    c_sub_border = "#CBD5E1" # Subtle inner border
    c_text = "#0F172A"       # Pure dark text
    c_muted = "#475569"      # Muted label text

    if lang == 'TR':
        stages = [
            {
                "num": "1. AŞAMA",
                "title": "KATILIMCI SEÇİMİ",
                "sub": "Başlangıç Tarama (N = 57)",
                "hdr_color": c_slate_hdr,
                "blocks": [
                    ("Dahil Edilme Kriterleri", [
                        "• DSM-5 resmi OSB tanısı",
                        "• Erkek, 5–15 yaş aralığı",
                        "• ≥ 6 ay düzenli hareket eğitimi",
                        "  (2 gün/hf, 60 dk/seans)",
                        "• Yönerge takip ve test uyumu",
                        "• Veli onamı ve etik kurul onayı"
                    ]),
                    ("Metodolojik Kontrol", [
                        "• ≥ 6 aylık düzenli katılım",
                        "  ile bazal motor plato kontrolü",
                        "• Mirwald PHV büyüme kontrolü"
                    ])
                ]
            },
            {
                "num": "2. AŞAMA",
                "title": "ÖN TEST ÖLÇÜMLERİ",
                "sub": "1. Hafta (N = 57)",
                "hdr_color": c_slate_hdr,
                "blocks": [
                    ("Antropometri ve Olgunlaşma", [
                        "• Boy, oturma boyu, vücut ağırlığı",
                        "• VKİ (kg/m²), kulaç ve karış (cm)",
                        "• Mirwald PHV olgunlaşma offseti"
                    ]),
                    ("Performans Test Bataryası", [
                        "• Görsel reaksiyon süresi (sn)",
                        "• Rölatif kavrama kuvveti (kg/kg)",
                        "• Bilateral asimetri indeksi (%)",
                        "• Dikey sıçrama (cm) ve Sayers (W)",
                        "• Pro-Agility 5-10-5 çeviklik (sn)",
                        "• Otur-eriş esneklik testi (cm)"
                    ])
                ]
            },
            {
                "num": "3. AŞAMA",
                "title": "MİKRO-DOZ ANTRENMAN",
                "sub": "2–13. Haftalar (12 Hf, 24 Seans)",
                "hdr_color": c_navy,
                "blocks": [
                    ("Seans Yapısı (2 Gün/Hafta)", [
                        "• 60 dk rutin + 15 dk ek beceri"
                    ]),
                    ("Temel (T1–38) ve Genel (G1–41)", [
                        "• Tırmanma, direnç bandı çekişi",
                        "• Kaykay itiş, reaktif çıkışlar",
                        "• Zikzak slalom, çeviklik merdiveni",
                        "• Engel sıçrama, denge-top kontrolü"
                    ]),
                    ("Bireysel İlerleme ve Puanlama", [
                        "• İcra başarısı (yapabildi/yapamadı)",
                        "• Tekrar hacmi (1–8 tekrar)",
                        "• Hareket kalitesi (1–5 puan)"
                    ])
                ]
            },
            {
                "num": "4. AŞAMA",
                "title": "SON TEST VE AKIŞ",
                "sub": "14. Hafta (N = 34)",
                "hdr_color": c_slate_hdr,
                "blocks": [
                    ("STROBE Kohort Akışı", [
                        "• Başlangıç kohortu: N = 57",
                        "• Son testi kaçıranlar: n = 23",
                        "  (devamsızlık / lojistik nedenler)",
                        "• Tamamlayan kohort: N = 34",
                        "  - Çocukluk (5–9 yaş): n = 17",
                        "  - Ergenlik (10–15 yaş): n = 17"
                    ]),
                    ("İstatistiksel Analiz", [
                        "• Bağımlı t-testi, Cohen's d, ηp²",
                        "• SWC bireysel yanıt (0.2 × SS)",
                        "• Çoklu doğrusal regresyon (OLS)"
                    ])
                ]
            }
        ]
    else:
        stages = [
            {
                "num": "STAGE 1",
                "title": "ELIGIBILITY & ENROLMENT",
                "sub": "Initial Screening (N = 57)",
                "hdr_color": c_slate_hdr,
                "blocks": [
                    ("Inclusion Criteria", [
                        "• DSM-5 confirmed ASD diagnosis",
                        "• Male, aged 5–15 years",
                        "• ≥ 6 months movement training",
                        "  (2 days/wk, 60 min/session)",
                        "• Test instruction compliance",
                        "• Parental consent & IRB approval"
                    ]),
                    ("Methodological Control", [
                        "• ≥ 6-month baseline attendance",
                        "  establishing neuromotor plateau",
                        "• Mirwald PHV maturity control"
                    ])
                ]
            },
            {
                "num": "STAGE 2",
                "title": "PRE-TEST ASSESSMENTS",
                "sub": "Week 1 (N = 57)",
                "hdr_color": c_slate_hdr,
                "blocks": [
                    ("Anthropometry & Maturation", [
                        "• Standing & sitting height, mass",
                        "• BMI (kg/m²), arm & hand span (cm)",
                        "• Mirwald PHV maturity offset (yrs)"
                    ]),
                    ("Performance Test Battery", [
                        "• Visual reaction time (s)",
                        "• Relative handgrip strength (kg/kg)",
                        "• Bilateral asymmetry index (%)",
                        "• Vertical jump (cm) & Sayers (W)",
                        "• Pro-Agility 5-10-5 shuttle (s)",
                        "• Sit-and-reach flexibility (cm)"
                    ])
                ]
            },
            {
                "num": "STAGE 3",
                "title": "MICRO-DOSED TRAINING",
                "sub": "Weeks 2–13 (12 Wks, 24 Sessions)",
                "hdr_color": c_navy,
                "blocks": [
                    ("Session Structure (2 Days/Wk)", [
                        "• 60 min routine + 15 min skill"
                    ]),
                    ("Basic (T1–38) & General (G1–41)", [
                        "• Rope/wall climb, band traction",
                        "• Scooter push, reactive sprints",
                        "• Zigzag slalom, agility ladder",
                        "• Hurdle hops, balance ball control"
                    ]),
                    ("Individual Progression Rubric", [
                        "• Task completion (pass/fail)",
                        "• Repetition volume (1–8 reps)",
                        "• Movement quality (1–5 scale)"
                    ])
                ]
            },
            {
                "num": "STAGE 4",
                "title": "POST-TEST & STROBE FLOW",
                "sub": "Week 14 (N = 34)",
                "hdr_color": c_slate_hdr,
                "blocks": [
                    ("STROBE Cohort Retention", [
                        "• Enrolled baseline cohort: N = 57",
                        "• Missed post-testing: n = 23",
                        "  (logistical / attendance reasons)",
                        "• Per-protocol cohort: N = 34",
                        "  - Childhood (5–9 yrs): n = 17",
                        "  - Adolescence (10–15 yrs): n = 17"
                    ]),
                    ("Statistical Analysis", [
                        "• Paired t-test, Cohen's d, ηp²",
                        "• SWC responder analysis (0.2 × SD)",
                        "• Multiple linear regression (OLS)"
                    ])
                ]
            }
        ]

    cols_x = [1.5, 26.5, 51.5, 76.5]
    col_w = 22.0
    col_y = 18.0
    col_h = 79.0

    # Arrows between columns
    arrow_y = 57.5
    for i in range(3):
        x_start = cols_x[i] + col_w + 0.2
        x_end = cols_x[i+1] - 0.2
        ax.annotate(
            '', xy=(x_end, arrow_y), xytext=(x_start, arrow_y),
            arrowprops=dict(arrowstyle="-|>", color=c_navy, lw=2.2, mutation_scale=16),
            zorder=5
        )

    for idx, st in enumerate(stages):
        cx = cols_x[idx]
        # Main card outline
        card = FancyBboxPatch(
            (cx, col_y), col_w, col_h,
            boxstyle="round,pad=0.3,rounding_size=1.2",
            facecolor="#FFFFFF", edgecolor=c_border, lw=1.4, zorder=2
        )
        ax.add_patch(card)

        # Header band
        hdr_h = 15.0
        hdr = FancyBboxPatch(
            (cx, col_y + col_h - hdr_h), col_w, hdr_h,
            boxstyle="round,pad=0.3,rounding_size=1.2",
            facecolor=st["hdr_color"], edgecolor="none", zorder=3
        )
        ax.add_patch(hdr)

        ax.text(
            cx + col_w / 2, col_y + col_h - 3.8, st["num"],
            ha="center", va="center", color="#CBD5E1",
            fontsize=8.0, fontweight="bold", fontfamily="sans-serif", zorder=4
        )
        ax.text(
            cx + col_w / 2, col_y + col_h - 8.2, st["title"],
            ha="center", va="center", color="#FFFFFF",
            fontsize=9.5, fontweight="bold", fontfamily="sans-serif", zorder=4
        )
        ax.text(
            cx + col_w / 2, col_y + col_h - 12.5, st["sub"],
            ha="center", va="center", color="#E2E8F0",
            fontsize=8.2, fontfamily="sans-serif", zorder=4
        )

        # Render inner blocks dynamically
        curr_y = col_y + col_h - hdr_h - 2.5
        for b_title, b_lines in st["blocks"]:
            n_lines = len(b_lines)
            box_h = 5.8 + n_lines * 4.7
            sub_box = FancyBboxPatch(
                (cx + 1.0, curr_y - box_h), col_w - 2.0, box_h,
                boxstyle="round,pad=0.2,rounding_size=0.8",
                facecolor=c_sub_bg, edgecolor=c_sub_border, lw=1.0, zorder=3
            )
            ax.add_patch(sub_box)

            ax.text(
                cx + 1.8, curr_y - 3.2, b_title,
                ha="left", va="center", color=c_navy,
                fontsize=8.4, fontweight="bold", fontfamily="sans-serif", zorder=4
            )

            line_y = curr_y - 7.2
            for line in b_lines:
                ax.text(
                    cx + 1.8, line_y, line,
                    ha="left", va="center", color=c_text if not line.startswith("  ") else c_muted,
                    fontsize=7.9, fontfamily="sans-serif", zorder=4
                )
                line_y -= 4.5

            curr_y -= (box_h + 2.2)

    save_clean_png(fig, output_path, dpi=300)
    print(f"Grafik üretildi ({lang}): {output_path}")


if __name__ == "__main__":
    base_dir = "/Users/alik/Desktop/anti_projeler/articles_all/autism_2026_testtekrar/04_manuscript"
    artifact_dir = "/Users/alik/.gemini/antigravity/brain/b070dc08-8fc0-46d0-9bc0-8e6ac510db31"
    for d in [base_dir, artifact_dir]:
        create_infographic(lang='TR', output_path=os.path.join(d, "Sekil_1_Arastirma_Tasarimi_TR.png"))
        create_infographic(lang='EN', output_path=os.path.join(d, "Figure_1_Study_Design_EN.png"))
