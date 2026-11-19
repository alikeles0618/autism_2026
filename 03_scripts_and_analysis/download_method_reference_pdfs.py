#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Otizm 2026 Mikro-Doz Motor Beceri Antrenmanı Çalışması
Yöntem ve Bulgular Bölümü Kaynak Makaleleri (Atıf 1 - Atıf 24)
Açık Erişim (Open Access) PDF İndirici ve Eksik Kaynaklar Raporlayıcı Betik.
"""

import os
import re
import json
import subprocess
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

BASE_DIR = "/Users/alik/Desktop/anti_projeler/articles_all/autism_2026_testtekrar"
PDF_DIR = os.path.join(BASE_DIR, "Yontem_Kaynak_Makaleler_PDF")
REPORT_JSON = os.path.join(PDF_DIR, "pdf_indirme_raporu.json")
REPORT_MD = os.path.join(PDF_DIR, "README_Indirilen_ve_Eksik_Kaynaklar.md")
REPORT_DOCX = os.path.join(PDF_DIR, "Eksik_PDF_Kaynaklar_Calisma_Listesi.docx")

os.makedirs(PDF_DIR, exist_ok=True)

USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

# Reference definitions with alternative Open Access endpoints, PMC IDs, and direct URLs
SOURCES = [
    {
        "id": "Atif_01_von_Elm_2007_STROBE",
        "marker": "[Atıf Gerekli - 1]",
        "topic": "STROBE Raporlama Kılavuzu",
        "authors": "von Elm, E., Altman, D. G., Egger, M., Pocock, S. J., Gøtzsche, P. C., & Vandenbroucke, J. P.",
        "year": 2007,
        "title": "The Strengthening the Reporting of Observational Studies in Epidemiology (STROBE) Statement: Guidelines for Reporting Observational Studies",
        "journal": "PLOS Medicine / The Lancet / Annals of Internal Medicine",
        "doi": "10.1371/journal.pmed.0040296",
        "pmcid": "PMC2020495",
        "urls": [
            "https://journals.plos.org/plosmedicine/article/file?id=10.1371/journal.pmed.0040296&type=printable",
            "https://europepmc.org/backend/ptpmcrender.fcgi?accid=PMC2020495&blobtype=pdf",
            "https://www.ncbi.nlm.nih.gov/pmc/articles/PMC2020495/pdf/pmed.0040296.pdf",
            "https://annals.org/aim/article-abstract/737187"
        ]
    },
    {
        "id": "Atif_02_Slade_2016_CERT",
        "marker": "[Atıf Gerekli - 2]",
        "topic": "CERT Egzersiz Raporlama Şablonu",
        "authors": "Slade, S. C., Dionne, C. E., Underwood, M., & Buchbinder, R.",
        "year": 2016,
        "title": "Consensus on Exercise Reporting Template (CERT): Explanation and Elaboration Statement",
        "journal": "British Journal of Sports Medicine, 50(23), 1428–1437",
        "doi": "10.1136/bjsports-2016-096651",
        "pmcid": "PMC5198642",
        "urls": [
            "https://bjsm.bmj.com/content/bjsports/50/23/1428.full.pdf",
            "https://europepmc.org/backend/ptpmcrender.fcgi?accid=PMC5198642&blobtype=pdf"
        ]
    },
    {
        "id": "Atif_03A_Pan_2017_Physical_Activity_ASD",
        "marker": "[Atıf Gerekli - 3]",
        "topic": "Nöromotor Adaptasyon ve Otizmde Egzersiz (≥6 Ay)",
        "authors": "Pan, C.-Y., Chu, C.-H., Tsai, C.-L., Sung, M.-C., Huang, C.-Y., & Ma, W.-Y.",
        "year": 2017,
        "title": "The Impacts of Physical Activity Intervention on Physical and Cognitive Outcomes in Children with Autism Spectrum Disorder",
        "journal": "Autism, 21(2), 190–202",
        "doi": "10.1177/1362361316633562",
        "pmid": "27154907",
        "urls": [
            "https://journals.sagepub.com/doi/pdf/10.1177/1362361316633562",
            "https://europepmc.org/backend/ptpmcrender.fcgi?accid=PMC5405820&blobtype=pdf"
        ]
    },
    {
        "id": "Atif_03B_Sale_1988_Neural_Adaptations",
        "marker": "[Atıf Gerekli - 3]",
        "topic": "Direnç Antrenmanında Nöral Adaptasyon ve Bazal Plato",
        "authors": "Sale, D. G.",
        "year": 1988,
        "title": "Neural Adaptation to Resistance Training",
        "journal": "Medicine & Science in Sports & Exercise, 20(5 Suppl), S135–S145",
        "doi": "10.1249/00005768-198810001-00009",
        "urls": []
    },
    {
        "id": "Atif_04_Lord_2018_ASD_Lancet_Review",
        "marker": "[Atıf Gerekli - 4]",
        "topic": "DSM-5 Otizm Spektrum Bozukluğu Tanı Ölçütleri",
        "authors": "Lord, C., Elsabbagh, M., Baird, G., & Veenstra-Vanderweele, J. / APA DSM-5",
        "year": 2018,
        "title": "Autism Spectrum Disorder",
        "journal": "The Lancet, 392(10146), 508–520 / DSM-5 (American Psychiatric Association, 2013)",
        "doi": "10.1016/S0140-6736(18)31129-2",
        "pmcid": "PMC7398158",
        "urls": [
            "https://europepmc.org/backend/ptpmcrender.fcgi?accid=PMC7398158&blobtype=pdf",
            "https://www.ncbi.nlm.nih.gov/pmc/articles/PMC7398158/pdf/nihms-1608669.pdf"
        ]
    },
    {
        "id": "Atif_05_WMA_2013_Declaration_of_Helsinki",
        "marker": "[Atıf Gerekli - 5]",
        "topic": "Helsinki Bildirgesi Araştırma Etiği İlkeleri",
        "authors": "World Medical Association",
        "year": 2013,
        "title": "World Medical Association Declaration of Helsinki: Ethical Principles for Medical Research Involving Human Subjects",
        "journal": "JAMA, 310(20), 2191–2194",
        "doi": "10.1001/jama.2013.281053",
        "urls": [
            "https://jamanetwork.com/journals/jama/articlepdf/1760318/jsc130006.pdf",
            "https://www.wma.net/wp-content/uploads/2016/11/DoH-Oct2013-JAMA.pdf"
        ]
    },
    {
        "id": "Atif_06_Faul_2007_GPower3",
        "marker": "[Atıf Gerekli - 6]",
        "topic": "Duyarlılık Güç Analizi (G*Power 3)",
        "authors": "Faul, F., Erdfelder, E., Lang, A.-G., & Buchner, A.",
        "year": 2007,
        "title": "G*Power 3: A Flexible Statistical Power Analysis Program for the Social, Behavioral, and Biomedical Sciences",
        "journal": "Behavior Research Methods, 39(2), 175–191",
        "doi": "10.3758/BF03193146",
        "urls": [
            "https://link.springer.com/content/pdf/10.3758/BF03193146.pdf"
        ]
    },
    {
        "id": "Atif_07A_Ruggeri_2020_Motor_Interventions_ASD",
        "marker": "[Atıf Gerekli - 7]",
        "topic": "Otizmde Motor Beceri ve Fiziksel Aktivite Müdahaleleri",
        "authors": "Ruggeri, A., Dancel, A., Johnson, R., & Sargent, B.",
        "year": 2020,
        "title": "The Effect of Motor and Physical Activity Intervention on Motor Outcomes of Children with Autism Spectrum Disorder: A Systematic Review",
        "journal": "Autism, 24(3), 544–568",
        "doi": "10.1177/1362361319885215",
        "pmcid": "PMC7080479",
        "urls": [
            "https://europepmc.org/backend/ptpmcrender.fcgi?accid=PMC7080479&blobtype=pdf",
            "https://www.ncbi.nlm.nih.gov/pmc/articles/PMC7080479/pdf/10.1177_1362361319885215.pdf"
        ]
    },
    {
        "id": "Atif_07B_Healy_2018_Physical_Activity_ASD_Meta",
        "marker": "[Atıf Gerekli - 7]",
        "topic": "Otizmde Fiziksel Aktivite Müdahaleleri Meta-Analizi",
        "authors": "Healy, S., Nacario, A., Braithwaite, R. E., & Hopper, C.",
        "year": 2018,
        "title": "The Effect of Physical Activity Interventions on Youth with Autism Spectrum Disorder: A Meta-Analysis",
        "journal": "Autism Research, 11(6), 818–833",
        "doi": "10.1002/aur.1955",
        "urls": [
            "https://onlinelibrary.wiley.com/doi/pdf/10.1002/aur.1955"
        ]
    },
    {
        "id": "Atif_08_Guadagnoli_2004_Challenge_Point",
        "marker": "[Atıf Gerekli - 8]",
        "topic": "Bireysel Meydan Okuma Noktası Çerçevesi (Challenge Point Framework)",
        "authors": "Guadagnoli, M. A., & Lee, T. D.",
        "year": 2004,
        "title": "Challenge Point: A Framework for Conceptualizing the Effects of Various Practice Conditions in Motor Learning",
        "journal": "Journal of Motor Behavior, 36(2), 212–224",
        "doi": "10.3200/JMBR.36.2.212-224",
        "urls": [
            "https://www.tandfonline.com/doi/pdf/10.3200/JMBR.36.2.212-224"
        ]
    },
    {
        "id": "Atif_09A_Chtourou_2012_Time_of_Day_Training",
        "marker": "[Atıf Gerekli - 9]",
        "topic": "Günün Saati (Sirkadiyen Ritim) ve Egzersiz Performansı",
        "authors": "Chtourou, H., & Souissi, N.",
        "year": 2012,
        "title": "The Effect of Training at a Specific Time of Day: A Review",
        "journal": "Journal of Strength and Conditioning Research, 26(7), 1984–2005",
        "doi": "10.1519/JSC.0b013e31825770a7",
        "urls": []
    },
    {
        "id": "Atif_09B_Atkinson_1996_Circadian_Performance",
        "marker": "[Atıf Gerekli - 9]",
        "topic": "Sporda Sirkadiyen Değişimler",
        "authors": "Atkinson, G., & Reilly, T.",
        "year": 1996,
        "title": "Circadian Variation in Sports Performance",
        "journal": "Sports Medicine, 21(4), 292–312",
        "doi": "10.2165/00007256-199621040-00005",
        "urls": []
    },
    {
        "id": "Atif_10A_Nuttall_2015_BMI_Critical_Review",
        "marker": "[Atıf Gerekli - 10]",
        "topic": "Antropometri ve Vücut Kütle İndeksi (VKİ) Standartları",
        "authors": "Nuttall, F. Q. / WHO",
        "year": 2015,
        "title": "Body Mass Index: Obesity, BMI, and Health: A Critical Review",
        "journal": "Nutrition Today, 50(3), 117–128",
        "doi": "10.1097/NT.0000000000000092",
        "pmcid": "PMC4890841",
        "urls": [
            "https://europepmc.org/backend/ptpmcrender.fcgi?accid=PMC4890841&blobtype=pdf",
            "https://www.ncbi.nlm.nih.gov/pmc/articles/PMC4890841/pdf/nutoday-50-117.pdf"
        ]
    },
    {
        "id": "Atif_10B_Marfell_Jones_2012_ISAK_Standards",
        "marker": "[Atıf Gerekli - 10]",
        "topic": "ISAK Antropometrik Ölçüm Standartları",
        "authors": "Marfell-Jones, M. J., Stewart, A. D., & de Ridder, J. H.",
        "year": 2012,
        "title": "International Standards for Anthropometric Assessment",
        "journal": "International Society for the Advancement of Kinanthropometry (ISAK)",
        "doi": "",
        "urls": []
    },
    {
        "id": "Atif_11A_Malina_2015_Biological_Maturation_Athletes",
        "marker": "[Atıf Gerekli - 11]",
        "topic": "Genç Sporcularda Biyolojik Olgunlaşma Değerlendirmesi",
        "authors": "Malina, R. M., Rogol, A. D., Cumming, S. P., Coelho e Silva, M. J., & Figueiredo, A. J.",
        "year": 2015,
        "title": "Biological Maturation of Youth Athletes: Assessment and Implications",
        "journal": "British Journal of Sports Medicine, 49(13), 852–859",
        "doi": "10.1136/bjsports-2015-094623",
        "urls": [
            "https://bjsm.bmj.com/content/bjsports/49/13/852.full.pdf"
        ]
    },
    {
        "id": "Atif_11B_Mirwald_2002_Maturity_Offset_Equation",
        "marker": "[Atıf Gerekli - 11]",
        "topic": "Mirwald Zirve Boy Hızı (PHV) Olgunlaşma Offseti Denklemi",
        "authors": "Mirwald, R. L., Baxter-Jones, A. D. G., Bailey, D. A., & Beunen, G. P.",
        "year": 2002,
        "title": "An Assessment of Maturity from Anthropometric Measurements",
        "journal": "Medicine & Science in Sports & Exercise, 34(4), 689–694",
        "doi": "10.1097/00005768-200204000-00020",
        "urls": []
    },
    {
        "id": "Atif_12A_Laurenti_2026_Reaction_Time_Coordination_ASD",
        "marker": "[Atıf Gerekli - 12]",
        "topic": "Otizmli Çocuklarda Spor Antrenmanı, Reaksiyon Süresi ve İnce Motor Koordinasyon",
        "authors": "Laurenti, F., Presta, V., Compiani, M., Zobbi, G., Ilari, B., Picchi, M. P., Maré, E., Severini, F., Guarnieri, A., Mazzei, S., di Martino, O., Pozzi, G., Condello, G., & Gobbi, G.",
        "year": 2026,
        "title": "The Effects of a Sport-Based Training Program on Reaction Time and Fine Motor Coordination in Children with Autism Spectrum Disorder: A Pilot Study",
        "journal": "Sports, 14(2), 80",
        "doi": "10.3390/sports14020080",
        "urls": [
            "https://doi.org/10.3390/sports14020080"
        ]
    },
    {
        "id": "Atif_13A_Roberts_2011_Grip_Strength_Measurement_Review",
        "marker": "[Atıf Gerekli - 13]",
        "topic": "Jamar Dinamometresi ve El Kavrama Kuvveti Standardizasyonu (ASHT)",
        "authors": "Roberts, H. C., Denison, H. J., Martin, H. J., Patel, H. P., Syddall, H., Cooper, C., & Sayer, A. A.",
        "year": 2011,
        "title": "A Review of the Measurement of Grip Strength in Clinical and Epidemiological Studies: Towards a Standardised Approach",
        "journal": "Age and Ageing, 40(4), 423–429",
        "doi": "10.1093/ageing/afr051",
        "pmcid": "PMC3114461",
        "urls": [
            "https://academic.oup.com/ageing/article-pdf/40/4/423/822765/afr051.pdf",
            "https://europepmc.org/backend/ptpmcrender.fcgi?accid=PMC3114461&blobtype=pdf"
        ]
    },
    {
        "id": "Atif_14A_Wind_2010_Grip_Strength_Pediatric_Predictor",
        "marker": "[Atıf Gerekli - 14]",
        "topic": "Çocuk ve Ergenlerde El Kavrama Kuvveti ve Vücut Kütlesi",
        "authors": "Wind, A. E., Takken, T., Helders, P. J. M., & Engelbert, R. H. H.",
        "year": 2010,
        "title": "Is Grip Strength a Predictor for Total Muscle Strength in Healthy Children, Adolescents, and Young Adults?",
        "journal": "European Journal of Pediatrics, 169(3), 281–287",
        "doi": "10.1007/s00431-009-1010-4",
        "pmcid": "PMC2810359",
        "urls": [
            "https://link.springer.com/content/pdf/10.1007/s00431-009-1010-4.pdf",
            "https://europepmc.org/backend/ptpmcrender.fcgi?accid=PMC2810359&blobtype=pdf"
        ]
    },
    {
        "id": "Atif_14B_Jaric_2002_Muscle_Strength_Body_Size_Scaling",
        "marker": "[Atıf Gerekli - 14]",
        "topic": "Kas Kuvveti Testlerinde Vücut Boyutuna Göre Normalizasyon",
        "authors": "Jaric, S.",
        "year": 2002,
        "title": "Muscle Strength Testing: Use of Normalization for Body Size",
        "journal": "Sports Medicine, 32(10), 615–631",
        "doi": "10.2165/00007256-200232100-00002",
        "urls": []
    },
    {
        "id": "Atif_15A_Bishop_2018_Interlimb_Asymmetries",
        "marker": "[Atıf Gerekli - 15]",
        "topic": "Bilateral Asimetri İndeksi Hesaplama Metotları",
        "authors": "Bishop, C., Read, P., Lake, J., Chavda, S., & Turner, A.",
        "year": 2018,
        "title": "Interlimb Asymmetries: Understanding How to Calculate Differences from Bilateral and Unilateral Tests",
        "journal": "Strength and Conditioning Journal, 40(4), 1–6",
        "doi": "10.1519/SSC.0000000000000371",
        "urls": []
    },
    {
        "id": "Atif_15B_Impellizzeri_2007_Bilateral_Strength_Asymmetry",
        "marker": "[Atıf Gerekli - 15]",
        "topic": "Dikey Sıçrama ve İki Yanlı Kuvvet Asimetrisi",
        "authors": "Impellizzeri, F. M., Rampinini, E., Maffiuletti, N., & Marcora, S. M.",
        "year": 2007,
        "title": "A Vertical Jump Force Test for Assessing Bilateral Strength Asymmetry in Athletes",
        "journal": "Medicine & Science in Sports & Exercise, 39(11), 2044–2050",
        "doi": "10.1249/mss.0b013e31814fb55c",
        "urls": []
    },
    {
        "id": "Atif_16A_Markovic_2004_Squat_CMJ_Reliability",
        "marker": "[Atıf Gerekli - 16]",
        "topic": "Karşıt Hareketli Aktif Dikey Sıçrama (CMJ) Güvenirliği",
        "authors": "Markovic, G., Dizdar, D., Jukic, I., & Cardinale, M.",
        "year": 2004,
        "title": "Reliability and Factorial Validity of Squat and Countermovement Jump Tests",
        "journal": "Journal of Strength and Conditioning Research, 18(3), 551–555",
        "doi": "10.1519/1533-4287(2004)18<551:RAFVOS>2.0.CO;2",
        "urls": []
    },
    {
        "id": "Atif_16B_Sayers_1999_Jump_Power_Equations",
        "marker": "[Atıf Gerekli - 16]",
        "topic": "Dikey Sıçramadan Sayers Zirve Mekanik Güç Denklemi",
        "authors": "Sayers, S. P., Harackiewicz, D. V., Harman, E. A., Frykman, P. N., & Rosenstein, M. T.",
        "year": 1999,
        "title": "Cross-Validation of Three Jump Power Equations",
        "journal": "Medicine & Science in Sports & Exercise, 31(4), 572–577",
        "doi": "10.1097/00005768-199904000-00013",
        "urls": []
    },
    {
        "id": "Atif_17_Stewart_2014_Change_of_Direction_Speed",
        "marker": "[Atıf Gerekli - 17]",
        "topic": "Pro-Agility (5-10-5) Çeviklik ve Yön Değiştirme Hızı",
        "authors": "Stewart, P. F., Turner, A. N., & Miller, S. C.",
        "year": 2014,
        "title": "Reliability, Factorial Validity, and Interrelationships of Five Commonly Used Change of Direction Speed Tests",
        "journal": "Scandinavian Journal of Medicine & Science in Sports, 24(3), 500–506",
        "doi": "10.1111/sms.12019",
        "urls": [
            "https://onlinelibrary.wiley.com/doi/pdf/10.1111/sms.12019"
        ]
    },
    {
        "id": "Atif_18_Castro_Pinero_2009_Sit_and_Reach_Validity",
        "marker": "[Atıf Gerekli - 18]",
        "topic": "Otur-Eriş Esneklik Testinin Çocuk ve Gençlerde Geçerliği",
        "authors": "Castro-Piñero, J., Chillón, P., Ortega, F. B., Montesinos, J. L., Sjöström, M., & Ruiz, J. R.",
        "year": 2009,
        "title": "Criterion-Related Validity of Sit-and-Reach and Modified Sit-and-Reach Test for Estimating Hamstring Flexibility in Children and Adolescents Aged 6–17 Years",
        "journal": "International Journal of Sports Medicine, 30(9), 658–662",
        "doi": "10.1055/s-0029-1224175",
        "urls": []
    },
    {
        "id": "Atif_19_Razali_2011_Shapiro_Wilk_Power",
        "marker": "[Atıf Gerekli - 19]",
        "topic": "Shapiro-Wilk Normallik Testi Güç Karşılaştırmaları",
        "authors": "Razali, N. M., & Wah, Y. B. / Shapiro & Wilk (1965)",
        "year": 2011,
        "title": "Power Comparisons of Shapiro-Wilk, Kolmogorov-Smirnov, Lilliefors and Anderson-Darling Tests",
        "journal": "Journal of Statistical Modeling and Analytics, 2(1), 21–33",
        "doi": "",
        "urls": [
            "http://education.um.edu.my/images/education/doc/Penyelidikan/Power%20comparisons%20of%20shapiro-wilk,%20kolmogorov-smirnov,%20lilliefors%20and%20anderson-darling%20tests.pdf",
            "https://www.nbi.dk/~waclaw/RazaliWah2011.pdf"
        ]
    },
    {
        "id": "Atif_20_Lakens_2013_Effect_Sizes_Primer",
        "marker": "[Atıf Gerekli - 20]",
        "topic": "Etki Büyüklüğü (Cohen's d ve ηp²) Hesaplama Rehberi",
        "authors": "Lakens, D. / Cohen, J.",
        "year": 2013,
        "title": "Calculating and Reporting Effect Sizes to Facilitate Cumulative Science: A Practical Primer for t-tests and ANOVAs",
        "journal": "Frontiers in Psychology, 4, 863",
        "doi": "10.3389/fpsyg.2013.00863",
        "pmcid": "PMC3840331",
        "urls": [
            "https://www.frontiersin.org/articles/10.3389/fpsyg.2013.00863/pdf",
            "https://europepmc.org/backend/ptpmcrender.fcgi?accid=PMC3840331&blobtype=pdf"
        ]
    },
    {
        "id": "Atif_21A_Hopkins_2004_Interpret_Changes_Athletic_Test",
        "marker": "[Atıf Gerekli - 21]",
        "topic": "En Küçük Anlamlı Değişim (SWC = 0.20 × SS_ön)",
        "authors": "Hopkins, W. G.",
        "year": 2004,
        "title": "How to Interpret Changes in an Athletic Performance Test",
        "journal": "Sportscience, 8, 1–7",
        "doi": "",
        "urls": [
            "https://www.sportsci.org/jour/04/wghtests.pdf"
        ]
    },
    {
        "id": "Atif_21B_Swinton_2018_Individual_Response_Framework",
        "marker": "[Atıf Gerekli - 21]",
        "topic": "Bireysel Yanıtlayıcı (Responder) İstatistiksel Çerçevesi",
        "authors": "Swinton, P. A., Hemingway, B. S., Saunders, B., Gualano, B., & Dolan, E.",
        "year": 2018,
        "title": "A Statistical Framework to Interpret Individual Response to Intervention: Paving the Way for Personalized Nutrition and Exercise Prescription",
        "journal": "Frontiers in Nutrition, 5, 41",
        "doi": "10.3389/fnut.2018.00041",
        "pmcid": "PMC6008442",
        "urls": [
            "https://www.frontiersin.org/articles/10.3389/fnut.2018.00041/pdf",
            "https://europepmc.org/backend/ptpmcrender.fcgi?accid=PMC6008442&blobtype=pdf"
        ]
    },
    {
        "id": "Atif_22_Tabachnick_2019_Multivariate_Statistics",
        "marker": "[Atıf Gerekli - 22]",
        "topic": "Çoklu Doğrusal Regresyon (OLS) ve Çok Değişkenli İstatistik",
        "authors": "Tabachnick, B. G., & Fidell, L. S. / Aiken & West (1991)",
        "year": 2019,
        "title": "Using Multivariate Statistics (7th ed.) / Multiple Regression",
        "journal": "Pearson / Sage Publications",
        "doi": "",
        "urls": []
    },
    {
        "id": "Atif_23A_Lloyd_2012_Youth_Physical_Development",
        "marker": "[Atıf Gerekli - 23]",
        "topic": "Gençlerde Fiziksel Gelişim Modeli ve Nöromotor Olgunlaşma",
        "authors": "Lloyd, R. S., & Oliver, J. L.",
        "year": 2012,
        "title": "The Youth Physical Development Model: A New Approach to Long-Term Athletic Development",
        "journal": "Strength and Conditioning Journal, 34(3), 61–72",
        "doi": "10.1519/SSC.0b013e31825760ea",
        "urls": [
            "https://journals.lww.com/nsca-scj/Fulltext/2012/06000/The_Youth_Physical_Development_Model__A_New.9.aspx"
        ]
    },
    {
        "id": "Atif_23B_Kail_1991_Developmental_Processing_Speed",
        "marker": "[Atıf Gerekli - 23]",
        "topic": "Çocukluk ve Ergenlikte İşlemleme Hızının Gelişimi",
        "authors": "Kail, R.",
        "year": 1991,
        "title": "Developmental Change in Speed of Processing During Childhood and Adolescence",
        "journal": "Psychological Bulletin, 109(3), 490–501",
        "doi": "10.1037/0033-2909.109.3.490",
        "urls": []
    },
    {
        "id": "Atif_24A_Hager_Ross_2002_Grip_Strength_Norms_Children",
        "marker": "[Atıf Gerekli - 24]",
        "topic": "Çocuklarda El Kavrama Kuvveti Normları ve El Boyutu",
        "authors": "Hager-Ross, C., & Rösblad, B.",
        "year": 2002,
        "title": "Norms for Grip Strength in Children Aged 4–16 Years",
        "journal": "Acta Paediatrica, 91(6), 617–625",
        "doi": "10.1111/j.1651-2227.2002.tb03290.x",
        "urls": [
            "https://onlinelibrary.wiley.com/doi/pdf/10.1111/j.1651-2227.2002.tb03290.x"
        ]
    },
    {
        "id": "Atif_24B_Nicolay_2005_Grip_Strength_Anthropometry",
        "marker": "[Atıf Gerekli - 24]",
        "topic": "El Kavrama Kuvveti ve Antropometrik Kaldıraç İlişkisi",
        "authors": "Nicolay, C. W., & Walker, A. L.",
        "year": 2005,
        "title": "Grip Strength and Endurance: Influences of Anthropometric Variation, Hand Dominance, and Gender",
        "journal": "International Journal of Industrial Ergonomics, 35(7), 605–618",
        "doi": "10.1016/j.ergon.2005.01.007",
        "urls": [
            "https://www.sciencedirect.com/science/article/pii/S016981410500020X"
        ]
    }
]


def try_fetch_pdf_curl(url, out_path):
    try:
        cmd = [
            "/usr/bin/curl",
            "-k",
            "-sL",
            "--max-time", "15",
            "-A", USER_AGENT,
            "-o", out_path,
            url,
        ]
        subprocess.run(cmd, check=False, timeout=18)
        if os.path.exists(out_path) and os.path.getsize(out_path) > 10000:
            with open(out_path, "rb") as f:
                header = f.read(1024)
            idx = header.find(b"%PDF-")
            if idx != -1:
                if idx > 0:
                    with open(out_path, "rb") as f:
                        full = f.read()
                    with open(out_path, "wb") as f:
                        f.write(full[idx:])
                return True
        if os.path.exists(out_path):
            os.remove(out_path)
    except Exception:
        if os.path.exists(out_path):
            os.remove(out_path)
    return False


def set_run_font(run, size_pt=11, bold=False, italic=False, color=RGBColor(0, 0, 0)):
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = color


def set_black_apa_borders(table):
    tblPr = table._tbl.tblPr
    tblBorders = tblPr.first_child_found_in('w:tblBorders')
    if tblBorders is not None:
        tblPr.remove(tblBorders)
    tblBorders = OxmlElement('w:tblBorders')
    for border_name, val, sz in [
        ('top', 'single', '8'),
        ('bottom', 'single', '8'),
        ('left', 'none', '0'),
        ('right', 'none', '0'),
        ('insideH', 'none', '0'),
        ('insideV', 'none', '0'),
    ]:
        b = OxmlElement(f'w:{border_name}')
        b.set(qn('w:val'), val)
        b.set(qn('w:sz'), sz)
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), '000000')
        tblBorders.append(b)
    tblLook = tblPr.first_child_found_in('w:tblLook')
    if tblLook is not None:
        tblLook.addprevious(tblBorders)
    else:
        tblPr.append(tblBorders)


def set_cell_border_bottom(cell, sz='6'):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in('w:tcBorders')
    if tcBorders is not None:
        tcPr.remove(tcBorders)
    tcBorders = OxmlElement('w:tcBorders')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), sz)
    bottom.set(qn('w:space'), '0')
    bottom.set(qn('w:color'), '000000')
    tcBorders.append(bottom)
    tcPr.append(tcBorders)


def set_cell_margins(cell, top=60, bottom=60, left=80, right=80):
    tcPr = cell._tc.get_or_add_tcPr()
    old_mar = tcPr.first_child_found_in('w:tcMar')
    if old_mar is not None:
        tcPr.remove(old_mar)
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)


def build_missing_docx(downloaded, missing):
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.9)
        section.bottom_margin = Inches(0.9)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.9)

    # Title
    p_t = doc.add_paragraph()
    p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t.paragraph_format.space_before = Pt(0)
    p_t.paragraph_format.space_after = Pt(8)
    r_t = p_t.add_run("YÖNTEM VE BULGULAR BÖLÜMLERİ EKSİK PDF KAYNAKLAR ÇALIŞMA LİSTESİ")
    set_run_font(r_t, size_pt=13, bold=True)

    # Intro
    p_i = doc.add_paragraph()
    p_i.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_i.paragraph_format.space_after = Pt(10)
    p_i.paragraph_format.line_spacing = 1.35
    r_i = p_i.add_run(
        f"Bu belge, makalenin Yöntem ve Bulgular bölümlerinde atıf yapılan kaynaklardan açık erişim (Open Access) "
        f"aracılığıyla otomatik olarak indirilen {len(downloaded)} adet tam metin PDF dışında kalan ve üniversite kütüphanesi, "
        f"kurumsal abonelik veya ResearchGate üzerinden manuel olarak indirilerek 'Yontem_Kaynak_Makaleler_PDF' klasörüne "
        f"eklenmesi gereken {len(missing)} kaynağı listelemektedir."
    )
    set_run_font(r_i, size_pt=11)

    # Table of missing items
    p_ht = doc.add_paragraph()
    p_ht.paragraph_format.space_before = Pt(8)
    p_ht.paragraph_format.space_after = Pt(6)
    r_ht = p_ht.add_run("1. Manuel İndirilmesi Gereken Kaynaklar ve İndirme Bilgileri")
    set_run_font(r_ht, size_pt=12, bold=True)

    headers = ["Atıf No", "Önerilen Dosya Adı", "Künye ve DOI", "Doğrudan Erişim Linki"]
    table = doc.add_table(rows=1 + len(missing), cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_black_apa_borders(table)
    col_widths = [1.10, 1.70, 2.30, 1.60]

    for j, h_text in enumerate(headers):
        cell = table.rows[0].cells[j]
        cell.width = Inches(col_widths[j])
        set_cell_border_bottom(cell, sz='6')
        set_cell_margins(cell)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(h_text)
        set_run_font(r, size_pt=9.5, bold=True)

    for i, item in enumerate(missing):
        row = table.rows[i + 1]
        c0 = row.cells[0]
        c1 = row.cells[1]
        c2 = row.cells[2]
        c3 = row.cells[3]

        for idx, cell in enumerate([c0, c1, c2, c3]):
            cell.width = Inches(col_widths[idx])
            set_cell_margins(cell)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.15

        r0 = c0.paragraphs[0].add_run(item["marker"])
        set_run_font(r0, size_pt=9.0, bold=True, color=RGBColor(192, 0, 0))

        r1 = c1.paragraphs[0].add_run(f"{item['id']}.pdf")
        set_run_font(r1, size_pt=9.0, bold=True)

        cite_str = f"{item['authors']} ({item['year']}). {item['title']}. {item['journal']}."
        if item.get("doi"):
            cite_str += f" DOI: {item['doi']}"
        r2 = c2.paragraphs[0].add_run(cite_str)
        set_run_font(r2, size_pt=8.5)

        link_str = f"https://doi.org/{item['doi']}" if item.get("doi") else "Kitap / Manuel Arama"
        r3 = c3.paragraphs[0].add_run(link_str)
        set_run_font(r3, size_pt=8.5, color=RGBColor(0, 51, 102))

    # Section 2: Instructions
    p_inst_h = doc.add_paragraph()
    p_inst_h.paragraph_format.space_before = Pt(14)
    p_inst_h.paragraph_format.space_after = Pt(4)
    r_ih = p_inst_h.add_run("2. Klasöre Ekleme Talimatı")
    set_run_font(r_ih, size_pt=12, bold=True)

    instructions = [
        "1. Yukarıdaki tabloda yer alan DOI linkine tıklayarak veya üniversite kütüphane veritabanınızdan (ScienceDirect, Wiley, Springer, Human Kinetics, Ovid LWW) tam metin PDF'i indirin.",
        "2. İndirdiğiniz PDF dosyasını tabloda belirtilen 'Önerilen Dosya Adı' (Örn: Atif_08_Guadagnoli_2004_Challenge_Point.pdf) ile yeniden adlandırın.",
        "3. Yeniden adlandırılan dosyayı doğrudan 'autism_2026_testtekrar/Yontem_Kaynak_Makaleler_PDF' klasörünün içine yapıştırın.",
        "4. Bu işlem tamamlandığında çalışmanın tüm yöntem kaynakları tek bir arşiv çatısı altında eksiksiz toplanmış olacaktır."
    ]
    for ins in instructions:
        p_ins = doc.add_paragraph()
        p_ins.paragraph_format.space_before = Pt(1)
        p_ins.paragraph_format.space_after = Pt(3)
        p_ins.paragraph_format.line_spacing = 1.2
        r_ins = p_ins.add_run(ins)
        set_run_font(r_ins, size_pt=10.0)

    # Metadata scrub
    cp = doc.core_properties
    cp.author = ""
    cp.last_modified_by = ""
    cp.comments = ""
    cp.title = ""
    doc.save(REPORT_DOCX)
    print("Saved Eksik_PDF_Kaynaklar_Calisma_Listesi.docx")


def main():
    downloaded = []
    missing = []

    print(f"--- YÖNTEM KAYNAKLARI PDF İNDİRME İŞLEMİ BAŞLATILDI (Toplam: {len(SOURCES)} Kaynak) ---")

    for item in SOURCES:
        fname = f"{item['id']}.pdf"
        out_path = os.path.join(PDF_DIR, fname)

        # Check if already downloaded/copied
        if os.path.exists(out_path) and os.path.getsize(out_path) > 5000:
            sz_kb = round(os.path.getsize(out_path) / 1024, 1)
            print(f"[MEVCUT] {item['marker']:<20} -> {fname} ({sz_kb} KB)")
            downloaded.append({
                "marker": item["marker"],
                "id": item["id"],
                "filename": fname,
                "authors": item["authors"],
                "year": item["year"],
                "title": item["title"],
                "journal": item["journal"],
                "doi": item.get("doi", ""),
                "size_kb": sz_kb,
                "status": "mevcut"
            })
            continue

        # Try downloading via URLs
        success = False
        for url in item.get("urls", []):
            if not url.startswith("http"):
                continue
            print(f"  Deniyor: {url[:70]}...")
            if try_fetch_pdf_curl(url, out_path):
                sz_kb = round(os.path.getsize(out_path) / 1024, 1)
                print(f"[İNDİRİLDİ] {item['marker']:<20} -> {fname} ({sz_kb} KB)")
                downloaded.append({
                    "marker": item["marker"],
                    "id": item["id"],
                    "filename": fname,
                    "authors": item["authors"],
                    "year": item["year"],
                    "title": item["title"],
                    "journal": item["journal"],
                    "doi": item.get("doi", ""),
                    "size_kb": sz_kb,
                    "status": "indirildi"
                })
                success = True
                break

        if not success:
            print(f"[EKSİK / MANUEL] {item['marker']:<20} -> {fname} (DOI: {item.get('doi', 'Yok')})")
            missing.append(item)

    # Save JSON report
    report_data = {
        "toplam_kaynak": len(SOURCES),
        "indirilen_veya_mevcut_sayisi": len(downloaded),
        "eksik_sayisi": len(missing),
        "indirilenler": downloaded,
        "eksikler": missing
    }
    with open(REPORT_JSON, "w", encoding="utf-8") as f:
        json.dump(report_data, f, ensure_ascii=False, indent=2)
    print(f"Saved JSON report: {REPORT_JSON}")

    # Build Markdown report
    md_lines = [
        "# Yöntem ve Bulgular Bölümü Kaynak Makaleler — PDF İndirme Raporu\n",
        f"Bu klasör (`autism_2026_testtekrar/Yontem_Kaynak_Makaleler_PDF`), makalede atıf yapılan **`[Atıf Gerekli - 1]` – `[Atıf Gerekli - 24]`** "
        f"arasındaki kaynakların tam metin PDF dosyalarını ve manuel eklenecek kaynaklar listesini içerir.\n",
        f"**Özet İstatistik:**",
        f"* **Toplam Taranan Kaynak Adayı:** {len(SOURCES)}",
        f"* **Klasörde Mevcut / İndirilen Tam Metin PDF:** {len(downloaded)}",
        f"* **Manuel İndirilmesi Gereken (Abonelikli / Kitap):** {len(missing)}\n",
        "---\n",
        f"## 1. Klasörde Mevcut Olan Tam Metin PDF Dosyaları ({len(downloaded)} Adet)\n",
        "| Atıf No | Dosya Adı | Kaynak Künyesi ve DOI | Boyut |",
        "| :--- | :--- | :--- | :--- |"
    ]
    for d in downloaded:
        doi_str = f" `DOI: {d['doi']}`" if d.get("doi") else ""
        md_lines.append(f"| **{d['marker']}** | `{d['filename']}` | {d['authors']} ({d['year']}). {d['title']}. *{d['journal']}*.{doi_str} | {d['size_kb']} KB |")

    md_lines.extend([
        "\n---\n",
        f"## 2. Üniversite / Kurumsal Abonelik ile Manuel İndirilecek Kaynaklar ({len(missing)} Adet)\n",
        "Aşağıdaki kaynaklar abonelik duvarı (Paywall) veya kitap bölümü olması nedeniyle açık erişimden otomatik indirilememiştir. "
        "Üniversite kütüphanesi erişimiyle indirilip bu klasöre eklenebilir:\n",
        "| Atıf No | Önerilen Dosya Adı | Kaynak Künyesi | Doğrudan Erişim Linki |",
        "| :--- | :--- | :--- | :--- |"
    ])
    for m in missing:
        doi_link = f"[https://doi.org/{m['doi']}](https://doi.org/{m['doi']})" if m.get("doi") else "Kitap / Manuel Veritabanı"
        doi_txt = f" DOI: {m['doi']}" if m.get("doi") else ""
        md_lines.append(f"| **{m['marker']}** | `{m['id']}.pdf` | {m['authors']} ({m['year']}). {m['title']}. *{m['journal']}*.{doi_txt} | {doi_link} |")

    md_lines.append("\n---\n*Rapor Otomatik Olarak Üretilmiştir.*")
    with open(REPORT_MD, "w", encoding="utf-8") as f:
        f.write("\n".join(md_lines) + "\n")
    print(f"Saved Markdown report: {REPORT_MD}")

    # Build Word checklist
    build_missing_docx(downloaded, missing)


if __name__ == "__main__":
    main()
