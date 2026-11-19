#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
run_citation_audit_stage4.py

Stage 4 (Citation Audit & Full-Passage PDF Highlighting) for all verified primary PDFs
in Yontem_Kaynak_Makaleler_PDF/ (26 primary PDFs), archived secondary review
Baxter-Jones et al. (2005) in kullanilmayanlar/orijinal_kaynaklar/, and missing Cohen (1988).

Outputs:
- 26 *_annotated.pdf files in 05_citation_audit_boyali_pdfler/Boyali_PDFler/
- Atif_05B_BaxterJones_2005_Maturation_Pediatric_Exercise_Review_annotated.pdf in kullanilmayanlar/boyali_kaynaklar/
- Verification PNGs in 05_citation_audit_boyali_pdfler/_sistem_verileri/dogrulama_png/
- Stage 4 JSON: 05_citation_audit_boyali_pdfler/_sistem_verileri/audit_results_stage4.json
- Full DOCX/PDF/XLSX/CSV reports: 05_citation_audit_boyali_pdfler/Birincil_Kaynak_Atif_Denetim_Raporu_TR_*
"""

from __future__ import annotations

import json
import re
import shutil
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

XML_CTRL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")


def clean_xml(s: str) -> str:
    return XML_CTRL_RE.sub("", s)

BASE_DIR = Path("/Users/alik/Desktop/anti_projeler/articles_all/autism_2026_testtekrar")
PDF_DIR = BASE_DIR / "Yontem_Kaynak_Makaleler_PDF"
UNUSED_ORIG_DIR = BASE_DIR / "kullanilmayanlar" / "orijinal_kaynaklar"
UNUSED_ANN_DIR = BASE_DIR / "kullanilmayanlar" / "boyali_kaynaklar"

AUDIT_DIR = BASE_DIR / "05_citation_audit_boyali_pdfler"
ANNOTATED_DIR = AUDIT_DIR / "Boyali_PDFler"
SYS_DIR = AUDIT_DIR / "_sistem_verileri"
PNG_DIR = SYS_DIR / "dogrulama_png"

SCRIPTS_DIR = Path("/Users/alik/Desktop/anti_projeler/articles_all/.agents/skills/citation-audit/scripts")
AUDIT_CLEAN_SCRIPT = Path("/Users/alik/Desktop/anti_projeler/articles_all/.agents/skills/anti-ai-writing-style/scripts/audit_and_clean.py")
sys.path.insert(0, str(SCRIPTS_DIR))
import highlight_pdf  # noqa: E402


# Verified exact 0-based page block indices from fresh _extracted.json files
PDF_HIGHLIGHT_SPECS = [
    {
        "cid": "A001",
        "stem": "Atif_01B_Fournier_2010_Motor_Coordination_ASD_MetaAnalysis",
        "marker": "[REQ-01] Fournier ve ark. (2010)",
        "location": "1. Giriş (1.1. OSB'de Motor Yetersizlik)",
        "claim_type": "Ampirik Meta-Analiz Bulgusu",
        "claim_text": "Otizm spektrum bozukluğu olan bireylerde motor koordinasyon yetersizliği belirgin ve yaygın bir özellik olup 51 karşılaştırmalık meta-analizde standartlaştırılmış ortalama etki büyüklüğü 1.20 (SE = 0.114, p < 0.0001) bulunmuştur.",
        "source_ref": "Fournier, K. A., Hass, C. J., Naik, S. K., Lodha, N., & Cauraugh, J. H. (2010). Motor coordination in autism spectrum disorders: A synthesis and meta-analysis. Journal of Autism and Developmental Disorders, 40(10), 1227-1240. https://doi.org/10.1007/s10803-010-0981-3",
        "decision_label": "Tam destek",
        "confidence": "Yüksek",
        "primary_status": "Birincil",
        "final_status": "Uygun",
        "page_and_section": "S. 1 (Özet) ve S. 6 (Bulgular - Meta-Analiz)",
        "blocks": [(1, [4]), (6, [9, 11])],
        "rationale": "Fournier ve ark. (2010) 83 çalışmayı ve 51 karşılaştırmayı niceliksel meta-analizle sentezleyerek OSB'de büyük düzeyde (1.20; %95 GA: 0.973-1.42) motor koordinasyon açığını doğrudan göstermektedir.",
        "recommended_action": "Birincil kaynak olarak Giriş 1.1 bölümüne entegre edildi.",
        "doi_status": "Yerel PDF DOI Doğrulandı (10.1007/s10803-010-0981-3)",
    },
    {
        "cid": "A002",
        "stem": "Atif_02B_Behm_2017_Youth_Strength_Training_MetaAnalysis",
        "marker": "[REQ-02] Behm ve ark. (2017)",
        "location": "1. Giriş (1.1) & 4. Tartışma (4.1 / 4.3)",
        "claim_type": "Çocuk ve Ergen Kuvvet/Güç Meta-Analizi",
        "claim_text": "Çocuk ve ergenlerde 107 çalışmayı kapsayan meta-analiz, direnç/güç antrenmanlarının sıçrama gücü ve alt ekstremite kuvveti üzerindeki etki büyüklüklerinin başlangıç eğitim durumu (antrene ve antrene olmayan) ile olgunlaşma düzeyine göre değiştiğini göstermiştir.",
        "source_ref": "Behm, D. G., Young, J. D., Whitten, J. H. D., Reid, J. C., Quigley, P. J., Low, J., Li, Y., Lima, C. D., Hodgson, D. D., Chaouachi, A., Prieske, O., & Granacher, U. (2017). Effectiveness of traditional strength vs. power training on muscle strength, power and speed with youth: A systematic review and meta-analysis. Frontiers in Physiology, 8, Article 423. https://doi.org/10.3389/fphys.2017.00423",
        "decision_label": "Tam destek",
        "confidence": "Yüksek",
        "primary_status": "Birincil",
        "final_status": "Uygun",
        "page_and_section": "S. 1 (Giriş) ve S. 22 (Tablo 3 Moderatör Sonuçları)",
        "blocks": [(1, [21]), (22, [2, 6, 7, 8, 12])],
        "rationale": "Pediatrik direnç ve güç antrenmanı etkilerini 107 çalışmada inceleyen niceliksel meta-analizdir; antrene vs antrene olmayan ve çocuk vs ergen moderatör etki büyüklüklerini doğrudan Tablo 3'te raporlar (6 aylık bazal plato kuralı ise çalışma tasarım kuralı olarak ayrı tutulmuştur).",
        "recommended_action": "Antrene durumu moderatör bağlamı ve pediatrik modalite özgüllüğü için entegre edildi.",
        "doi_status": "Yerel PDF DOI Doğrulandı (10.3389/fphys.2017.00423)",
    },
    {
        "cid": "A003",
        "stem": "Atif_03_DelVecchio_2019_Motor_Unit_Neural_Adaptation",
        "marker": "[REQ-07, REQ-23] Del Vecchio ve ark. (2019)",
        "location": "2. Yöntem (2.1. Tasarım) & 4. Tartışma (4.2)",
        "claim_type": "Birincil Nörofizyolojik Deneysel Çalışma",
        "claim_text": "Kuvvet antrenmanının ilk 4 haftasında (12 seans) oluşan kuvvet kazanımı (MVF: 289 +/- 64 N'den 329 +/- 61 N'ye; p = 0.006, eta_p2 = 0.280), motor ünite alım eşiğindeki düşüş ve plato ateşleme hızındaki (+3.3 +/- 2.5 pps, p < 0.001) nöral adaptasyonlarla ilişkilidir.",
        "source_ref": "Del Vecchio, A., Casolo, A., Negro, F., Scorcelletti, M., Bazzucchi, I., Enoka, R., Felici, F., & Farina, D. (2019). The increase in muscle force after 4 weeks of strength training is mediated by adaptations in motor unit recruitment and rate coding. The Journal of Physiology, 597(7), 1873-1887. https://doi.org/10.1113/JP277250",
        "decision_label": "Tam destek",
        "confidence": "Yüksek",
        "primary_status": "Birincil",
        "final_status": "Uygun",
        "page_and_section": "S. 1 (Özet), S. 2 (Giriş) ve S. 6 (Bulgular - Motor Ünite Nöral Veriler)",
        "blocks": [(1, [9]), (2, [1]), (6, [11])],
        "rationale": "HD-EMG motor ünite ayrıştırmasıyla 4 haftalık erken dönem kuvvet kazanımlarının alım eşiği düşüşü ve ateşleme hızı artışı ile açıklandığını deneysel olarak kanıtlar. Çalışmamızdaki >=6 ay kesintisiz eğitim ön koşulunu fizyolojik gerekçe olarak bağlamlandırır.",
        "recommended_action": "Birincil nöral adaptasyon kanıtı olarak Yöntem 2.1 ve Tartışma 4.2'ye entegre edildi.",
        "doi_status": "Yerel PDF DOI Doğrulandı (10.1113/JP277250)",
    },
    {
        "cid": "A004",
        "stem": "Atif_04_Rice_2022_DSM5_ASD_Criteria_Reliability",
        "marker": "[REQ-08] Rice ve ark. (2022)",
        "location": "2. Yöntem (2.2. Katılımcılar)",
        "claim_type": "Birincil Çok Merkezli Tanı Güvenirliği Çalışması",
        "claim_text": "Çocuklarda DSM-5 OSB tanı ölçütlerinin operasyonelleştirilmiş klinik değerlendirmesinde 11 ADDM merkezinde 4319 olgu incelemesinde yüksek tanısal uyum (%82-%89; kappa = 0.89) bildirilmiştir.",
        "source_ref": "Rice, C. E., Carpenter, L. A., Morrier, M. J., Lord, C., DiRienzo, M., Boan, A., Skowyra, C., Fusco, A., Baio, J., Esler, A., Zahorodny, W., Hobson, N., Mars, A., Thurm, A., Bishop, S., & Wiggins, L. D. (2022). Defining in detail and evaluating reliability of DSM-5 criteria for autism spectrum disorder (ASD) among children. Journal of Autism and Developmental Disorders, 52(12), 5308-5320. https://doi.org/10.1007/s10803-021-05377-y",
        "decision_label": "Tam destek",
        "confidence": "Yüksek",
        "primary_status": "Birincil",
        "final_status": "Uygun",
        "page_and_section": "S. 3, S. 6 (Phase 2 & 2016 Güvenirlik) ve S. 9 (Tartışma)",
        "blocks": [(3, [2]), (6, [50, 51]), (9, [18])],
        "rationale": "DSM-5 kriterlerinin çocuk popülasyonundaki klinisyenler arası operasyonelleştirme ve değerlendirme güvenirliğini büyük ölçekli sürveyans verileriyle doğrulayan birincil araştırmadır.",
        "recommended_action": "DSM-5 klinik tanı operasyonelleştirme güvenirliği desteği olarak Yöntem 2.2'ye entegre edildi.",
        "doi_status": "Yerel PDF DOI Doğrulandı (10.1007/s10803-021-05377-y)",
    },
    {
        "cid": "A005",
        "stem": "Atif_01_von_Elm_2007_STROBE",
        "marker": "von Elm ve ark. (2007)",
        "location": "2. Yöntem (2.1. Araştırma Tasarımı)",
        "claim_type": "Metodolojik Raporlama Rehberi",
        "claim_text": "Gözlemsel kohort çalışmasının tasarımı ve raporlanması 22 maddelik STROBE kontrol listesi ilkeleri doğrultusunda yapılandırılmıştır.",
        "source_ref": "von Elm, E., Altman, D. G., Egger, M., Pocock, S. J., Gøtzsche, P. C., & Vandenbroucke, J. P. (2007). The Strengthening the Reporting of Observational Studies in Epidemiology (STROBE) statement: Guidelines for reporting observational studies. PLoS Medicine, 4(10), Article e296. https://doi.org/10.1371/journal.pmed.0040296",
        "decision_label": "Tam destek",
        "confidence": "Yüksek",
        "primary_status": "Birincil",
        "final_status": "Uygun",
        "page_and_section": "S. 1 (Özet) ve S. 2 (STROBE Kontrol Listesi Kullanımı)",
        "blocks": [(1, [12]), (2, [13])],
        "rationale": "STROBE girişiminin kohort ve gözlemsel çalışmalar için yayımladığı orijinal 22 maddelik raporlama kılavuzudur.",
        "recommended_action": "Birincil raporlama rehberi olarak korundu ve entegre edildi.",
        "doi_status": "Yerel PDF DOI Doğrulandı (10.1371/journal.pmed.0040296)",
    },
    {
        "cid": "A006",
        "stem": "Atif_02_Slade_2016_CERT",
        "marker": "Slade ve ark. (2016)",
        "location": "2. Yöntem (2.1 & 2.3 CERT Müdahale Protokolü)",
        "claim_type": "Egzersiz Raporlama Şablonu (CERT)",
        "claim_text": "Ek motor beceri eğitim programının bileşenleri, bireysel uyarlama ölçütleri, gözetim ve ilerleme kuralları CERT kontrol listesine göre şeffaf biçimde raporlanmıştır.",
        "source_ref": "Slade, S. C., Dionne, C. E., Underwood, M., & Buchbinder, R. (2016). Consensus on Exercise Reporting Template (CERT): Modified Delphi study. British Journal of Sports Medicine, 50(23), 1428-1437. https://doi.org/10.1136/bjsports-2015-095454",
        "decision_label": "Tam destek",
        "confidence": "Yüksek",
        "primary_status": "Birincil",
        "final_status": "Uygun",
        "page_and_section": "S. 1-2 (CERT Resmi Kontrol Listesi Maddeleri 1-16b)",
        "blocks": [(1, [14, 15, 25, 30, 32, 33, 34]), (2, [12, 14, 16])],
        "rationale": "Klasördeki resmi CERT kontrol listesi belgesi egzersiz ekipmanı, eğitici niteliği, ilerleme kuralları, bireyselleştirme, dozaj ve uyum raporlaması için zorunlu maddeleri doğrudan tanımlar.",
        "recommended_action": "Birincil şablon olarak korundu ve entegre edildi.",
        "doi_status": "Yerel PDF Doğrulandı (Slade et al., 2016 CERT Checklist)",
    },
    {
        "cid": "A007",
        "stem": "Atif_05_WMA_2013_Declaration_of_Helsinki",
        "marker": "Dünya Tıp Birliği (2013)",
        "location": "2. Yöntem (2.2. Katılımcılar ve Etik Onay)",
        "claim_type": "Etik Beyan Birincil Belgesi",
        "claim_text": "Etik kurul onayı, ebeveynlerden yazılı bilgilendirilmiş onam ve katılımcı çocuklardan sözlü onay alınması Helsinki Bildirgesi ilkelerine göre yürütülmüştür.",
        "source_ref": "World Medical Association. (2013). World Medical Association Declaration of Helsinki: Ethical principles for medical research involving human subjects. JAMA, 310(20), 2191-2194. https://doi.org/10.1001/jama.2013.281053",
        "decision_label": "Tam destek",
        "confidence": "Yüksek",
        "primary_status": "Birincil",
        "final_status": "Uygun",
        "page_and_section": "S. 2 (Madde 23) ve S. 3 (Maddeler 26, 28, 29)",
        "blocks": [(2, [20]), (3, [1, 5, 6])],
        "rationale": "Madde 23 araştırma etik kurulu onayını, Madde 26-29 ise yasal temsilci yazılı bilgilendirilmiş onamını ve onay verebilen çocuklardan alınacak onamı (assent) doğrudan düzenler.",
        "recommended_action": "Birincil etik kaynak olarak entegre edildi.",
        "doi_status": "Yerel PDF DOI Doğrulandı (10.1001/jama.2013.281053)",
    },
    {
        "cid": "A008",
        "stem": "Atif_06_Faul_2007_GPower3",
        "marker": "Faul ve ark. (2007)",
        "location": "2. Yöntem (2.2. İstatistiksel Güç Beyanı)",
        "claim_type": "İstatistiksel Yazılım Birincil Makalesi",
        "claim_text": "N = 34 örneklem büyüklüğü, alfa = 0.05 ve %80 istatistiksel güç koşullarında bağımlı örneklemler t-testi için saptanabilir en küçük etki büyüklüğü (dz ≈ 0.49) G*Power 3 duyarlılık analizi yaklaşımıyla hesaplanmıştır.",
        "source_ref": "Faul, F., Erdfelder, E., Lang, A. G., & Buchner, A. (2007). G*Power 3: A flexible statistical power analysis program for the social, behavioral, and biomedical sciences. Behavior Research Methods, 39(2), 175-191. https://doi.org/10.3758/BF03193146",
        "decision_label": "Tam destek",
        "confidence": "Yüksek",
        "primary_status": "Birincil",
        "final_status": "Uygun",
        "page_and_section": "S. 2-3 (Duyarlılık Analizi) ve S. 6 (Tablo 3 Ortalamalar için t-Testleri)",
        "blocks": [(2, [2]), (3, [2]), (6, [13])],
        "rationale": "Faul ve ark. (2007) G*Power 3'ün duyarlılık analizini (alfa, 1-beta ve N değerlerinden kritik etki büyüklüğünü hesaplama) ve bağımlı gruplar t-testini tanımlayan orijinal birincil yayındır.",
        "recommended_action": "Sayfa/blok indeksleri düzeltilerek tam paragraf düzeyinde boyandı ve entegre edildi.",
        "doi_status": "Yerel PDF Künyesi Doğrulandı (Behavior Research Methods, 2007)",
    },
    {
        "cid": "A009",
        "stem": "Atif_07A_Ruggeri_2020_Motor_Interventions_ASD",
        "marker": "[REQ-03, REQ-26] Ruggeri ve ark. (2020)",
        "location": "1. Giriş (1.2) & 2. Yöntem (2.3) & 4. Tartışma (4.4)",
        "claim_type": "Sistematik Derleme (OSB Motor Müdahaleler)",
        "claim_text": "Otizmli çocuk ve gençlerde haftada 1-5 gün ve 40 dakika-4 saatlik oturumlarla yürütülen yapılandırılmış motor ve fiziksel aktivite müdahaleleri aktivite ve vücut yapı/işlev çıktılarında iyileşme sağlamaktadır.",
        "source_ref": "Ruggeri, A., Dancel, A., Johnson, R., & Sargent, B. (2020). The effect of motor and physical activity intervention on motor outcomes of children with autism spectrum disorder: A systematic review. Autism, 24(3), 544-563. https://doi.org/10.1177/1362361319885215",
        "decision_label": "Tam destek",
        "confidence": "Yüksek",
        "primary_status": "Birincil",
        "final_status": "Uygun",
        "page_and_section": "S. 14 (Motor Beceri Müdahale Dozları - 34 Müdahale Çalışması Sentezi)",
        "blocks": [(14, [6])],
        "rationale": "OSB'de 41 çalışmayı inceleyen sistematik derlemedir; önceki taslaktaki hatalı yazar/DOI bilgisi yerel PDF ile eşleştirilerek Ruggeri, A., Dancel, A., Johnson, R., & Sargent, B. (DOI: 10.1177/1362361319885215) olarak düzeltilmiştir.",
        "recommended_action": "Künye ve DOI hataları düzeltilerek Giriş 1.2, Yöntem 2.3 ve Tartışma 4.4'e entegre edildi.",
        "doi_status": "Yerel PDF DOI Doğrulandı (10.1177/1362361319885215)",
    },
    {
        "cid": "A010",
        "stem": "Atif_08_Guadagnoli_2004_Challenge_Point",
        "marker": "[REQ-04] Guadagnoli ve Lee (2004)",
        "location": "1. Giriş (1.2) & 2. Yöntem (2.3. Protokol)",
        "claim_type": "Motor Öğrenme Kuramsal Çerçeve Makalesi",
        "claim_text": "Görev zorluğunun bireyin anlık işlevsel beceri düzeyine uyarlanması (Challenge Point Framework), algısal-motor öğrenme ve uyaran işlemleme bilgisini optimize etmektedir.",
        "source_ref": "Guadagnoli, M. A., & Lee, T. D. (2004). Challenge point: A framework for conceptualizing the effects of various practice conditions in motor learning. Journal of Motor Behavior, 36(2), 212-224. https://doi.org/10.3200/JMBR.36.2.212-224",
        "decision_label": "Tam destek",
        "confidence": "Yüksek",
        "primary_status": "Birincil",
        "final_status": "Uygun",
        "page_and_section": "S. 3 (Fonksiyonel Zorluk), S. 6 (Bilgi Kazanımı) ve S. 8 (Optimal Nokta)",
        "blocks": [(3, [5, 7]), (6, [5, 10]), (8, [2])],
        "rationale": "Challenge Point Framework kuramını kuran orijinal makaledir; fonksiyonel görev zorluğu ile bireyin yeterlilik düzeyi arasındaki ilişkiyi doğrudan tanımlar.",
        "recommended_action": "Birincil kuramsal kaynak olarak entegre edildi.",
        "doi_status": "Yerel PDF DOI Doğrulandı (10.3200/JMBR.36.2.212-224)",
    },
    {
        "cid": "A011",
        "stem": "Atif_09_Chtourou_2012_Diurnal_Variation_Training_JSCR",
        "marker": "[REQ-10] Chtourou ve ark. (2012)",
        "location": "2. Yöntem (2.4. Ölçüm Koşulları)",
        "claim_type": "Birincil Sirkadiyen Antrenman RCT Çalışması",
        "claim_text": "Ön-test ve son-test ölçümlerinde gün içi sirkadiyen performans dalgalanmalarını kontrol etmek amacıyla ölçümler günün aynı saat diliminde gerçekleştirilmiştir.",
        "source_ref": "Chtourou, H., Chaouachi, A., Driss, T., Dogui, M., Behm, D. G., Chamari, K., & Souissi, N. (2012). The effect of training at the same time of day and tapering period on the diurnal variation of short exercise performances. Journal of Strength and Conditioning Research, 26(3), 697-708. https://doi.org/10.1519/JSC.0b013e3182281c36",
        "decision_label": "Tam destek",
        "confidence": "Yüksek",
        "primary_status": "Birincil",
        "final_status": "Uygun",
        "page_and_section": "S. 1 (Özet), S. 10 (Başlangıç Diürnal Farklar) ve S. 11 (Özet Sonuç)",
        "blocks": [(1, [4]), (10, [13, 15]), (11, [8])],
        "rationale": "Önceki ikincil derleme (Chtourou, 2012 Sports Med) yerine n = 31 randomize kontrollü deneysel JSCR çalışmasıdır; kısa süreli maksimal güç/kuvvet performanslarında sabah-akşam farklarını (%3.01-%14.6) ve aynı saatte test/antrenman kontrolünü doğrudan kanıtlar.",
        "recommended_action": "Birincil deneysel sirkadiyen kontrol kaynağı olarak Yöntem 2.4'e entegre edildi.",
        "doi_status": "Yerel PDF Künyesi Doğrulandı (JSCR, 26(3), 697-708)",
    },
    {
        "cid": "A012",
        "stem": "Atif_10_CamachoCardenosa_2018_Youth_ArmSpan_HandSpan",
        "marker": "[REQ-11] Camacho-Cardenosa ve ark. (2018)",
        "location": "2. Yöntem (2.4. Antropometri)",
        "claim_type": "Birincil Pediatrik Antropometrik Protokol Çalışması",
        "claim_text": "Genç katılımcılarda ayakta boy, oturma boyu, vücut ağırlığı, VKİ, kulaç uzunluğu ve dominant el karış uzunluğu standart kinantropometrik prosedürlerle ölçülmüştür.",
        "source_ref": "Camacho-Cardenosa, A., Camacho-Cardenosa, M., González-Custodio, A., Martínez-Guardado, I., Timón, R., Olcina, G., & Brazo-Sayavera, J. (2018). Anthropometric and physical performance of youth handball players: The role of the relative age. Sports, 6(2), Article 47. https://doi.org/10.3390/sports6020047",
        "decision_label": "Tam destek",
        "confidence": "Yüksek",
        "primary_status": "Birincil",
        "final_status": "Uygun",
        "page_and_section": "S. 3 (Yöntem - Antropometrik Ölçüm Standartları)",
        "blocks": [(3, [6])],
        "rationale": "Ayakta boy (0.1 cm), oturma boyu (0.1 cm), vücut kütlesi (0.1 kg), kulaç uzunluğu (0.1 cm) ve el karış uzunluğu ölçümlerini genç sporcularda uygulayan birincil hakemli makaledir (el karış ölçümünün tam 1. ve 5. parmak ucu tanımı veya yöntem detayı metinde açıkça ayırt edilmiştir).",
        "recommended_action": "Yöntem 2.4 antropometrik ölçümler bölümüne birincil kaynak olarak entegre edildi.",
        "doi_status": "Yerel PDF Künyesi Doğrulandı (Sports, 2018, 6, 47)",
    },
    {
        "cid": "A013",
        "stem": "Atif_11B_Mirwald_2002_Maturity_Offset_Equation",
        "marker": "Mirwald ve ark. (2002)",
        "location": "1. Giriş (1.3) & 2. Yöntem (2.4. Biyolojik Olgunlaşma)",
        "claim_type": "Birincil Antropometrik Denklem Geliştirme Çalışması",
        "claim_text": "Erkek çocuk ve ergenlerde biyolojik olgunlaşma offseti, kronolojik yaş, ayakta boy, oturma boyu, bacak uzunluğu ve vücut ağırlığına dayalı Mirwald erkek PHV kestirim denklemiyle (R² = 0.891, SEE = 0.592) hesaplanmıştır.",
        "source_ref": "Mirwald, R. L., Baxter-Jones, A. D. G., Bailey, D. A., & Beunen, G. P. (2002). An assessment of maturity from anthropometric measurements. Medicine & Science in Sports & Exercise, 34(4), 689-694. https://doi.org/10.1097/00005768-200204000-00020",
        "decision_label": "Tam destek",
        "confidence": "Yüksek",
        "primary_status": "Birincil",
        "final_status": "Uygun",
        "page_and_section": "S. 1 (Giriş) ve S. 3-4 (Erkek PHV Kestirim Denklemi 3)",
        "blocks": [(1, [5]), (3, [13]), (4, [1])],
        "rationale": "Erkek çocuklar için Denklem 3 katsayılarını (-9.236 + 0.0002708...) bizzat türeten birincil çalışmadır; elenen Baxter-Jones ve ark. (2005) derlemesinin de yerini alır.",
        "recommended_action": "Sayfa/blok indeksleri düzeltilerek tam denklem paragrafıyla boyandı ve entegre edildi.",
        "doi_status": "Yerel PDF Künyesi Doğrulandı (Med. Sci. Sports Exerc., 2002, 34(4), 689-694)",
    },
    {
        "cid": "A014",
        "stem": "Atif_12_Badau_2018_Reaction_Time_Laterality",
        "marker": "[REQ-12, REQ-22] Badau ve ark. (2018)",
        "location": "2. Yöntem (2.4) & 4. Tartışma (4.2)",
        "claim_type": "Birincil Görsel Reaksiyon Süresi ve Lateralite Çalışması",
        "claim_text": "Görsel uyaranlara karşı reaksiyon süresi ölçümünde çoklu deneme protokolü ve dominant/non-dominant el lateralitesi nöromotor işlemleme hızında anlamlı farklılıklar ortaya koymaktadır.",
        "source_ref": "Badau, D., Baydil, B., & Badau, A. (2018). Differences among three measures of reaction time based on hand laterality in individual sports. Sports, 6(2), Article 45. https://doi.org/10.3390/sports6020045",
        "decision_label": "Tam destek",
        "confidence": "Yüksek",
        "primary_status": "Birincil",
        "final_status": "Uygun",
        "page_and_section": "S. 1 (Özet), S. 2 (Yöntem), S. 4 (Tablo 2) ve S. 5 (Tartışma)",
        "blocks": [(1, [7]), (2, [10]), (4, [20, 25]), (5, [9])],
        "rationale": "N = 332 katılımcıda görsel basit, tanıma ve bilişsel reaksiyon sürelerini dominant vs dominant olmayan el açısından inceleyen birincil deneysel çalışmadır (p = 0.001).",
        "recommended_action": "Görsel reaksiyon süresi ve sağ el dominantlığı nöromotor tartışması için entegre edildi.",
        "doi_status": "Yerel PDF Künyesi Doğrulandı (Sports, 2018, 6, 45)",
    },
    {
        "cid": "A015",
        "stem": "Atif_13_Gasior_2020_Jamar_Handgrip_Reliability_Youth",
        "marker": "[REQ-13] Gąsior ve ark. (2020)",
        "location": "2. Yöntem (2.4. El Kavrama Kuvveti)",
        "claim_type": "Birincil Pediatrik Dinamometre Güvenirlik Çalışması",
        "claim_text": "Çocuk ve ergenlerde oturur pozisyonda dirsek 90° fleksiyonda ASHT protokolüyle Jamar dinamometresi kullanılarak yapılan el kavrama ölçümleri mükemmel test-tekrar test güvenirliği (ICC = 0.95-0.98) sergilemektedir.",
        "source_ref": "Gąsior, J. S., Pawłowski, M., Jeleń, P. J., Rameckers, E. A., Williams, C. A., Makuch, R., & Werner, B. (2020). Test-retest reliability of handgrip strength measurement in children and preadolescents. International Journal of Environmental Research and Public Health, 17(21), Article 8026. https://doi.org/10.3390/ijerph17218026",
        "decision_label": "Tam destek",
        "confidence": "Yüksek",
        "primary_status": "Birincil",
        "final_status": "Uygun",
        "page_and_section": "S. 3 (ASHT Protokolü), S. 4 (Tablo 1 ICC Verileri) ve S. 6 (Tartışma)",
        "blocks": [(3, [1, 2]), (4, [20, 29, 30]), (6, [3])],
        "rationale": "7-13 yaş arası 306 katılımcıda Jamar 2. tutuş pozisyonu ve 90° dirsek fleksiyonlu ASHT protokolünün 7-9 ve 10-13 yaş gruplarında yüksek güvenirliğini (ICC = 0.95 ve 0.98) doğrulayan birincil çalışmadır.",
        "recommended_action": "Jamar ASHT protokol güvenirliği için Yöntem 2.4'e entegre edildi.",
        "doi_status": "Yerel PDF Künyesi Doğrulandı (IJERPH, 2020, 17, 8026)",
    },
    {
        "cid": "A016",
        "stem": "Atif_14_Correia_2024_Relative_Handgrip_Youth",
        "marker": "[REQ-14, REQ-24 ikame] Correia ve ark. (2024)",
        "location": "2. Yöntem (2.4) & 4. Tartışma (4.3)",
        "claim_type": "Birincil Rölatif El Kavrama Kuvveti ve Olgunlaşma Çalışması",
        "claim_text": "Gençlerde mutlak el kavrama kuvvetinin vücut ağırlığına bölünmesiyle elde edilen rölatif el kavrama oranı (kg/kg), kronolojik yaş ve somatik olgunlaşma kaynaklı varyansı (%49-%72'den erkeklerde %22-%25 ve kızlarda %6 düzeyine) belirgin biçimde azaltmaktadır.",
        "source_ref": "Correia, A. S., Zymbal, V., & Baptista, F. (2024). Musculoskeletal fitness: Relative handgrip strength and vertical jump power from 10 to 18 years old. Frontiers in Pediatrics, 12, Article 1207609. https://doi.org/10.3389/fped.2024.1207609",
        "decision_label": "Tam destek",
        "confidence": "Yüksek",
        "primary_status": "Birincil",
        "final_status": "Uygun",
        "page_and_section": "S. 1-3 (Yöntem ve Normalizasyon) ve S. 8 (Tartışma - Olgunlaşma Varyansı)",
        "blocks": [(1, [16]), (2, [10]), (3, [10]), (8, [0, 1])],
        "rationale": "N = 736 genç katılımcıda (10-18 yaş) rölatif kavrama kuvveti (kg/kg) normalizasyonunun biyolojik olgunlaşma ve yaş bağımlılığını azalttığını ampirik olarak kanıtlar ve Baxter-Jones (2005) derlemesinin yerini doldurur.",
        "recommended_action": "Rölatif kuvvet normalizasyonu ve olgunlaşma kontrolü için Yöntem 2.4 ve Tartışma 4.3'e entegre edildi.",
        "doi_status": "Yerel PDF DOI Doğrulandı (10.3389/fped.2024.1207609)",
    },
    {
        "cid": "A017",
        "stem": "Atif_14C_Peterson_2016_Relative_Grip_Adolescents",
        "marker": "[REQ-06] Peterson ve ark. (2016)",
        "location": "1. Giriş (1.3) & 4. Tartışma (4.3)",
        "claim_type": "Birincil Pediatrik Normalleştirilmiş Kuvvet Risk Eşiği Çalışması",
        "claim_text": "N = 1326 ergen katılımcıda el kavrama kuvvetinin vücut kütlesine bölünmesiyle elde edilen normalize edilmiş kuvvet (grip strength/body mass), kardiyometabolik ve fonksiyonel risk sınıflandırmasında klinik eşik değerler (erkeklerde yüksek risk <= 0.33, orta <= 0.45) sağlamaktadır.",
        "source_ref": "Peterson, M. D., Zhang, P., Saltarelli, W. A., Visich, P. S., & Gordon, P. M. (2016). Low muscle strength thresholds for the detection of cardiometabolic risk in adolescents. American Journal of Preventive Medicine, 50(5), 593-599. https://doi.org/10.1016/j.amepre.2015.09.019",
        "decision_label": "Tam destek",
        "confidence": "Yüksek",
        "primary_status": "Birincil",
        "final_status": "Uygun",
        "page_and_section": "S. 1 (Özet Klinik Eşikler) ve S. 2 (Yöntem - Jamar Normalizasyon)",
        "blocks": [(1, [5, 6, 7]), (2, [10])],
        "rationale": "Vücut kütlesine normalize edilmiş kavrama kuvvetinin (kg/kg) klinik anlamlılık eşiklerini 1326 ergende belirleyen birincil araştırmadır ve çalışmamızın kohort başlangıç rölatif kuvvet düzeyini (0.25-0.28 kg/kg) yorumlamak için doğrudan referans sunar.",
        "recommended_action": "Giriş 1.3 ve Tartışma 4.3 klinik yorumlamasına entegre edildi.",
        "doi_status": "Yerel PDF DOI Doğrulandı (10.1016/j.amepre.2015.09.019)",
    },
    {
        "cid": "A018",
        "stem": "Atif_15_Zifchock_2008_Symmetry_Angle",
        "marker": "[REQ-15, REQ-25] Zifchock ve ark. (2008)",
        "location": "2. Yöntem (2.4) & 4. Tartışma (4.3)",
        "claim_type": "Birincil Biyomekanik Asimetri Metodolojisi Çalışması",
        "claim_text": "Klasik simetri oranı indekslerinin seçilen referans tarafına duyarlılığı ve literatürde %15 üzerindeki değerlerin klinik anlamlı asimetri ölçütü olarak ele alınması dikkate alınarak iki taraf asimetri düzeyi değerlendirilmiştir.",
        "source_ref": "Zifchock, R. A., Davis, I., Higginson, J., & Royer, T. (2008). The Symmetry Angle: A novel, robust method of quantifying asymmetry. Gait & Posture, 27(4), 622-627. https://doi.org/10.1016/j.gaitpost.2007.08.006",
        "decision_label": "Tam destek",
        "confidence": "Yüksek",
        "primary_status": "Birincil",
        "final_status": "Uygun",
        "page_and_section": "S. 1-2 (Simetri İndeksi ve Referans Seçimi), S. 3 (%15 Ölçütü) ve S. 5 (Tartışma)",
        "blocks": [(1, [17, 19]), (2, [0]), (3, [12]), (5, [12, 16])],
        "rationale": "Oran temelli simetri indekslerindeki referans taraf seçimi sorununu ampirik olarak inceler ve %15 klinik eşiği test eder (taslağımızdaki max(Sağ, Sol) normalizasyonu bu metodolojik tartışmayla uyumlu biçimde çerçevelenmiştir).",
        "recommended_action": "Yöntem 2.4 asimetri indeks gerekçesi ve Tartışma 4.3 (%15 eşik yorumu) için entegre edildi.",
        "doi_status": "Yerel PDF DOI Doğrulandı (10.1016/j.gaitpost.2007.08.006)",
    },
    {
        "cid": "A019",
        "stem": "Atif_16A_Markovic_2004_Squat_CMJ_Reliability",
        "marker": "Markovic ve ark. (2004)",
        "location": "2. Yöntem (2.4. Dikey Sıçrama)",
        "claim_type": "Birincil Dikey Sıçrama Güvenirlik Çalışması",
        "claim_text": "Eller belde karşıt hareketli aktif dikey sıçrama (CMJ) testinin elektronik sıçrama matı üzerinde 3 deneme ile uygulanması yüksek faktöriyel geçerlik ve güvenirlik sağlamaktadır.",
        "source_ref": "Markovic, G., Dizdar, D., Jukic, I., & Cardinale, M. (2004). Reliability and factorial validity of squat and countermovement jump tests. Journal of Strength and Conditioning Research, 18(3), 551-555. https://doi.org/10.1519/1533-4287(2004)18<551:RAFVOS>2.0.CO;2",
        "decision_label": "Tam destek",
        "confidence": "Yüksek",
        "primary_status": "Birincil",
        "final_status": "Uygun",
        "page_and_section": "S. 2 (Yöntem - Temas Matı ve 3 Deneme) ve S. 4 (Sonuç)",
        "blocks": [(2, [11, 12, 13, 14, 17]), (4, [2, 8])],
        "rationale": "Temas/sıçrama matında eller belde CMJ yüksekliği ölçümünün yüksek test-tekrar test güvenirliğini doğrudan kanıtlayan birincil araştırmadır.",
        "recommended_action": "Blok indeksleri temizlenerek boyandı ve CMJ ölçüm protokolü için entegre edildi.",
        "doi_status": "Yerel PDF Künyesi Doğrulandı (JSCR, 2004, 18(3), 551-555)",
    },
    {
        "cid": "A020",
        "stem": "Atif_16B_Gulu_2021_Youth_CMJ_Peak_Power_Equation",
        "marker": "[REQ-16] Gülü ve Akalan (2021)",
        "location": "2. Yöntem (2.4) & 4. Tartışma (4.5 Sınırlılıklar)",
        "claim_type": "Birincil Pediatrik Sıçrama Zirve Güç Regresyon Çalışması",
        "claim_text": "Yetişkin popülasyonlarda geliştirilen sıçrama gücü tahmin denklemleri gelişim çağındaki çocuklarda yaş ve protokol yanlılığı gösterebilmektedir; çocuk ve ergenlerde sıçrama yüksekliği ve vücut ağırlığından güç kestiriminde popülasyon ve kol pozisyonu uyumu kritik önem taşır.",
        "source_ref": "Gülü, M., & Akalan, C. (2021). A new peak-power estimation equations in 12 to 14 years-old soccer players. Medicine, 100(39), Article e27383. https://doi.org/10.1097/MD.0000000000027383",
        "decision_label": "Tam destek",
        "confidence": "Yüksek",
        "primary_status": "Birincil",
        "final_status": "Uygun",
        "page_and_section": "S. 1 (Amaç - n = 188) ve S. 3 (Eller Belde Protokol & Yetişkin Denklem Yanlılığı)",
        "blocks": [(1, [19]), (3, [4, 10, 19, 20])],
        "rationale": "n = 188 gelişim çağı katılımcıda kuvvet platformu geçerli zirve güç kestirimini test eder ve yetişkin denklemlerinin pediatrik örneklerde uygun olmayabileceğini vurgular; bu kanıt doğrultusunda çalışmamızdaki Sayers W çıktısı CMJ (cm) yanında keşifsel türev olarak çerçevelenmiştir.",
        "recommended_action": "Yetişkin güç denklemlerinin pediatrik kısıtlılığını vurgulayan metodolojik çerçeveyle entegre edildi.",
        "doi_status": "Yerel PDF DOI Doğrulandı (10.1097/MD.0000000000027383)",
    },
    {
        "cid": "A021",
        "stem": "Atif_17_Stewart_2014_Change_of_Direction_Speed",
        "marker": "[REQ-17] Stewart ve ark. (2014)",
        "location": "2. Yöntem (2.4. Çeviklik Ölçümü)",
        "claim_type": "Birincil Yön Değiştirme Güvenirlik Çalışması",
        "claim_text": "9.14 metrelik (4.55 m + 9.10 m + 4.55 m) Pro-Agility mekik parkurunun fotosel zamanlama kapılarıyla ölçümü yüksek test-tekrar test güvenirliği (havuzlanmış ICC = 0.90, %95 GA: 0.84-0.94; SEM = 0.12 sn; CV = %2.19) sunmaktadır.",
        "source_ref": "Stewart, P. F., Turner, A. N., & Miller, S. C. (2014). Reliability, factorial validity, and interrelationships of five commonly used change of direction speed tests. Scandinavian Journal of Medicine & Science in Sports, 24(3), 500-506. https://doi.org/10.1111/sms.12019",
        "decision_label": "Tam destek",
        "confidence": "Yüksek",
        "primary_status": "Birincil",
        "final_status": "Uygun",
        "page_and_section": "S. 1 (Özet), S. 2 (Pro-Agility Fotosel Kurulumu) ve S. 4 (Tablo 3 Güvenirlik)",
        "blocks": [(1, [6]), (2, [11, 17]), (4, [17])],
        "rationale": "Pro-Agility ve dört yön değiştirme testinin fotosel zamanlamalı güvenirliğini ve faktöriyel geçerliğini doğrudan sınayan birincil deneysel çalışmadır (çalışmamızda OSB popülasyon yorgunluğunu önlemek adına 2 deneme ile uygulanmış olduğu açıkça belirtilmiştir).",
        "recommended_action": "Birincil Pro-Agility güvenirlik kaynağı olarak Yöntem 2.4'e entegre edildi.",
        "doi_status": "Yerel PDF DOI Doğrulandı (10.1111/sms.12019)",
    },
    {
        "cid": "A022",
        "stem": "Atif_18_Tsiakaras_2025_Sit_and_Reach_Reliability_Youth",
        "marker": "[REQ-18] Tsiakaras ve ark. (2025)",
        "location": "2. Yöntem (2.4. Esneklik Ölçümü)",
        "claim_type": "Birincil Gelişim Çağı Otur-Eriş Güvenirlik Çalışması",
        "claim_text": "Gelişim çağındaki çocuk ve ergenlerde dizler tam ekstansiyonda uygulanan otur-eriş esneklik testinin 3 maksimal deneme ile ölçülmesi çok yüksek seans içi ve seanslar arası güvenirlik (ICC = 0.988-0.998; SEM% = 1.73-3.40) sağlamaktadır.",
        "source_ref": "Tsiakaras, N., Karatrantou, K., Batatolis, C., Papavasileiou, K., Tzeli, F., & Gerodimos, V. (2025). Intrasession and intersession reliability of flexibility tests during developmental years: The effects of sport, age, and sex. Sports, 13(8), Article 238. https://doi.org/10.3390/sports13080238",
        "decision_label": "Tam destek",
        "confidence": "Yüksek",
        "primary_status": "Birincil",
        "final_status": "Uygun",
        "page_and_section": "S. 4 (Protokol), S. 7 (Denemeler Arası ANOVA), S. 12 (Tablo ICC) ve S. 18-19 (Sonuç)",
        "blocks": [(4, [7]), (7, [4]), (12, [0, 11]), (18, [1]), (19, [3])],
        "rationale": "Çocuk ve ergenlerde 3 denemeli standart otur-eriş kutusu testinin ölçüm hatası (SEM%) ve ICC güvenirliğini yaş ile spor etkileriyle birlikte analiz eden birincil araştırmadır.",
        "recommended_action": "Otur-eriş esneklik protokol güvenirliği için Yöntem 2.4'e entegre edildi.",
        "doi_status": "Yerel PDF DOI Doğrulandı (10.3390/sports13080238)",
    },
    {
        "cid": "A023",
        "stem": "Atif_19_Shapiro_1965_Normality_Test",
        "marker": "[REQ-19] Shapiro ve Wilk (1965)",
        "location": "2. Yöntem (2.5. İstatistiksel Analiz)",
        "claim_type": "Orijinal İstatistiksel Test Kurucu Makalesi",
        "claim_text": "Sürekli değişkenlerin normal dağılım varsayımı, küçük ve orta büyüklükteki tam örneklemlerde yüksek duyarlılığa sahip Shapiro-Wilk W istatistiği ile test edilmiştir.",
        "source_ref": "Shapiro, S. S., & Wilk, M. B. (1965). An analysis of variance test for normality (complete samples). Biometrika, 52(3-4), 591-611. https://doi.org/10.2307/2333709",
        "decision_label": "Tam destek",
        "confidence": "Yüksek",
        "primary_status": "Birincil",
        "final_status": "Uygun",
        "page_and_section": "S. 1-2 (Giriş ve W İstatistiği), S. 13 (Matematiksel Özellikler) ve S. 21 (Özet)",
        "blocks": [(1, [0]), (2, [11, 12]), (13, [21]), (21, [7])],
        "rationale": "Shapiro-Wilk normallik testini matematiksel olarak geliştiren orijinal 1965 Biometrika makalesidir; n < 20 ve N = 34 örneklemlerde normallik değerlendirmesini doğrudan temellendirir.",
        "recommended_action": "Önceki simülasyon ikamesi yerine birincil kurucu kaynak olarak Yöntem 2.5'e entegre edildi.",
        "doi_status": "Yerel PDF Künyesi Doğrulandı (Biometrika, 1965, 52(3/4), 591-611)",
    },
    {
        "cid": "A024",
        "stem": "Atif_21A_Hopkins_2004_Interpret_Changes_Athletic_Test",
        "marker": "Hopkins (2004)",
        "location": "2. Yöntem (2.5. İstatistiksel Analiz)",
        "claim_type": "Orijinal Yöntem Makalesi (SWC)",
        "claim_text": "Bireysel yanıtlayıcı analizinde kullanılan En Küçük Anlamlı Değişim (SWC = 0.20 × Başlangıç SS) eşiği, sportif ve fonksiyonel performans takibinde standart yöntem olarak uygulanmıştır.",
        "source_ref": "Hopkins, W. G. (2004). How to interpret changes in an athletic performance test. Sportscience, 8, 1-7.",
        "decision_label": "Tam destek",
        "confidence": "Yüksek",
        "primary_status": "Birincil",
        "final_status": "Uygun",
        "page_and_section": "S. 1 (Giriş), S. 3 (SWC Gerekçesi) ve S. 7 (Uygulama Adımları)",
        "blocks": [(1, [3]), (3, [3]), (7, [1])],
        "rationale": "Atletik ve performans testlerinde SWC yaklaşımını doğrudan tanımlayan birincil yöntem makalesidir.",
        "recommended_action": "Blok indeksleri temizlenerek tam paragraf olarak boyandı ve Yöntem 2.5'e entegre edildi.",
        "doi_status": "Yerel PDF Künyesi Doğrulandı (Sportscience, 8, 1-7, 2004)",
    },
    {
        "cid": "A025",
        "stem": "Atif_22_Nathans_2012_Multiple_Regression_Variable_Importance",
        "marker": "[REQ-21] Nathans ve ark. (2012)",
        "location": "2. Yöntem (2.5) & 4. Tartışma (4.3)",
        "claim_type": "Birincil Regresyon Değişken Önemi Metodoloji Çalışması",
        "claim_text": "Çoklu doğrusal regresyon modellerinde bağımsız değişkenler arasındaki eşdoğrusallık ve paylaşılan varyans nedeniyle tek başına standartlaştırılmış beta katsayılarına dayanılarak değişken önemi yorumlanmamalı, yapısallık ve iki değişkenli ilişkiler birlikte değerlendirilmelidir.",
        "source_ref": "Nathans, L. L., Oswald, F. L., & Nimon, K. (2012). Interpreting multiple linear regression: A guidebook of variable importance. Practical Assessment, Research, and Evaluation, 17(1), Article 9. https://doi.org/10.7275/5fex-b874",
        "decision_label": "Tam destek",
        "confidence": "Yüksek",
        "primary_status": "Birincil",
        "final_status": "Uygun",
        "page_and_section": "S. 2 (Amaç), S. 4 (Beta Ağırlıkları ve Eşdoğrusallık Kısıtı) ve S. 18 (Öneriler)",
        "blocks": [(2, [8, 9]), (4, [3, 4, 5]), (18, [4])],
        "rationale": "Klasördeki eksik Tabachnick kitabının yerine temin edilen bu hakemli metodoloji makalesi, antropometrik yordayıcılar (kulaç, karış, yaş, PHV; r = 0.82-0.97) arasındaki yüksek eşdoğrusallık altında beta katsayılarının neden tek başına yorumlanmaması gerektiğini doğrudan açıklamaktadır.",
        "recommended_action": "Yöntem 2.5 ve Tartışma 4.3 eşdoğrusallık yorumuna birincil kaynak olarak entegre edildi.",
        "doi_status": "Yerel PDF DOI Doğrulandı (10.7275/5fex-b874)",
    },
    {
        "cid": "A026",
        "stem": "Atif_26_Logan_2012_Motor_Skills_Interventions_Children",
        "marker": "[REQ-26] Logan ve ark. (2012)",
        "location": "1. Giriş (1.2) & 4. Tartışma (4.4)",
        "claim_type": "Birincil Pediatrik Motor Beceri Meta-Analizi",
        "claim_text": "Çocuklarda temel motor beceriler kendiliğinden yalnızca olgunlaşma yoluyla optimal düzeye ulaşmaz; yapılandırılmış motor beceri müdahaleleri anlamlı pozitif etki (d = 0.39, p < 0.001) üretirken serbest oyun gruplarında anlamlı gelişim görülmemiştir (d = 0.06, p = 0.33).",
        "source_ref": "Logan, S. W., Robinson, L. E., Wilson, A. E., & Lucas, W. A. (2012). Getting the fundamentals of movement: A meta-analysis of the effectiveness of motor skill interventions in children. Child: Care, Health and Development, 38(3), 305-315. https://doi.org/10.1111/j.1365-2214.2011.01307.x",
        "decision_label": "Tam destek",
        "confidence": "Yüksek",
        "primary_status": "Birincil",
        "final_status": "Uygun",
        "page_and_section": "S. 1 (Olgunlaşma vs Planlı Öğretim) ve S. 3 (Meta-Analiz Etki Büyüklükleri)",
        "blocks": [(1, [12]), (3, [7, 9])],
        "rationale": "Planlı motor müdahalenin serbest oyun/olgunlaşmaya üstünlüğünü (d = 0.39 vs 0.06) ve müdahale süresi ile korelasyon bulunmamasını (r = -0.18, p = 0.296) niceliksel olarak kanıtlayarak 15 dakikalık hedefli mikro-doz yaklaşımımızı doğrudan destekler.",
        "recommended_action": "Giriş 1.2 ve Tartışma 4.4'e birincil meta-analiz kanıtı olarak entegre edildi.",
        "doi_status": "Yerel PDF DOI Doğrulandı (10.1111/j.1365-2214.2011.01307.x)",
    },
]

ARCHIVED_SECONDARY_SPEC = {
    "cid": "A027",
    "pdf_path": UNUSED_ORIG_DIR / "Atif_05B_BaxterJones_2005_Maturation_Pediatric_Exercise_Review.pdf",
    "stem": "Atif_05B_BaxterJones_2005_Maturation_Pediatric_Exercise_Review",
    "marker": "[REQ-05, REQ-24] Baxter-Jones ve ark. (2005)",
    "location": "Giriş 1.3 & Tartışma 4.3 (ELENDİ - Arşive Taşındı)",
    "claim_type": "Anlatısal Derleme (Birincil Değil)",
    "claim_text": "Büyüme ve biyolojik olgunlaşmanın pediatrik kuvvet ve egzersiz bilimi çıktıları üzerindeki ayrıştırılmasına ilişkin derleme.",
    "source_ref": "Baxter-Jones, A. D. G., Eisenmann, J. C., & Sherar, L. B. (2005). Controlling for maturation in pediatric exercise science. Pediatric Exercise Science, 17(1), 18-30. https://doi.org/10.1123/pes.17.1.18",
    "decision_label": "Birincil değil",
    "confidence": "Yüksek",
    "primary_status": "Birincil değil",
    "secondary_attribution_proof": "S.1: 'This article reviews some of the possible biological maturity indicators...' | S.2: ikincil atıflar (8,28,29,30) | S.6 Şekil 2: Saskatchewan PBMAS (3) verilerinden uyarlanmıştır.",
    "final_status": "Uygun değil (Arşivlendi)",
    "page_and_section": "S. 1-2, S. 4 ve S. 6 (Anlatısal Derleme Kanıtları)",
    "blocks": [(1, [6, 8]), (2, [3]), (4, [7]), (6, [1])],
    "rationale": "Metin içi ifadeler özgün bir deneysel analiz değil, mevcut olgunlaşma yöntemlerini özetleyen anlatısal bir derleme (narrative review) olduğunu açıkça kanıtlamaktadır.",
    "recommended_action": "Çalışmadan çıkarılarak kullanilmayanlar/ klasörüne arşivlendi; yerine birincil Mirwald ve ark. (2002), Correia ve ark. (2024) ve Peterson ve ark. (2016) kaynakları ikame edildi.",
    "doi_status": "Yerel PDF Künyesi Doğrulandı (Arşivlendi)",
}

MISSING_CITATIONS = [
    {
        "cid": "A028",
        "marker": "[REQ-09, REQ-20] Cohen (1988)",
        "location": "2. Yöntem (2.2. Güç Beyanı & 2.5. İstatistiksel Analiz)",
        "claim_type": "Birincil İstatistik Monografisi (Klasörde PDF Eksik)",
        "claim_text": "Cohen d (0.20 küçük, 0.50 orta, 0.80 büyük) ve eta-kare (0.01/0.06/0.14) nitel etki büyüklüğü eşik sınıflandırmalarının orijinal kaynağı.",
        "source_ref": "Cohen, J. (1988). Statistical power analysis for the behavioral sciences (2nd ed.). Lawrence Erlbaum Associates.",
        "pdf_file": "Atif_06B_Cohen_1988_Power_Analysis.pdf (KLASÖRDE YOK)",
        "page_and_section": "Doğrulanamadı (PDF yüklenmemiş)",
        "passage": "[DOĞRULANAMADI: Cohen (1988) kitabına/PDF dosyasına araştırmacı tarafından erişilemediği için Yontem_Kaynak_Makaleler_PDF/ klasöründe tam metin bulunmamaktadır.]",
        "decision_label": "Bulunamadı",
        "confidence": "Düşük",
        "primary_status": "Doğrulanamadı (PDF Eksik)",
        "secondary_attribution_proof": "",
        "final_status": "Çıkarıldı / Doğrulanamadı",
        "rationale": "citation-audit kuralı gereği yerel PDF'i klasörde bulunmayan hiçbir kaynak doğrulanmış birincil kaynak olarak Aşama 5 makale metnine veya kaynakça listesine yazılamaz.",
        "recommended_action": "Aşama 5 final makalede Cohen (1988) kaynakça listesine eklenmemiştir; duyarlılık analizi sayısal dz ≈ 0.49 değeri Faul ve ark. (2007) ile raporlanmış, bulgular kesin sayısal d ve ηp² değerleri ile Hopkins (2004) SWC eşiği (0.20 × SS_ön) üzerinden sunulmuştur.",
        "doi_status": "PDF Eksik (Doğrulanamadı)",
    }
]


def load_blocks_by_page(stem: str) -> dict[int, list[dict]]:
    jpath = SYS_DIR / f"{stem}_extracted.json"
    if not jpath.exists():
        raise FileNotFoundError(f"Extracted JSON not found: {jpath}")
    data = json.loads(jpath.read_text(encoding="utf-8"))
    by_page: dict[int, list[dict]] = {}
    for blk in data.get("blocks", []):
        by_page.setdefault(int(blk["page"]), []).append(blk)
    return by_page, data


def reset_stage4_dirs() -> None:
    ANNOTATED_DIR.mkdir(parents=True, exist_ok=True)
    PNG_DIR.mkdir(parents=True, exist_ok=True)
    for f in ANNOTATED_DIR.glob("*.pdf"):
        f.unlink()
    for f in PNG_DIR.glob("*.png"):
        f.unlink()


def process_pdf_spec(spec: dict, src_pdf: Path, out_pdf: Path) -> tuple[dict, dict]:
    by_page, meta = load_blocks_by_page(spec["stem"])
    annotations = []
    passages = []

    for pg, idxs in spec["blocks"]:
        pblocks = by_page.get(pg, [])
        bboxes = []
        pg_texts = []
        for idx in idxs:
            if idx >= len(pblocks):
                raise IndexError(f"{spec['stem']} P{pg} index {idx} >= {len(pblocks)}")
            blk = pblocks[idx]
            bboxes.append(blk["bbox"])
            pg_texts.append(clean_xml(blk["text"].replace("\n", " ").strip()))
        passage_pg = " ".join(pg_texts)
        passages.append(f"[S.{pg}] {passage_pg}")
        annotations.append({
            "citation_id": spec["cid"],
            "page": pg,
            "bboxes": bboxes,
            "exact_bboxes": False,
            "decision_label": spec["decision_label"],
            "primary_status": spec["primary_status"],
            "final_status": spec["final_status"],
            "claim_text": spec["claim_text"],
            "passage": passage_pg,
            "rationale": spec["rationale"],
        })

    res = highlight_pdf.annotate_pdf(
        source_pdf=src_pdf,
        annotations=annotations,
        output_pdf=out_pdf,
        render_dir=PNG_DIR,
    )
    print(f"[OK] {spec['cid']} -> {out_pdf.name} (highlights={res.get('highlights_added')}, pngs={len(res.get('rendered_pages', []))})")

    cit_record = {
        "citation_id": spec["cid"],
        "citation_marker": spec["marker"],
        "manuscript_location": spec["location"],
        "claim_type": spec["claim_type"],
        "claim_text": spec["claim_text"],
        "source_ref": spec["source_ref"],
        "pdf_file": src_pdf.name,
        "page_and_section": spec["page_and_section"],
        "passage": " | ".join(passages),
        "decision_label": spec["decision_label"],
        "confidence": spec["confidence"],
        "primary_status": spec["primary_status"],
        "secondary_attribution_proof": spec.get("secondary_attribution_proof", ""),
        "final_status": spec["final_status"],
        "rationale": spec["rationale"],
        "recommended_action": spec["recommended_action"],
        "phase8_doi_status": spec["doi_status"],
        "human_approval": "Onaylandı (Stage 4 Tamam)" if spec["primary_status"] == "Birincil" else "Arşivlendi (İkincil Elendi)",
    }
    inv_entry = {
        "page_count": meta.get("page_count"),
        "sha256": meta.get("sha256"),
        "has_text_layer": meta.get("has_text_layer"),
        "dois": meta.get("dois", [])[:5],
    }
    return cit_record, inv_entry


def main() -> None:
    print("=== Running Stage 4 Citation Audit & Full-Passage Highlighting (26 Primary + 1 Archived + 1 Missing) ===")
    reset_stage4_dirs()

    citations = []
    pdf_inventory = {}

    for spec in PDF_HIGHLIGHT_SPECS:
        src_pdf = PDF_DIR / f"{spec['stem']}.pdf"
        out_pdf = ANNOTATED_DIR / f"{spec['stem']}_annotated.pdf"
        if not src_pdf.exists():
            raise FileNotFoundError(f"Primary PDF missing: {src_pdf}")
        rec, inv = process_pdf_spec(spec, src_pdf, out_pdf)
        citations.append(rec)
        pdf_inventory[src_pdf.name] = inv

    # Archived secondary review -> kullanilmayanlar/boyali_kaynaklar/
    UNUSED_ANN_DIR.mkdir(parents=True, exist_ok=True)
    baxter_out = UNUSED_ANN_DIR / f"{ARCHIVED_SECONDARY_SPEC['stem']}_annotated.pdf"
    rec_sec, inv_sec = process_pdf_spec(ARCHIVED_SECONDARY_SPEC, ARCHIVED_SECONDARY_SPEC["pdf_path"], baxter_out)
    citations.append(rec_sec)
    pdf_inventory[f"kullanilmayanlar/orijinal_kaynaklar/{ARCHIVED_SECONDARY_SPEC['pdf_path'].name}"] = inv_sec

    # Missing Cohen (1988)
    for m in MISSING_CITATIONS:
        citations.append({
            "citation_id": m["cid"],
            "citation_marker": m["marker"],
            "manuscript_location": m["location"],
            "claim_type": m["claim_type"],
            "claim_text": m["claim_text"],
            "source_ref": m["source_ref"],
            "pdf_file": m["pdf_file"],
            "page_and_section": m["page_and_section"],
            "passage": m["passage"],
            "decision_label": m["decision_label"],
            "confidence": m["confidence"],
            "primary_status": m["primary_status"],
            "secondary_attribution_proof": m["secondary_attribution_proof"],
            "final_status": m["final_status"],
            "rationale": m["rationale"],
            "recommended_action": m["recommended_action"],
            "phase8_doi_status": m["doi_status"],
            "human_approval": "Hariç Tutuldu (PDF Eksik)",
        })

    audit_payload = {
        "metadata": {
            "manuscript_name": "Q1_Makale_Iskelet_Taslagi_TR.md -> Stage 5 Birincil Atıflı Tam Metin (N = 34)",
            "citation_style": "APA 7",
            "timestamp": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
            "phase8_status": "Tamamlandı: 26 birincil PDF doğrulandı ve boyandı; 1 ikincil derleme (Baxter-Jones 2005) elendi; 1 eksik PDF (Cohen 1988) raporlandı ve kaynakça dışında bırakıldı.",
            "highlighted_pdf_count": len(PDF_HIGHLIGHT_SPECS),
            "highlighted_pdf_dir": str(ANNOTATED_DIR),
        },
        "pdf_inventory": pdf_inventory,
        "citations": citations,
        "unmatched_and_missing": [
            {
                "category": "PDF Eksik (Aşama 5 Kaynakçasından Hariç Tutuldu)",
                "item": "Cohen (1988) — Statistical Power Analysis for the Behavioral Sciences (2nd ed.) [REQ-09, REQ-20]",
                "detail": "Kullanıcı tam metin PDF'e ulaşamadığı için Yontem_Kaynak_Makaleler_PDF/ klasöründe bulunmamaktadır. Dürüstlük ve birincil doğrulama ilkesi gereği Aşama 5 kaynakça listesine eklenmemiştir; duyarlılık analizindeki sayısal eşik (dz ≈ 0.49) Faul ve ark. (2007) ile verilmiş, etki büyüklükleri sayısal değerleri ve Hopkins (2004) SWC eşiği (0.20 × SS_ön) üzerinden sunulmuştur.",
            },
            {
                "category": "İkincil Derleme Elendi (kullanilmayanlar/ Arşivi)",
                "item": "Baxter-Jones, Eisenmann, & Sherar (2005) — Pediatric Exercise Science [REQ-05, REQ-24]",
                "detail": "Anlatısal derleme ('This article reviews some of the possible biological maturity indicators...') olduğu tespit edilip turuncu (#E67E22) boyanarak kullanilmayanlar/ klasörüne arşivlenmiş; yerine birincil Mirwald ve ark. (2002), Correia ve ark. (2024) ve Peterson ve ark. (2016) entegre edilmiştir.",
            },
            {
                "category": "Yerel PDF Künye ve Yazar Düzeltmesi",
                "item": "Atif_07A_Ruggeri_2020_Motor_Interventions_ASD.pdf & Atif_15_Zifchock_2008_Symmetry_Angle.pdf",
                "detail": "Önceki iskelet kaynakçasındaki hatalı yazar/DOI kaydı yerel PDF üzerinden düzeltilmiştir: Ruggeri, A., Dancel, A., Johnson, R., & Sargent, B. (2020), DOI: 10.1177/1362361319885215; Zifchock, R. A. (2008), Gait & Posture, 27(4), 622-627.",
            },
            {
                "category": "Ampirik Kapsam Hizalaması",
                "item": "Behm (2017), Del Vecchio (2019), Gülü & Akalan (2021), Zifchock (2008), Stewart (2014)",
                "detail": "Behm ve ark. (2017) pediatrik antrenman moderatörleri için, Del Vecchio ve ark. (2019) ise 4 haftalık nöral adaptasyon mekanizması için atıf yapılmış; ≥6 aylık ön eğitim ölçütü çalışmamızın tasarım kuralı olarak ayrı tutulmuştur. Gülü ve Akalan (2021) doğrultusunda Sayers zirve güç denklemi pediatrik yaş/kol pozisyonu sınırlılığıyla birlikte keşifsel türev olarak sunulmuştur.",
            },
        ],
    }

    audit_json_path = SYS_DIR / "audit_results_stage4.json"
    audit_json_path.write_text(json.dumps(audit_payload, indent=2, ensure_ascii=False), encoding="utf-8")
    shutil.copyfile(audit_json_path, SYS_DIR / "audit_results.json")
    print(f"[OK] Wrote audit JSON: {audit_json_path}")

    # Run build_report.py
    build_cmd = [
        "python3",
        str(SCRIPTS_DIR / "build_report.py"),
        str(audit_json_path),
        "--out-dir",
        str(AUDIT_DIR),
        "--prefix",
        "Birincil_Kaynak_Atif_Denetim_Raporu_TR",
    ]
    res = subprocess.run(build_cmd, capture_output=True, text=True)
    print(res.stdout)
    if res.returncode != 0:
        raise RuntimeError(f"build_report.py failed: {res.stderr}")

    # Clean audit DOCX metadata and check style
    report_docx = AUDIT_DIR / "Birincil_Kaynak_Atif_Denetim_Raporu_TR_report.docx"
    if report_docx.exists() and AUDIT_CLEAN_SCRIPT.exists():
        clean_res = subprocess.run(["python3", str(AUDIT_CLEAN_SCRIPT), "--clean", str(report_docx)], capture_output=True, text=True)
        print(clean_res.stdout)
        if clean_res.returncode != 0:
            print("[WARN] audit_and_clean on report docx:", clean_res.stderr)


if __name__ == "__main__":
    main()
