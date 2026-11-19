#!/usr/bin/env python3
"""
Phase 5, 6 & 7 Complete Citation-Audit Pipeline for autism_2026_testtekrar.
- Loads Phase 0/3 extracted JSONs from 05_citation_audit_boyali_pdfler/_sistem_verileri/
- Builds complete 44-unit audit_results.json across [Atıf Gerekli - 1] .. [Atıf Gerekli - 24]
- Highlights all 34 highlightable Atif_* PDFs into 05_citation_audit_boyali_pdfler/Boyali_PDFler/<stem>_annotated.pdf
  using full-paragraph / full-sentence / full-row bounding boxes (strict Core Principle 8 compliance)
- Renders verification PNGs into 05_citation_audit_boyali_pdfler/_sistem_verileri/dogrulama_png/
- Generates DOCX, PDF, XLSX, and CSV audit reports in 05_citation_audit_boyali_pdfler/
"""

import json
import re
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

XML_CTRL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")

SKILL_SCRIPTS = Path("/Users/alik/Desktop/anti_projeler/articles_all/.agents/skills/citation-audit/scripts")
sys.path.insert(0, str(SKILL_SCRIPTS))
from highlight_pdf import annotate_pdf  # noqa: E402

PROJECT_DIR = Path("/Users/alik/Desktop/anti_projeler/articles_all/autism_2026_testtekrar")
PDF_DIR = PROJECT_DIR / "Yontem_Kaynak_Makaleler_PDF"
OUT_ROOT = PROJECT_DIR / "05_citation_audit_boyali_pdfler"
SYS_DIR = OUT_ROOT / "_sistem_verileri"
BOYALI_DIR = OUT_ROOT / "Boyali_PDFler"
PNG_DIR = SYS_DIR / "dogrulama_png"

BOYALI_DIR.mkdir(parents=True, exist_ok=True)
PNG_DIR.mkdir(parents=True, exist_ok=True)


def load_extracted_map() -> dict[str, dict]:
    mp = {}
    for jf in sorted(SYS_DIR.glob("*_extracted.json")):
        data = json.loads(jf.read_text(encoding="utf-8"))
        mp[data["file_name"]] = data
    return mp


def get_block_info(ext: dict, page: int, block_nos: list[int]) -> tuple[list[list[float]], str]:
    """Returns exact bboxes and concatenated verbatim text for the given (page, block_nos)."""
    bboxes = []
    texts = []
    block_lookup = {(b["page"], b["block_no"]): b for b in ext["blocks"]}
    for bno in block_nos:
        b = block_lookup.get((page, bno))
        if b:
            bboxes.append([round(float(v), 2) for v in b["bbox"]])
            clean_t = XML_CTRL_RE.sub(" ", b["text"])
            texts.append(" ".join(clean_t.split()))
    return bboxes, " ".join(texts)


def main() -> None:
    ext_map = load_extracted_map()
    state = json.loads((SYS_DIR / "state.json").read_text(encoding="utf-8"))

    # Helper to fetch PDF filename by prefix
    def pdf_by_prefix(pref: str) -> str:
        for fn in sorted(ext_map.keys()):
            if fn.startswith(pref):
                return fn
        raise KeyError(f"No PDF starting with {pref}")

    # Build pdf_inventory for audit_results.json
    pdf_inventory = {}
    for fname, edata in sorted(ext_map.items()):
        pdf_inventory[fname] = {
            "page_count": edata["page_count"],
            "sha256": edata["sha256"],
            "has_text_layer": edata["has_text_layer"],
            "dois": edata.get("dois", []),
        }

    # Define all 44 atomic citation audit entries and their multi-page highlight targets
    # Each highlight target in `_hl_targets` is {"page": int, "blocks": list[int], "exact_bboxes": bool (optional)}
    raw_specs = [
        {
            "citation_id": "A001",
            "citation_marker": "[Atıf Gerekli - 1] -> von Elm ve ark. (2007)",
            "manuscript_location": "Bölüm 2.1 (Araştırma Modeli ve Raporlama Standartları)",
            "claim_type": "Metodolojik Raporlama Rehberi",
            "claim_text": "Çalışmanın metodolojik kurgusu ve raporlama süreçleri, Gözlemsel Kohort Araştırmalarının Raporlanmasının Güçlendirilmesi (STROBE) rehberi ilkelerine tam uyumlu olarak yürütülmüştür.",
            "source_ref": "von Elm, E., Altman, D. G., Egger, M., Pocock, S. J., Gøtzsche, P. C., & Vandenbroucke, J. P. (2007). PLoS Medicine, 4(10), e296.",
            "pdf_prefix": "Atif_01_",
            "page_and_section": "Sayfa 1 (Özet) ve Sayfa 2 (Aims and Use of the STROBE Statement)",
            "decision_label": "Tam destek",
            "confidence": "Yüksek",
            "primary_status": "Birincil",
            "secondary_attribution_proof": "",
            "final_status": "Uygun",
            "rationale": "STROBE Girişimi'nin resmi birincil kılavuz makalesidir; kohort, olgu-kontrol ve kesitsel gözlemsel çalışmaların raporlama standartlarını doğrudan tanımlar (Not: Kaynakçadaki dergi bilgisi eldeki PDF ile uyumlu olacak şekilde PLoS Medicine olarak güncellenmiştir).",
            "recommended_action": "Birincil kaynak olarak makale taslağına entegre edildi (von Elm ve ark., 2007).",
            "hl_targets": [
                {"page": 1, "blocks": [12]},
                {"page": 2, "blocks": [5]},
            ],
        },
        {
            "citation_id": "A002",
            "citation_marker": "[Atıf Gerekli - 2] -> Slade ve ark. (2016)",
            "manuscript_location": "Bölüm 2.1 ve Bölüm 2.3 (CERT Egzersiz Raporlama Şablonu)",
            "claim_type": "Metodolojik Raporlama Rehberi",
            "claim_text": "Egzersiz Raporlama Uzlaşı Şablonu (CERT) kontrol listesi ilkelerine tam uyumlu olarak yürütülmüş; haftalık sıklık, seans süresi, eğitmen-katılımcı oranı ve bireyselleştirilmiş ilerleme kuralları CERT doğrultusunda yapılandırılmıştır.",
            "source_ref": "Slade, S. C., Dionne, C. E., Underwood, M., & Buchbinder, R. (2016). Consensus on Exercise Reporting Template (CERT). British Journal of Sports Medicine, 50(23), 1428–1437.",
            "pdf_prefix": "Atif_02_",
            "page_and_section": "Sayfa 1 (Madde 1–11) ve Sayfa 2 (Madde 12–16b, Dozaj ve Bireyselleştirme)",
            "decision_label": "Tam destek",
            "confidence": "Yüksek",
            "primary_status": "Birincil",
            "secondary_attribution_proof": "",
            "final_status": "Uygun",
            "rationale": "Resmi 16 maddelik CERT kontrol listesi belgesidir; ekipman (Madde 1), eğitmen niteliği (Madde 2), grup/bireysel uygulama (Madde 3–4), ilerleme karar kuralı (Madde 7a–7b), dozaj/sıklık/süre (Madde 13) ve bireyselleştirme (Madde 14a–14b) kriterlerini birincil olarak ortaya koymaktadır.",
            "recommended_action": "Birincil kaynak olarak makale taslağına entegre edildi (Slade ve ark., 2016).",
            "hl_targets": [
                {"page": 1, "blocks": [0, 1, 15, 30, 33]},
                {"page": 2, "blocks": [3, 4, 8, 9, 12, 29, 30]},
            ],
        },
        {
            "citation_id": "A003",
            "citation_marker": "[Atıf Gerekli - 3] -> Škarabot ve ark. (2021)",
            "manuscript_location": "Bölüm 2.2 (Katılımcılar — ≥6 Ay Antrenman Geçmişi ve Erken Nöral Adaptasyon Dışlama)",
            "claim_type": "Fizyolojik Mekanizma / Dahil Etme Kriteri",
            "claim_text": "Çalışma öncesinde en az 6 ay süreyle düzenli hareket eğitimi programına devam etmiş olmak (erken dönem nöromüsküler öğrenme etkisini dışlamak ve bazal platoya ulaşmış bir örneklemde ek antrenman etkisini izole etmek amacıyla).",
            "source_ref": "Škarabot, J., Brownstein, C. G., Casolo, A., Del Vecchio, A., & Ansdell, P. (2021). European Journal of Applied Physiology, 121(3), 675–685.",
            "pdf_prefix": "Atif_03B_",
            "page_and_section": "Sayfa 1–2 (Invited Review — Introduction)",
            "decision_label": "Birincil değil",
            "confidence": "Yüksek",
            "primary_status": "Birincil değil",
            "secondary_attribution_proof": "Following a period of resistance training, the initial increases in force production (< 2–4 weeks) are thought to be primarily underpinned by neural adaptations (Moritani and DeVries 1979; for reviews see Enoka 1988; Sale 1988; Carroll et al. 2001; Folland and Williams 2007).",
            "final_status": "Uygun değil",
            "rationale": "Makale davetli anlatısal derlemedir ('INVITED REVIEW'); erken dönem kuvvet artışlarının nöral adaptasyonlardan kaynaklandığı iddiasını Moritani & DeVries (1979), Enoka (1988) ve Sale (1988) kaynaklarına atfetmektedir. Ayrıca OSB'de ≥6 ay bazal plato kriterini test etmemektedir.",
            "recommended_action": "Makale taslağında 'BİRİNCİL KAYNAK DEĞİL' olarak işaretlendi; erken nöral adaptasyon için birincil deneysel kaynak (örn. Moritani & DeVries, 1979) temin edilmeli veya metodolojik gerekçe kaynak atfı olmadan kendi çalışma tasarım kuralımız olarak ifade edilmelidir.",
            "hl_targets": [
                {"page": 1, "blocks": [10]},
                {"page": 2, "blocks": [2]},
            ],
        },
        {
            "citation_id": "A004",
            "citation_marker": "[Atıf Gerekli - 3] -> Pan ve ark. (2017)",
            "manuscript_location": "Bölüm 2.2 (Katılımcılar — ≥6 Ay Antrenman Geçmişi ve Bazal Plato)",
            "claim_type": "Metodolojik Dahil Etme Kriteri",
            "claim_text": "Çalışma öncesinde en az 6 ay süreyle düzenli hareket eğitimi programına devam etmiş olmak (erken dönem nöromüsküler öğrenme etkisini dışlamak ve bazal platoya ulaşmış bir örneklemde ek antrenman etkisini izole etmek amacıyla).",
            "source_ref": "Pan, C. Y., Chu, C. H., Tsai, C. L., Sung, M. C., Huang, C. Y., & Ma, W. Y. (2017). Autism, 21(2), 190–202.",
            "pdf_prefix": "Atif_03A_",
            "page_and_section": "Sayfa 4 (Methods — Participants) ve Sayfa 5 (Physical activity intervention)",
            "decision_label": "Kısmi destek",
            "confidence": "Yüksek",
            "primary_status": "Birincil",
            "secondary_attribution_proof": "",
            "final_status": "Düzeltme gerekli",
            "rationale": "Pan ve ark. (2017) OSB'li çocuklarda 12 haftalık masa tenisi/motor beceri müdahalesini inceleyen özgün birincil araştırmadır ve katılımcıların eşzamanlı rehabilitasyon aldığını raporlar; ancak dahil etme kriterlerinde (Sayfa 4) 'çalışma öncesi en az 6 ay düzenli hareket eğitimi alma veya bazal platoya ulaşma' şartı yer almamaktadır.",
            "recommended_action": "Makale taslağında 'KISMİ DESTEK (Kapsam Farkı)' olarak işaretlendi; yalnızca eşzamanlı terapi alan OSB'li çocuklarda ek motor müdahale bağlamı için kullanılmalı, ≥6 ay ön eğitim şartına kanıt olarak bağlanmamalıdır.",
            "hl_targets": [
                {"page": 4, "blocks": [2]},
                {"page": 5, "blocks": [1, 3]},
            ],
        },
        {
            "citation_id": "A005",
            "citation_marker": "[Atıf Gerekli - 4] -> American Psychiatric Association (2013) [DSM-5]",
            "manuscript_location": "Bölüm 2.2 (Katılımcılar — DSM-5 Resmi Klinik OSB Tanısı)",
            "claim_type": "Tanı Kriteri / Klinik Sınıflama",
            "claim_text": "Katılımcı dahil etme ölçütleri: (1) DSM-5 tanı ölçütlerine göre uzman çocuk ve ergen psikiyatristi tarafından konulmuş resmi OSB tanısına sahip olmak.",
            "source_ref": "American Psychiatric Association. (2013). Diagnostic and statistical manual of mental disorders (5th ed.).",
            "pdf_prefix": None,
            "page_and_section": "PDF mevcut değil",
            "decision_label": "Doğrulanamadı",
            "confidence": "Yüksek",
            "primary_status": "Doğrulanamadı (PDF eksik)",
            "secondary_attribution_proof": "",
            "final_status": "Doğrulanamadı",
            "rationale": "DSM-5 el kitabı OSB tanı ölçütlerinin resmi birincil kaynağıdır; ancak Yontem_Kaynak_Makaleler_PDF klasöründe PDF dosyası bulunmadığından metin katmanı üzerinden doğrudan pasaj doğrulaması yapılamamıştır.",
            "recommended_action": "Makale taslağında 'DOĞRULANAMADI (Birincil Kaynak PDF Eksik: APA, 2013 DSM-5)' olarak işaretlendi; resmi tanı el kitabı PDF'i/bölümü doğrulama arşivine eklendiğinde entegre edilebilir.",
            "hl_targets": [],
        },
        {
            "citation_id": "A006",
            "citation_marker": "[Atıf Gerekli - 4] -> Lord ve ark. (2018)",
            "manuscript_location": "Bölüm 2.2 (Katılımcılar — DSM-5 Resmi Klinik OSB Tanısı)",
            "claim_type": "Tanı Kriteri / Klinik Sınıflama",
            "claim_text": "Katılımcı dahil etme ölçütleri: (1) DSM-5 tanı ölçütlerine göre uzman çocuk ve ergen psikiyatristi tarafından konulmuş resmi OSB tanısına sahip olmak.",
            "source_ref": "Lord, C., Elsabbagh, M., Baird, G., & Veenstra-Vanderweele, J. (2018). Autism spectrum disorder. The Lancet, 392(10146), 508–520.",
            "pdf_prefix": "Atif_04_",
            "page_and_section": "Sayfa 1 (Signs, symptoms, and general diagnostic issues) ve Sayfa 2 (Panel 1)",
            "decision_label": "Birincil değil",
            "confidence": "Yüksek",
            "primary_status": "Birincil değil",
            "secondary_attribution_proof": "Panel 1: Signs and symptoms of autism spectrum disorder as described in DSM-5 (299.0)^6 ... DSM-5 criteria,^6 published in 2013.",
            "final_status": "Uygun değil",
            "rationale": "Makale The Lancet dergisinde yayımlanmış klinik derlemedir ('Seminar'); DSM-5 tanı ölçütlerini 6 numaralı kaynak (American Psychiatric Association, 2013) üzerinden aktarmaktadır.",
            "recommended_action": "Makale taslağında 'BİRİNCİL KAYNAK DEĞİL' olarak işaretlendi; DSM-5 tanı kriterleri için doğrudan APA (2013) birincil kaynağı kullanılmalıdır.",
            "hl_targets": [
                {"page": 1, "blocks": [8]},
                {"page": 2, "blocks": [9, 10, 4]},
            ],
        },
        {
            "citation_id": "A007",
            "citation_marker": "[Atıf Gerekli - 5] -> World Medical Association (2013)",
            "manuscript_location": "Bölüm 2.2 (Etik Kurul Onayı, Ebeveyn Yazılı Onamı ve Çocuk Sözlü Rızası)",
            "claim_type": "Etik Standart / Araştırma Etiği",
            "claim_text": "Çalışma protokolü Üniversite Girişimsel Olmayan Klinik Araştırmalar Etik Kurulu tarafından onaylanmış ve tüm süreçler Helsinki Bildirgesi'ne uygun olarak yürütülmüştür. Veri toplama öncesinde tüm ebeveynlerden yazılı aydınlatılmış onam, katılımcılardan ise gelişimsel düzeylerine uygun sözlü rıza alınmıştır.",
            "source_ref": "World Medical Association. (2013). World Medical Association Declaration of Helsinki: Ethical principles for medical research involving human subjects. JAMA, 310(20), 2191–2194.",
            "pdf_prefix": "Atif_05_",
            "page_and_section": "Sayfa 2 (Madde 23) ve Sayfa 3 (Madde 26, 28 ve 29)",
            "decision_label": "Tam destek",
            "confidence": "Yüksek",
            "primary_status": "Birincil",
            "secondary_attribution_proof": "",
            "final_status": "Uygun",
            "rationale": "Dünya Tıp Birliği'nin resmi 2013 Helsinki Bildirgesi metnidir; bağımsız etik kurul onayını (Madde 23), yazılı aydınlatılmış onamı (Madde 26), yasal temsilci onamını (Madde 28) ve katılımcının ek rızasını/assent (Madde 29) birincil norm olarak düzenlemektedir.",
            "recommended_action": "Birincil kaynak olarak makale taslağına entegre edildi (World Medical Association, 2013).",
            "hl_targets": [
                {"page": 2, "blocks": [21]},
                {"page": 3, "blocks": [2, 6, 7]},
            ],
        },
        {
            "citation_id": "A008",
            "citation_marker": "[Atıf Gerekli - 6] -> Faul ve ark. (2007)",
            "manuscript_location": "Bölüm 2.2 (Duyarlılık Güç Analizi — Sensitivity Power Analysis)",
            "claim_type": "İstatistiksel Yöntem / Yazılım ve Güç Analizi",
            "claim_text": "N = 34 örneklem büyüklüğü için yapılan duyarlılık güç analizi (α = 0.05, çift yönlü, 1 - β = 0.80), eşleştirilmiş ölçümlerde dz ≈ 0.49 büyüklüğündeki orta düzey zaman etkilerini saptamak için yeterli istatistiksel güce sahip olduğunu doğrulamıştır.",
            "source_ref": "Faul, F., Erdfelder, E., Lang, A. G., & Buchner, A. (2007). G*Power 3: A flexible statistical power analysis program for the social, behavioral, and biomedical sciences. Behavior Research Methods, 39(2), 175–191.",
            "pdf_prefix": "Atif_06_",
            "page_and_section": "Sayfa 3 (Sensitivity Analyses) ve Sayfa 6 (Tests for Means — Eşleştirilmiş t-testi dz)",
            "decision_label": "Tam destek",
            "confidence": "Yüksek",
            "primary_status": "Birincil",
            "secondary_attribution_proof": "",
            "final_status": "Uygun",
            "rationale": "G*Power 3 yazılımının ve beş güç analizi türünün orijinal tanıtım makalesidir; verilen α, 1-β ve N değerlerinden kritik etki büyüklüğünün hesaplandığı duyarlılık analizini (Sayfa 3) ve bağımlı örneklemler t-testinde dz parametresini (Sayfa 6) doğrudan tanımlar.",
            "recommended_action": "Duyarlılık güç analizi hesaplaması için birincil kaynak olarak makale taslağına entegre edildi (Faul ve ark., 2007).",
            "hl_targets": [
                {"page": 3, "blocks": [3]},
                {"page": 6, "blocks": [9]},
            ],
        },
        {
            "citation_id": "A009",
            "citation_marker": "[Atıf Gerekli - 6] -> Cohen (1988)",
            "manuscript_location": "Bölüm 2.2 (Duyarlılık Güç Analizi — dz ≈ 0.49 Orta Düzey Etki Yorumu)",
            "claim_type": "İstatistiksel Eşik / Yorumlama",
            "claim_text": "dz ≈ 0.49 büyüklüğündeki orta düzey (moderate) zaman etkilerini saptamak için yeterli istatistiksel güç.",
            "source_ref": "Cohen, J. (1988). Statistical power analysis for the behavioral sciences (2nd ed.). Lawrence Erlbaum Associates.",
            "pdf_prefix": None,
            "page_and_section": "PDF mevcut değil",
            "decision_label": "Doğrulanamadı",
            "confidence": "Yüksek",
            "primary_status": "Doğrulanamadı (PDF eksik)",
            "secondary_attribution_proof": "",
            "final_status": "Doğrulanamadı",
            "rationale": "d ≈ 0.50 değerinin 'orta düzey (medium/moderate)' etki olarak sınıflandırılmasının birincil kaynağı Cohen (1988) kitabıdır; ancak PDF klasöründe yer almadığı için metin içi doğrulama yapılamamıştır.",
            "recommended_action": "Makale taslağında 'DOĞRULANAMADI (Birincil Kaynak PDF Eksik: Cohen, 1988)' olarak işaretlendi.",
            "hl_targets": [],
        },
        {
            "citation_id": "A010",
            "citation_marker": "[Atıf Gerekli - 7] -> Ruggeri ve ark. (2020)",
            "manuscript_location": "Bölüm 2.3 (Mikro-Doz Antrenman Müdahalesi — 12 Hafta, Haftada 2 Gün, 60 dk Motor Beceri Dozajı)",
            "claim_type": "Sentez / Müdahale Dozaj Çerçevesi",
            "claim_text": "Katılımcılar 12 hafta boyunca haftada 2 gün ve günde 60 dakika süreyle yapılandırılmış hareket eğitimi ve motor beceri programına katılmıştır.",
            "source_ref": "Ruggeri, A., Dancel, A., Johnson, R., & Sargent, B. (2020). The effect of motor and physical activity intervention on motor outcomes of children with autism spectrum disorder: A systematic review. Autism, 24(3), 549–563.",
            "pdf_prefix": "Atif_07A_",
            "page_and_section": "Sayfa 14 (Results — Motor skill interventions)",
            "decision_label": "Tam destek",
            "confidence": "Yüksek",
            "primary_status": "Birincil",
            "secondary_attribution_proof": "",
            "final_status": "Uygun",
            "rationale": "Sistematik derlemenin kendi Bulgular (Results) sentezinde OSB'li çocuklara yönelik motor beceri müdahalelerinin 3–12 hafta sürdüğü, haftada 1–5 kez ve seans başına 40 dk–4 saat (toplam 9–40 seans) uygulandığı doğrudan ortaya konulmuştur (primary_source_criteria.md Madde 2.2 uyarınca sistematik derlemenin kendi sentez bulgusu birincildir).",
            "recommended_action": "OSB motor beceri müdahale dozajı çerçevesi için birincil sentez kaynağı olarak makale taslağına entegre edildi (Ruggeri ve ark., 2020).",
            "hl_targets": [
                {"page": 14, "blocks": [6]},
            ],
        },
        {
            "citation_id": "A011",
            "citation_marker": "[Atıf Gerekli - 7] -> Healy ve ark. (2018)",
            "manuscript_location": "Bölüm 2.3 (Mikro-Doz Antrenman Müdahalesi — 12 Hafta, Haftada 2 Gün, 60 dk)",
            "claim_type": "Meta-Analiz / Müdahale Dozaj Çerçevesi",
            "claim_text": "12 hafta boyunca haftada 2 gün (seanslar arası en az 48 saat toparlanma süresi) ve günde 60 dakika süreyle yapılandırılmış hareket eğitimi ve motor beceri programı.",
            "source_ref": "Healy, S., Nacario, A., Braithwaite, R. E., & Hopper, C. (2018). The effect of physical activity interventions on youth with autism spectrum disorder: A meta-analysis. Autism Research, 11(6), 818–833.",
            "pdf_prefix": "Atif_07B_",
            "page_and_section": "Sayfa 5 (Results — Outcome Analyses) ve Sayfa 7 (Moderator Analyses)",
            "decision_label": "Kısmi destek",
            "confidence": "Yüksek",
            "primary_status": "Birincil",
            "secondary_attribution_proof": "",
            "final_status": "Düzeltme gerekli",
            "rationale": "Healy ve ark. (2018) kendi meta-analiz bulgularında (birincil sentez) fiziksel aktivite müdahalelerinin kas kuvveti (g = 0.78), lokomotor (g = 1.60) ve manipülatif beceriler (g = 2.76) üzerinde anlamlı etkisini ve süre moderatörlerini (<10, 10–16, >16 hafta) raporlamaktadır; ancak seans başına 60 dakika veya seanslar arası 48 saat toparlanma kuralına ilişkin veri sunmamaktadır.",
            "recommended_action": "Makale taslağında 'KISMİ DESTEK (Kapsam Farkı)' olarak işaretlendi; seans sıklığı/48 saat kuralı yerine yalnızca OSB'de fiziksel aktivite müdahalelerinin genel motor/kuvvet etkinliği bağlamında kullanılmalıdır.",
            "hl_targets": [
                {"page": 5, "blocks": [5]},
                {"page": 7, "blocks": [2, 3]},
            ],
        },
        {
            "citation_id": "A012",
            "citation_marker": "[Atıf Gerekli - 8] -> Guadagnoli ve Lee (2004)",
            "manuscript_location": "Bölüm 2.3 (Optimal Zorluk Noktası / Challenge Point Framework ile Bireyselleştirilmiş İlerleme)",
            "claim_type": "Kuramsal Çerçeve / Motor Öğrenme Modeli",
            "claim_text": "Egzersiz zorluk düzeyi, Optimal Zorluk Noktası (Challenge Point Framework) ilkeleri doğrultusunda her çocuğun bireysel motor yeterliliğine göre kademeli olarak uyarlanmıştır.",
            "source_ref": "Guadagnoli, M. A., & Lee, T. D. (2004). Challenge point: A framework for conceptualizing the effects of various practice conditions in motor learning. Journal of Motor Behavior, 36(2), 212–224.",
            "pdf_prefix": "Atif_08_",
            "page_and_section": "Sayfa 3 (Task Difficulty and Skill), Sayfa 6 (Şekil 2 & Optimal Challenge Points) ve Sayfa 8 (Şekil 4)",
            "decision_label": "Tam destek",
            "confidence": "Yüksek",
            "primary_status": "Birincil",
            "secondary_attribution_proof": "",
            "final_status": "Uygun",
            "rationale": "Optimal Zorluk Noktası (Challenge Point Framework) kuramını literatüre kazandıran orijinal kuramsal makaledir; fonksiyonel görev zorluğunun bireyin beceri düzeyi yükseldikçe kademeli olarak artırılması gerektiğini doğrudan formüle etmektedir.",
            "recommended_action": "Birincil kuramsal kaynak olarak makale taslağına entegre edildi (Guadagnoli ve Lee, 2004).",
            "hl_targets": [
                {"page": 3, "blocks": [7]},
                {"page": 6, "blocks": [5, 6]},
                {"page": 8, "blocks": [2]},
            ],
        },
        {
            "citation_id": "A013",
            "citation_marker": "[Atıf Gerekli - 9] -> Atkinson ve Reilly (1996)",
            "manuscript_location": "Bölüm 2.4 (Sirkadiyen Ritim, Salon Sıcaklığı 20–22°C ve Günün Saati ±30 dk Kontrolü)",
            "claim_type": "Fizyolojik Mekanizma / Ölçüm Standardizasyonu",
            "claim_text": "Sirkadiyen ritim kaynaklı performans dalgalanmalarını kontrol altına almak amacıyla ön-test ve son-test oturumları her katılımcı için günün aynı zaman diliminde (10:00–12:00, ±30 dakika toleransla), standart kapalı spor salonu koşullarında (sıcaklık: 20–22°C) gerçekleştirilmiştir.",
            "source_ref": "Atkinson, G., & Reilly, T. (1996). Circadian variation in sports performance. Sports Medicine, 21(4), 292–312.",
            "pdf_prefix": "Atif_09B_",
            "page_and_section": "Sayfa 1 (Summary), Sayfa 4 (Bölüm 2) ve Sayfa 8 (Bölüm 4.3.1)",
            "decision_label": "Birincil değil",
            "confidence": "Yüksek",
            "primary_status": "Birincil değil",
            "secondary_attribution_proof": "1°C rise in body temperature, nerve conduction velocity increases by 2.4 m/sec.[52] ... Circadian rhythms in psychomotor performance, particularly tasks that entail cognitive operations seem especially prone to a post-lunch 'dip'.[5]",
            "final_status": "Uygun değil",
            "rationale": "Makale anlatısal derlemedir (Sports Medicine review); reaksiyon süresi, esneklik ve kas kuvvetindeki sirkadiyen değişimleri önceki birincil deneysel çalışmalara ([5], [52], [54] vb.) atfetmekte ve spesifik olarak '20–22°C salon sıcaklığı / ±30 dk tolerans' kuralı tanımlamamaktadır.",
            "recommended_action": "Makale taslağında 'BİRİNCİL KAYNAK DEĞİL' olarak işaretlendi; sirkadiyen performans değişimi için özgün deneysel çalışma eklenmeli veya ölçüm standardizasyonu kendi iç protokol kuralımız olarak atıfsız yazılmalıdır.",
            "hl_targets": [
                {"page": 1, "blocks": [10, 11]},
                {"page": 8, "blocks": [5, 6, 11]},
            ],
        },
        {
            "citation_id": "A014",
            "citation_marker": "[Atıf Gerekli - 9] -> Chtourou ve Souissi (2012)",
            "manuscript_location": "Bölüm 2.4 (Sirkadiyen Ritim, Salon Sıcaklığı 20–22°C ve Günün Saati ±30 dk Kontrolü)",
            "claim_type": "Fizyolojik Mekanizma / Ölçüm Standardizasyonu",
            "claim_text": "Sirkadiyen ritim kaynaklı performans dalgalanmalarını kontrol altına almak amacıyla ön-test ve son-test oturumları her katılımcı için günün aynı zaman diliminde (10:00–12:00, ±30 dakika toleransla) gerçekleştirilmiştir.",
            "source_ref": "Chtourou, H., & Souissi, N. (2012). The effect of training at a specific time of day: A review. Journal of Strength and Conditioning Research, 26(7), 1984–2005.",
            "pdf_prefix": "Atif_09A_",
            "page_and_section": "Sayfa 1 (Introduction) ve Sayfa 2 (Diurnal Variations of Sport Performances)",
            "decision_label": "Birincil değil",
            "confidence": "Yüksek",
            "primary_status": "Birincil değil",
            "secondary_attribution_proof": "In this context, it has been well established that short-term maximal performances fluctuate with time of day, with morning nadirs and afternoon maximum values (16–18,108–111) ... in adults (16–19,21) or youth subjects (105,106).",
            "final_status": "Uygun değil",
            "rationale": "Başlığında açıkça belirtildiği üzere anlatısal derleme makalesidir ('A REVIEW'); gün içi anaerobik güç ve kuvvet dalgalanmalarını önceki deneysel çalışmalara (16–21, 105–111) atfetmektedir.",
            "recommended_action": "Makale taslağında 'BİRİNCİL KAYNAK DEĞİL' olarak işaretlendi; çocuk/ergenlerde günün saati etkisini doğrudan test eden birincil deneysel kaynak (örn. Souissi ve ark., 2010/2012) tercih edilmelidir.",
            "hl_targets": [
                {"page": 1, "blocks": [8, 9]},
                {"page": 2, "blocks": [0, 1]},
            ],
        },
        {
            "citation_id": "A015",
            "citation_marker": "[Atıf Gerekli - 10] -> Nuttall (2015)",
            "manuscript_location": "Bölüm 2.4.1 (ISAK Antropometri Protokolü — Boy, Vücut Kütlesi, Kulaç, Karış ve VKİ = kg/m²)",
            "claim_type": "Ölçüm Protokolü ve Matematiksel Formül",
            "claim_text": "Antropometrik ölçümler ISAK protokollerine göre alınmış; kulaç uzunluğu ve el karış uzunluğu ölçülmüş; VKİ = Vücut Kütlesi (kg) / [Boy Uzunluğu (m)]² formülüyle hesaplanmıştır.",
            "source_ref": "Nuttall, F. Q. (2015). Body mass index: Obesity, BMI, and health: A critical review. Nutrition Today, 50(3), 117–128.",
            "pdf_prefix": "Atif_10A_",
            "page_and_section": "Sayfa 3 (Adoption of the BMI as an Index of Obesity)",
            "decision_label": "Birincil değil",
            "confidence": "Yüksek",
            "primary_status": "Birincil değil",
            "secondary_attribution_proof": "In 1972, Keys et al^16 severely criticized the validity of Metropolitan Life Insurance published data ... Keys et al coined the term body mass index ... the ratio of weight to height squared (Quetelet index).",
            "final_status": "Uygun değil",
            "rationale": "Makale eleştirel anlatısal derlemedir ('A Critical Review'); VKİ formülünü Quetelet (1835) ve Keys ve ark. (1972) kaynaklarına atfetmekte olup ISAK kulaç/karış ölçüm protokollerine dair hiçbir içerik barındırmamaktadır.",
            "recommended_action": "Makale taslağında 'BİRİNCİL KAYNAK DEĞİL' olarak işaretlendi; ISAK resmi ölçüm el kitabı (örn. Marfell-Jones ve ark., 2012 / Stewart ve ark., 2011) birincil kaynak olarak sağlanmalıdır.",
            "hl_targets": [
                {"page": 3, "blocks": [0, 1]},
            ],
        },
        {
            "citation_id": "A016",
            "citation_marker": "[Atıf Gerekli - 10] -> World Health Organization / de Onis ve ark. (2007)",
            "manuscript_location": "Bölüm 2.4.1 (ISAK Antropometri Protokolü ve VKİ Hesaplaması)",
            "claim_type": "Referans Büyüme Standardı / Antropometri",
            "claim_text": "Antropometrik ölçümler ISAK protokollerine göre alınmış; kulaç uzunluğu, el karış uzunluğu ve VKİ = kg/m² hesaplanmıştır.",
            "source_ref": "de Onis, M., Onyango, A. W., Borghi, E., Siyam, A., Nishida, C., & Siekmann, J. (2007). Development of a WHO growth reference for school-aged children and adolescents. Bulletin of the World Health Organization, 85(9), 660–667.",
            "pdf_prefix": "Atif_10B_",
            "page_and_section": "Sayfa 1 (Objective & Introduction) ve Sayfa 2 (Reconstruction of 5–19 yr BMI reference)",
            "decision_label": "Kısmi destek",
            "confidence": "Yüksek",
            "primary_status": "Birincil",
            "secondary_attribution_proof": "",
            "final_status": "Düzeltme gerekli",
            "rationale": "5–19 yaş okul çağı çocuk ve ergenler için DSÖ (WHO) VKİ-yaş büyüme referans eğrilerini oluşturan resmi birincil çalışmadır; ancak ISAK antropometrik ölçüm tekniklerini (kulaç uzunluğu, oturma boyu, el karış uzunluğu) içermemektedir.",
            "recommended_action": "Makale taslağında 'KISMİ DESTEK (Yalnızca 5–19 Yaş DSÖ VKİ Referansı İçin Birincil; ISAK Protokolü Eksik)' olarak işaretlendi.",
            "hl_targets": [
                {"page": 1, "blocks": [1, 11]},
                {"page": 2, "blocks": [3]},
            ],
        },
        {
            "citation_id": "A017",
            "citation_marker": "[Atıf Gerekli - 11] -> Mirwald ve ark. (2002)",
            "manuscript_location": "Bölüm 2.4.1 (Mirwald Erkek PHV Olgunluk Sapması Denklemi)",
            "claim_type": "Matematiksel Regresyon Denklemi",
            "claim_text": "Biyolojik olgunlaşma durumu, oturma boyu, bacak uzunluğu, kronolojik yaş, boy uzunluğu ve vücut kütlesi kullanılarak Mirwald erkek zirve boy hızı (PHV) regresyon denklemi ile hesaplanmıştır: Olgunluk Sapması (yıl) = -9.236 + [0.0002708 × (Bacak Uzunluğu × Oturma Boyu)] - [0.001663 × (Yaş × Bacak Uzunluğu)] + [0.007216 × (Yaş × Oturma Boyu)] + [0.02292 × (Vücut Kütlesi / Boy Uzunluğu × 100)].",
            "source_ref": "Mirwald, R. L., Baxter-Jones, A. D. G., Bailey, D. A., & Beunen, G. P. (2002). An assessment of maturity from anthropometric measurements. Medicine & Science in Sports & Exercise, 34(4), 689–694.",
            "pdf_prefix": "Atif_11B_",
            "page_and_section": "Sayfa 1 (Özet), Sayfa 3 (Results — Eq. 3 başlangıcı) ve Sayfa 4 (Eq. 3 devamı, R² = 0.891, SEE = 0.592)",
            "decision_label": "Tam destek",
            "confidence": "Yüksek",
            "primary_status": "Birincil",
            "secondary_attribution_proof": "",
            "final_status": "Uygun",
            "rationale": "Erkek çocuklar için girişimsel olmayan PHV olgunluk sapması regresyon denklemini (Eq. 3) türeten ve çapraz doğrulayan orijinal çalışmadır; makaledeki tüm katsayılar (-9.236, +0.0002708, -0.001663, +0.007216, +0.02292) PDF'teki Eq. 3 ile birebir örtüşmektedir.",
            "recommended_action": "Birincil kaynak olarak makale taslağına entegre edildi (Mirwald ve ark., 2002).",
            "hl_targets": [
                {"page": 1, "blocks": [5]},
                {"page": 3, "blocks": [9]},
                {"page": 4, "blocks": [0]},
            ],
        },
        {
            "citation_id": "A018",
            "citation_marker": "[Atıf Gerekli - 11] -> Malina ve ark. (2015)",
            "manuscript_location": "Bölüm 2.4.1 (Mirwald Erkek PHV Olgunluk Sapması Denklemi)",
            "claim_type": "Matematiksel Regresyon Denklemi",
            "claim_text": "Biyolojik olgunlaşma durumu Mirwald erkek PHV regresyon denklemi ile hesaplanmış ve Tahmini PHV Yaşı = Kronolojik Yaş - Olgunluk Sapması olarak türetilmiştir.",
            "source_ref": "Malina, R. M., Rogol, A. D., Cumming, S. P., Coelho e Silva, M. J., & Figueiredo, A. J. (2015). Biological maturation of youth athletes: Assessment and implications. British Journal of Sports Medicine, 49(13), 852–859.",
            "pdf_prefix": "Atif_11A_",
            "page_and_section": "Sayfa 9–10 (Non-Invasive Estimates — kaynak [30] atfı)",
            "decision_label": "Birincil değil",
            "confidence": "Yüksek",
            "primary_status": "Birincil değil",
            "secondary_attribution_proof": "Sex-specific equations incorporating CA, height, weight, sitting height and estimated leg length are used to predict maturity offset.[30] Age at PHV is estimated as CA minus offset.",
            "final_status": "Uygun değil",
            "rationale": "Makale derleme niteliğindedir; olgunluk sapması denklemlerini doğrudan [30] numaralı kaynağa (Mirwald ve ark., 2002) atfetmektedir. Denklem için birincil kaynak zaten mevcut olduğundan (Mirwald ve ark., 2002) ikincil derleme atfına gerek yoktur.",
            "recommended_action": "Makale taslağında 'BİRİNCİL KAYNAK DEĞİL' olarak işaretlendi; denklem atfında yalnızca Mirwald ve ark. (2002) korunmalıdır.",
            "hl_targets": [
                {"page": 9, "blocks": [23, 24]},
                {"page": 10, "blocks": [2, 3, 4, 5, 6, 7, 8, 9]},
            ],
        },
        {
            "citation_id": "A019",
            "citation_marker": "[Atıf Gerekli - 12] -> Magill ve Anderson (2017)",
            "manuscript_location": "Bölüm 2.4.2 (Elektronik LED Görsel Basit Reaksiyon Süresi Testi — 1 Deneme + 3 Ölçüm Ortalaması)",
            "claim_type": "Ölçüm Protokolü",
            "claim_text": "Görsel basit reaksiyon süresi (sn), elektronik LED ışıklı reaksiyon ölçüm sistemi kullanılarak baskın el ile değerlendirilmiştir. Bir deneme ölçümünün ardından rastgele zaman aralıklı 3 geçerli ölçüm alınmış ve bu 3 ölçümün aritmetik ortalaması kaydedilmiştir.",
            "source_ref": "Magill, R. A., & Anderson, D. I. (2017). Motor learning and control: Concepts and applications (11th ed.). McGraw-Hill.",
            "pdf_prefix": None,
            "page_and_section": "PDF mevcut değil (Ders kitabı)",
            "decision_label": "Doğrulanamadı",
            "confidence": "Yüksek",
            "primary_status": "Doğrulanamadı (PDF eksik)",
            "secondary_attribution_proof": "",
            "final_status": "Doğrulanamadı",
            "rationale": "Klasörde PDF dosyası bulunmamaktadır; ayrıca primary_source_criteria.md Madde 3.2 uyarınca ders kitapları deneysel ölçüm protokolleri için kural olarak birincil kaynak kabul edilmez.",
            "recommended_action": "Makale taslağında 'DOĞRULANAMADI / DERS KİTABI (PDF Eksik ve Birincil Kaynak Değil)' olarak işaretlendi; kullanılan LED cihazın geçerlik-güvenirlik çalışması eklenmelidir.",
            "hl_targets": [],
        },
        {
            "citation_id": "A020",
            "citation_marker": "[Atıf Gerekli - 12] -> Laurenti ve ark. (2026)",
            "manuscript_location": "Bölüm 2.4.2 (Elektronik LED Görsel Basit Reaksiyon Süresi Testi — 1 Deneme + 3 Ölçüm Ortalaması)",
            "claim_type": "Ölçüm Protokolü",
            "claim_text": "Görsel basit reaksiyon süresi (sn), elektronik LED ışıklı reaksiyon ölçüm sistemi kullanılarak baskın el ile değerlendirilmiştir. Bir deneme ölçümünün ardından rastgele zaman aralıklı 3 geçerli ölçüm alınmış ve 3 ölçümün aritmetik ortalaması saniye (sn) cinsinden kaydedilmiştir.",
            "source_ref": "Laurenti, F., Presta, V., Compiani, M., et al. (2026). The effects of a sport-based training program on reaction time and fine motor coordination in children with autism spectrum disorder: A pilot study. Sports, 14(2), 80.",
            "pdf_prefix": "Atif_12A_",
            "page_and_section": "Sayfa 3–4 (2.2.1. Reactive Tests) ve Sayfa 6 (3. Results)",
            "decision_label": "Kısmi destek",
            "confidence": "Yüksek",
            "primary_status": "Birincil",
            "secondary_attribution_proof": "Two reactive tests were conducted using a lighting system (WittySEM, Microgate Srl, Bolzano, Italy), typically used to investigate reaction time and speed processing in sport settings [22]... The reaction time was assessed with a simple and a complex reactive test [26].",
            "final_status": "Düzeltme gerekli",
            "rationale": "Laurenti ve ark. (2026) OSB'li çocuklarda LED tabanlı görsel reaksiyon sistemini (WittySEM) saniye cinsinden uygulayan birincil pilot araştırmadır; ancak tekli uyaran ve '3 ölçüm ortalaması' yerine ardışık 10 yeşil LED uyaranının toplam süresini (~16.5 sn) ve '2 denemenin en iyisini' kaydetmiş, ayrıca protokolü [22, 26] kaynaklarına atfetmiştir.",
            "recommended_action": "Makale taslağında 'KISMİ DESTEK (Protokol Farkı: 10 Uyaran Toplamı ve En İyi/2 Deneme vs. Tek Uyaran ve Ortalama/3 Deneme)' olarak işaretlendi.",
            "hl_targets": [
                {"page": 3, "blocks": [7]},
                {"page": 4, "blocks": [1, 3]},
                {"page": 6, "blocks": [4]},
            ],
        },
        {
            "citation_id": "A021",
            "citation_marker": "[Atıf Gerekli - 13] -> Fess (1992) [ASHT El Kitabı]",
            "manuscript_location": "Bölüm 2.4.2 (Jamar El Dinamometresi ve ASHT Standart Oturma Pozisyonu — 3 Ölçüm, 60 sn Dinlenme)",
            "claim_type": "Ölçüm Protokolü",
            "claim_text": "İzometrik el kavrama kuvveti (kg), kalibre edilmiş hidrolik el dinamometresi (Jamar) ile Amerikan El Terapistleri Derneği (ASHT) standart pozisyonunda (dirsek 90° fleksiyonda, önkol nötral) her el için 3 kez ve 60 sn dinlenme ile ölçülmüştür.",
            "source_ref": "Fess, E. E. (1992). Grip strength. In Clinical assessment recommendations (2nd ed., pp. 41–45). American Society of Hand Therapists.",
            "pdf_prefix": None,
            "page_and_section": "PDF mevcut değil",
            "decision_label": "Doğrulanamadı",
            "confidence": "Yüksek",
            "primary_status": "Doğrulanamadı (PDF eksik)",
            "secondary_attribution_proof": "",
            "final_status": "Doğrulanamadı",
            "rationale": "ASHT standart kavrama pozisyonunun resmi birincil kaynağı Fess (1992) el kitabıdır; ancak PDF klasöründe yer almadığı için metin katmanından doğrulanamamıştır.",
            "recommended_action": "Makale taslağında 'DOĞRULANAMADI (Birincil Kaynak PDF Eksik: Fess, 1992 ASHT)' olarak işaretlendi.",
            "hl_targets": [],
        },
        {
            "citation_id": "A022",
            "citation_marker": "[Atıf Gerekli - 13] -> Roberts ve ark. (2011)",
            "manuscript_location": "Bölüm 2.4.2 (Jamar El Dinamometresi ve ASHT Standart Oturma Pozisyonu — 3 Ölçüm, 60 sn Dinlenme)",
            "claim_type": "Ölçüm Protokolü",
            "claim_text": "İzometrik el kavrama kuvveti (kg), Jamar dinamometresi ile ASHT standart pozisyonunda (sandalyede dik oturuş, dirsek 90° fleksiyonda, önkol nötral) her el için 3 kez ve ölçümler arasında 60 sn dinlenme verilerek ölçülmüştür.",
            "source_ref": "Roberts, H. C., Denison, H. J., Martin, H. J., Patel, H. P., Syddall, H., Cooper, C., & Sayer, A. A. (2011). A review of the measurement of grip strength in clinical and epidemiological studies. Age and Ageing, 40(4), 423–429.",
            "pdf_prefix": "Atif_13A_",
            "page_and_section": "Sayfa 4 (Posture, Interval between measurements, Number of assessments)",
            "decision_label": "Birincil değil",
            "confidence": "Yüksek",
            "primary_status": "Birincil değil",
            "secondary_attribution_proof": "The American Society of Hand Therapists (ASHT) recommends standardised positioning: subject seated, shoulders adducted and neutrally rotated, elbow flexed at 90°, forearm in neutral and wrist between 0 and 30° of dorsiflexion and 0 and 15° of ulnar deviation [12, 38]... Watanabe et al. [43] compared ... 1 min intervals... The ASHT protocol uses the mean of three trials of grip strength in each hand [38].",
            "final_status": "Uygun değil",
            "rationale": "Makale derleme çalışmasıdır ('A review of the measurement of grip strength'); ASHT duruş ve 3 ölçüm kuralını [38] numaralı kaynağa (Fess, 1992), 1 dakikalık dinlenme aralığını ise [43] numaralı kaynağa (Watanabe ve ark., 2005) atfetmektedir. Üstelik ASHT 3 ölçüm ortalamasını önerirken çalışmamızda en yüksek değer (maksimum) alınmıştır.",
            "recommended_action": "Makale taslağında 'BİRİNCİL KAYNAK DEĞİL' olarak işaretlendi; ASHT pozisyonu için doğrudan Fess (1992) ve pediatrik maksimum deneme tercihi için ilgili birincil geçerlik çalışması kullanılmalıdır.",
            "hl_targets": [
                {"page": 4, "blocks": [10, 14, 23]},
            ],
        },
        {
            "citation_id": "A023",
            "citation_marker": "[Atıf Gerekli - 14] -> Jaric (2002)",
            "manuscript_location": "Bölüm 2.4.2 (Vücut Kütlesine Göre Normalize Edilmiş Bağıl Kuvvet Formülü — kg/kg)",
            "claim_type": "Matematiksel Formül / Normalizasyon",
            "claim_text": "Vücut kütlesi artışından bağımsız nöromüsküler kuvvet gelişimini incelemek amacıyla Bağıl Kavrama Kuvveti = [(Sağ + Sol Kavrama) / 2] / Vücut Kütlesi (kg/kg) olarak hesaplanmıştır.",
            "source_ref": "Jaric, S. (2002). Muscle strength testing: Use of normalisation for body size. Sports Medicine, 32(10), 615–631.",
            "pdf_prefix": "Atif_14B_",
            "page_and_section": "Sayfa 7 (Tablo II ve Bölüm 2.2) ve Sayfa 14 (Bölüm 5. Conclusion)",
            "decision_label": "Birincil değil",
            "confidence": "Yüksek",
            "primary_status": "Birincil değil",
            "secondary_attribution_proof": "More often researchers applied so-called ratio standards by normalising recorded strength per kilogram of body mass, or fat-free mass (see data in the fourth column of table II).[11,27,30,72]",
            "final_status": "Uygun değil",
            "rationale": "Makale anlatısal derlemedir; kuvvetin doğrudan vücut kütlesine (m^1) bölünmesini önceki çalışmalara ([11, 27, 30, 72]) atfetmekte ve kendi sonuç bölümünde (Sayfa 14) doğrusal kg/kg oranı yerine allometrik m^0.67 ölçeklemesini önermektedir (yani kg/kg oranının farklı vücut boyutlarında yanıltıcı olabileceğini savunmaktadır).",
            "recommended_action": "Makale taslağında 'BİRİNCİL KAYNAK DEĞİL (ve Allometrik Ölçekleme Savunusu)' olarak işaretlendi; doğrusal kg/kg bağıl kavrama kuvveti indeksini pediatrik örneklemde birincil olarak tanımlayan kaynak seçilmelidir.",
            "hl_targets": [
                {"page": 7, "blocks": [0, 1]},
                {"page": 14, "blocks": [6, 7]},
            ],
        },
        {
            "citation_id": "A024",
            "citation_marker": "[Atıf Gerekli - 14] -> Wind ve ark. (2010)",
            "manuscript_location": "Bölüm 2.4.2 (Vücut Kütlesine Göre Normalize Edilmiş Bağıl Kavrama Kuvveti — kg/kg)",
            "claim_type": "Matematiksel Formül / Normalizasyon",
            "claim_text": "Bağıl Kavrama Kuvveti = [(Sağ + Sol Kavrama) / 2] / Vücut Kütlesi (kg/kg) olarak hesaplanmıştır.",
            "source_ref": "Wind, A. E., Takken, T., Helders, P. J. M., & Engelbert, R. H. H. (2010). Is grip strength a predictor for total muscle strength in healthy children, adolescents, and young adults? European Journal of Pediatrics, 169(3), 281–287.",
            "pdf_prefix": "Atif_14A_",
            "page_and_section": "Sayfa 2 (Muscle strength & Statistical analysis) ve Sayfa 3 (Results)",
            "decision_label": "Kısmi destek",
            "confidence": "Yüksek",
            "primary_status": "Birincil",
            "secondary_attribution_proof": "",
            "final_status": "Düzeltme gerekli",
            "rationale": "8–20 yaş arası 384 çocuk ve gençte kavrama kuvveti ile toplam kas kuvveti ilişkisini vücut ağırlığına göre kontrol ederek inceleyen birincil pediatrik araştırmadır; ancak ölçümü Newton (Citec el dinamometresi, break-test) cinsinden yapmış ve ağırlık etkisini [(Sağ+Sol)/2]/VK (kg/kg) formülü yerine kısmi korelasyon ve LMS persentil eğrileriyle modellemiştir.",
            "recommended_action": "Makale taslağında 'KISMİ DESTEK (Formül Farkı: LMS/Kısmi Korelasyon vs. kg/kg Oranı)' olarak işaretlendi.",
            "hl_targets": [
                {"page": 2, "blocks": [13, 15]},
                {"page": 3, "blocks": [6, 7]},
            ],
        },
        {
            "citation_id": "A025",
            "citation_marker": "[Atıf Gerekli - 15] -> Impellizzeri ve ark. (2007)",
            "manuscript_location": "Bölüm 2.4.2 (Bilateral Fonksiyonel Asimetri İndeksi Formülü)",
            "claim_type": "Matematiksel Formül",
            "claim_text": "Ekstremiteler arası fonksiyonel kavrama asimetrisi Asimetri (%) = [|Sağ - Sol| / maks(Sağ, Sol)] × 100 formülü ile belirlenmiştir.",
            "source_ref": "Impellizzeri, F. M., Rampinini, E., Maffiuletti, N., & Marcora, S. M. (2007). A vertical jump force test for assessing bilateral strength asymmetry in athletes. Medicine & Science in Sports & Exercise, 39(11), 2044–2050.",
            "pdf_prefix": "Atif_15B_",
            "page_and_section": "Sayfa 1 (Özet) ve Sayfa 4 (Bilateral Strength Asymmetry Calculation)",
            "decision_label": "Birincil değil",
            "confidence": "Yüksek",
            "primary_status": "Birincil değil",
            "secondary_attribution_proof": "In the literature (19,20,30), relative lower-limb strength asymmetry is calculated in different ways: 1) (injured - noninjured)/injured × 100; 2) (right - left)/left × 100; 3) (stronger - weaker)/stronger × 100.",
            "final_status": "Uygun değil",
            "rationale": "Impellizzeri ve ark. (2007) çift kuvvet platformlu dikey sıçrama testi için özgün çalışma olsa da, (güçlü - zayıf)/güçlü × 100 asimetri formülünü kendi türetmemiş; açıkça önceki literatüre (19, 20, 30; örn. Newton ve ark., 2006; Sapega, 1990) atfetmiştir.",
            "recommended_action": "Makale taslağında 'BİRİNCİL KAYNAK DEĞİL' olarak işaretlendi; formülün orijinal kaynağı (örn. Sapega, 1990 / Nunn & Mayhew, 1988) bulunmalı veya formül standart matematiksel tanım olarak açıkça verilmelidir.",
            "hl_targets": [
                {"page": 4, "blocks": [3]},
            ],
        },
        {
            "citation_id": "A026",
            "citation_marker": "[Atıf Gerekli - 15] -> Bishop ve ark. (2018)",
            "manuscript_location": "Bölüm 2.4.2 (Bilateral Fonksiyonel Asimetri İndeksi Formülü)",
            "claim_type": "Matematiksel Formül",
            "claim_text": "Ekstremiteler arası fonksiyonel kavrama asimetrisi Asimetri (%) = [|Sağ - Sol| / maks(Sağ, Sol)] × 100 formülü ile belirlenmiştir.",
            "source_ref": "Bishop, C., Read, P., Lake, J., Chavda, S., & Turner, A. (2018). Interlimb asymmetries: Understanding how to calculate differences from bilateral and unilateral tests. Strength and Conditioning Journal, 40(4), 1–6.",
            "pdf_prefix": "Atif_15A_",
            "page_and_section": "Sayfa 2 (Tablo 1 — BSA satırı) ve Sayfa 5 (Summary)",
            "decision_label": "Birincil değil",
            "confidence": "Yüksek",
            "primary_status": "Birincil değil",
            "secondary_attribution_proof": "Table 1: BSA | (Stronger limb - Weaker limb)/Stronger limb × 100 | Reference: Nunn et al. (21), Impellizzeri et al. (14).",
            "final_status": "Uygun değil",
            "rationale": "Makale metodolojik derleme makalesidir; Tablo 1'de (Stronger - Weaker)/Stronger × 100 (BSA) denklemini Nunn ve ark. (1988) [21] ile Impellizzeri ve ark. (2007) [14] kaynaklarına atfetmektedir (yalnızca unilateral testlerde BSA denkleminin seçilmesi önerisi bakımından yazarların kendi görüşünü sunar).",
            "recommended_action": "Makale taslağında 'BİRİNCİL KAYNAK DEĞİL' olarak işaretlendi.",
            "hl_targets": [
                {"page": 2, "blocks": [6, 7, 8, 12]},
                {"page": 5, "blocks": [1]},
            ],
        },
        {
            "citation_id": "A027",
            "citation_marker": "[Atıf Gerekli - 16] -> Markovic ve ark. (2004)",
            "manuscript_location": "Bölüm 2.4.2 (Eller Belde Aktif Dikey Sıçrama / CMJ Sıçrama Matı Protokolü — 3 Ölçüm)",
            "claim_type": "Ölçüm Protokolü ve Güvenirlik",
            "claim_text": "Alt ekstremite patlayıcı kuvveti, kol salınımı kısıtlanarak (eller belde sabitlenmiş pozisyonda) elektronik sıçrama matı üzerinde gerçekleştirilen aktif dikey sıçrama (CMJ) testi ile değerlendirilmiştir. Katılımcılar 3 geçerli sıçrama gerçekleştirmiştir.",
            "source_ref": "Markovic, G., Dizdar, D., Jukic, I., & Cardinale, M. (2004). Reliability and factorial validity of squat and countermovement jump tests. Journal of Strength and Conditioning Research, 18(3), 551–555.",
            "pdf_prefix": "Atif_16A_",
            "page_and_section": "Sayfa 2 (Testing Procedure — SJ and CMJ, 3 deneme, 1 dk dinlenme) ve Sayfa 4 (Sonuç — α = 0.98)",
            "decision_label": "Tam destek",
            "confidence": "Yüksek",
            "primary_status": "Birincil",
            "secondary_attribution_proof": "",
            "final_status": "Uygun",
            "rationale": "Dijital zamanlayıcılı temas/sıçrama matı üzerinde eller belde uygulanan 3 denemeli CMJ protokolünün güvenirlik (Cronbach α = 0.98, ICC = 0.98) ve faktöriyel geçerliğini doğrudan kanıtlayan birincil deneysel araştırmadır.",
            "recommended_action": "CMJ temas matı test protokolü cümlesine birincil kaynak olarak entegre edildi (Markovic ve ark., 2004).",
            "hl_targets": [
                {"page": 2, "blocks": [12, 13, 14, 15, 18]},
                {"page": 4, "blocks": [9]},
            ],
        },
        {
            "citation_id": "A028",
            "citation_marker": "[Atıf Gerekli - 16] -> Duncan ve ark. (2008)",
            "manuscript_location": "Bölüm 2.4.2 (Sayers Zirve Anaerobik Güç Denklemi: P_zirve = 60.7 × CMJ + 45.3 × VK - 2055)",
            "claim_type": "Matematiksel Regresyon Denklemi",
            "claim_text": "Zirve anaerobik güç çıktısı (W), Sayers dikey sıçrama güç denklemi kullanılarak P_zirve (W) = (60.7 × Sıçrama Yüksekliği [cm]) + (45.3 × Vücut Kütlesi [kg]) - 2055 formülüyle hesaplanmıştır.",
            "source_ref": "Duncan, M. J., Lyons, M., & Nevill, A. M. (2008). Evaluation of peak power prediction equations in male basketball players. Journal of Strength and Conditioning Research, 22(4), 1379–1381.",
            "pdf_prefix": "Atif_16B_",
            "page_and_section": "Sayfa 1 (Özet & Giriş) ve Sayfa 3 (Tartışma & Kaynak 12: Sayers ve ark., 1999)",
            "decision_label": "Birincil değil",
            "confidence": "Yüksek",
            "primary_status": "Birincil değil",
            "secondary_attribution_proof": "This study compared peak power estimated using 4 commonly used regression equations (Harman et al., Sayers squat jump [SJ], Sayers countermovement jump [CMJ], and Canavan and Vescovi)... 12. Sayers, SP, Harackiewicz, DV, Harman, EA, Frykman, PN, and Rosenstein, MT. Cross-validation of three jump power equations. Med Sci Sports Exerc 31: 572–577, 1999.",
            "final_status": "Uygun değil",
            "rationale": "Duncan ve ark. (2008) Sayers CMJ denklemini türeten birincil kaynak değildir; denklemi 12 numaralı kaynağa (Sayers ve ark., 1999) atfetmekte, denklem katsayılarını (60.7, 45.3, -2055) metin içinde açık olarak basmamakta ve ergen basketbolcularda denklemin gerçek gücü düşük tahmin ettiğini raporlamaktadır.",
            "recommended_action": "Makale taslağında Sayers formülü cümlesinde 'BİRİNCİL KAYNAK DEĞİL (Sayers ve ark., 1999 Birincil Kaynağı Gerekli)' olarak işaretlendi.",
            "hl_targets": [
                {"page": 1, "blocks": [4, 9]},
                {"page": 3, "blocks": [0, 2, 6]},
            ],
        },
        {
            "citation_id": "A029",
            "citation_marker": "[Atıf Gerekli - 17] -> Harman ve Garhammer (2008)",
            "manuscript_location": "Bölüm 2.4.2 (Pro-Agility 5-10-5 Çeviklik Testi — 9.14 m Parkur ve Fotosel Protokolü)",
            "claim_type": "Ölçüm Protokolü",
            "claim_text": "Yön değiştirme hızı ve çeviklik performansı, toplam 9.14 m (10 yarda) genişliğindeki standart Pro-Agility (5-10-5) parkurunda çift ışınlı elektronik fotosel kapıları ile ölçülmüştür. 2 geçerli denemeden en hızlı süre kaydedilmiştir.",
            "source_ref": "Harman, E., & Garhammer, J. (2008). Administration, scoring, and interpretation of selected tests. In T. R. Baechle & R. W. Earle (Eds.), Essentials of strength training and conditioning (3rd ed., pp. 249–292). Human Kinetics.",
            "pdf_prefix": None,
            "page_and_section": "PDF mevcut değil (Kitap bölümü)",
            "decision_label": "Doğrulanamadı",
            "confidence": "Yüksek",
            "primary_status": "Doğrulanamadı (PDF eksik)",
            "secondary_attribution_proof": "",
            "final_status": "Doğrulanamadı",
            "rationale": "Klasörde PDF dosyası bulunmamaktadır; ayrıca ders/el kitabı bölümü niteliğindedir.",
            "recommended_action": "Makale taslağında 'DOĞRULANAMADI / KİTAP BÖLÜMÜ (PDF Eksik)' olarak işaretlendi.",
            "hl_targets": [],
        },
        {
            "citation_id": "A030",
            "citation_marker": "[Atıf Gerekli - 17] -> Stewart ve ark. (2014)",
            "manuscript_location": "Bölüm 2.4.2 (Pro-Agility 5-10-5 Çeviklik Testi — Fotosel Güvenirliği ve 2 Deneme)",
            "claim_type": "Ölçüm Protokolü ve Güvenirlik",
            "claim_text": "Yön değiştirme hızı ve çeviklik performansı, toplam 9.14 m (10 yarda) genişliğindeki standart Pro-Agility (5-10-5) parkurunda çift ışınlı elektronik fotosel kapıları ile ölçülmüştür. En az 3 dakika tam dinlenme aralığıyla yapılan 2 geçerli denemeden en hızlı süre kaydedilmiştir.",
            "source_ref": "Stewart, P. F., Turner, A. N., & Miller, S. C. (2014). Reliability, factorial validity, and interrelationships of five commonly used change of direction speed tests. Scandinavian Journal of Medicine & Science in Sports, 24(3), 500–506.",
            "pdf_prefix": "Atif_17_",
            "page_and_section": "Sayfa 2 (Procedures & Pro-Agility test) ve Sayfa 3 (Results — ICC = 0.90)",
            "decision_label": "Kısmi destek",
            "confidence": "Yüksek",
            "primary_status": "Birincil",
            "secondary_attribution_proof": "The Pro-Agility test was set up and administered using the protocol outlined by Harman et al. (2000; Fig. 1c).",
            "final_status": "Düzeltme gerekli",
            "rationale": "Elektronik fotosel kapılarıyla ölçülen Pro-Agility (5-10-5) testinin yüksek güvenirliğini (ICC = 0.90) ve faktöriyel geçerliğini kanıtlayan birincil araştırmadır; ancak parkur kurulumunu Harman ve ark. (2000) kaynağına atfetmekte (4.55 m + 9.10 m + 4.55 m yazımıyla) ve çift ışınlı 2 deneme yerine tek ışınlı Brower fotosellerle 3 deneme uygulamaktadır.",
            "recommended_action": "Makale taslağında 'KISMİ DESTEK (Fotosel Güvenirliği İçin Birincil; Parkur Kurulumu İkincil ve Deneme Sayısı Farklı [3 vs. 2])' olarak işaretlendi.",
            "hl_targets": [
                {"page": 2, "blocks": [3, 6, 12]},
                {"page": 3, "blocks": [3, 4]},
            ],
        },
        {
            "citation_id": "A031",
            "citation_marker": "[Atıf Gerekli - 18] -> Wells ve Dillon (1952)",
            "manuscript_location": "Bölüm 2.4.2 (Otur-Eriş Esneklik Testi — ±0.5 cm Sehpa, 3 Denemenin En İyisi)",
            "claim_type": "Ölçüm Protokolü",
            "claim_text": "Hamstring ve lomber ekstansör esnekliği standart otur-eriş sehpası (±0.5 cm hassasiyet) kullanılarak ölçülmüştür. En az 2 sn korunan 3 deneme arasından ulaşılan en yüksek erişim mesafesi (cm) analizlere dahil edilmiştir.",
            "source_ref": "Wells, K. F., & Dillon, E. K. (1952). The sit and reach—A test of back and leg flexibility. Research Quarterly, 23(1), 115–118.",
            "pdf_prefix": None,
            "page_and_section": "PDF mevcut değil",
            "decision_label": "Doğrulanamadı",
            "confidence": "Yüksek",
            "primary_status": "Doğrulanamadı (PDF eksik)",
            "secondary_attribution_proof": "",
            "final_status": "Doğrulanamadı",
            "rationale": "Otur-eriş (Sit-and-Reach) testini geliştiren orijinal birincil makaledir; ancak PDF klasöründe yer almadığından metin üzerinden doğrulanamamıştır.",
            "recommended_action": "Makale taslağında 'DOĞRULANAMADI (Birincil Kaynak PDF Eksik: Wells ve Dillon, 1952)' olarak işaretlendi.",
            "hl_targets": [],
        },
        {
            "citation_id": "A032",
            "citation_marker": "[Atıf Gerekli - 18] -> Castro-Piñero ve ark. (2009)",
            "manuscript_location": "Bölüm 2.4.2 (Otur-Eriş Esneklik Testi — Hamstring/Lomber Esneklik ve 3 Denemenin En İyisi)",
            "claim_type": "Ölçüm Protokolü ve Pediatrik Geçerlik",
            "claim_text": "Hamstring ve lomber ekstansör esnekliği standart otur-eriş sehpası (±0.5 cm hassasiyet) kullanılarak ölçülmüştür. 3 deneme arasından ulaşılan en yüksek erişim mesafesi (cm) analizlere dahil edilmiştir.",
            "source_ref": "Castro-Piñero, J., Chillón, P., Ortega, F. B., Montesinos, J. L., Sjöström, M., & Ruiz, J. R. (2009). Criterion-related validity of sit-and-reach and modified sit-and-reach test for estimating hamstring flexibility in children and adolescents aged 6–17 years. International Journal of Sports Medicine, 30(9), 658–662.",
            "pdf_prefix": "Atif_18_",
            "page_and_section": "Sayfa 1 (Giriş — [40] atfı), Sayfa 2 (Sit-and-reach test — 2 deneme ortalaması) ve Sayfa 4 (Sonuç)",
            "decision_label": "Kısmi destek",
            "confidence": "Yüksek",
            "primary_status": "Birincil",
            "secondary_attribution_proof": "The classical sit-and-reach test (SRT) originally designed by Wells and Dillon [40] is frequently used to estimate hamstring flexibility.",
            "final_status": "Düzeltme gerekli",
            "rationale": "6–17 yaş çocuk ve ergenlerde otur-eriş testini inceleyen birincil araştırmadır; ancak testin tasarımını Wells ve Dillon [40] kaynağına atfetmekte, 3 denemenin en iyisi yerine '2 ölçümün ortalamasını' almakta ve sonuç bölümünde otur-eriş testinin çocuk/ergenlerde hamstring esnekliğini tahmin etmede zayıf ölçüt geçerliğine (weak validity) sahip olduğunu bildirmektedir.",
            "recommended_action": "Makale taslağında 'KISMİ DESTEK (Test Kökeni İkincil [Wells & Dillon, 1952], Deneme Sayısı Farklı [2 Ortalama vs. 3 Maksimum] ve Zayıf Geçerlik Sonucu)' olarak işaretlendi.",
            "hl_targets": [
                {"page": 1, "blocks": [8]},
                {"page": 2, "blocks": [5]},
                {"page": 4, "blocks": [3]},
            ],
        },
        {
            "citation_id": "A033",
            "citation_marker": "[Atıf Gerekli - 19] -> Shapiro ve Wilk (1965)",
            "manuscript_location": "Bölüm 2.5 (İstatistiksel Analiz — Shapiro-Wilk Normallik Testi)",
            "claim_type": "İstatistiksel Test",
            "claim_text": "Sürekli değişkenlerin ve ön-son test fark (Δ) puanlarının dağılım normalliği Shapiro-Wilk testi ve Q-Q grafikleri ile incelenmiştir.",
            "source_ref": "Shapiro, S. S., & Wilk, M. B. (1965). An analysis of variance test for normality (complete samples). Biometrika, 52(3–4), 591–611.",
            "pdf_prefix": None,
            "page_and_section": "PDF mevcut değil",
            "decision_label": "Doğrulanamadı",
            "confidence": "Yüksek",
            "primary_status": "Doğrulanamadı (PDF eksik)",
            "secondary_attribution_proof": "",
            "final_status": "Doğrulanamadı",
            "rationale": "Shapiro-Wilk W normallik testini geliştiren orijinal birincil makaledir; ancak PDF klasöründe yer almadığından doğrulanamamıştır.",
            "recommended_action": "Makale taslağında 'DOĞRULANAMADI (Birincil Kaynak PDF Eksik: Shapiro ve Wilk, 1965)' olarak işaretlendi.",
            "hl_targets": [],
        },
        {
            "citation_id": "A034",
            "citation_marker": "[Atıf Gerekli - 19] -> Razali ve Wah (2011)",
            "manuscript_location": "Bölüm 2.5 (İstatistiksel Analiz — Shapiro-Wilk Normallik Testi ve Q-Q Grafikleri)",
            "claim_type": "İstatistiksel Test / Monte Carlo Güç Karşılaştırması",
            "claim_text": "Sürekli değişkenlerin ve ön-son test fark (Δ) puanlarının dağılım normalliği Shapiro-Wilk testi ve Q-Q grafikleri ile incelenmiştir.",
            "source_ref": "Razali, N. M., & Wah, Y. B. (2011). Power comparisons of Shapiro-Wilk, Kolmogorov-Smirnov, Lilliefors and Anderson-Darling tests. Journal of Statistical Modeling and Analytics, 2(1), 21–33.",
            "pdf_prefix": "Atif_19_",
            "page_and_section": "Sayfa 1 (Özet), Sayfa 5 (Shapiro-Wilk Test — Shapiro and Wilk, 1965 atfı) ve Sayfa 12 (Sonuç)",
            "decision_label": "Birincil değil",
            "confidence": "Yüksek",
            "primary_status": "Birincil değil",
            "secondary_attribution_proof": "Shapiro and Wilk (1965) test was originally restricted for sample size of less than 50. This test was the first test that was able to detect departures from normality due to either skewness or kurtosis, or both (Althouse et al., 1998).",
            "final_status": "Uygun değil",
            "rationale": "Shapiro-Wilk testinin kendisi için birincil kaynak değildir; W istatistiğini doğrudan Shapiro ve Wilk (1965) ile Royston (1982) kaynaklarına atfetmektedir (yalnızca küçük örneklemlerde Shapiro-Wilk'in en yüksek istatistiksel güce sahip olduğunu gösteren Monte Carlo simülasyonu ve Q-Q grafikleriyle birlikte kullanım önerisi bakımından birincildir).",
            "recommended_action": "Makale taslağında 'BİRİNCİL KAYNAK DEĞİL (Testin Kökeni Shapiro ve Wilk, 1965; Razali ve Wah, 2011 Yalnızca Küçük Örneklem Güç Karşılaştırması İçin Birincildir)' olarak işaretlendi.",
            "hl_targets": [
                {"page": 1, "blocks": [9]},
                {"page": 5, "blocks": [2]},
                {"page": 12, "blocks": [2]},
            ],
        },
        {
            "citation_id": "A035",
            "citation_marker": "[Atıf Gerekli - 20] -> Cohen (1988)",
            "manuscript_location": "Bölüm 2.5 (Etki Büyüklüğü Yorumlama Eşikleri: d_av = 0.20/0.50/0.80 ve ηp² = 0.01/0.06/0.14)",
            "claim_type": "İstatistiksel Eşik / Sınıflama",
            "claim_text": "Etki büyüklükleri d_av için 0.20 (küçük), 0.50 (orta), ≥0.80 (büyük); ηp² için 0.01 (küçük), 0.06 (orta), ≥0.14 (büyük) olarak yorumlanmıştır.",
            "source_ref": "Cohen, J. (1988). Statistical power analysis for the behavioral sciences (2nd ed.). Lawrence Erlbaum Associates.",
            "pdf_prefix": None,
            "page_and_section": "PDF mevcut değil",
            "decision_label": "Doğrulanamadı",
            "confidence": "Yüksek",
            "primary_status": "Doğrulanamadı (PDF eksik)",
            "secondary_attribution_proof": "",
            "final_status": "Doğrulanamadı",
            "rationale": "Standart etki büyüklüğü eşiklerinin (0.20/0.50/0.80 ve 0.01/0.06/0.14) orijinal birincil kaynağı Cohen (1988) monografisidir; ancak klasörde PDF dosyası bulunmamaktadır.",
            "recommended_action": "Makale taslağında 'DOĞRULANAMADI (Birincil Kaynak PDF Eksik: Cohen, 1988)' olarak işaretlendi.",
            "hl_targets": [],
        },
        {
            "citation_id": "A036",
            "citation_marker": "[Atıf Gerekli - 20] -> Lakens (2013)",
            "manuscript_location": "Bölüm 2.5 (Etki Büyüklüğü Hesaplamaları ve Eşikleri — Cohen's d_av ve Kısmi Eta-Kare ηp²)",
            "claim_type": "İstatistiksel Formül ve Yorumlama Eşikleri",
            "claim_text": "Standartlaştırılmış etki büyüklüğü olarak grup içi ortalama standart sapmaya dayalı Cohen's d_av ve ANOVA modellerinde kısmi eta-kare (ηp²) hesaplanmıştır. Etki büyüklükleri d_av için 0.20 (küçük), 0.50 (orta), ≥0.80 (büyük); ηp² için 0.01 (küçük), 0.06 (orta), ≥0.14 (büyük) olarak yorumlanmıştır.",
            "source_ref": "Lakens, D. (2013). Calculating and reporting effect sizes to facilitate cumulative science: A practical primer for t-tests and ANOVAs. Frontiers in Psychology, 4, 863.",
            "pdf_prefix": "Atif_20_",
            "page_and_section": "Sayfa 3 (Interpreting Cohen's d), Sayfa 5 (Eq. 10 — Cohen's d_av), Sayfa 6 (Eq. 13 — ηp²) ve Sayfa 7 (Cohen 1988 η² eşikleri)",
            "decision_label": "Birincil değil",
            "confidence": "Yüksek",
            "primary_status": "Birincil değil",
            "secondary_attribution_proof": "Cumming (2012) refers to this approach as Cohen's d_av (p. 5)... A commonly used interpretation is to refer to effect sizes as small (d = 0.2), medium (d = 0.5), and large (d = 0.8) based on benchmarks suggested by Cohen (1988) (p. 3)... Cohen (1988) has provided benchmarks to define small (η² = 0.01), medium (η² = 0.06), and large (η² = 0.14) effects (p. 7).",
            "final_status": "Uygun değil",
            "rationale": "Lakens (2013) uygulamalı metodolojik rehberdir ('A practical primer'); hem d_av adlandırmasını Cumming (2012)'ye hem de d (0.20/0.50/0.80) ve η² (0.01/0.06/0.14) yorumlama eşiklerini açıkça Cohen (1988) kaynağına atfetmektedir (üstelik Sayfa 7'de η² eşiklerinin tekrarlı ölçümlü ηp² için doğrudan kullanılmaması gerektiğini belirtmektedir).",
            "recommended_action": "Makale taslağında 'BİRİNCİL KAYNAK DEĞİL (Eşikler Cohen, 1988'e; d_av Adlandırması Cumming, 2012'ye Atfediliyor)' olarak işaretlendi.",
            "hl_targets": [
                {"page": 3, "blocks": [32]},
                {"page": 5, "blocks": [1, 2, 3, 4]},
                {"page": 7, "blocks": [4, 5]},
            ],
        },
        {
            "citation_id": "A037",
            "citation_marker": "[Atıf Gerekli - 21] -> Hopkins (2004)",
            "manuscript_location": "Bölüm 2.5 (En Küçük Anlamlı Değişim — SWC = 0.20 × Ön-Test SS)",
            "claim_type": "Matematiksel Formül / Bireysel Yanıt Sınıflaması",
            "claim_text": "Bireysel düzeyde klinik/pratik anlamlı yanıt oranlarını belirlemek üzere her parametre için En Küçük Anlamlı Değişim (Smallest Worthwhile Change; SWC = 0.20 × Ön-Test SS) eşiği hesaplanmıştır.",
            "source_ref": "Hopkins, W. G. (2004). How to interpret changes in an athletic performance test. Sportscience, 8, 1–7.",
            "pdf_prefix": "Atif_21A_",
            "page_and_section": "Sayfa 1 (Özet — one-fifth of the between-athlete SD) ve Sayfa 3 (What's a Worthwhile Enhancement — 0.20 × between-subject SD)",
            "decision_label": "Tam destek",
            "confidence": "Yüksek",
            "primary_status": "Birincil",
            "secondary_attribution_proof": "",
            "final_status": "Uygun",
            "rationale": "Spor bilimleri performans testlerinde doğrudan müsabaka derecesi olmadığında En Küçük Anlamlı Değişim (SWC) eşiğinin varsayılan olarak bireyler arası standart sapmanın beşte biri (0.20 × SD) alınması kuralını literatüre kazandıran orijinal metodolojik makaledir.",
            "recommended_action": "SWC = 0.20 × Ön-Test SS eşiği için birincil kaynak olarak makale taslağına entegre edildi (Hopkins, 2004).",
            "hl_targets": [
                {"page": 1, "blocks": [3]},
                {"page": 3, "blocks": [5, 6]},
            ],
        },
        {
            "citation_id": "A038",
            "citation_marker": "[Atıf Gerekli - 21] -> Swinton ve ark. (2018)",
            "manuscript_location": "Bölüm 2.5 (En Küçük Anlamlı Değişim — SWC = 0.20 × Ön-Test SS ve Yanıt Veren Sınıflaması)",
            "claim_type": "Matematiksel Formül / Bireysel Yanıt Çerçevesi",
            "claim_text": "En Küçük Anlamlı Değişim (SWC = 0.20 × Ön-Test SS) eşiği hesaplanmış ve bu eşiği gelişim yönünde aşan katılımcılar 'yanıt veren (responder)' olarak sınıflandırılmıştır.",
            "source_ref": "Swinton, P. A., Hemingway, B. S., Saunders, B., Gualano, B., & Dolan, E. (2018). A statistical framework to interpret individual response to intervention: Paving the way for personalized nutrition and exercise prescription. Frontiers in Nutrition, 5, 41.",
            "pdf_prefix": "Atif_21B_",
            "page_and_section": "Sayfa 8 (Threshold-based classification — kaynak [27] atfı)",
            "decision_label": "Birincil değil",
            "confidence": "Yüksek",
            "primary_status": "Birincil değil",
            "secondary_attribution_proof": "A standardized approach relies on calculation of a fraction (typically 0.2) of the baseline between-individual SD (27)... [Ref 27: Hopkins WG. How to interpret changes in an athletic performance test. Sportscience (2004) 8:1–7.]",
            "final_status": "Uygun değil",
            "rationale": "Makale istatistiksel çerçeve derlemesidir; 0.2 × bazal bireyler arası SS (SWC) formülünü açıkça [27] numaralı kaynağa (Hopkins, 2004) atfetmektedir.",
            "recommended_action": "Makale taslağında 'BİRİNCİL KAYNAK DEĞİL' olarak işaretlendi; SWC formülü için yalnızca birincil kaynak olan Hopkins (2004) entegre edilmiştir.",
            "hl_targets": [
                {"page": 8, "blocks": [28]},
            ],
        },
        {
            "citation_id": "A039",
            "citation_marker": "[Atıf Gerekli - 22] -> Aiken ve West (1991)",
            "manuscript_location": "Bölüm 2.5 (OLS Çoklu Doğrusal Regresyon ve Pearson Korelasyon Analizi)",
            "claim_type": "İstatistiksel Yöntem",
            "claim_text": "Kronolojik yaş ve başlangıç performansının net gelişim (Δ) üzerindeki bağımsız yordayıcı rollerini belirlemek amacıyla Sıradan En Küçük Kareler (OLS) çoklu doğrusal regresyon modelleri (Δ = β0 + β1(Yaş) + β2(Ön-Test)) kurulmuş; antropometrik boyutlar ile kazanımlar arasındaki ilişkiler Pearson korelasyon katsayısı (r) ile değerlendirilmiştir.",
            "source_ref": "Aiken, L. S., & West, S. G. (1991). Multiple regression: Testing and interpreting interactions. Sage Publications.",
            "pdf_prefix": None,
            "page_and_section": "PDF mevcut değil",
            "decision_label": "Doğrulanamadı",
            "confidence": "Yüksek",
            "primary_status": "Doğrulanamadı (PDF eksik)",
            "secondary_attribution_proof": "",
            "final_status": "Doğrulanamadı",
            "rationale": "Klasörde PDF dosyası bulunmamaktadır (ayrıca kurulan regresyon modelinde çarpımsal etkileşim terimi yer almamaktadır).",
            "recommended_action": "Makale taslağında 'DOĞRULANAMADI (PDF Eksik: Aiken ve West, 1991)' olarak işaretlendi; standart OLS regresyon ve Pearson korelasyon analizleri istatistikte genel geçer temel yöntemler olduğundan atıfsız da bırakılabilir.",
            "hl_targets": [],
        },
        {
            "citation_id": "A040",
            "citation_marker": "[Atıf Gerekli - 22] -> Tabachnick ve Fidell (2019)",
            "manuscript_location": "Bölüm 2.5 (OLS Çoklu Doğrusal Regresyon ve Pearson Korelasyon Analizi)",
            "claim_type": "İstatistiksel Yöntem",
            "claim_text": "OLS çoklu doğrusal regresyon modelleri kurulmuş ve Pearson korelasyon katsayısı (r) ile değerlendirilmiştir.",
            "source_ref": "Tabachnick, B. G., & Fidell, L. S. (2019). Using multivariate statistics (7th ed.). Pearson.",
            "pdf_prefix": "Atif_22_",
            "page_and_section": "Sayfa 1–14 (Yalnızca Kapak, İçindekiler ve Ek İstatistik Tabloları)",
            "decision_label": "Bulunamadı",
            "confidence": "Yüksek",
            "primary_status": "Birincil değil",
            "secondary_attribution_proof": "",
            "final_status": "Uygun değil",
            "rationale": "İki yönlü uyumsuzluk tespit edilmiştir: (1) Kaynak genel istatistik ders kitabıdır (primary_source_criteria.md Madde 3.2 uyarınca birincil kaynak değildir); (2) Klasördeki 14 sayfalık PDF dosyası yalnızca ön kapak, içindekiler (TOC) ve ek kritik değer tablolarını içermekte olup çoklu regresyon (Bölüm 5) veya korelasyon metinlerini barındırmamaktadır. Core Principle 2 ve 8 gereği rastgele içindekiler başlığı boyanmamıştır.",
            "recommended_action": "Makale taslağında 'BİRİNCİL KAYNAK DEĞİL / PASAJ BULUNAMADI (Ders Kitabı Ön Sayfa Özeti)' olarak işaretlendi.",
            "hl_targets": [],
        },
        {
            "citation_id": "A041",
            "citation_marker": "[Atıf Gerekli - 23] -> Luna ve ark. (2004)",
            "manuscript_location": "Bölüm 3.1 (Bulgular — Yaş Grupları Başlangıç Karşılaştırması: Ergenlikte Hızlı Reaksiyon Süresi)",
            "claim_type": "Ampirik Karşılaştırma / Gelişimsel Bulgular",
            "claim_text": "Ergen grubun reaksiyon süresi çocuk gruba kıyasla anlamlı düzeyde daha hızlı bulunmuştur (p = 0.015). Buna karşın vücut kütlesine göre normalize edilmiş bağıl kavrama kuvveti (p = 0.308) ve Pro-Agility çeviklik süresi (p = 0.175) yaş grupları arasında anlamlı farklılık göstermemiştir.",
            "source_ref": "Luna, B., Garver, K. E., Urban, T. A., Lazar, N. A., & Sweeney, J. A. (2004). Maturation of cognitive processes from late childhood to adulthood. Child Development, 75(5), 1357–1372.",
            "pdf_prefix": "Atif_23B_",
            "page_and_section": "Sayfa 1 (Özet), Sayfa 9 (Results — VGS Latency) ve Sayfa 11 (Development of processing speed)",
            "decision_label": "Kısmi destek",
            "confidence": "Yüksek",
            "primary_status": "Birincil",
            "secondary_attribution_proof": "",
            "final_status": "Düzeltme gerekli",
            "rationale": "8–30 yaş arası sağlıklı bireylerde görsel uyarana tepki hızının (işlemleme hızı) çocukluktan ~15 yaşa kadar anlamlı biçimde hızlandığını gösteren birincil araştırmadır; ancak (1) OSB örneklemini, bağıl kavrama kuvvetini ve çevikliği içermemekte, (2) Q1 makale yazım standardı gereği Bulgular (Results) bölümünde dış literatür atfı yapılmayıp bu karşılaştırmanın Tartışma (Discussion) bölümüne taşınması gerekmektedir.",
            "recommended_action": "Makale taslağında 'KISMİ DESTEK (Yalnızca Sağlıklı Ergenlerde İşlemleme Hızı Olgunlaşmasını Destekler; Tartışma Bölümüne Taşınmalı)' olarak işaretlendi.",
            "hl_targets": [
                {"page": 1, "blocks": [2]},
                {"page": 11, "blocks": [4, 5]},
            ],
        },
        {
            "citation_id": "A042",
            "citation_marker": "[Atıf Gerekli - 23] -> Lloyd ve Oliver (2012)",
            "manuscript_location": "Bölüm 3.1 (Bulgular — Yaş Grupları Başlangıç Karşılaştırması: Kuvvet ve Çeviklik Gelişimi)",
            "claim_type": "Kuramsal Gelişim Modeli / Anlatısal Derleme",
            "claim_text": "Ergen grubun reaksiyon süresi çocuk gruba kıyasla daha hızlı bulunmuş; bağıl kavrama kuvveti ve Pro-Agility çeviklik süresi yaş grupları arasında anlamlı farklılık göstermemiştir.",
            "source_ref": "Lloyd, R. S., & Oliver, J. L. (2012). The youth physical development model: A new approach to long-term athletic development. Strength and Conditioning Journal, 34(3), 61–72.",
            "pdf_prefix": "Atif_23A_",
            "page_and_section": "Sayfa 2 (Windows of Opportunity), Sayfa 4 (Muscular Strength) ve Sayfa 5 (Agility)",
            "decision_label": "Birincil değil",
            "confidence": "Yüksek",
            "primary_status": "Birincil değil",
            "secondary_attribution_proof": "Viru et al. (110) identified that spurts in speed and endurance occurred before and around PHV, respectively, whereas accelerated gains in strength qualities occurred after PHV (110).",
            "final_status": "Uygun değil",
            "rationale": "Makale kavramsal/anlatısal derlemedir (Youth Physical Development Model); olgunlaşma dönemlerine göre hız, kuvvet ve çeviklik gelişim pencerelerini önceki çalışmalara ([21], [106], [110] vb.) atfetmekte ve OSB örneklemi içermemektedir. Ayrıca Bulgular (Results) bölümünde kendi verilerimizin sonuna literatür atfı eklenmesi Q1 formatına uygun değildir.",
            "recommended_action": "Makale taslağında 'BİRİNCİL KAYNAK DEĞİL (Anlatısal Model Derlemesi; Bulgular Bölümü Yerine Tartışmada Birincil Veri İle Kullanılmalı)' olarak işaretlendi.",
            "hl_targets": [
                {"page": 2, "blocks": [3, 4]},
                {"page": 5, "blocks": [6, 7]},
            ],
        },
        {
            "citation_id": "A043",
            "citation_marker": "[Atıf Gerekli - 24] -> Häger-Ross ve Rösblad (2002)",
            "manuscript_location": "Bölüm 3.4 (Bulgular — Başlangıç El Karış ve Kulaç Uzunluğunun 12 Haftalık Kavrama Kazanımı İle Korelasyonu)",
            "claim_type": "Ampirik Karşılaştırma / Antropometri-Kuvvet İlişkisi",
            "claim_text": "İki değişkenli korelasyon analizlerinde başlangıç el karış uzunluğu ve kulaç uzunluğu hem sağ hem sol el mutlak kavrama kuvveti kazanımı ile pozitif ilişkili bulunmuştur.",
            "source_ref": "Häger-Ross, C., & Rösblad, B. (2002). Norms for grip strength in children aged 4–16 years. Acta Paediatrica, 91(6), 617–625.",
            "pdf_prefix": "Atif_24A_",
            "page_and_section": "Sayfa 4 (Results — Anthropometric correlations & multiple regression)",
            "decision_label": "Kısmi destek",
            "confidence": "Yüksek",
            "primary_status": "Birincil",
            "secondary_attribution_proof": "",
            "final_status": "Düzeltme gerekli",
            "rationale": "4–16 yaş arası 530 sağlıklı çocukta el boyutu (el uzunluğu / hand length) ile kesitsel zirve kavrama kuvveti arasındaki güçlü pozitif korelasyonu (erkeklerde r = 0.91) gösteren birincil araştırmadır; ancak (1) el karış/kulaç uzunluğunun 12 haftalık boylamsal antrenman kazanımı (Δ) ile ilişkisini veya OSB'yi incelememekte, (2) bu atfın Bulgular yerine Tartışma bölümünde mekanistik yorum olarak yer alması gerekmektedir.",
            "recommended_action": "Makale taslağında 'KISMİ DESTEK (Kesitsel El Uzunluğu-Kuvvet İlişkisi İçin Birincil; Boylamsal Antrenman Kazanımı Değil ve Tartışmaya Taşınmalı)' olarak işaretlendi.",
            "hl_targets": [
                {"page": 4, "blocks": [4, 5, 6]},
            ],
        },
        {
            "citation_id": "A044",
            "citation_marker": "[Atıf Gerekli - 24] -> Nicolay ve Walker (2005)",
            "manuscript_location": "Bölüm 3.4 (Bulgular — Başlangıç El Karış ve Kulaç Uzunluğunun 12 Haftalık Kavrama Kazanımı İle Korelasyonu)",
            "claim_type": "Ampirik Karşılaştırma / Antropometri-Kuvvet Biyomekaniği",
            "claim_text": "İki değişkenli korelasyon analizlerinde başlangıç el karış uzunluğu ve kulaç uzunluğu hem sağ hem sol el mutlak kavrama kuvveti kazanımı ile pozitif ilişkili bulunmuştur.",
            "source_ref": "Nicolay, C. W., & Walker, A. L. (2005). Grip strength and endurance: Influences of anthropometric variation, hand dominance, and gender. International Journal of Industrial Ergonomics, 35(7), 605–618.",
            "pdf_prefix": "Atif_24B_",
            "page_and_section": "Sayfa 8 (3.5. Association between anthropometric variables and measures of grip force) ve Sayfa 10 (4.2. Anatomical variation)",
            "decision_label": "Kısmi destek",
            "confidence": "Yüksek",
            "primary_status": "Birincil",
            "secondary_attribution_proof": "",
            "final_status": "Düzeltme gerekli",
            "rationale": "El ve önkol antropometrik boyutlarının (avuç genişliği, el uzunluğu, önkol çevresi) maksimum kavrama kuvvetiyle güçlü pozitif korelasyonunu (r > 0.70–0.82) ve dinamometre tutamak mekaniğini gösteren birincil araştırmadır; ancak sağlıklı yetişkinlerde (18–33 yaş) kesitsel olarak yürütülmüştür (OSB'li gençlerde boylamsal antrenman kazanımını incelememektedir) ve Bulgular yerine Tartışma bölümünde kullanılmalıdır.",
            "recommended_action": "Makale taslağında 'KISMİ DESTEK (Yetişkin Kesitsel Biyomekanik Çalışması; Boylamsal OSB Kazanımı Değil ve Tartışmaya Taşınmalı)' olarak işaretlendi.",
            "hl_targets": [
                {"page": 8, "blocks": [4, 5]},
                {"page": 10, "blocks": [4]},
            ],
        },
    ]

    # Populate citations list and group highlight annotations per PDF
    citations_out = []
    pdf_annotations: dict[str, list[dict]] = {}

    for spec in raw_specs:
        pref = spec["pdf_prefix"]
        if pref is not None:
            fname = pdf_by_prefix(pref)
            edata = ext_map[fname]
            all_passages = []
            for t in spec["hl_targets"]:
                bboxes, ptxt = get_block_info(edata, t["page"], t["blocks"])
                if ptxt:
                    all_passages.append(f"[S.{t['page']}] {ptxt}")
                if bboxes and spec["decision_label"] in ("Tam destek", "Kısmi destek", "Birincil değil", "Çelişki", "Düşük güvenli aday"):
                    pdf_annotations.setdefault(fname, []).append(
                        {
                            "citation_id": spec["citation_id"],
                            "page": t["page"],
                            "passage": ptxt,
                            "bbox": bboxes[0],
                            "bboxes": bboxes,
                            "exact_bboxes": bool(t.get("exact_bboxes", False)),
                            "decision_label": spec["decision_label"],
                            "primary_status": spec["primary_status"],
                            "final_status": spec["final_status"],
                            "claim_text": spec["claim_text"],
                            "rationale": spec["rationale"],
                        }
                    )
            passage_joined = " | ".join(all_passages) if all_passages else "İlgili bölüm metni PDF içerisinde yer almamaktadır (Yalnızca kapak ve içindekiler sayfaları mevcuttur)."
        else:
            fname = "PDF Mevcut Değil"
            passage_joined = "PDF dosyası Yontem_Kaynak_Makaleler_PDF klasöründe bulunmadığından metin doğrulaması yapılamamıştır."

        citations_out.append(
            {
                "citation_id": spec["citation_id"],
                "citation_marker": spec["citation_marker"],
                "manuscript_location": spec["manuscript_location"],
                "claim_type": spec["claim_type"],
                "claim_text": spec["claim_text"],
                "source_ref": spec["source_ref"],
                "pdf_file": fname,
                "page_and_section": spec["page_and_section"],
                "passage": passage_joined,
                "decision_label": spec["decision_label"],
                "confidence": spec["confidence"],
                "primary_status": spec["primary_status"],
                "secondary_attribution_proof": spec["secondary_attribution_proof"],
                "final_status": spec["final_status"],
                "rationale": spec["rationale"],
                "recommended_action": spec["recommended_action"],
                "phase8_doi_status": "Yerel PDF DOI/Künye Doğrulandı" if pref is not None else "Doğrulanamadı (PDF Eksik / Çevrimdışı)",
                "human_approval": "Bekliyor",
            }
        )

    # Run Phase 6 PDF highlighting on all highlightable PDFs
    highlight_summary = []
    for fname, anns in sorted(pdf_annotations.items()):
        src_pdf = PDF_DIR / fname
        out_pdf = BOYALI_DIR / f"{src_pdf.stem}_annotated.pdf"
        res = annotate_pdf(
            source_pdf=src_pdf,
            annotations=anns,
            output_pdf=out_pdf,
            render_dir=PNG_DIR,
        )
        # Scrub metadata on the annotated PDF while preserving annotations
        import pymupdf as fitz
        doc = fitz.open(str(out_pdf))
        doc.set_metadata({"title": "", "author": "", "subject": "", "keywords": "", "creator": "", "producer": ""})
        tmp_pdf = out_pdf.with_suffix(".clean.pdf")
        doc.save(str(tmp_pdf), garbage=3, deflate=True)
        doc.close()
        tmp_pdf.replace(out_pdf)
        highlight_summary.append(res)
        if res["applied_highlights"] < 1:
            raise RuntimeError(f"Zero highlights applied to {fname}!")

    # Unmatched & Missing inventory findings (Phase 2)
    unmatched_and_missing = [
        {
            "category": "PDF Eksik (Birincil Tanı/Ölçüm Kaynağı)",
            "item": "American Psychiatric Association (2013) — DSM-5 [A005]",
            "detail": "Atıf Rehberinde [Atıf Gerekli - 4] için listelenmiş ancak klasörde PDF'i yoktur; Lord ve ark. (2018) ikincil derleme olduğundan DSM-5 birincil kaynağı gereklidir.",
        },
        {
            "category": "PDF Eksik (Birincil İstatistik Monografisi)",
            "item": "Cohen (1988) — Statistical Power Analysis (2nd ed.) [A009, A035]",
            "detail": "Duyarlılık analizindeki orta düzey etki yorumu ve d / η² yorumlama eşikleri (0.20/0.50/0.80 ve 0.01/0.06/0.14) için birincil kaynaktır; klasörde PDF'i yoktur.",
        },
        {
            "category": "PDF Eksik (Birincil Klinik Ölçüm El Kitabı)",
            "item": "Fess (1992) — ASHT Clinical Assessment Recommendations [A021]",
            "detail": "Jamar el dinamometresi ASHT standart oturma pozisyonu için birincil kaynaktır; klasörde PDF'i yoktur (Roberts ve ark., 2011 bu kaynağa atıf yapan derlemedir).",
        },
        {
            "category": "PDF Eksik (Birincil Test Geliştirme Makalesi)",
            "item": "Wells ve Dillon (1952) — Research Quarterly [A031]",
            "detail": "Otur-eriş (Sit-and-Reach) esneklik testini tanımlayan orijinal birincil makaledir; klasörde PDF'i yoktur.",
        },
        {
            "category": "PDF Eksik (Birincil İstatistik Makalesi)",
            "item": "Shapiro ve Wilk (1965) — Biometrika [A033]",
            "detail": "Shapiro-Wilk W normallik testini türeten orijinal makaledir; klasörde PDF'i yoktur (Razali ve Wah, 2011 testin gücünü karşılaştıran simülasyon çalışmasıdır).",
        },
        {
            "category": "PDF Eksik / Ders Kitabı",
            "item": "Magill ve Anderson (2017) [A019]; Harman ve Garhammer (2008) [A029]; Aiken ve West (1991) [A039]",
            "detail": "Atıf Rehberinde yer alan ancak PDF klasöründe bulunmayan ders kitabı/monografi kaynaklarıdır.",
        },
        {
            "category": "Eksik İçerikli PDF (Yalnızca Kapak ve İçindekiler)",
            "item": "Atif_22_Tabachnick_2019_Multivariate_Statistics.pdf [A040]",
            "detail": "14 sayfalık PDF yalnızca ön kapak, içindekiler ve ek tabloları içermekte olup Bölüm 5 (Çoklu Regresyon) metnini barındırmamaktadır; ayrıca ders kitabı olduğundan birincil kaynak değildir (boyanmadı).",
        },
        {
            "category": "Künye Düzeltmesi (Dergi Eşleşmesi)",
            "item": "Atif_01_von_Elm_2007_STROBE.pdf & Atif_16B_Duncan_2008_Jump_Peak_Power_Equations.pdf",
            "detail": "von Elm ve ark. (2007) yerel PDF'i The Lancet değil PLoS Medicine (4(10), e296) baskısıdır; Duncan ve ark. (2008) yerel PDF'i Pediatric Exercise Science değil Journal of Strength and Conditioning Research (22(4), 1379–1381) makalesidir. Kaynakça künyeleri yerel PDF'lerle birebir uyumlu hale getirilmiştir.",
        },
        {
            "category": "Klasörde Ek/Atıfsız Kaynak PDF",
            "item": "Ek_Referans_Travers_2013_Motor_Difficulties_ASD.pdf",
            "detail": "Travers ve ark. (2013; J Autism Dev Disord, 43, 1568–1583) klasörde ek arka plan kaynağı olarak yer almakta olup [Atıf Gerekli 1–24] işaretleriyle eşleşmemektedir (boyanmadı).",
        },
    ]

    ts = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    audit_results = {
        "metadata": {
            "manuscript_name": "Q1_Master_Makale_Taslagi_TR.md / Q1_Master_Manuscript_EN.md (OSB 12 Haftalık Mikro-Doz Motor Beceri Antrenmanı, N = 34)",
            "citation_style": "APA 7",
            "timestamp": ts,
            "phase8_status": "Yerel/Çevrimdışı mod (Kullanıcı talimatı gereği ağ indirmesi yapılmadı; tüm yerel PDF DOI ve künyeleri doğrulandı)",
            "highlighted_pdf_count": len(highlight_summary),
            "highlighted_pdf_dir": str(BOYALI_DIR),
        },
        "pdf_inventory": pdf_inventory,
        "citations": citations_out,
        "unmatched_and_missing": unmatched_and_missing,
    }

    audit_json_path = SYS_DIR / "audit_results.json"
    audit_json_path.write_text(json.dumps(audit_results, ensure_ascii=False, indent=2), encoding="utf-8")

    # Update state.json
    state["phase"] = "completed_phase_6_and_7"
    state["updated_at"] = ts
    state["highlighted_pdfs"] = len(highlight_summary)
    (SYS_DIR / "state.json").write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")

    # Call build_report.py to create DOCX, PDF, XLSX, CSV in 05_citation_audit_boyali_pdfler/
    cmd = [
        sys.executable,
        str(SKILL_SCRIPTS / "build_report.py"),
        str(audit_json_path),
        "--out-dir",
        str(OUT_ROOT),
        "--prefix",
        "Birincil_Kaynak_Atif_Denetim_Raporu_TR",
    ]
    subprocess.run(cmd, check=True)

    print(f"SUCCESS: Audited {len(citations_out)} citation units; highlighted {len(highlight_summary)} PDFs into {BOYALI_DIR}")


if __name__ == "__main__":
    main()
