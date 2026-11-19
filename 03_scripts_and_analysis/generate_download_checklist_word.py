#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_download_checklist_word.py

Otizm 2026 Test-Tekrar Araştırması için İndirilecek Birincil Kaynakların
ayrıntılı çalışma listesini içeren APA 7 formatında Word belgesi (.docx) ve
Markdown fihristini (.md) üretir.
"""

import os
import shutil
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

BASE_DIR = "/Users/alik/Desktop/anti_projeler/articles_all/autism_2026_testtekrar"
OUTPUT_DOCX = os.path.join(BASE_DIR, "Indirilecek_Birincil_Kaynaklar_Listesi.docx")
OUTPUT_MD = os.path.join(BASE_DIR, "Indirilecek_Birincil_Kaynaklar_Listesi.md")
PDF_DIR_DOCX = os.path.join(BASE_DIR, "Yontem_Kaynak_Makaleler_PDF", "Indirilecek_Birincil_Kaynaklar_Listesi.docx")

BLACK_COLOR = RGBColor(0, 0, 0)
DARK_BLUE = RGBColor(20, 50, 100)

def set_run_font(run, size_pt=11, bold=False, italic=False, color=BLACK_COLOR):
    run.font.name = "Times New Roman"
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.first_child_found_in('w:rFonts')
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), "Times New Roman")
    rFonts.set(qn('w:hAnsi'), "Times New Roman")
    rFonts.set(qn('w:cs'), "Times New Roman")

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

def set_cell_margins(cell, top=60, bottom=60, left=70, right=70):
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

def build_document():
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Başlık
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(4)
    r_t = p_title.add_run("OTİZM 2026 TEST-TEKRAR MOTOR BECERİ ARAŞTIRMASI")
    set_run_font(r_t, size_pt=14, bold=True, color=BLACK_COLOR)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(14)
    r_s = p_sub.add_run("İndirilmesi Gereken Birincil Kaynaklar ve Standart Dosya Adlandırma Çalışma Listesi")
    set_run_font(r_s, size_pt=11.5, bold=False, italic=True, color=BLACK_COLOR)

    # Açıklama Kutusu
    p_intro = doc.add_paragraph()
    p_intro.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_intro.paragraph_format.space_before = Pt(0)
    p_intro.paragraph_format.space_after = Pt(8)
    p_intro.paragraph_format.line_spacing = 1.15
    r_in = p_intro.add_run(
        "Bu çalışma listesi, makale taslağında yer alan 24 atıf işaretleyicisinin Citation-Audit denetimi sonucunda "
        "ikincil derleme (narrative review), aktarma alıntı veya eksiklik nedeniyle elenen kaynakların yerine indirilecek "
        "gerçek birincil kaynakları içermektedir. İndirilen PDF dosyalarının doğrudan "
    )
    set_run_font(r_in, size_pt=10.5, bold=False)
    r_in_p = p_intro.add_run("autism_2026_testtekrar/Yontem_Kaynak_Makaleler_PDF/ ")
    set_run_font(r_in_p, size_pt=10.5, bold=True)
    r_in_2 = p_intro.add_run(
        "klasörüne belirtilen standart dosya adlarıyla kaydedilmesi gerekmektedir. Dosyalar eklendikten sonra Citation-Audit "
        "otomasyonu bu belgeleri tarayacak, boyalı PDF'leri üretecek ve makale taslağına tam entegrasyonu tamamlayacaktır."
    )
    set_run_font(r_in_2, size_pt=10.5, bold=False)

    # Özet İndirme Tablosu
    p_th = doc.add_paragraph()
    p_th.paragraph_format.space_before = Pt(8)
    p_th.paragraph_format.space_after = Pt(4)
    r_th1 = p_th.add_run("Tablo 1. ")
    set_run_font(r_th1, size_pt=11, bold=True)
    r_th2 = p_th.add_run("İndirilecek Birincil Kaynaklar Hızlı Kontrol Listesi")
    set_run_font(r_th2, size_pt=11, bold=False)

    headers = ["No", "Atıf İşareti", "Bölüm", "Hedef Yöntem / Ölçüm", "Önerilen Birincil Kaynak", "Kaydedilecek Dosya Adı", "Durum"]
    table_data = [
        ["1", "[Atıf 3]", "Bölüm 2.1", ">=6 Ay Bazal Plato Kriteri", "Moritani & deVries (1979)", "Atif_03_Moritani_1979_Neural_Factors.pdf", "[  ]"],
        ["2", "[Atıf 4]", "Bölüm 2.2", "DSM-5 OSB Tanı Ölçütleri", "Amerikan Psikiyatri Birliği (APA, 2013)", "Atif_04_APA_2013_DSM5_Manual.pdf", "[  ]"],
        ["3", "[Atıf 6B]", "Bölüm 2.2", "Cohen Orta Etki Eşiği (dz=0.50)", "Cohen (1988) İstatistik Kitabı", "Atif_06B_Cohen_1988_Power_Analysis.pdf", "[  ]"],
        ["4", "[Atıf 9]", "Bölüm 2.4", "Sirkadiyen Sabitleme (10:00-12:00)", "Chtourou ve ark. (2012, JSCR)", "Atif_09_Chtourou_2012_Diurnal_JSCR.pdf", "[  ]"],
        ["5", "[Atıf 10]", "Bölüm 2.4", "ISAK Kulaç ve El Karış Standartı", "Marfell-Jones ve ark. (2012, ISAK)", "Atif_10_Marfell_Jones_2012_ISAK_Standards.pdf", "[  ]"],
        ["6", "[Atıf 12]", "Bölüm 2.4", "Görsel Reaksiyon Testi (3 Deneme)", "Badau ve ark. (2018) / Eckner (2010)", "Atif_12_Badau_2018_Reaction_Time.pdf", "[  ]"],
        ["7", "[Atıf 13]", "Bölüm 2.4", "Jamar Dinamometre ASHT Pozisyonu", "Mathiowetz ve ark. (1986)", "Atif_13_Mathiowetz_1986_Grip_Norms.pdf", "[  ]"],
        ["8", "[Atıf 14]", "Bölüm 2.4", "Rölatif Kavrama Oranı (kg/kg)", "Peterson ve ark. (2016)", "Atif_14_Peterson_2016_Relative_Grip.pdf", "[  ]"],
        ["9", "[Atıf 15]", "Bölüm 2.4", "Bilateral Asimetri İndeksi (%)", "Zifchock ve ark. (2008)", "Atif_15_Zifchock_2008_Symmetry_Angle.pdf", "[  ]"],
        ["10", "[Atıf 16B]", "Bölüm 2.4", "Sayers Zirve Güç Formülü", "Sayers ve ark. (1999)", "Atif_16B_Sayers_1999_Jump_Power.pdf", "[  ]"],
        ["11", "[Atıf 17]", "Bölüm 2.4", "Pro-Agility 5-10-5 Fotosel Testi", "Stewart ve ark. (2014) / Sassi (2009)", "Atif_17_Stewart_2014_Change_of_Direction_Speed.pdf", "[  ]"],
        ["12", "[Atıf 18]", "Bölüm 2.4", "Otur-Eriş (Sit-and-Reach) Testi", "Wells ve Dillon (1952)", "Atif_18_Wells_1952_Sit_and_Reach.pdf", "[  ]"],
        ["13", "[Atıf 19]", "Bölüm 2.5", "Shapiro-Wilk Normallik Testi", "Shapiro ve Wilk (1965)", "Atif_19_Shapiro_1965_Normality_Test.pdf", "[  ]"],
        ["14", "[Atıf 20]", "Bölüm 2.5", "Cohen d ve Kısmi Eta-Kare Eşikleri", "Cohen (1988) İstatistik Kitabı", "Atif_20_Cohen_1988_Power_Analysis.pdf", "[  ]"],
        ["15", "[Atıf 22]", "Bölüm 2.5", "Çoklu Regresyon & Multikolinearite", "Aiken ve West (1991)", "Atif_22_Aiken_1991_Multiple_Regression.pdf", "[  ]"],
    ]
    col_widths = [0.4, 0.75, 0.75, 1.3, 1.3, 1.5, 0.5]

    table = doc.add_table(rows=1 + len(table_data), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_black_apa_borders(table)

    for j, h in enumerate(headers):
        cell = table.rows[0].cells[j]
        cell.width = Inches(col_widths[j])
        set_cell_border_bottom(cell, sz='6')
        set_cell_margins(cell)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT if j in [3, 4, 5] else WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.05
        r = p.add_run(h)
        set_run_font(r, size_pt=8.5, bold=True)

    for i, row in enumerate(table_data):
        row_cells = table.rows[i + 1].cells
        is_last = (i == len(table_data) - 1)
        for j, val in enumerate(row):
            cell = row_cells[j]
            cell.width = Inches(col_widths[j])
            set_cell_margins(cell)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if is_last:
                set_cell_border_bottom(cell, sz='8')
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if j in [3, 4, 5] else WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.05
            r = p.add_run(val)
            set_run_font(r, size_pt=8.5, bold=(j == 1))

    p_note = doc.add_paragraph()
    p_note.paragraph_format.space_before = Pt(3)
    p_note.paragraph_format.space_after = Pt(14)
    r_n1 = p_note.add_run("Açıklama: ")
    set_run_font(r_n1, size_pt=8.5, bold=True)
    r_n2 = p_note.add_run(
        "Cohen (1988) kaynağı hem Atıf 6B (orta düzey etki boyutu) hem de Atıf 20 (Cohen d ve eta-kare eşikleri) için ortak kaynaktır; "
        "tek bir PDF indirilmesi iki atıfı birden doğrulamaya yetecektir. Atıf 23 ve Atıf 24 ise Q1 dergi kuralları gereği Bulgular "
        "metninden tamamen kaldırılmıştır."
    )
    set_run_font(r_n2, size_pt=8.5, bold=False)

    # Ayrıntılı Kartlar Başlığı
    p_h2 = doc.add_paragraph()
    p_h2.paragraph_format.space_before = Pt(10)
    p_h2.paragraph_format.space_after = Pt(8)
    p_h2.paragraph_format.keep_with_next = True
    r_h2 = p_h2.add_run("İNDİRİLECEK BİRİNCİL KAYNAKLARIN DETAYLI KÜNYELERİ VE GEREKÇELERİ")
    set_run_font(r_h2, size_pt=12, bold=True, color=BLACK_COLOR)

    # 15 Detaylı Kart Verisi
    details = [
        {
            "num": "1",
            "marker": "[Atıf 3]",
            "section": "Bölüm 2.1 (Araştırma Tasarımı)",
            "claim": "Erken nöromotor adaptasyon platosu ve test öğrenme etkisini dışlamak amacıyla en az 6 ay kesintisiz hareket eğitimi almış olma kriteri.",
            "removed": "Skarabot ve ark. (2021) ikincil derlemedir ve erken nöral adaptasyonu Moritani ve deVries'e (1979) atfeder; Pan ve ark. (2017) ise 6 aylık ön plato kriterini test etmemektedir.",
            "primary_apa": "Moritani, T., & deVries, H. A. (1979). Neural factors versus hypertrophy in the time course of muscle strength gain. American Journal of Physical Medicine, 58(3), 115–130.",
            "alt_apa": "Sale, D. G. (1988). Neural adaptation to resistance training. Medicine and Science in Sports and Exercise, 20(5 Suppl), S135–S145. https://doi.org/10.1249/00005768-198810001-00009",
            "access": "PMID: 453338 (Moritani 1979) | PMID: 3057313 (Sale 1988)",
            "filename": "Atif_03_Moritani_1979_Neural_Factors.pdf",
        },
        {
            "num": "2",
            "marker": "[Atıf 4]",
            "section": "Bölüm 2.2 (Katılımcılar ve Güç Beyanı)",
            "claim": "Katılımcıların çocuk psikiyatristleri/nörologları tarafından DSM-5 tanı ölçütlerine göre resmi OSB tanısına sahip olması.",
            "removed": "Lord ve ark. (2018) Lancet derlemesidir ve DSM-5 tanı kriterlerini Amerikan Psikiyatri Birliği'ne (APA, 2013) atfeder.",
            "primary_apa": "American Psychiatric Association. (2013). Diagnostic and statistical manual of mental disorders (5th ed.). American Psychiatric Publishing. https://doi.org/10.1176/appi.books.9780890425596",
            "alt_apa": "-",
            "access": "DOI: 10.1176/appi.books.9780890425596 | ISBN: 978-0-89042-555-8",
            "filename": "Atif_04_APA_2013_DSM5_Manual.pdf",
        },
        {
            "num": "3",
            "marker": "[Atıf 6B]",
            "section": "Bölüm 2.2 (Katılımcılar ve Güç Beyanı)",
            "claim": "G*Power 3 hassasiyet analizinde hesaplanan dz = 0.49 değerinin orta düzey etki büyüklüğü (medium effect size) olarak tanımlanması.",
            "removed": "Faul ve ark. (2007) G*Power 3 için birincildir ve entegre edilmiştir; ancak dz = 0.50 orta etki eşiğinin birincil kaynağı olan Cohen (1988) kitap PDF'i eksiktir.",
            "primary_apa": "Cohen, J. (1988). Statistical power analysis for the behavioral sciences (2nd ed.). Lawrence Erlbaum Associates.",
            "alt_apa": "-",
            "access": "ISBN: 978-0-8058-0283-2 | Sayfa 24-27 (Effect Size Conventions)",
            "filename": "Atif_06B_Cohen_1988_Power_Analysis.pdf",
        },
        {
            "num": "4",
            "marker": "[Atıf 9]",
            "section": "Bölüm 2.4 (Veri Toplama Araçları ve Ölçüm Koşulları)",
            "claim": "Sirkadiyen ritim kaynaklı gün içi dalgalanmaları kontrol altına almak amacıyla ön-test ve son-testin günün aynı saat diliminde (10:00-12:00) yapılması.",
            "removed": "Atkinson ve Reilly (1996) ile Chtourou ve Souissi (2012) anlatısal derlemelerdir; çocuk veya sporcu kohortunda sirkadiyen performans dalgalanmasını deneysel olarak test etmez.",
            "primary_apa": "Chtourou, H., Chaouachi, A., Driss, T., Dogui, M., Behm, D. G., Chamari, K., & Souissi, N. (2012). The effect of training at the same time of day and tapering period on the diurnal variation of short exercise performances. Journal of Strength and Conditioning Research, 26(3), 697–708. https://doi.org/10.1519/JSC.0b013e3182281c36",
            "alt_apa": "Teo, W., Newton, M. J., & McGuigan, M. R. (2011). Circadian rhythms in exercise performance: Implications for hormonal and muscular adaptation. Journal of Sports Science & Medicine, 10(4), 600–606.",
            "access": "DOI: 10.1519/JSC.0b013e3182281c36 | PMID: 21997453",
            "filename": "Atif_09_Chtourou_2012_Diurnal_JSCR.pdf",
        },
        {
            "num": "5",
            "marker": "[Atıf 10]",
            "section": "Bölüm 2.4 (Veri Toplama Araçları ve Ölçüm Koşulları)",
            "claim": "Kulaç uzunluğu ve dominant el karış uzunluğunun Uluslararası Kinantropometri İlerleme Derneği (ISAK) standart protokolüne göre ölçülmesi.",
            "removed": "Nuttall (2015) VKİ tarihçesi derlemesidir; DSÖ (de Onis ve ark., 2007) VKİ persentilleri için birincildir ancak kulaç/el ölçüm protokolünü kapsamaz; ISAK el kitabı PDF'i eksiktir.",
            "primary_apa": "Marfell-Jones, M. J., Stewart, A. D., & de Ridder, J. H. (2012). International standards for anthropometric assessment. International Society for the Advancement of Kinanthropometry (ISAK).",
            "alt_apa": "Marfell-Jones, M., Olds, T., Stewart, A., & Carter, L. (2006). International standards for anthropometric assessment. ISAK.",
            "access": "ISBN: 978-0-620-36207-8 | ISAK Standartları Kılavuzu",
            "filename": "Atif_10_Marfell_Jones_2012_ISAK_Standards.pdf",
        },
        {
            "num": "6",
            "marker": "[Atıf 12]",
            "section": "Bölüm 2.4 (Veri Toplama Araçları ve Ölçüm Koşulları)",
            "claim": "Standart LED görsel uyaranlı elektronik reaksiyon zamanı testi uygulanarak 3 geçerli denemenin ortalamasının alınması.",
            "removed": "Laurenti ve ark. (2026) pilot çalışması 10 uyarılı toplam süre/en iyi 2 deneme protokolü kullanmaktadır (metindeki tek uyarı ve 3 deneme ile kısmi uyumlu); Magill ve Anderson (2017) ders kitabı PDF'i eksiktir.",
            "primary_apa": "Badau, D., Baydil, B., & Badau, A. (2018). Differences among three measures of reaction time based on hand laterality in individual sports. Sports, 6(2), Article 45. https://doi.org/10.3390/sports6020045",
            "alt_apa": "Eckner, J. T., Kutcher, J. S., & Richardson, J. K. (2010). Pilot evaluation of a novel clinical test of reaction time in National Collegiate Athletic Association Division I football players. Journal of Athletic Training, 45(4), 327–332. https://doi.org/10.4085/1062-6050-45.4.327",
            "access": "DOI: 10.3390/sports6020045 | PMID: 29910349",
            "filename": "Atif_12_Badau_2018_Reaction_Time.pdf",
        },
        {
            "num": "7",
            "marker": "[Atıf 13]",
            "section": "Bölüm 2.4 (Veri Toplama Araçları ve Ölçüm Koşulları)",
            "claim": "Jamar dinamometresi ile ASHT standartlarında (dik oturuş, omuz adduksiyonda, dirsek 90 derece fleksiyonda, önkol nötral) 3 maksimal deneme yapılması.",
            "removed": "Roberts ve ark. (2011) anlatısal bir derlemedir ve ASHT oturma pozisyonunu Fess'e (1992) ile Mathiowetz'e (1985/1986) atfeder; Fess (1992) ASHT kılavuzu PDF'i eksiktir.",
            "primary_apa": "Mathiowetz, V., Wiemer, D. M., & Federman, S. M. (1986). Grip and pinch strength: Norms for 6- to 19-year-olds. The American Journal of Occupational Therapy, 40(10), 705–711. https://doi.org/10.5014/ajot.40.10.705",
            "alt_apa": "Fess, E. E. (1992). Clinical assessment recommendations (2nd ed.). American Society of Hand Therapists.",
            "access": "DOI: 10.5014/ajot.40.10.705 | PMID: 3777107",
            "filename": "Atif_13_Mathiowetz_1986_Grip_Norms.pdf",
        },
        {
            "num": "8",
            "marker": "[Atıf 14]",
            "section": "Bölüm 2.4 (Veri Toplama Araçları ve Ölçüm Koşulları)",
            "claim": "Katılımcıların geniş vücut ağırlığı varyansını normalize etmek için iki el ortalamasının vücut ağırlığına bölünerek Rölatif El Kavrama Kuvveti (kg/kg) oranının hesaplanması.",
            "removed": "Jaric (2002) doğrusal oran ölçeklemesini (kg/kg) eleştirip allometrik kuvvet ölçeklemesi öneren teorik bir derlemedir; Wind ve ark. (2010) ise yalnızca persentil eğrileri sunmaktadır.",
            "primary_apa": "Peterson, M. D., Zhang, P., Saltarelli, W. A., Visich, P. S., & Gordon, P. M. (2016). Low muscle strength thresholds for the detection of cardiometabolic risk in adolescents. American Journal of Preventive Medicine, 50(5), 593–599. https://doi.org/10.1016/j.amepre.2015.09.019",
            "alt_apa": "Ramirez-Velez, R., et al. (2017). Handgrip strength cutoff points for identifying cardiometabolic risk in children and adolescents. Frontiers in Pediatrics, 5, Article 115.",
            "access": "DOI: 10.1016/j.amepre.2015.09.019 | PMID: 26585050",
            "filename": "Atif_14_Peterson_2016_Relative_Grip.pdf",
        },
        {
            "num": "9",
            "marker": "[Atıf 15]",
            "section": "Bölüm 2.4 (Veri Toplama Araçları ve Ölçüm Koşulları)",
            "claim": "Bilateral Fonksiyonel Asimetri İndeksi (%) = [|Sağ El - Sol El| / maks(Sağ El, Sol El)] x 100 formülüyle iki ekstremite arasındaki kuvvet dengesizliğinin belirlenmesi.",
            "removed": "Bishop ve ark. (2018) ve Impellizzeri ve ark. (2007) derleme ve aktarma çalışmalardır; asimetri denklem formüllerini önceki literatüre (Nunn 1988, Zifchock 2008) atfeder.",
            "primary_apa": "Zifchock, R. A., Davis, I., Higginson, J., & Royer, T. (2008). The Symmetry Angle: A novel, robust method of quantifying asymmetry. Gait & Posture, 27(4), 622–627. https://doi.org/10.1016/j.gaitpost.2007.08.006",
            "alt_apa": "Herzog, W., Nigg, B. M., Read, L. J., & Olsson, E. (1989). Asymmetries in ground reaction force patterns in normal human gait. Medicine and Science in Sports and Exercise, 21(1), 110–114.",
            "access": "DOI: 10.1016/j.gaitpost.2007.08.006 | PMID: 17919915",
            "filename": "Atif_15_Zifchock_2008_Symmetry_Angle.pdf",
        },
        {
            "num": "10",
            "marker": "[Atıf 16B]",
            "section": "Bölüm 2.4 (Veri Toplama Araçları ve Ölçüm Koşulları)",
            "claim": "Sayers denklemi kullanılarak dikey sıçrama yüksekliği ve vücut ağırlığından alt ekstremite zirve mekanik gücünün kestirilmesi: Zirve Güç (W) = 60.7 x Sıçrama + 45.3 x Ağırlık - 2055.",
            "removed": "Duncan ve ark. (2008) aktarma makaledir ve Sayers güç formülü katsayılarını metinde vermez; CMJ temas matı güvenirliği için Markovic (2004) birincil olarak entegre edilmiştir ancak Sayers formül makalesi eksiktir.",
            "primary_apa": "Sayers, S. P., Harackiewicz, D. V., Harman, E. A., Frykman, P. N., & Rosenstein, M. T. (1999). Cross-validation of three jump power equations. Medicine and Science in Sports and Exercise, 31(4), 572–577. https://doi.org/10.1097/00005768-199904000-00013",
            "alt_apa": "-",
            "access": "DOI: 10.1097/00005768-199904000-00013 | PMID: 10211854",
            "filename": "Atif_16B_Sayers_1999_Jump_Power.pdf",
        },
        {
            "num": "11",
            "marker": "[Atıf 17]",
            "section": "Bölüm 2.4 (Veri Toplama Araçları ve Ölçüm Koşulları)",
            "claim": "9.14 metrelik Pro-Agility (5-10-5 yarda) mekik parkurunda çift ışınlı fotosel kapıları ile çeviklik süresinin kaydedilmesi.",
            "removed": "Harman ve Garhammer (2008) NSCA ders kitabı bölümü PDF'i eksiktir; Stewart ve ark. (2014) fotosel güvenirliği için birincildir ancak parkuru Harman'a atfeder.",
            "primary_apa": "Stewart, P. F., Turner, A. N., & Miller, S. C. (2014). Reliability, factorial validity, and interrelationships of five commonly used change of direction speed tests. Scandinavian Journal of Medicine & Science in Sports, 24(3), 500–506. https://doi.org/10.1111/sms.12019",
            "alt_apa": "Sassi, R. H., et al. (2009). Relative and absolute reliability of a modified agility T-test and its relationship with vertical jump. Journal of Strength and Conditioning Research, 23(6), 1644–1651.",
            "access": "DOI: 10.1111/sms.12019 | PMID: 23198858",
            "filename": "Atif_17_Stewart_2014_Change_of_Direction_Speed.pdf",
        },
        {
            "num": "12",
            "marker": "[Atıf 18]",
            "section": "Bölüm 2.4 (Veri Toplama Araçları ve Ölçüm Koşulları)",
            "claim": "Hamstring ve lomber ekstansör esnekliğinin standart otur-eriş (sit-and-reach) sehpası ile ölçülmesi (3 denemenin en iyisi).",
            "removed": "Castro-Pinero ve ark. (2009) derlemedir ve otur-eriş testini Wells ve Dillon'a (1952) atfeder; Wells ve Dillon (1952) birincil makale PDF'i eksiktir.",
            "primary_apa": "Wells, K. F., & Dillon, E. K. (1952). The sit and reach: A test of back and leg flexibility. Research Quarterly of the American Association for Health, Physical Education and Recreation, 23(1), 115–118. https://doi.org/10.1080/10671188.1952.10761965",
            "alt_apa": "-",
            "access": "DOI: 10.1080/10671188.1952.10761965",
            "filename": "Atif_18_Wells_1952_Sit_and_Reach.pdf",
        },
        {
            "num": "13",
            "marker": "[Atıf 19]",
            "section": "Bölüm 2.5 (İstatistiksel Analiz)",
            "claim": "Sürekli değişkenlerin normallik dağılım varsayımının basıklık-çarpıklık, Q-Q grafikleri ve Shapiro-Wilk testi ile doğrulanması.",
            "removed": "Razali ve Wah (2011) simülasyon güç analizi makalesidir ve W testini doğrudan Shapiro ve Wilk'e (1965) atfeder; Shapiro ve Wilk (1965) orijinal test makalesi PDF'i eksiktir.",
            "primary_apa": "Shapiro, S. S., & Wilk, M. B. (1965). An analysis of variance test for normality (complete samples). Biometrika, 52(3/4), 591–611. https://doi.org/10.2307/2333709",
            "alt_apa": "-",
            "access": "DOI: 10.2307/2333709 | JSTOR: 2333709",
            "filename": "Atif_19_Shapiro_1965_Normality_Test.pdf",
        },
        {
            "num": "14",
            "marker": "[Atıf 20]",
            "section": "Bölüm 2.5 (İstatistiksel Analiz)",
            "claim": "Standartlaştırılmış etki büyüklükleri Cohen d (0.20 küçük, 0.50 orta, 0.80 büyük) ve kısmi eta-kare (0.01 küçük, 0.06 orta, >= 0.14 büyük) eşik katsayıları.",
            "removed": "Lakens (2013) pratik eğitim kılavuzudur ve eşik değerleri doğrudan Cohen'e (1988) atfeder; Cohen (1988) temel kitap PDF'i eksiktir.",
            "primary_apa": "Cohen, J. (1988). Statistical power analysis for the behavioral sciences (2nd ed.). Lawrence Erlbaum Associates.",
            "alt_apa": "-",
            "access": "ISBN: 978-0-8058-0283-2 | Sayfa 284-288 (ANOVA ve eta-kare)",
            "filename": "Atif_20_Cohen_1988_Power_Analysis.pdf",
        },
        {
            "num": "15",
            "marker": "[Atıf 22]",
            "section": "Bölüm 2.5 (İstatistiksel Analiz)",
            "claim": "Kuvvet kazanımları üzerinde antropometrik kaldıraç ve olgunlaşma etkileşimini açıklamak amacıyla OLS hiyerarşik regresyon modellerinin kurulması.",
            "removed": "Tabachnick ve Fidell (2019) ders kitabının klasördeki dosyası yalnızca içindekiler tablosudur; Aiken ve West (1991) etkileşim el kitabı PDF'i eksiktir.",
            "primary_apa": "Aiken, L. S., & West, S. G. (1991). Multiple regression: Testing and interpreting interactions. SAGE Publications.",
            "alt_apa": "Cohen, J., Cohen, P., West, S. G., & Aiken, L. S. (2003). Applied multiple regression/correlation analysis for the behavioral sciences (3rd ed.). Lawrence Erlbaum Associates.",
            "access": "ISBN: 978-0-7619-0712-1 | SAGE Publications",
            "filename": "Atif_22_Aiken_1991_Multiple_Regression.pdf",
        },
    ]

    for item in details:
        p_c = doc.add_paragraph()
        p_c.paragraph_format.space_before = Pt(8)
        p_c.paragraph_format.space_after = Pt(2)
        p_c.paragraph_format.keep_with_next = True
        
        r_num = p_c.add_run(f"MADDE {item['num']}: {item['marker']} - {item['section']}\n")
        set_run_font(r_num, size_pt=11, bold=True, color=DARK_BLUE)
        
        p_desc = doc.add_paragraph()
        p_desc.paragraph_format.left_indent = Inches(0.25)
        p_desc.paragraph_format.space_before = Pt(0)
        p_desc.paragraph_format.space_after = Pt(6)
        p_desc.paragraph_format.line_spacing = 1.15
        
        # Alanlar
        fields = [
            ("Hedef İddia / Ölçüm:", item['claim']),
            ("Çıkarılan İkincil Kaynak:", item['removed']),
            ("İndirilecek Birincil Kaynak:", item['primary_apa']),
            ("Alternatif Birincil Kaynak:", item['alt_apa']),
            ("Erişim / Tanımlayıcı:", item['access']),
            ("Kaydedilecek Standart Dosya Adı:", item['filename']),
            ("İndirme Durumu:", "[  ] İndirildi ve Yontem_Kaynak_Makaleler_PDF/ klasörüne eklendi"),
        ]
        for f_label, f_val in fields:
            r_l = p_desc.add_run(f"{f_label} ")
            set_run_font(r_l, size_pt=10, bold=True)
            r_v = p_desc.add_run(f"{f_val}\n")
            is_fname = (f_label == "Kaydedilecek Standart Dosya Adı:")
            set_run_font(r_v, size_pt=10, bold=is_fname, italic=False)

    doc.save(OUTPUT_DOCX)
    print(f"Saved {OUTPUT_DOCX}")

    # Ayrıca Yontem_Kaynak_Makaleler_PDF klasörüne de kopyala
    shutil.copyfile(OUTPUT_DOCX, PDF_DIR_DOCX)
    print(f"Copied to {PDF_DIR_DOCX}")

    # Markdown Fihristi Üret
    md_lines = [
        "# Otizm 2026 Test-Tekrar Motor Beceri Araştırması",
        "## İndirilmesi Gereken Birincil Kaynaklar Çalışma Listesi\n",
        "Bu çalışma listesi, Citation-Audit denetimi sonucunda ikincil derleme veya eksiklik nedeniyle elenen kaynakların yerine ",
        "doğrudan araştırmacı tarafından indirilip eklenecek birincil kaynakların künyelerini ve standart dosya adlarını içerir.\n",
        f"**Hedef İndirme Klasörü:** `{BASE_DIR}/Yontem_Kaynak_Makaleler_PDF/`\n",
        "### Tablo 1. İndirilecek Birincil Kaynaklar Hızlı Kontrol Listesi\n",
        "| No | Atıf İşareti | Bölüm | Hedef Yöntem / Ölçüm | Önerilen Birincil Kaynak | Kaydedilecek Standart Dosya Adı | İndirme Durumu |",
        "| :---: | :---: | :---: | :--- | :--- | :--- | :---: |",
    ]
    for row in table_data:
        md_lines.append(f"| {row[0]} | **{row[1]}** | {row[2]} | {row[3]} | {row[4]} | `{row[5]}` | {row[6]} |")

    md_lines.extend([
        "\n---\n",
        "### İndirilecek Birincil Kaynakların Ayrıntılı Künyeleri ve Dosya Adları\n",
    ])
    for item in details:
        md_lines.extend([
            f"#### Madde {item['num']}: {item['marker']} ({item['section']})",
            f"- **Hedef İddia / Ölçüm:** {item['claim']}",
            f"- **Çıkarılan İkincil Kaynak:** {item['removed']}",
            f"- **İndirilecek Birincil Kaynak (APA 7):** {item['primary_apa']}",
            f"- **Alternatif Kaynak:** {item['alt_apa']}",
            f"- **Erişim / Tanımlayıcı:** {item['access']}",
            f"- **Kaydedilecek Standart Dosya Adı:** `{item['filename']}`",
            f"- **İndirme Onayı:** [ ] İndirildi\n",
        ])

    with open(OUTPUT_MD, "w", encoding="utf-8") as f:
        f.write("\n".join(md_lines) + "\n")
    print(f"Saved {OUTPUT_MD}")

def scrub_docx_metadata(docx_path):
    doc = docx.Document(docx_path)
    core = doc.core_properties
    core.author = ""
    core.last_modified_by = ""
    core.comments = ""
    core.category = ""
    core.keywords = ""
    core.subject = ""
    core.title = ""
    doc.save(docx_path)

if __name__ == "__main__":
    build_document()
    scrub_docx_metadata(OUTPUT_DOCX)
    scrub_docx_metadata(PDF_DIR_DOCX)
    print("Metadata scrub completed successfully.")
