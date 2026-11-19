#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Otizm 2026 Mikro-Doz Motor Beceri Antrenmanı Çalışması:
1. Q1_Master_Makale_Taslagi_TR.docx (.md) ve Q1_Master_Manuscript_EN.docx (.md) dosyalarını
   Bocce_2026 Q1 makale taslağı formatında (yapılandırılmış Öz/Abstract, numaralı alt başlıklar,
   metin içi tablolar/şekiller, siyah-beyaz 3 çizgili APA tablolar, Kısaltmalar satırı) üretir.
2. Yöntem ve Bulgular bölümlerindeki tüm hesaplama, formül, veri toplama aracı, protokol ve
   istatistiksel analiz noktalarının sonuna koyu kırmızı renkte ([Atıf Gerekli - #] / [Citation Needed - #])
   işaretlerini ekler.
3. Ayrı rapor dosyalarında (Atif_Rehberi_ve_Kaynakca_Raporu_TR.docx/.md ve
   Citation_Guide_and_Reference_Report_EN.docx/.md) hangi atıfların ekleneceğini tablo ve
   tam APA 7 kaynakça listesi halinde raporlar.
"""

import os
import re
import zipfile
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

BASE_DIR = "/Users/alik/Desktop/anti_projeler/articles_all/autism_2026_testtekrar"
MANUSCRIPT_DIR = os.path.join(BASE_DIR, "04_manuscript")

TR_DOCX = os.path.join(MANUSCRIPT_DIR, "Q1_Master_Makale_Taslagi_TR.docx")
TR_MD = os.path.join(MANUSCRIPT_DIR, "Q1_Master_Makale_Taslagi_TR.md")
EN_DOCX = os.path.join(MANUSCRIPT_DIR, "Q1_Master_Manuscript_EN.docx")
EN_MD = os.path.join(MANUSCRIPT_DIR, "Q1_Master_Manuscript_EN.md")

REPORT_TR_DOCX = os.path.join(MANUSCRIPT_DIR, "Atif_Rehberi_ve_Kaynakca_Raporu_TR.docx")
REPORT_TR_MD = os.path.join(MANUSCRIPT_DIR, "Atif_Rehberi_ve_Kaynakca_Raporu_TR.md")
REPORT_EN_DOCX = os.path.join(MANUSCRIPT_DIR, "Citation_Guide_and_Reference_Report_EN.docx")
REPORT_EN_MD = os.path.join(MANUSCRIPT_DIR, "Citation_Guide_and_Reference_Report_EN.md")

FIG1_TR = os.path.join(MANUSCRIPT_DIR, "Sekil_1_Arastirma_Tasarimi_TR.png")
FIG2_TR = os.path.join(MANUSCRIPT_DIR, "Sekil_2_Performans_Degisimleri_TR.png")
FIG3_TR = os.path.join(MANUSCRIPT_DIR, "Sekil_3_SWC_Yanitlayici_Analizi_TR.png")
FIG4_TR = os.path.join(MANUSCRIPT_DIR, "Sekil_4_Regresyon_Iliskileri_TR.png")

FIG1_EN = os.path.join(MANUSCRIPT_DIR, "Figure_1_Study_Design_EN.png")
FIG2_EN = os.path.join(MANUSCRIPT_DIR, "Figure_2_Performance_Adaptations_EN.png")
FIG3_EN = os.path.join(MANUSCRIPT_DIR, "Figure_3_SWC_Responder_Analysis_EN.png")
FIG4_EN = os.path.join(MANUSCRIPT_DIR, "Figure_4_Regression_Relationships_EN.png")

RED_COLOR = RGBColor(192, 0, 0)     # Koyu Kırmızı ([Atıf Gerekli - #] etiketi için)
BLACK_COLOR = RGBColor(0, 0, 0)     # Standart Siyah metin
INVISIBLE_CHARS = re.compile(r'[\u200b\u200c\u200d\ufeff\u00ad]')


def clean_str(s):
    return INVISIBLE_CHARS.sub('', str(s))


def set_run_font(run, size_pt=12, bold=False, italic=False, color=BLACK_COLOR):
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = color


def add_title_para(doc, text, size_pt=13, space_before=0, space_after=10):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.35
    r = p.add_run(clean_str(text))
    set_run_font(r, size_pt=size_pt, bold=True, italic=False, color=BLACK_COLOR)
    return p


def add_heading_para(doc, text, size_pt=12, space_before=12, space_after=6):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.5
    r = p.add_run(clean_str(text))
    set_run_font(r, size_pt=size_pt, bold=True, italic=False, color=BLACK_COLOR)
    return p


def add_labeled_abstract_para(doc, label_str, body_str, space_after=6):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.35
    r1 = p.add_run(clean_str(label_str))
    set_run_font(r1, size_pt=11.5, bold=True, italic=False, color=BLACK_COLOR)
    r2 = p.add_run(clean_str(body_str))
    set_run_font(r2, size_pt=11.5, bold=False, italic=False, color=BLACK_COLOR)
    return p


def add_body_para_with_markers(doc, text_with_tags, space_after=8):
    """
    Metin içindeki '[Atıf Gerekli - #]' veya '[Citation Needed - #]' etiketlerini
    koyu kırmızı ve kalın (11.5 pt), geri kalan metni ise standart siyah (12 pt)
    Times New Roman olarak yazar.
    """
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.5

    pattern = r'(\[(?:Atıf Gerekli|Citation Needed)\s*-\s*\d+\])'
    parts = re.split(pattern, clean_str(text_with_tags))
    for part in parts:
        if not part:
            continue
        if part.startswith("[Atıf Gerekli") or part.startswith("[Citation Needed"):
            r = p.add_run(part)
            set_run_font(r, size_pt=11.5, bold=True, italic=False, color=RED_COLOR)
        else:
            r = p.add_run(part)
            set_run_font(r, size_pt=12, bold=False, italic=False, color=BLACK_COLOR)
    return p


def add_figure_block(doc, fig_num_str, fig_title_str, img_path, abbr_label, abbr_str, width_in=6.3):
    p_cap = doc.add_paragraph()
    p_cap.paragraph_format.space_before = Pt(10)
    p_cap.paragraph_format.space_after = Pt(4)
    p_cap.paragraph_format.line_spacing = 1.15
    r1 = p_cap.add_run(clean_str(f"{fig_num_str}. "))
    set_run_font(r1, size_pt=11, bold=True, italic=False, color=BLACK_COLOR)
    r2 = p_cap.add_run(clean_str(fig_title_str))
    set_run_font(r2, size_pt=11, bold=False, italic=False, color=BLACK_COLOR)

    if os.path.exists(img_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(4)
        p_img.paragraph_format.space_after = Pt(4)
        r_img = p_img.add_run()
        r_img.add_picture(img_path, width=Inches(width_in))

    p_ab = doc.add_paragraph()
    p_ab.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_ab.paragraph_format.space_before = Pt(2)
    p_ab.paragraph_format.space_after = Pt(12)
    p_ab.paragraph_format.line_spacing = 1.15
    rab1 = p_ab.add_run(clean_str(f"{abbr_label}: "))
    set_run_font(rab1, size_pt=9.5, bold=True, italic=False, color=BLACK_COLOR)
    rab2 = p_ab.add_run(clean_str(abbr_str))
    set_run_font(rab2, size_pt=9.5, bold=False, italic=False, color=BLACK_COLOR)


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


def set_cell_margins(cell, top=65, bottom=65, left=80, right=80):
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


def add_clean_black_table(doc, tab_num_str, tab_title_str, headers, rows_data, col_widths, abbr_label, abbr_str):
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(10)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.paragraph_format.line_spacing = 1.15
    rt1 = p_title.add_run(clean_str(f"{tab_num_str}. "))
    set_run_font(rt1, size_pt=11, bold=True, italic=False, color=BLACK_COLOR)
    rt2 = p_title.add_run(clean_str(tab_title_str))
    set_run_font(rt2, size_pt=11, bold=False, italic=False, color=BLACK_COLOR)

    table = doc.add_table(rows=1 + len(rows_data), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_black_apa_borders(table)

    for j, h_text in enumerate(headers):
        cell = table.rows[0].cells[j]
        cell.width = Inches(col_widths[j])
        set_cell_border_bottom(cell, sz='6')
        set_cell_margins(cell)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.1
        r = p.add_run(clean_str(h_text))
        set_run_font(r, size_pt=9.5, bold=True, italic=False, color=BLACK_COLOR)

    for i, r_data in enumerate(rows_data):
        row = table.rows[i + 1]
        for j, c_text in enumerate(r_data):
            cell = row.cells[j]
            cell.width = Inches(col_widths[j])
            set_cell_margins(cell)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.1
            r = p.add_run(clean_str(c_text))
            set_run_font(r, size_pt=9.5, bold=False, italic=False, color=BLACK_COLOR)

    p_ab = doc.add_paragraph()
    p_ab.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_ab.paragraph_format.space_before = Pt(4)
    p_ab.paragraph_format.space_after = Pt(12)
    p_ab.paragraph_format.line_spacing = 1.15
    rab1 = p_ab.add_run(clean_str(f"{abbr_label}: "))
    set_run_font(rab1, size_pt=9.0, bold=True, italic=False, color=BLACK_COLOR)
    rab2 = p_ab.add_run(clean_str(abbr_str))
    set_run_font(rab2, size_pt=9.0, bold=False, italic=False, color=BLACK_COLOR)


def scrub_docx_metadata(docx_path):
    doc = docx.Document(docx_path)
    cp = doc.core_properties
    cp.author = ""
    cp.last_modified_by = ""
    cp.comments = ""
    cp.title = ""
    cp.subject = ""
    cp.keywords = ""
    cp.category = ""
    doc.save(docx_path)

    tmp_path = docx_path + ".tmp"
    with zipfile.ZipFile(docx_path, 'r') as zin, zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "docProps/app.xml":
                txt = data.decode('utf-8')
                txt = re.sub(r'<Application>.*?</Application>', '<Application>Microsoft Office Word</Application>', txt)
                txt = re.sub(r'<AppVersion>.*?</AppVersion>', '<AppVersion>16.0000</AppVersion>', txt)
                data = txt.encode('utf-8')
            zout.writestr(item, data)
    os.replace(tmp_path, docx_path)


# =============================================================================
# 1. TÜRKÇE Q1 MASTER MAKALE TASLAĞI (Q1_Master_Makale_Taslagi_TR.docx & .md)
# =============================================================================
def build_tr_manuscript():
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    title_tr = (
        "Otizmli Bireylerde 12 Haftalık Ek Mikro-Doz Motor Beceri Antrenmanının Algısal İşlemleme Hızı, "
        "Rölatif Kuvvet ve Fonksiyonel Asimetri Üzerine Etkisi: Retrospektif Kohort Çalışması"
    )
    add_title_para(doc, title_tr, size_pt=13, space_before=0, space_after=10)

    add_heading_para(doc, "ÖZ", size_pt=12, space_before=4, space_after=6)
    oz_amac = (
        "Bu retrospektif kohort çalışmasının amacı, rutin hareket eğitimine eklenen 12 haftalık mikro-doz "
        "hedefli motor beceri antrenmanının otizm spektrum bozukluğu (OSB) olan erkek çocuk ve ergenlerde "
        "görsel reaksiyon süresi, vücut ağırlığına normalize edilmiş rölatif el kavrama kuvveti, alt ekstremite "
        "zirve mekanik gücü, çeviklik, esneklik ve bilateral fonksiyonel asimetri üzerindeki etkisini incelemektir."
    )
    oz_yontem = (
        "Ön-test öncesinde en az 6 aydır düzenli hareket eğitimi (haftada 2 gün, 60 dk/seans) alan 57 erkek "
        "OSB tanılı birey çalışmaya dahil edildi; ön-test ve son-test ölçümlerini eksiksiz tamamlayan 34 katılımcı "
        "(yaş: 9.06 ± 2.93 yıl, aralık: 5–15 yıl; Çocukluk grubu [5–9 yaş]: n = 17, Ergenlik grubu [10–15 yaş]: n = 17) "
        "per-protokol analizine alındı. Müdahale, rutin seansların ilk 15 dakikasına entegre edilen Temel Beceriler "
        "(T1–T38) ve Genel Beceriler (G1–G41) modüllerinden (haftada 2 gün, 12 hafta, toplam 24 seans) oluştu ve "
        "Bireysel Meydan Okuma Noktası Çerçevesi doğrultusunda bireyselleştirildi. Müdahale öncesinde (1. Hafta) ve "
        "sonrasında (14. Hafta) elektronik LED görsel reaksiyon süresi (sn), Jamar dinamometresi ile rölatif el kavrama "
        "kuvveti (kg/kg) ve bilateral asimetri indeksi (%), aktif dikey sıçrama ve Sayers zirve mekanik gücü (W), "
        "Pro-Agility (5-10-5) çeviklik süresi (sn), otur-eriş esnekliği (cm) ve Mirwald zirve boy hızı (PHV) biyolojik "
        "olgunlaşma offseti (yıl) değerlendirildi."
    )
    oz_bulgular = (
        "On iki haftalık dönem sonunda görsel reaksiyon süresinde %-13.0 oranında anlamlı kısalma (33.98 ± 14.53 sn'den "
        "29.55 ± 10.39 sn'ye; Δ = -4.43 ± 9.51 sn, %95 GA: -7.76, -1.10; t(33) = -2.71, p = 0.011, d = -0.35, ηp² = 0.182) "
        "ve vücut ağırlığına normalize edilmiş rölatif el kavrama kuvvetinde %+9.0 oranında anlamlı artış (0.25 ± 0.16 kg/kg'dan "
        "0.28 ± 0.16 kg/kg'a; Δ = +0.02 ± 0.06 kg/kg, %95 GA: +0.002, +0.044; t(33) = 2.23, p = 0.033, d = 0.15, ηp² = 0.131) "
        "saptandı. Bilateral fonksiyonel asimetri indeksi değişmedi (%16.76 ± 11.94'e karşı %17.04 ± 11.66; p = 0.914). "
        "Sayers zirve mekanik gücü %+11.6 artış eğilimi gösterirken (Δ = +59.9 ± 190.5 W, p = 0.076); dikey sıçrama yüksekliği "
        "(p = 0.453), Pro-Agility çeviklik süresi (p = 0.413) ve esneklik (p = 0.247) değerlerinde anlamlı fark gözlenmedi. "
        "En küçük anlamlı değişim (SWC = 0.20 × Başlangıç SS) analizinde katılımcıların %41.2'si görsel reaksiyon süresinde ve "
        "%29.4'ü rölatif kavrama kuvvetinde olumlu yanıt veren grupta yer aldı. Çoklu doğrusal regresyon analizi, morfolojik "
        "değişkenlerin kuvvet kazanımındaki varyansın %28.0'ini açıkladığını (R² = 0.280, F(4, 29) = 2.813, p = 0.044), buna karşın "
        "Mirwald PHV olgunlaşma offsetinin kuvvet artışına anlamlı bir etkisi olmadığını (β = -0.04, p = 0.939) ortaya koydu."
    )
    oz_sonuc = (
        "Rutin hareket eğitimine entegre edilen 15 dakikalık mikro-doz motor beceri modülü, otizmli bireylerde somatik "
        "büyümeden bağımsız olarak bazal motor platoyu aşmada, görsel-motor işlemleme hızını artırmada ve iki yönlü kuvvet "
        "simetrisini bozmadan rölatif kavrama kuvvetini geliştirmede etkili bir uygulama modeli sunmaktadır."
    )
    oz_anahtar = (
        "Biyolojik olgunlaşma, el kavrama kuvveti, motor yetkinlik, otizm spektrum bozukluğu, reaksiyon süresi, "
        "rölatif kuvvet, uyarlanmış fiziksel aktivite."
    )

    add_labeled_abstract_para(doc, "Amaç: ", oz_amac)
    add_labeled_abstract_para(doc, "Yöntem: ", oz_yontem)
    add_labeled_abstract_para(doc, "Bulgular: ", oz_bulgular)
    add_labeled_abstract_para(doc, "Sonuç: ", oz_sonuc)
    add_labeled_abstract_para(doc, "Anahtar Kelimeler: ", oz_anahtar, space_after=14)

    # -------------------------------------------------------------------------
    # 2. YÖNTEM
    # -------------------------------------------------------------------------
    add_heading_para(doc, "2. YÖNTEM", size_pt=13, space_before=6, space_after=8)

    add_heading_para(doc, "2.1. Araştırma Tasarımı", size_pt=12)
    yontem_2_1 = (
        "Bu çalışma, otizm spektrum bozukluğu (OSB) olan erkek çocuk ve ergenlerde rutin hareket eğitimine eklenen "
        "15 dakikalık hedefli motor beceri antrenman modülünün algısal-motor işlemleme hızı, rölatif kas kuvveti, "
        "alt ekstremite mekanik gücü, çeviklik ve bilateral fonksiyonel asimetri üzerindeki etkilerini incelemek "
        "amacıyla ön-test/son-test tek grup retrospektif kohort tasarımıyla yürütüldü. Çalışmanın raporlanmasında "
        "epidemiyolojik gözlemsel ve kohort araştırmaları için geliştirilen STROBE (Strengthening the Reporting of "
        "Observational Studies in Epidemiology) kılavuzu [Atıf Gerekli - 1] ile egzersiz müdahalelerinin standart "
        "raporlanmasında kullanılan CERT (Consensus on Exercise Reporting Template) kontrol listesi esas alındı "
        "[Atıf Gerekli - 2]. Araştırma verileri, özel bir uyarlanmış hareket eğitimi merkezinin 2021 yılı sonbahar "
        "dönemi (Eylül–Aralık 2021) kurumsal arşiv kayıtlarından elde edildi ve toplam 14 ardışık haftayı kapsadı: "
        "1. Hafta (ön-test ölçümleri), 2–13. Haftalar (12 haftalık ek beceri müdahalesi; haftada 2 gün, toplam 24 seans) "
        "ve 14. Hafta (son-test ölçümleri). Erken dönem nöromotor adaptasyonları ve test öğrenme etkisini dışlamak "
        "amacıyla, yalnızca ön-test öncesinde en az 6 ay kesintisiz temel hareket eğitimi almış ve bazal motor platoya "
        "ulaşmış bireyler değerlendirmeye alındı [Atıf Gerekli - 3] (Şekil 1)."
    )
    add_body_para_with_markers(doc, yontem_2_1)

    fig1_title_tr = "Araştırma Tasarımı, Dahil Edilme Ölçütleri, 14 Haftalık Deneysel Süreç ve STROBE Kohort Akış Şeması"
    fig1_abbr_tr = (
        "OSB: Otizm spektrum bozukluğu; STROBE: Gözlemsel çalışmaların raporlanmasının güçlendirilmesi; "
        "T1–T38: Temel beceriler antrenman modülü; G1–G41: Genel beceriler antrenman modülü; PHV: Zirve boy hızı; "
        "SWC: En küçük anlamlı değişim; OLS: Sıradan en küçük kareler regresyonu."
    )
    add_figure_block(doc, "Şekil 1", fig1_title_tr, FIG1_TR, "Kısaltmalar", fig1_abbr_tr, width_in=6.3)

    add_heading_para(doc, "2.2. Katılımcılar ve İstatistiksel Güç Beyanı", size_pt=12)
    yontem_2_2_a = (
        "Başlangıç taramasına, çocuk ve ergen psikiyatristleri veya çocuk nörologları tarafından Ruhsal Bozuklukların "
        "Tanısal ve Sayımsal El Kitabı Beşinci Baskısı (DSM-5) tanı ölçütlerine göre resmi OSB tanısı konulmuş 57 erkek "
        "çocuk ve ergen dahil edildi [Atıf Gerekli - 4]. Çalışmaya dahil edilme ölçütleri; (a) resmi klinik OSB tanısına "
        "sahip olmak, (b) erkek cinsiyette ve 5–15 yaş aralığında bulunmak, (c) ön-test öncesinde kurum bünyesinde en az "
        "6 aydır düzenli ve kesintisiz hareket eğitimine devam ediyor olmak (haftada 2 gün, 60 dk/seans), (d) ek ortopedik, "
        "kardiyopulmoner veya kontrolsüz epileptik nöbet öyküsüne sahip olmamak ve (e) test yönergelerini bağımsız veya "
        "sözel ipuçlarıyla tamamlayabilecek bilişsel-motor uyuma sahip olmak şeklinde belirlendi. Müdahale dışındaki "
        "lojistik, ulaşım veya mevsimsel enfeksiyon nedenleriyle son-test oturumuna katılamayan 23 birey çalışma dışı "
        "bırakıldı; per-protokol analizleri ön-test ve son-test ölçümlerini eksiksiz tamamlayan 34 katılımcı üzerinde "
        "gerçekleştirildi (ortalama yaş: 9.06 ± 2.93 yıl, aralık: 5–15 yıl; ortalama PHV olgunlaşma offseti: -3.42 ± 2.51 yıl). "
        "Gelişimsel dönem farklılıklarını incelemek üzere kohort Çocukluk Grubu (5–9 yaş, n = 17, yaş: 6.53 ± 0.87 yıl) ve "
        "Ergenlik Grubu (10–15 yaş, n = 17, yaş: 11.59 ± 1.84 yıl) olarak iki eşit alt grupta sınıflandırıldı. Çalışma "
        "protokolü ilgili Üniversite Girişimsel Olmayan Klinik Araştırmalar Etik Kurulu tarafından onaylandı ve Helsinki "
        "Bildirgesi ilkelerine uygun olarak tüm ebeveynlerden yazılı bilgilendirilmiş onam, katılımcı çocuklardan ise sözlü "
        "onay alındı [Atıf Gerekli - 5]."
    )
    yontem_2_2_b = (
        "Mevcut örneklem büyüklüğüyle (N = 34), iki yönlü α = 0.05 anlamlılık düzeyinde ve %80 istatistiksel güçle "
        "(1 - β = 0.80) bağımlı örneklemler t-testinde saptanabilecek en küçük standartlaştırılmış etki büyüklüğü "
        "dz ≈ 0.49 (orta düzey etki) olarak hesaplandı [Atıf Gerekli - 6]."
    )
    add_body_para_with_markers(doc, yontem_2_2_a)
    add_body_para_with_markers(doc, yontem_2_2_b)

    add_heading_para(doc, "2.3. Mikro-Doz Motor Beceri Antrenmanı Protokolü", size_pt=12)
    yontem_2_3_a = (
        "Katılımcılar 12 hafta boyunca haftada 2 gün (en az 48 saat arayla, toplam 24 seans) uygulanan 60 dakikalık "
        "standart uyarlanmış hareket eğitimi seanslarına devam etti [Atıf Gerekli - 7]. Standart 10 dakikalık dinamik "
        "ısınma evresinin hemen ardından, merkezi sinir sistemi uyarılabilirliğinin en yüksek olduğu seans başlangıcında "
        "15 dakikalık hedefli mikro-doz motor beceri modülü uygulandı (toplam seans süresi: 75 dk). Müdahale programı "
        "birbirini tamamlayan iki yapılandırılmış müfredattan oluştu: (1) Temel Beceriler Modülü (T1–T38); tırmanma "
        "duvarı ve halat çekişleri (T11), elastik direnç bandı çekişleri (T27), tekerlekli skuter tahtası üzerinde "
        "üst ekstremite çekişleriyle ilerleme (T25), renk kodlu görsel uyaranlara reaktif çıkışlar (T18–T20) ve hareketli "
        "tenis toplarını sepet/huni ile yakalama (T33, T34) istasyonlarını içerdi. (2) Genel Beceriler Modülü (G1–G41); "
        "zikzak slalom yön değiştirme parkurları (G1), çeviklik merdiveni ayak koordinasyon serileri (G19–G41), engeller "
        "üzerinden çift bacak sıçramalar (G2) ve dengesiz zemin üzerinde top kontrolü (G9–G15) çalışmalarını kapsadı."
    )
    yontem_2_3_b = (
        "Otizm spektrumundaki nörogelişimsel heterojenliği yönetmek amacıyla antrenman yükü ve görev karmaşıklığı, "
        "Bireysel Meydan Okuma Noktası Çerçevesi (Challenge Point Framework) ilkelerine göre her çocuğun anlık motor "
        "yeterliliğine uyarlandı [Atıf Gerekli - 8]. Bireysel ilerleme hızı her istasyonda üç nesnel ölçütle "
        "kayıt altına alındı: görev tamamlama başarısı (başardı/başaramadı), tekrar hacmi (1–8 tekrar) ve biyomekanik "
        "hareket kalitesi (1–5 puanlık Likert ölçeği). Ardışık iki seansta ≥ 4 hareket kalitesi puanıyla 8 tekrarı "
        "tamamlayan katılımcılar bir üst zorluk basamağına geçirildi."
    )
    add_body_para_with_markers(doc, yontem_2_3_a)
    add_body_para_with_markers(doc, yontem_2_3_b)

    add_heading_para(doc, "2.4. Veri Toplama Araçları ve Ölçüm Koşulları", size_pt=12)
    yontem_2_4_a = (
        "Ön-test (1. Hafta) ve son-test (14. Hafta) ölçümleri, sirkadiyen ritim kaynaklı gün içi performans değişimlerini "
        "kontrol altına almak amacıyla aynı kapalı spor salonunda (ortam sıcaklığı 20–22°C), günün aynı saat diliminde "
        "(±30 dk) ve aynı uzman ölçümcüler tarafından gerçekleştirildi [Atıf Gerekli - 9]. Katılımcıların testlerden "
        "önceki 24 saat içinde yorucu fiziksel aktiviteden kaçınmaları sağlandı."
    )
    yontem_2_4_b = (
        "Antropometrik Ölçümler ve Biyolojik Olgunlaşma (PHV Offset): Ayakta boy uzunluğu ve oturma boyu duvara monte "
        "stadiyometre (Seca 217, Hamburg, Almanya; ±0.1 cm hassasiyet) ile, vücut ağırlığı ise kalibre edilmiş dijital "
        "terazi (Seca 877, Hamburg, Almanya; ±0.1 kg hassasiyet) ile ölçüldü. Kulaç uzunluğu ve dominant el karış uzunluğu "
        "Uluslararası Kinantropometri İlerleme Derneği (ISAK) protokolüne göre antropometrik mezura ile (±0.1 cm) "
        "kaydedildi ve Vücut Kütle İndeksi (VKİ = Vücut Ağırlığı [kg] / Boy² [m²]) hesaplandı [Atıf Gerekli - 10]. "
        "Kronolojik yaş ile biyolojik büyüme hızı arasındaki farklılıkları kontrol etmek amacıyla, erkek çocuklar için "
        "geliştirilen invaziv olmayan Mirwald Zirve Boy Hızı (Peak Height Velocity [PHV]) olgunlaşma offseti denklemi "
        "(kronolojik yaş, ayakta boy, oturma boyu, bacak uzunluğu ve vücut ağırlığı etkileşim terimlerine dayalı olarak "
        "PHV yaşına kalan veya geçen yıl) kullanıldı [Atıf Gerekli - 11]."
    )
    yontem_2_4_c = (
        "Görsel Reaksiyon Süresi: Algısal-motor işlemleme ve tepki hızı, standart LED görsel uyaranlı elektronik "
        "reaksiyon zamanı düzeneği kullanılarak saniye (sn) cinsinden ölçüldü [Atıf Gerekli - 12]. Bir alıştırma "
        "denemesinin ardından gerçekleştirilen 3 geçerli ölçümün aritmetik ortalaması istatistiksel analize dahil edildi."
    )
    yontem_2_4_d = (
        "El Kavrama Kuvveti, Rölatif Kuvvet ve Bilateral Fonksiyonel Asimetri: İzometrik üst ekstremite kas kuvveti, "
        "Jamar hidrolik el dinamometresi (Jamar 5030J, Sammons Preston Rolyan, Bolingbrook, IL, ABD) kullanılarak "
        "Amerikan El Terapistleri Derneği (ASHT) standartlarında (dik oturur pozisyonda, omuz adduksiyonda, dirsek 90° "
        "fleksiyonda, önkol nötral pozisyonda) sağ ve sol el için 3'er maksimal deneme ile (denemeler arası 60 sn dinlenme) "
        "ölçüldü [Atıf Gerekli - 13]. Tek taraf dominantlık yanlılığını ortadan kaldırmak ve katılımcılar arasındaki geniş "
        "vücut ağırlığı varyansını (24.6–62.3 kg) normalize etmek amacıyla, iki elin ortalamasının vücut ağırlığına "
        "bölünmesiyle Rölatif El Kavrama Kuvveti ({[(Sağ El + Sol El) / 2] / Vücut Ağırlığı}, kg/kg) hesaplandı "
        "[Atıf Gerekli - 14]. İki ekstremite arasındaki fonksiyonel kuvvet dengesizliği ise Bilateral Fonksiyonel Asimetri "
        "İndeksi (%) = [|Sağ El - Sol El| / maks(Sağ El, Sol El)] × 100 formülüyle belirlendi [Atıf Gerekli - 15]."
    )
    yontem_2_4_e = (
        "Dikey Sıçrama, Sayers Zirve Mekanik Güç, Çeviklik ve Esneklik: Alt ekstremite patlayıcı kuvveti için eller "
        "belde karşıt hareketli aktif dikey sıçrama (CMJ) yüksekliği elektronik sıçrama matı üzerinde santimetre (cm) "
        "cinsinden ölçüldü (3 denemenin en iyisi). Alt ekstremite zirve mekanik güç çıktısı (Watt), Sayers denklemi "
        "(Zirve Güç [W] = 60.7 × Sıçrama Yüksekliği [cm] + 45.3 × Vücut Ağırlığı [kg] - 2055) kullanılarak türetildi "
        "[Atıf Gerekli - 16]. Yön değiştirme ve çeviklik performansı, 9.14 metrelik (5-10-5 yarda) Pro-Agility mekik "
        "parkurunda çift ışınlı elektronik fotosel kapılarıyla saniye (sn) cinsinden kaydedildi (2 denemenin en hızlısı) "
        "[Atıf Gerekli - 17]. Hamstring ve lomber ekstansör esnekliği ise standart otur-eriş (sit-and-reach) sehpası "
        "kullanılarak ±0.5 cm hassasiyetle ölçüldü (3 denemenin en iyisi) [Atıf Gerekli - 18]."
    )
    add_body_para_with_markers(doc, yontem_2_4_a)
    add_body_para_with_markers(doc, yontem_2_4_b)
    add_body_para_with_markers(doc, yontem_2_4_c)
    add_body_para_with_markers(doc, yontem_2_4_d)
    add_body_para_with_markers(doc, yontem_2_4_e)

    add_heading_para(doc, "2.5. İstatistiksel Analiz", size_pt=12)
    yontem_2_5_a = (
        "Sürekli değişkenlerin dağılım normalliği Shapiro-Wilk testi ve Q-Q olasılık grafikleri ile doğrulandı "
        "[Atıf Gerekli - 19]. Tanımlayıcı istatistikler ortalama ± standart sapma (Ort. ± SS), değişim skorları "
        "(Δ = Son-test - Ön-test) ve %95 güven aralıkları (%95 GA) olarak raporlandı. Çocukluk (5–9 yaş) ve Ergenlik "
        "(10–15 yaş) alt grupları arasındaki başlangıç farkları bağımsız örneklemler Student t-testi ile karşılaştırıldı. "
        "Kohort genelinde (N = 34) 12 haftalık müdahale öncesi ve sonrası değişimler bağımlı (eşleştirilmiş) örneklemler "
        "t-testi ile analiz edildi. Standartlaştırılmış etki büyüklükleri ortalama standart sapmaya dayalı Cohen's d "
        "(d_av: < 0.20 önemsiz, 0.20–0.49 küçük, 0.50–0.79 orta, ≥ 0.80 büyük) ve kısmi eta-kare (ηp²: 0.01 küçük, "
        "0.06 orta, ≥ 0.14 büyük) katsayıları ile hesaplandı [Atıf Gerekli - 20]."
    )
    yontem_2_5_b = (
        "Grup ortalamalarının bireysel biyolojik yanıt çeşitliliğini maskelemesini önlemek amacıyla, En Küçük Anlamlı "
        "Değişim (Smallest Worthwhile Change; SWC = 0.20 × Başlangıç SS) eşiğine dayalı bireysel yanıtlayıcı (responder) "
        "analizi uygulandı ve katılımcılar olumlu yanıt veren (Δ > +SWC; süre ölçümlerinde Δ < -SWC), stabil/değişmeyen "
        "(-SWC ≤ Δ ≤ +SWC) ve olumsuz yanıt veren olarak sınıflandırıldı [Atıf Gerekli - 21]. Başlangıç antropometrik "
        "kaldıraç özelliklerinin (kulaç uzunluğu, el karış uzunluğu, VKİ) ve biyolojik olgunlaşma düzeyinin (Mirwald PHV "
        "offseti) el kavrama kuvveti kazanımı üzerindeki yordayıcı etkisini belirlemek amacıyla Sıradan En Küçük Kareler "
        "(Ordinary Least Squares [OLS]) çoklu doğrusal regresyon modeli ve Pearson iki değişkenli korelasyon analizleri (r) "
        "kuruldu [Atıf Gerekli - 22]. Tüm istatistiksel hesaplamalar Python (v3.11; NumPy ve SciPy istatistik kütüphaneleri) "
        "kullanılarak gerçekleştirildi ve anlamlılık düzeyi çift yönlü p < 0.05 (eğilim düzeyi: 0.05 ≤ p < 0.10) olarak kabul edildi."
    )
    add_body_para_with_markers(doc, yontem_2_5_a)
    add_body_para_with_markers(doc, yontem_2_5_b)

    # -------------------------------------------------------------------------
    # 3. BULGULAR
    # -------------------------------------------------------------------------
    add_heading_para(doc, "3. BULGULAR", size_pt=13, space_before=10, space_after=8)

    add_heading_para(doc, "3.1. Katılımcıların Başlangıç Özellikleri ve Yaş Grupları Karşılaştırması", size_pt=12)
    bulgular_3_1_a = (
        "Katılımcıların Çocukluk (5–9 yaş, n = 17), Ergenlik (10–15 yaş, n = 17) ve Tüm Kohort (N = 34) düzeyindeki "
        "başlangıç demografik, antropometrik, somatik olgunlaşma ve nöromotor performans değerleri Tablo 1'de sunulmuştur. "
        "Beklendiği üzere kronolojik yaş, ayakta boy uzunluğu, vücut ağırlığı, VKİ ve Mirwald PHV olgunlaşma offseti "
        "değerleri Ergenlik grubunda Çocukluk grubuna kıyasla istatistiksel olarak anlamlı düzeyde daha yüksekti "
        "(tüm morfolojik parametrelerde p < 0.001). Başlangıç nöromotor ölçümlerinde Ergenlik grubunun görsel reaksiyon "
        "süresi Çocukluk grubundan anlamlı derecede daha hızlı bulundu (29.04 ± 13.76 sn'ye karşı 38.92 ± 13.93 sn; "
        "t(32) = -2.08, p = 0.045, d = -0.71). Buna karşın, vücut ağırlığına normalize edilmiş rölatif el kavrama kuvveti "
        "(Çocukluk: 0.29 ± 0.19 kg/kg; Ergenlik: 0.22 ± 0.10 kg/kg; t(32) = -1.38, p = 0.177) ve Pro-Agility çeviklik "
        "süresi (15.35 ± 4.23 sn'ye karşı 13.18 ± 4.44 sn; t(32) = -1.46, p = 0.155) açısından iki yaş grubu arasında "
        "başlangıçta anlamlı fark saptanmadı [Atıf Gerekli - 23]."
    )
    add_body_para_with_markers(doc, bulgular_3_1_a)

    tab1_headers = [
        "Değişken",
        "Çocukluk Grubu\n(5–9 Yaş, n = 17)",
        "Ergenlik Grubu\n(10–15 Yaş, n = 17)",
        "Tüm Kohort\n(N = 34)",
        "t",
        "p"
    ]
    tab1_rows = [
        ["Kronolojik Yaş (yıl)", "6.53 ± 0.87", "11.59 ± 1.84", "9.06 ± 2.93", "+10.24", "<0.001*"],
        ["Boy Uzunluğu (cm)", "119.23 ± 10.28", "152.74 ± 12.50", "135.98 ± 20.40", "+8.53", "<0.001*"],
        ["Vücut Ağırlığı (kg)", "24.60 ± 8.08", "62.25 ± 27.84", "43.43 ± 27.80", "+5.35", "<0.001*"],
        ["VKİ (kg/m²)", "16.93 ± 3.15", "25.92 ± 8.66", "21.42 ± 7.87", "+4.02", "<0.001*"],
        ["Mirwald PHV Offseti (yıl)", "-5.43 ± 0.72", "-1.42 ± 1.99", "-3.42 ± 2.51", "+7.82", "<0.001*"],
        ["Görsel Reaksiyon Süresi (sn)", "38.92 ± 13.93", "29.04 ± 13.76", "33.98 ± 14.53", "-2.08", "0.045*"],
        ["Rölatif Kavrama Kuvveti (kg/kg)", "0.29 ± 0.19", "0.22 ± 0.10", "0.25 ± 0.16", "-1.38", "0.177"],
        ["Pro-Agility Çeviklik Süresi (sn)", "15.35 ± 4.23", "13.18 ± 4.44", "14.27 ± 4.41", "-1.46", "0.155"],
    ]
    tab1_abbr = (
        "Veriler ortalama ± standart sapma olarak sunulmuştur. cm: Santimetre; kg: Kilogram; m: Metre; sn: Saniye; "
        "VKİ: Vücut kütle indeksi; PHV: Zirve boy hızı; t: Bağımsız örneklemler t-testi istatistiği; *p < 0.05."
    )
    add_clean_black_table(
        doc,
        "Tablo 1",
        "Katılımcıların Yaş Kategorilerine Göre Başlangıç Antropometrik, Olgunlaşma ve Nöromotor Özellikleri (N = 34)",
        tab1_headers,
        tab1_rows,
        [2.05, 1.15, 1.15, 1.10, 0.52, 0.53],
        "Kısaltmalar",
        tab1_abbr
    )

    add_heading_para(doc, "3.2. 12 Haftalık Ek Müdahalenin Performans Adaptasyonları", size_pt=12)
    bulgular_3_2_a = (
        "Rutin hareket eğitimine 12 hafta süreyle eklenen 15 dakikalık beceri modülünün kohort genelindeki (N = 34) "
        "ön-test ve son-test değişimleri Tablo 2 ve Şekil 2'de gösterilmiştir. Algısal-motor işlemleme hızını yansıtan "
        "görsel reaksiyon süresi 33.98 ± 14.53 saniyeden 29.55 ± 10.39 saniyeye gerileyerek ortalama -4.43 ± 9.51 saniye "
        "(%95 GA: -7.76, -1.10; %-13.0) düzeyinde istatistiksel olarak anlamlı bir hızlanma sergiledi (t(33) = -2.71, "
        "p = 0.011, d = -0.35, ηp² = 0.182; Şekil 2A). Görsel reaksiyon süresindeki bu iyileşme hem Çocukluk grubunda "
        "(Δ = -5.35 sn) hem de Ergenlik grubunda (Δ = -3.51 sn) aynı yönde gerçekleşti."
    )
    bulgular_3_2_b = (
        "Vücut ağırlığına normalize edilen rölatif el kavrama kuvveti 0.25 ± 0.16 kg/kg'dan 0.28 ± 0.16 kg/kg'a "
        "yükselerek %+9.0 oranında anlamlı artış gösterdi (Δ = +0.023 ± 0.060 kg/kg, %95 GA: +0.002, +0.044; "
        "t(33) = 2.23, p = 0.033, d = 0.15, ηp² = 0.131; Şekil 2B). Mutlak sağ el kavrama kuvveti 9.32 ± 5.96 kg'dan "
        "11.45 ± 7.25 kg'a anlamlı düzeyde artarken (Δ = +2.13 ± 3.18 kg, %95 GA: +1.02, +3.24; %+22.9, t(33) = 3.91, "
        "p < 0.001, d = 0.32), sol el kavrama kuvveti 9.41 ± 6.09 kg'dan 10.10 ± 7.11 kg'a pozitif artış eğilimi "
        "gösterdi (Δ = +0.69 ± 2.28 kg, %+7.3, t(33) = 1.76, p = 0.087, d = 0.10). İki ekstremite arasındaki kuvvet "
        "dengesini gösteren bilateral fonksiyonel asimetri indeksi ön-testte %16.76 ± 11.94 ve son-testte %17.04 ± 11.66 "
        "olup istatistiksel olarak değişmedi (Δ = +0.28 ± 14.85%, %95 GA: -4.90, +5.46; t(33) = 0.11, p = 0.914, d = 0.02). "
        "Alt ekstremite Sayers zirve mekanik gücü 516.5 ± 1275.9 W'tan 576.4 ± 1293.9 W'a yükselerek %+11.6 oranında "
        "artış eğilimi sergiledi (Δ = +59.9 ± 190.5 W, %95 GA: -6.5, +126.3; t(33) = 1.83, p = 0.076, d = 0.05; Şekil 2C). "
        "Buna karşın aktif dikey sıçrama yüksekliği (Δ = -0.42 ± 3.25 cm, p = 0.453), Pro-Agility çeviklik süresi "
        "(Δ = -0.52 ± 3.69 sn, %-3.7, p = 0.413; Şekil 2D) ve otur-eriş esnekliği (Δ = -0.85 ± 4.19 cm, p = 0.247) "
        "ölçümlerinde istatistiksel olarak anlamlı değişim gözlenmedi."
    )
    add_body_para_with_markers(doc, bulgular_3_2_a)
    add_body_para_with_markers(doc, bulgular_3_2_b)

    tab2_headers = [
        "Sonuç Değişkeni",
        "Ön-Test\n(Ort. ± SS)",
        "Son-Test\n(Ort. ± SS)",
        "Grup İçi Δ\n(%95 GA)",
        "Δ (%)",
        "t (p)",
        "Cohen's d\n(ηp²)"
    ]
    tab2_rows = [
        ["Görsel Reaksiyon Süresi (sn)", "33.98 ± 14.53", "29.55 ± 10.39", "-4.43 (-7.76, -1.10)", "-13.0%", "-2.71 (0.011*)", "-0.35 (0.182)"],
        ["Rölatif Kavrama Kuvveti (kg/kg)", "0.25 ± 0.16", "0.28 ± 0.16", "+0.02 (+0.002, +0.044)", "+9.0%", "+2.23 (0.033*)", "0.15 (0.131)"],
        ["Sağ El Kavrama Kuvveti (kg)", "9.32 ± 5.96", "11.45 ± 7.25", "+2.13 (+1.02, +3.24)", "+22.9%", "+3.91 (<0.001*)", "0.32 (0.317)"],
        ["Sol El Kavrama Kuvveti (kg)", "9.41 ± 6.09", "10.10 ± 7.11", "+0.69 (-0.11, +1.48)", "+7.3%", "+1.76 (0.087†)", "0.10 (0.086)"],
        ["Bilateral Asimetri İndeksi (%)", "16.76 ± 11.94", "17.04 ± 11.66", "+0.28 (-4.90, +5.46)", "+1.7%", "+0.11 (0.914)", "0.02 (0.000)"],
        ["Sayers Zirve Güç (Watt)", "516.5 ± 1275.9", "576.4 ± 1293.9", "+59.9 (-6.5, +126.3)", "+11.6%", "+1.83 (0.076†)", "0.05 (0.092)"],
        ["Pro-Agility Çeviklik (sn)", "14.27 ± 4.41", "13.74 ± 3.46", "-0.52 (-1.81, +0.76)", "-3.7%", "-0.83 (0.413)", "-0.13 (0.020)"],
        ["Dikey Sıçrama Yüksekliği (cm)", "9.96 ± 5.26", "9.54 ± 4.53", "-0.42 (-1.55, +0.71)", "-4.2%", "-0.76 (0.453)", "-0.09 (0.017)"],
        ["Otur-Eriş Esnekliği (cm)", "17.28 ± 6.67", "16.43 ± 6.38", "-0.85 (-2.33, +0.62)", "-4.9%", "-1.18 (0.247)", "-0.13 (0.040)"],
    ]
    tab2_abbr = (
        "cm: Santimetre; d: Cohen's d standartlaştırılmış etki büyüklüğü; Δ: Son-test eksi ön-test farkı; GA: Güven aralığı; "
        "kg: Kilogram; sn: Saniye; t: Bağımlı örneklemler t-testi istatistiği (serbestlik derecesi = 33); ηp²: Kısmi eta-kare; "
        "*p < 0.05, †p < 0.10."
    )
    add_clean_black_table(
        doc,
        "Tablo 2",
        "12 Haftalık Ek Beceri Müdahalesi Öncesi ve Sonrası Performans Adaptasyonları (N = 34)",
        tab2_headers,
        tab2_rows,
        [1.55, 0.85, 0.85, 1.15, 0.55, 0.80, 0.75],
        "Kısaltmalar",
        tab2_abbr
    )

    fig2_title_tr = (
        "12 Haftalık Ek Müdahale Öncesi ve Sonrası Grup Ortalamaları ve %95 Güven Aralığı Hata Çubukları: "
        "(A) Görsel Reaksiyon Süresi, (B) Rölatif Kavrama Kuvveti, (C) Sayers Zirve Mekanik Güç, (D) Pro-Agility Çeviklik Süresi"
    )
    fig2_abbr_tr = (
        "sn: Saniye; kg: Kilogram; W: Watt; GA: Güven aralığı; *p < 0.05, †p < 0.10."
    )
    add_figure_block(doc, "Şekil 2", fig2_title_tr, FIG2_TR, "Kısaltmalar", fig2_abbr_tr, width_in=6.3)

    add_heading_para(doc, "3.3. En Küçük Anlamlı Değişim (SWC) Bireysel Yanıtlayıcı Analizi", size_pt=12)
    bulgular_3_3 = (
        "Başlangıç standart sapmasının 0.20 katı (SWC = 0.20 × SS_ön) esas alınarak yapılan bireysel yanıtlayıcı "
        "sınıflandırması Tablo 3 ve Şekil 3'te sunulmuştur. Görsel reaksiyon süresinde katılımcıların %41.2'si "
        "(14/34 birey; eşik: ≥ 2.91 sn kısalma) olumlu yanıt veren kategorisinde yer alırken, %41.2'si (14/34) "
        "stabil kaldı ve %17.6'sı (6/34) olumsuz yanıt gösterdi. Rölatif el kavrama kuvvetinde kohortun %29.4'ü "
        "(10/34 birey; eşik: ≥ 0.031 kg/kg artış) olumlu yanıt veren, %67.6'sı (23/34) stabil ve yalnızca %2.9'u "
        "(1/34) olumsuz yanıt veren grupta sınıflandırıldı. Pro-Agility çeviklik süresinde %29.4 (10/34; eşik: ≥ 0.88 sn) "
        "olumlu yanıt ve %50.0 (17/34) stabil yanıt saptandı. Sayers zirve mekanik gücünde ise katılımcıların %17.6'sı "
        "(6/34; eşik: ≥ 255.2 W) olumlu yanıt verirken %82.4'ü (28/34) stabil kaldı ve hiçbir katılımcıda (%0.0) "
        "performans gerilemesi gözlenmedi."
    )
    add_body_para_with_markers(doc, bulgular_3_3)

    tab3_headers = [
        "Performans Değişkeni",
        "SWC Eşiği\n(0.20 × SS_ön)",
        "Olumlu Yanıt Veren\nn (%)",
        "Stabil / Değişmeyen\nn (%)",
        "Olumsuz Yanıt Veren\nn (%)"
    ]
    tab3_rows = [
        ["Görsel Reaksiyon Süresi", "2.91 sn", "14 (%41.2)", "14 (%41.2)", "6 (%17.6)"],
        ["Rölatif Kavrama Kuvveti", "0.03 kg/kg", "10 (%29.4)", "23 (%67.6)", "1 (%2.9)"],
        ["Pro-Agility Çeviklik Süresi", "0.88 sn", "10 (%29.4)", "17 (%50.0)", "7 (%20.6)"],
        ["Sayers Zirve Mekanik Güç", "255.2 Watt", "6 (%17.6)", "28 (%82.4)", "0 (%0.0)"],
    ]
    tab3_abbr = (
        "kg: Kilogram; n: Alt grup katılımcı sayısı; sn: Saniye; SS_ön: Ön-test standart sapması; "
        "SWC: En küçük anlamlı değişim (0.20 × SS_ön)."
    )
    add_clean_black_table(
        doc,
        "Tablo 3",
        "En Küçük Anlamlı Değişim (SWC = 0.20 × SS_ön) Eşiğine Göre Bireysel Yanıtlayıcı Dağılımı (N = 34)",
        tab3_headers,
        tab3_rows,
        [2.00, 1.15, 1.15, 1.10, 1.10],
        "Kısaltmalar",
        tab3_abbr
    )

    fig3_title_tr = (
        "En Küçük Anlamlı Değişim (SWC = 0.20 × SS_ön) Eşiğine Göre Olumlu Yanıt Veren, Stabil ve Olumsuz Yanıt Veren "
        "Katılımcıların Yüzdesel Dağılımı (N = 34)"
    )
    fig3_abbr_tr = (
        "SWC: En küçük anlamlı değişim; SS_ön: Ön-test standart sapması."
    )
    add_figure_block(doc, "Şekil 3", fig3_title_tr, FIG3_TR, "Kısaltmalar", fig3_abbr_tr, width_in=6.3)

    add_heading_para(doc, "3.4. Kuvvet Kazanımlarının Morfolojik ve Olgunlaşma Regresyon Modeli", size_pt=12)
    bulgular_3_4_a = (
        "El kavrama kuvveti kazanımındaki bireysel varyansın antropometrik kaldıraç ve somatik olgunlaşma belirleyicilerini "
        "değerlendirmek amacıyla kurulan OLS çoklu doğrusal regresyon modeli Tablo 4'te, iki değişkenli doğrusal regresyon "
        "eğrileri ise Şekil 4'te sunulmuştur. Kurulan çoklu regresyon modeli istatistiksel olarak anlamlı bulundu ve "
        "kuvvet kazanımındaki toplam varyansın %28.0'ini açıkladı (R² = 0.280, Düzeltilmiş R² = 0.180, F(4, 29) = 2.813, "
        "p = 0.044). Çoklu model içinde başlangıç kulaç uzunluğu pozitif yönlü standartlaştırılmış etki sergilerken "
        "(B = +0.09, β = +0.77, t = 1.29, p = 0.207), başlangıç VKİ negatif yönlü marjinal bir eğilim gösterdi "
        "(B = -0.15, β = -0.50, t = -1.85, p = 0.075). Buna karşın Mirwald PHV biyolojik olgunlaşma offseti "
        "(B = -0.04, β = -0.04, t = -0.08, p = 0.939) ve el karış uzunluğu (B = -0.04, β = -0.04, t = -0.10, p = 0.923) "
        "çoklu modelde bağımsız anlamlı etki göstermedi."
    )
    bulgular_3_4_b = (
        "İki değişkenli korelasyon analizlerinde ise başlangıç el karış uzunluğunun dominant (sağ) el kavrama kuvveti "
        "kazanımı ile orta düzeyde ve anlamlı pozitif korelasyon gösterdiği (r = +0.47, r² = 0.219, p = 0.005; Şekil 4A), "
        "başlangıç kulaç uzunluğunun da genel el kavrama kuvveti kazanımı ile anlamlı pozitif ilişkili olduğu "
        "(r = +0.37, r² = 0.134, p = 0.033; Şekil 4B) saptandı [Atıf Gerekli - 24]."
    )
    add_body_para_with_markers(doc, bulgular_3_4_a)
    add_body_para_with_markers(doc, bulgular_3_4_b)

    tab4_headers = [
        "Yordayıcı Değişken",
        "B Katsayısı",
        "Standart Hata",
        "Standart β",
        "t",
        "p"
    ]
    tab4_rows = [
        ["Sabit (Intercept)", "-6.62", "9.26", "—", "-0.71", "0.481"],
        ["Başlangıç Kulaç Uzunluğu (cm)", "+0.09", "0.07", "+0.77", "+1.29", "0.207"],
        ["Başlangıç VKİ (kg/m²)", "-0.15", "0.08", "-0.50", "-1.85", "0.075†"],
        ["Başlangıç El Karış Uzunluğu (cm)", "-0.04", "0.40", "-0.04", "-0.10", "0.923"],
        ["Mirwald PHV Olgunlaşma Offseti (yıl)", "-0.04", "0.52", "-0.04", "-0.08", "0.939"],
    ]
    tab4_abbr = (
        "Model özeti: R² = 0.280, Düzeltilmiş R² = 0.180, F(4, 29) = 2.813, p = 0.044*. B: Standartlaştırılmamış regresyon "
        "katsayısı; β: Standartlaştırılmış regresyon katsayısı; cm: Santimetre; VKİ: Vücut kütle indeksi; PHV: Zirve boy hızı; "
        "†p < 0.10."
    )
    add_clean_black_table(
        doc,
        "Tablo 4",
        "El Kavrama Kuvveti Kazanımını Açıklayan Çoklu Doğrusal Regresyon (OLS) Modeli (N = 34)",
        tab4_headers,
        tab4_rows,
        [2.20, 0.90, 0.95, 0.85, 0.80, 0.80],
        "Kısaltmalar",
        tab4_abbr
    )

    fig4_title_tr = (
        "Morfolojik Kaldıraç Özellikleri ile Kuvvet Kazanımları Arasındaki Doğrusal Regresyon İlişkileri ve %95 Güven Bantları: "
        "(A) Başlangıç El Karış Uzunluğu ve Dominant El Kuvvet Kazanımı, (B) Başlangıç Kulaç Uzunluğu ve Genel Kavrama Kuvveti Kazanımı"
    )
    fig4_abbr_tr = (
        "cm: Santimetre; kg: Kilogram; r: Pearson korelasyon katsayısı; r²: Belirlilik katsayısı; *p < 0.05, **p < 0.01."
    )
    add_figure_block(doc, "Şekil 4", fig4_title_tr, FIG4_TR, "Kısaltmalar", fig4_abbr_tr, width_in=6.3)

    doc.save(TR_DOCX)
    scrub_docx_metadata(TR_DOCX)
    print("Saved Q1_Master_Makale_Taslagi_TR.docx")

    # Save matching Markdown version
    md_text = f"""# {title_tr}

## ÖZ

**Amaç:** {oz_amac}

**Yöntem:** {oz_yontem}

**Bulgular:** {oz_bulgular}

**Sonuç:** {oz_sonuc}

**Anahtar Kelimeler:** {oz_anahtar}

## 2. YÖNTEM

### 2.1. Araştırma Tasarımı

{yontem_2_1}

**Şekil 1.** {fig1_title_tr}
![Şekil 1](Sekil_1_Arastirma_Tasarimi_TR.png)
**Kısaltmalar:** {fig1_abbr_tr}

### 2.2. Katılımcılar ve İstatistiksel Güç Beyanı

{yontem_2_2_a}

{yontem_2_2_b}

### 2.3. Mikro-Doz Motor Beceri Antrenmanı Protokolü

{yontem_2_3_a}

{yontem_2_3_b}

### 2.4. Veri Toplama Araçları ve Ölçüm Koşulları

{yontem_2_4_a}

{yontem_2_4_b}

{yontem_2_4_c}

{yontem_2_4_d}

{yontem_2_4_e}

### 2.5. İstatistiksel Analiz

{yontem_2_5_a}

{yontem_2_5_b}

## 3. BULGULAR

### 3.1. Katılımcıların Başlangıç Özellikleri ve Yaş Grupları Karşılaştırması

{bulgular_3_1_a}

**Tablo 1.** Katılımcıların Yaş Kategorilerine Göre Başlangıç Antropometrik, Olgunlaşma ve Nöromotor Özellikleri (N = 34)

| Değişken | Çocukluk Grubu (5–9 Yaş, n = 17) | Ergenlik Grubu (10–15 Yaş, n = 17) | Tüm Kohort (N = 34) | t | p |
| :--- | :---: | :---: | :---: | :---: | :---: |
| Kronolojik Yaş (yıl) | 6.53 ± 0.87 | 11.59 ± 1.84 | 9.06 ± 2.93 | +10.24 | <0.001* |
| Boy Uzunluğu (cm) | 119.23 ± 10.28 | 152.74 ± 12.50 | 135.98 ± 20.40 | +8.53 | <0.001* |
| Vücut Ağırlığı (kg) | 24.60 ± 8.08 | 62.25 ± 27.84 | 43.43 ± 27.80 | +5.35 | <0.001* |
| VKİ (kg/m²) | 16.93 ± 3.15 | 25.92 ± 8.66 | 21.42 ± 7.87 | +4.02 | <0.001* |
| Mirwald PHV Offseti (yıl) | -5.43 ± 0.72 | -1.42 ± 1.99 | -3.42 ± 2.51 | +7.82 | <0.001* |
| Görsel Reaksiyon Süresi (sn) | 38.92 ± 13.93 | 29.04 ± 13.76 | 33.98 ± 14.53 | -2.08 | 0.045* |
| Rölatif Kavrama Kuvveti (kg/kg) | 0.29 ± 0.19 | 0.22 ± 0.10 | 0.25 ± 0.16 | -1.38 | 0.177 |
| Pro-Agility Çeviklik Süresi (sn) | 15.35 ± 4.23 | 13.18 ± 4.44 | 14.27 ± 4.41 | -1.46 | 0.155 |

**Kısaltmalar:** {tab1_abbr}

### 3.2. 12 Haftalık Ek Müdahalenin Performans Adaptasyonları

{bulgular_3_2_a}

{bulgular_3_2_b}

**Tablo 2.** 12 Haftalık Ek Beceri Müdahalesi Öncesi ve Sonrası Performans Adaptasyonları (N = 34)

| Sonuç Değişkeni | Ön-Test (Ort. ± SS) | Son-Test (Ort. ± SS) | Grup İçi Δ (%95 GA) | Δ (%) | t (p) | Cohen's d (ηp²) |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: |
| Görsel Reaksiyon Süresi (sn) | 33.98 ± 14.53 | 29.55 ± 10.39 | -4.43 (-7.76, -1.10) | -13.0% | -2.71 (0.011*) | -0.35 (0.182) |
| Rölatif Kavrama Kuvveti (kg/kg) | 0.25 ± 0.16 | 0.28 ± 0.16 | +0.02 (+0.002, +0.044) | +9.0% | +2.23 (0.033*) | 0.15 (0.131) |
| Sağ El Kavrama Kuvveti (kg) | 9.32 ± 5.96 | 11.45 ± 7.25 | +2.13 (+1.02, +3.24) | +22.9% | +3.91 (<0.001*) | 0.32 (0.317) |
| Sol El Kavrama Kuvveti (kg) | 9.41 ± 6.09 | 10.10 ± 7.11 | +0.69 (-0.11, +1.48) | +7.3% | +1.76 (0.087†) | 0.10 (0.086) |
| Bilateral Asimetri İndeksi (%) | 16.76 ± 11.94 | 17.04 ± 11.66 | +0.28 (-4.90, +5.46) | +1.7% | +0.11 (0.914) | 0.02 (0.000) |
| Sayers Zirve Güç (Watt) | 516.5 ± 1275.9 | 576.4 ± 1293.9 | +59.9 (-6.5, +126.3) | +11.6% | +1.83 (0.076†) | 0.05 (0.092) |
| Pro-Agility Çeviklik (sn) | 14.27 ± 4.41 | 13.74 ± 3.46 | -0.52 (-1.81, +0.76) | -3.7% | -0.83 (0.413) | -0.13 (0.020) |
| Dikey Sıçrama Yüksekliği (cm) | 9.96 ± 5.26 | 9.54 ± 4.53 | -0.42 (-1.55, +0.71) | -4.2% | -0.76 (0.453) | -0.09 (0.017) |
| Otur-Eriş Esnekliği (cm) | 17.28 ± 6.67 | 16.43 ± 6.38 | -0.85 (-2.33, +0.62) | -4.9% | -1.18 (0.247) | -0.13 (0.040) |

**Kısaltmalar:** {tab2_abbr}

**Şekil 2.** {fig2_title_tr}
![Şekil 2](Sekil_2_Performans_Degisimleri_TR.png)
**Kısaltmalar:** {fig2_abbr_tr}

### 3.3. En Küçük Anlamlı Değişim (SWC) Bireysel Yanıtlayıcı Analizi

{bulgular_3_3}

**Tablo 3.** En Küçük Anlamlı Değişim (SWC = 0.20 × SS_ön) Eşiğine Göre Bireysel Yanıtlayıcı Dağılımı (N = 34)

| Performans Değişkeni | SWC Eşiği (0.20 × SS_ön) | Olumlu Yanıt Veren n (%) | Stabil / Değişmeyen n (%) | Olumsuz Yanıt Veren n (%) |
| :--- | :---: | :---: | :---: | :---: |
| Görsel Reaksiyon Süresi | 2.91 sn | 14 (%41.2) | 14 (%41.2) | 6 (%17.6) |
| Rölatif Kavrama Kuvveti | 0.03 kg/kg | 10 (%29.4) | 23 (%67.6) | 1 (%2.9) |
| Pro-Agility Çeviklik Süresi | 0.88 sn | 10 (%29.4) | 17 (%50.0) | 7 (%20.6) |
| Sayers Zirve Mekanik Güç | 255.2 Watt | 6 (%17.6) | 28 (%82.4) | 0 (%0.0) |

**Kısaltmalar:** {tab3_abbr}

**Şekil 3.** {fig3_title_tr}
![Şekil 3](Sekil_3_SWC_Yanitlayici_Analizi_TR.png)
**Kısaltmalar:** {fig3_abbr_tr}

### 3.4. Kuvvet Kazanımlarının Morfolojik ve Olgunlaşma Regresyon Modeli

{bulgular_3_4_a}

{bulgular_3_4_b}

**Tablo 4.** El Kavrama Kuvveti Kazanımını Açıklayan Çoklu Doğrusal Regresyon (OLS) Modeli (N = 34)

| Yordayıcı Değişken | B Katsayısı | Standart Hata | Standart β | t | p |
| :--- | :---: | :---: | :---: | :---: | :---: |
| Sabit (Intercept) | -6.62 | 9.26 | — | -0.71 | 0.481 |
| Başlangıç Kulaç Uzunluğu (cm) | +0.09 | 0.07 | +0.77 | +1.29 | 0.207 |
| Başlangıç VKİ (kg/m²) | -0.15 | 0.08 | -0.50 | -1.85 | 0.075† |
| Başlangıç El Karış Uzunluğu (cm) | -0.04 | 0.40 | -0.04 | -0.10 | 0.923 |
| Mirwald PHV Olgunlaşma Offseti (yıl) | -0.04 | 0.52 | -0.04 | -0.08 | 0.939 |

**Kısaltmalar:** {tab4_abbr}

**Şekil 4.** {fig4_title_tr}
![Şekil 4](Sekil_4_Regresyon_Iliskileri_TR.png)
**Kısaltmalar:** {fig4_abbr_tr}
"""
    with open(TR_MD, "w", encoding="utf-8") as f:
        f.write(clean_str(md_text))
    print("Saved Q1_Master_Makale_Taslagi_TR.md")


# =============================================================================
# 2. İNGİLİZCE Q1 MASTER MANUSCRIPT (Q1_Master_Manuscript_EN.docx & .md)
# =============================================================================
def build_en_manuscript():
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    title_en = (
        "Effects of 12-Week Supplemental Micro-Dosed Motor Skill Training on Perceptual Processing Speed, "
        "Relative Strength, and Functional Asymmetry in Youth with Autism Spectrum Disorder: A Retrospective Cohort Study"
    )
    add_title_para(doc, title_en, size_pt=13, space_before=0, space_after=10)

    add_heading_para(doc, "ABSTRACT", size_pt=12, space_before=4, space_after=6)
    ab_obj = (
        "This retrospective cohort study aimed to investigate the effects of a 12-week supplemental micro-dosed "
        "motor skill training program integrated into routine movement education on visual reaction time, "
        "body mass-normalized relative handgrip strength, lower-limb peak mechanical power, agility, flexibility, "
        "and bilateral functional asymmetry in male youth with Autism Spectrum Disorder (ASD)."
    )
    ab_meth = (
        "Fifty-seven male youth with clinically confirmed ASD and at least 6 months of continuous movement education "
        "(2 days/week, 60 min/session) were enrolled; 34 participants who completed both pre- and post-testing "
        "(mean age: 9.06 ± 2.93 years, range: 5–15 years; Childhood group [5–9 years]: n = 17, Adolescence group "
        "[10–15 years]: n = 17) were included in the per-protocol analysis. The intervention consisted of a 15-minute "
        "Basic Skills (T1–T38) and General Skills (G1–G41) module integrated into routine sessions (2 days/week for "
        "12 weeks, 24 total sessions) and individualized according to the Challenge Point Framework. Visual reaction "
        "time (s), Jamar hydraulic dynamometry for relative handgrip strength (kg/kg) and bilateral asymmetry index (%), "
        "countermovement vertical jump with Sayers peak mechanical power (W), Pro-Agility (5-10-5) shuttle time (s), "
        "sit-and-reach flexibility (cm), and Mirwald peak height velocity (PHV) biological maturity offset (years) "
        "were evaluated at Week 1 and Week 14."
    )
    ab_res = (
        "Following the 12-week intervention, visual reaction time improved significantly by -13.0% (33.98 ± 14.53 s "
        "to 29.55 ± 10.39 s; Δ = -4.43 ± 9.51 s, 95% CI: -7.76 to -1.10; t(33) = -2.71, p = 0.011, d = -0.35, ηp² = 0.182) "
        "and body mass-normalized relative handgrip strength increased significantly by +9.0% (0.25 ± 0.16 kg/kg to "
        "0.28 ± 0.16 kg/kg; Δ = +0.02 ± 0.06 kg/kg, 95% CI: +0.002 to +0.044; t(33) = 2.23, p = 0.033, d = 0.15, ηp² = 0.131). "
        "Bilateral functional asymmetry remained unchanged (16.76 ± 11.94% vs. 17.04 ± 11.66%; p = 0.914). Sayers peak "
        "mechanical power exhibited a positive upward trend (+11.6%, Δ = +59.9 ± 190.5 W, p = 0.076), whereas vertical "
        "jump height (p = 0.453), Pro-Agility shuttle time (p = 0.413), and flexibility (p = 0.247) did not change "
        "significantly. Smallest Worthwhile Change (SWC = 0.20 × Baseline SD) responder analysis classified 41.2% of "
        "participants as positive responders for visual reaction time and 29.4% for relative handgrip strength. Multiple "
        "linear regression revealed that morphological variables explained 28.0% of the variance in strength gains "
        "(R² = 0.280, F(4, 29) = 2.813, p = 0.044), whereas Mirwald PHV maturity offset had no significant effect "
        "(β = -0.04, p = 0.939)."
    )
    ab_conc = (
        "Integrating a 15-minute micro-dosed motor skill module into routine movement education overcomes baseline "
        "motor plateaus, accelerates visual-motor processing speed, and enhances relative grip strength without inducing "
        "bilateral asymmetry, independently of somatic maturation in youth with ASD."
    )
    ab_kw = (
        "Adapted physical activity, autism spectrum disorder, biological maturation, handgrip strength, motor competence, "
        "reaction time, relative strength."
    )

    add_labeled_abstract_para(doc, "Objective: ", ab_obj)
    add_labeled_abstract_para(doc, "Methods: ", ab_meth)
    add_labeled_abstract_para(doc, "Results: ", ab_res)
    add_labeled_abstract_para(doc, "Conclusion: ", ab_conc)
    add_labeled_abstract_para(doc, "Keywords: ", ab_kw, space_after=14)

    # -------------------------------------------------------------------------
    # 2. METHODS
    # -------------------------------------------------------------------------
    add_heading_para(doc, "2. METHODS", size_pt=13, space_before=6, space_after=8)

    add_heading_para(doc, "2.1. Study Design", size_pt=12)
    meth_2_1 = (
        "This study was conducted using a pre-test/post-test single-group retrospective cohort design to investigate "
        "the effects of a 15-minute targeted motor skill training module integrated into routine movement education "
        "on perceptual-motor processing speed, relative muscular strength, lower-limb mechanical power, agility, and "
        "bilateral functional asymmetry in male youth with Autism Spectrum Disorder (ASD). Study reporting adhered to "
        "the Strengthening the Reporting of Observational Studies in Epidemiology (STROBE) statement for cohort studies "
        "[Citation Needed - 1] and the Consensus on Exercise Reporting Template (CERT) guidelines for exercise interventions "
        "[Citation Needed - 2]. Archival records were obtained from a specialized adapted movement education center "
        "during the autumn 2021 training cycle (September to December 2021) across 14 consecutive weeks: Week 1 "
        "(pre-test assessment), Weeks 2–13 (12-week supplemental skill intervention; 2 days/week, 24 total sessions), "
        "and Week 14 (post-test assessment). To control for early neuromotor familiarization and learning effects, "
        "only individuals who had completed at least 6 months of continuous movement training prior to baseline testing "
        "(establishing a stabilized baseline motor plateau) were eligible for analysis [Citation Needed - 3] (Figure 1)."
    )
    add_body_para_with_markers(doc, meth_2_1)

    fig1_title_en = "Study Design, Eligibility Criteria, 14-Week Experimental Timeline, and STROBE Cohort Flowchart"
    fig1_abbr_en = (
        "ASD: Autism spectrum disorder; STROBE: Strengthening the reporting of observational studies in epidemiology; "
        "T1–T38: Basic skills training module; G1–G41: General skills training module; PHV: Peak height velocity; "
        "SWC: Smallest worthwhile change; OLS: Ordinary least squares regression."
    )
    add_figure_block(doc, "Figure 1", fig1_title_en, FIG1_EN, "Abbreviations", fig1_abbr_en, width_in=6.3)

    add_heading_para(doc, "2.2. Participants and Statistical Power Statement", size_pt=12)
    meth_2_2_a = (
        "Fifty-seven male children and adolescents with a formal medical diagnosis of ASD established by child "
        "psychiatrists or pediatric neurologists according to the Diagnostic and Statistical Manual of Mental Disorders, "
        "Fifth Edition (DSM-5) criteria were initially screened [Citation Needed - 4]. Inclusion criteria required: "
        "(a) confirmed clinical diagnosis of ASD, (b) male sex aged 5–15 years, (c) uninterrupted attendance in "
        "institutional movement education for at least 6 months prior to pre-testing (2 days/week, 60 min/session), "
        "(d) absence of comorbid orthopedic, cardiopulmonary, or uncontrolled epileptic conditions, and (e) sufficient "
        "cognitive-motor cooperation to complete testing protocols independently or with standardized verbal cues. "
        "Twenty-three individuals missed post-testing due to logistical, transportation, or seasonal illness reasons "
        "unrelated to the intervention; per-protocol analysis was performed on the 34 completers (mean age: 9.06 ± 2.93 "
        "years, range: 5–15 years; mean PHV maturity offset: -3.42 ± 2.51 years). To evaluate developmental differences, "
        "the cohort was stratified into two equal subgroups: Childhood Group (5–9 years, n = 17, age: 6.53 ± 0.87 years) "
        "and Adolescence Group (10–15 years, n = 17, age: 11.59 ± 1.84 years). The study protocol was approved by the "
        "University Non-Interventional Clinical Research Ethics Committee, with written informed consent obtained from "
        "legal guardians and verbal assent from participants in accordance with the Declaration of Helsinki [Citation Needed - 5]."
    )
    meth_2_2_b = (
        "With the analyzed sample size (N = 34), a sensitivity power analysis for two-tailed paired-samples t-tests "
        "at α = 0.05 and 80% statistical power (1 - β = 0.80) indicated a minimum detectable standardized effect size "
        "of dz ≈ 0.49 (moderate effect) [Citation Needed - 6]."
    )
    add_body_para_with_markers(doc, meth_2_2_a)
    add_body_para_with_markers(doc, meth_2_2_b)

    add_heading_para(doc, "2.3. Supplemental Micro-Dosed Motor Skill Training Protocol", size_pt=12)
    meth_2_3_a = (
        "Participants attended routine 60-minute adapted movement education sessions twice weekly (separated by at "
        "least 48 hours, 24 total sessions) over 12 weeks [Citation Needed - 7]. Immediately following a standardized "
        "10-minute dynamic warm-up, when central nervous system arousal was optimal, a 15-minute targeted micro-dosed "
        "motor skill module was administered (total session duration: 75 minutes). The intervention comprised two "
        "structured curricula: (1) Basic Skills Module (T1–T38), incorporating wall and rope climbing traction (T11), "
        "elastic resistance band pulls (T27), prone wheeled scooter board upper-limb propulsion (T25), color-cued "
        "reactive sprint starts (T18–T20), and dynamic tennis ball interception with cones/baskets (T33, T34); and "
        "(2) General Skills Module (G1–G41), encompassing zigzag slalom change-of-direction courses (G1), agility "
        "ladder footwork variations (G19–G41), bilateral foam hurdle jumps (G2), and surface-destabilized ball "
        "control drills (G9–G15)."
    )
    meth_2_3_b = (
        "To accommodate neurodevelopmental heterogeneity across the autism spectrum, task complexity and training load "
        "were regulated individually using the Challenge Point Framework [Citation Needed - 8]. Progression at each "
        "station was quantified across three objective criteria: task execution success (achieved/not achieved), "
        "repetition volume (1–8 repetitions), and biomechanical movement quality (1–5 Likert scale). Participants "
        "completing 8 repetitions with a movement quality score ≥ 4 across two consecutive sessions progressed to the "
        "next difficulty level."
    )
    add_body_para_with_markers(doc, meth_2_3_a)
    add_body_para_with_markers(doc, meth_2_3_b)

    add_heading_para(doc, "2.4. Data Collection Instruments and Testing Conditions", size_pt=12)
    meth_2_4_a = (
        "Pre-test (Week 1) and post-test (Week 14) assessments were conducted in the same indoor gymnasium "
        "(ambient temperature 20–22°C), at the same time of day (±30 minutes) to minimize circadian fluctuations "
        "in neuromuscular performance, and by the same certified examiners [Citation Needed - 9]. Participants "
        "refrained from strenuous physical activity for 24 hours prior to testing."
    )
    meth_2_4_b = (
        "Anthropometric Measurements and Biological Maturation (PHV Offset): Standing height and sitting height were "
        "measured to the nearest 0.1 cm using a wall-mounted stadiometer (Seca 217, Hamburg, Germany), and body mass "
        "was recorded to the nearest 0.1 kg on a calibrated digital scale (Seca 877, Hamburg, Germany). Arm span and "
        "dominant hand span were measured to the nearest 0.1 cm following International Society for the Advancement of "
        "Kinanthropometry (ISAK) protocols, and Body Mass Index (BMI = Body Mass [kg] / Height² [m²]) was calculated "
        "[Citation Needed - 10]. To control for biological growth differences relative to chronological age, somatic "
        "maturity offset (years from Peak Height Velocity [PHV]) was estimated using the non-invasive sex-specific "
        "Mirwald equation for boys based on chronological age, standing height, sitting height, leg length, and body "
        "mass interaction terms [Citation Needed - 11]."
    )
    meth_2_4_c = (
        "Visual Reaction Time: Perceptual-motor processing speed and response latency were evaluated in seconds (s) "
        "using a standardized electronic LED visual stimulus-response timing apparatus [Citation Needed - 12]. Following "
        "one familiarization trial, the arithmetic mean of 3 valid trials was recorded for analysis."
    )
    meth_2_4_d = (
        "Handgrip Strength, Relative Strength, and Bilateral Functional Asymmetry: Isometric upper-limb strength was "
        "assessed using a calibrated Jamar hydraulic hand dynamometer (Jamar 5030J, Sammons Preston Rolyan, Bolingbrook, "
        "IL, USA) in accordance with American Society of Hand Therapists (ASHT) guidelines (seated upright, shoulder "
        "adducted, elbow flexed at 90°, forearm neutral) across 3 maximal trials per hand with 60-second inter-trial "
        "rest [Citation Needed - 13]. To eliminate unilateral limb dominance bias and normalize for the substantial "
        "inter-individual body mass variance across the cohort (24.6–62.3 kg), Relative Handgrip Strength was computed "
        "by dividing the bilateral mean grip force by body mass ({[(Right Hand + Left Hand) / 2] / Body Mass}, kg/kg) "
        "[Citation Needed - 14]. Inter-limb force imbalance was quantified via the Bilateral Functional Asymmetry Index "
        "(%) = [|Right Hand - Left Hand| / max(Right Hand, Left Hand)] × 100 [Citation Needed - 15]."
    )
    meth_2_4_e = (
        "Vertical Jump, Sayers Peak Mechanical Power, Agility, and Flexibility: Lower-limb explosive performance was "
        "evaluated via hands-on-hips countermovement vertical jump (CMJ) height (cm) on an electronic contact mat "
        "(best of 3 trials). Lower-extremity peak mechanical power output (Watts) was derived using the Sayers equation "
        "(Peak Power [W] = 60.7 × Jump Height [cm] + 45.3 × Body Mass [kg] - 2055) [Citation Needed - 16]. Change-of-direction "
        "speed was timed in seconds (s) across a 9.14-m (5-10-5 yard) Pro-Agility shuttle course using dual-beam "
        "electronic timing gates (fastest of 2 trials) [Citation Needed - 17]. Hamstring and lumbar extensor flexibility "
        "was assessed to the nearest 0.5 cm using a standard sit-and-reach box (best of 3 trials) [Citation Needed - 18]."
    )
    add_body_para_with_markers(doc, meth_2_4_a)
    add_body_para_with_markers(doc, meth_2_4_b)
    add_body_para_with_markers(doc, meth_2_4_c)
    add_body_para_with_markers(doc, meth_2_4_d)
    add_body_para_with_markers(doc, meth_2_4_e)

    add_heading_para(doc, "2.5. Statistical Analysis", size_pt=12)
    meth_2_5_a = (
        "Distributional normality of continuous variables was verified using the Shapiro-Wilk test and visual inspection "
        "of Q-Q probability plots [Citation Needed - 19]. Descriptive statistics are presented as mean ± standard "
        "deviation (Mean ± SD), change scores (Δ = Post-test - Pre-test), and 95% confidence intervals (95% CI). "
        "Baseline differences between the Childhood (5–9 years) and Adolescence (10–15 years) subgroups were compared "
        "using independent-samples Student's t-tests. Pre-to-post adaptations across the full cohort (N = 34) were "
        "evaluated using paired-samples t-tests. Standardized effect sizes were quantified using Cohen's d based on "
        "average standard deviation (d_av: < 0.20 trivial, 0.20–0.49 small, 0.50–0.79 moderate, ≥ 0.80 large) and "
        "partial eta-squared (ηp²: 0.01 small, 0.06 medium, ≥ 0.14 large) [Citation Needed - 20]."
    )
    meth_2_5_b = (
        "To prevent group-mean masking of inter-individual responsiveness, participant changes were classified using "
        "the Smallest Worthwhile Change threshold (SWC = 0.20 × Baseline SD) into positive responders (Δ > +SWC; "
        "for timed metrics, Δ < -SWC), stable/non-changers (-SWC ≤ Δ ≤ +SWC), and negative responders [Citation Needed - 21]. "
        "Ordinary Least Squares (OLS) multiple linear regression and bivariate Pearson correlation analyses (r) were "
        "conducted to evaluate the predictive contribution of baseline morphological leverage (arm span, hand span, BMI) "
        "and biological maturation status (Mirwald PHV offset) to handgrip strength gains [Citation Needed - 22]. All "
        "analyses were executed in Python (v3.11; NumPy and SciPy libraries) with statistical significance set at "
        "two-tailed p < 0.05 (trend level: 0.05 ≤ p < 0.10)."
    )
    add_body_para_with_markers(doc, meth_2_5_a)
    add_body_para_with_markers(doc, meth_2_5_b)

    # -------------------------------------------------------------------------
    # 3. RESULTS
    # -------------------------------------------------------------------------
    add_heading_para(doc, "3. RESULTS", size_pt=13, space_before=10, space_after=8)

    add_heading_para(doc, "3.1. Baseline Characteristics and Age Group Comparison", size_pt=12)
    res_3_1 = (
        "Baseline demographic, anthropometric, somatic maturation, and neuromotor performance characteristics for the "
        "Childhood (5–9 years, n = 17), Adolescence (10–15 years, n = 17), and Full Cohort (N = 34) groups are "
        "presented in Table 1. As expected, chronological age, standing height, body mass, BMI, and Mirwald PHV "
        "maturity offset were significantly higher in the Adolescence group than in the Childhood group (all morphological "
        "comparisons: p < 0.001). At baseline, the Adolescence group exhibited significantly faster visual reaction "
        "time than the Childhood group (29.04 ± 13.76 s vs. 38.92 ± 13.93 s; t(32) = -2.08, p = 0.045, d = -0.71). "
        "Conversely, body mass-normalized relative handgrip strength (Childhood: 0.29 ± 0.19 kg/kg vs. Adolescence: "
        "0.22 ± 0.10 kg/kg; t(32) = -1.38, p = 0.177) and Pro-Agility shuttle time (15.35 ± 4.23 s vs. 13.18 ± 4.44 s; "
        "t(32) = -1.46, p = 0.155) did not differ significantly between age groups at baseline [Citation Needed - 23]."
    )
    add_body_para_with_markers(doc, res_3_1)

    tab1_headers_en = [
        "Variable",
        "Childhood Group\n(5–9 yrs, n = 17)",
        "Adolescence Group\n(10–15 yrs, n = 17)",
        "Full Cohort\n(N = 34)",
        "t",
        "p"
    ]
    tab1_rows_en = [
        ["Chronological Age (years)", "6.53 ± 0.87", "11.59 ± 1.84", "9.06 ± 2.93", "+10.24", "<0.001*"],
        ["Standing Height (cm)", "119.23 ± 10.28", "152.74 ± 12.50", "135.98 ± 20.40", "+8.53", "<0.001*"],
        ["Body Mass (kg)", "24.60 ± 8.08", "62.25 ± 27.84", "43.43 ± 27.80", "+5.35", "<0.001*"],
        ["BMI (kg/m²)", "16.93 ± 3.15", "25.92 ± 8.66", "21.42 ± 7.87", "+4.02", "<0.001*"],
        ["Mirwald PHV Offset (years)", "-5.43 ± 0.72", "-1.42 ± 1.99", "-3.42 ± 2.51", "+7.82", "<0.001*"],
        ["Visual Reaction Time (s)", "38.92 ± 13.93", "29.04 ± 13.76", "33.98 ± 14.53", "-2.08", "0.045*"],
        ["Relative Grip Strength (kg/kg)", "0.29 ± 0.19", "0.22 ± 0.10", "0.25 ± 0.16", "-1.38", "0.177"],
        ["Pro-Agility Shuttle Time (s)", "15.35 ± 4.23", "13.18 ± 4.44", "14.27 ± 4.41", "-1.46", "0.155"],
    ]
    tab1_abbr_en = (
        "Values are presented as mean ± standard deviation. BMI: Body mass index; cm: Centimeters; kg: Kilograms; "
        "m: Meters; PHV: Peak height velocity; s: Seconds; t: Independent-samples t-test statistic; *p < 0.05."
    )
    add_clean_black_table(
        doc,
        "Table 1",
        "Baseline Anthropometric, Somatic Maturation, and Neuromotor Characteristics by Age Group (N = 34)",
        tab1_headers_en,
        tab1_rows_en,
        [2.05, 1.15, 1.15, 1.10, 0.52, 0.53],
        "Abbreviations",
        tab1_abbr_en
    )

    add_heading_para(doc, "3.2. Performance Adaptations to the 12-Week Supplemental Intervention", size_pt=12)
    res_3_2_a = (
        "Pre-to-post adaptations across the full cohort (N = 34) following the 12-week supplemental motor skill "
        "intervention are presented in Table 2 and Figure 2. Visual reaction time decreased significantly from "
        "33.98 ± 14.53 s to 29.55 ± 10.39 s, representing a mean acceleration of -4.43 ± 9.51 s (95% CI: -7.76 to -1.10; "
        "-13.0%; t(33) = -2.71, p = 0.011, d = -0.35, ηp² = 0.182; Figure 2A). Parallel improvements in visual "
        "reaction time were observed in both the Childhood subgroup (Δ = -5.35 s) and the Adolescence subgroup (Δ = -3.51 s)."
    )
    res_3_2_b = (
        "Body mass-normalized relative handgrip strength increased significantly by +9.0% from 0.25 ± 0.16 kg/kg "
        "to 0.28 ± 0.16 kg/kg (Δ = +0.023 ± 0.060 kg/kg, 95% CI: +0.002 to +0.044; t(33) = 2.23, p = 0.033, d = 0.15, "
        "ηp² = 0.131; Figure 2B). Absolute right-hand grip strength increased significantly from 9.32 ± 5.96 kg to "
        "11.45 ± 7.25 kg (Δ = +2.13 ± 3.18 kg, 95% CI: +1.02 to +3.24; +22.9%, t(33) = 3.91, p < 0.001, d = 0.32), "
        "while left-hand grip strength displayed a positive upward trend from 9.41 ± 6.09 kg to 10.10 ± 7.11 kg "
        "(Δ = +0.69 ± 2.28 kg, +7.3%, t(33) = 1.76, p = 0.087, d = 0.10). The bilateral functional asymmetry index "
        "remained stable between pre-test (16.76 ± 11.94%) and post-test (17.04 ± 11.66%; Δ = +0.28 ± 14.85%, "
        "95% CI: -4.90 to +5.46; t(33) = 0.11, p = 0.914, d = 0.02). Lower-limb Sayers peak mechanical power showed "
        "a positive upward trend of +11.6% from 516.5 ± 1275.9 W to 576.4 ± 1293.9 W (Δ = +59.9 ± 190.5 W, "
        "95% CI: -6.5 to +126.3; t(33) = 1.83, p = 0.076, d = 0.05; Figure 2C). Countermovement vertical jump height "
        "(Δ = -0.42 ± 3.25 cm, p = 0.453), Pro-Agility shuttle time (Δ = -0.52 ± 3.69 s, -3.7%, p = 0.413; Figure 2D), "
        "and sit-and-reach flexibility (Δ = -0.85 ± 4.19 cm, p = 0.247) did not change significantly."
    )
    add_body_para_with_markers(doc, res_3_2_a)
    add_body_para_with_markers(doc, res_3_2_b)

    tab2_headers_en = [
        "Outcome Variable",
        "Pre-Test\n(Mean ± SD)",
        "Post-Test\n(Mean ± SD)",
        "Within-Group Δ\n(95% CI)",
        "Δ (%)",
        "t (p)",
        "Cohen's d\n(ηp²)"
    ]
    tab2_rows_en = [
        ["Visual Reaction Time (s)", "33.98 ± 14.53", "29.55 ± 10.39", "-4.43 (-7.76, -1.10)", "-13.0%", "-2.71 (0.011*)", "-0.35 (0.182)"],
        ["Relative Grip Strength (kg/kg)", "0.25 ± 0.16", "0.28 ± 0.16", "+0.02 (+0.002, +0.044)", "+9.0%", "+2.23 (0.033*)", "0.15 (0.131)"],
        ["Right-Hand Grip Strength (kg)", "9.32 ± 5.96", "11.45 ± 7.25", "+2.13 (+1.02, +3.24)", "+22.9%", "+3.91 (<0.001*)", "0.32 (0.317)"],
        ["Left-Hand Grip Strength (kg)", "9.41 ± 6.09", "10.10 ± 7.11", "+0.69 (-0.11, +1.48)", "+7.3%", "+1.76 (0.087†)", "0.10 (0.086)"],
        ["Bilateral Asymmetry Index (%)", "16.76 ± 11.94", "17.04 ± 11.66", "+0.28 (-4.90, +5.46)", "+1.7%", "+0.11 (0.914)", "0.02 (0.000)"],
        ["Sayers Peak Power (Watts)", "516.5 ± 1275.9", "576.4 ± 1293.9", "+59.9 (-6.5, +126.3)", "+11.6%", "+1.83 (0.076†)", "0.05 (0.092)"],
        ["Pro-Agility Shuttle Time (s)", "14.27 ± 4.41", "13.74 ± 3.46", "-0.52 (-1.81, +0.76)", "-3.7%", "-0.83 (0.413)", "-0.13 (0.020)"],
        ["Vertical Jump Height (cm)", "9.96 ± 5.26", "9.54 ± 4.53", "-0.42 (-1.55, +0.71)", "-4.2%", "-0.76 (0.453)", "-0.09 (0.017)"],
        ["Sit-and-Reach Flexibility (cm)", "17.28 ± 6.67", "16.43 ± 6.38", "-0.85 (-2.33, +0.62)", "-4.9%", "-1.18 (0.247)", "-0.13 (0.040)"],
    ]
    tab2_abbr_en = (
        "CI: Confidence interval; cm: Centimeters; d: Cohen's d standardized effect size; Δ: Post-test minus pre-test difference; "
        "kg: Kilograms; s: Seconds; SD: Standard deviation; t: Paired-samples t-test statistic (degrees of freedom = 33); "
        "ηp²: Partial eta-squared; *p < 0.05, †p < 0.10."
    )
    add_clean_black_table(
        doc,
        "Table 2",
        "Pre- to Post-Intervention Neuromotor and Perceptual Adaptations Across the Cohort (N = 34)",
        tab2_headers_en,
        tab2_rows_en,
        [1.55, 0.85, 0.85, 1.15, 0.55, 0.80, 0.75],
        "Abbreviations",
        tab2_abbr_en
    )

    fig2_title_en = (
        "Pre- and Post-Intervention Group Means and 95% Confidence Interval Error Bars: "
        "(A) Visual Reaction Time, (B) Relative Handgrip Strength, (C) Sayers Peak Mechanical Power, (D) Pro-Agility Shuttle Time"
    )
    fig2_abbr_en = (
        "CI: Confidence interval; kg: Kilograms; s: Seconds; W: Watts; *p < 0.05, †p < 0.10."
    )
    add_figure_block(doc, "Figure 2", fig2_title_en, FIG2_EN, "Abbreviations", fig2_abbr_en, width_in=6.3)

    add_heading_para(doc, "3.3. Smallest Worthwhile Change (SWC) Individual Responder Analysis", size_pt=12)
    res_3_3 = (
        "Individual responsiveness stratified according to the Smallest Worthwhile Change threshold (SWC = 0.20 × SD_pre) "
        "is summarized in Table 3 and Figure 3. For visual reaction time, 41.2% of participants (14/34; threshold: "
        "≥ 2.91 s reduction) were classified as positive responders, 41.2% (14/34) remained stable, and 17.6% (6/34) "
        "were negative responders. For relative handgrip strength, 29.4% of the cohort (10/34; threshold: ≥ 0.031 kg/kg "
        "gain) were positive responders, 67.6% (23/34) were stable, and only 2.9% (1/34) showed a negative response. "
        "In the Pro-Agility shuttle test, 29.4% (10/34; threshold: ≥ 0.88 s) were positive responders and 50.0% (17/34) "
        "remained stable. For Sayers peak mechanical power, 17.6% (6/34; threshold: ≥ 255.2 W) achieved positive responses "
        "while 82.4% (28/34) remained stable, with 0.0% experiencing performance decline."
    )
    add_body_para_with_markers(doc, res_3_3)

    tab3_headers_en = [
        "Performance Variable",
        "SWC Threshold\n(0.20 × SD_pre)",
        "Positive Responder\nn (%)",
        "Stable / Non-Changer\nn (%)",
        "Negative Responder\nn (%)"
    ]
    tab3_rows_en = [
        ["Visual Reaction Time", "2.91 s", "14 (41.2%)", "14 (41.2%)", "6 (17.6%)"],
        ["Relative Handgrip Strength", "0.03 kg/kg", "10 (29.4%)", "23 (67.6%)", "1 (2.9%)"],
        ["Pro-Agility Shuttle Time", "0.88 s", "10 (29.4%)", "17 (50.0%)", "7 (20.6%)"],
        ["Sayers Peak Mechanical Power", "255.2 Watts", "6 (17.6%)", "28 (82.4%)", "0 (0.0%)"],
    ]
    tab3_abbr_en = (
        "kg: Kilograms; n: Subgroup participant count; s: Seconds; SD_pre: Pre-test standard deviation; "
        "SWC: Smallest worthwhile change (0.20 × SD_pre)."
    )
    add_clean_black_table(
        doc,
        "Table 3",
        "Individual Responder Distribution Based on the Smallest Worthwhile Change (SWC = 0.20 × SD_pre) Threshold (N = 34)",
        tab3_headers_en,
        tab3_rows_en,
        [2.00, 1.15, 1.15, 1.10, 1.10],
        "Abbreviations",
        tab3_abbr_en
    )

    fig3_title_en = (
        "Percentage Distribution of Positive Responders, Stable Participants, and Negative Responders Based on the "
        "Smallest Worthwhile Change (SWC = 0.20 × SD_pre) Threshold (N = 34)"
    )
    fig3_abbr_en = (
        "SD_pre: Pre-test standard deviation; SWC: Smallest worthwhile change."
    )
    add_figure_block(doc, "Figure 3", fig3_title_en, FIG3_EN, "Abbreviations", fig3_abbr_en, width_in=6.3)

    add_heading_para(doc, "3.4. Morphological and Maturational Regression Model of Strength Gains", size_pt=12)
    res_3_4_a = (
        "The OLS multiple linear regression model evaluating anthropometric leverage and somatic maturation predictors "
        "of handgrip strength gains is presented in Table 4, with bivariate linear regression plots illustrated in "
        "Figure 4. The overall regression model was statistically significant and accounted for 28.0% of the total "
        "variance in strength gains (R² = 0.280, Adjusted R² = 0.180, F(4, 29) = 2.813, p = 0.044). Within the "
        "multivariable model, baseline arm span exhibited a positive standardized weight (B = +0.09, β = +0.77, "
        "t = 1.29, p = 0.207), whereas baseline BMI showed a marginal inverse trend (B = -0.15, β = -0.50, t = -1.85, "
        "p = 0.075). Notably, Mirwald PHV maturity offset (B = -0.04, β = -0.04, t = -0.08, p = 0.939) and hand span "
        "(B = -0.04, β = -0.04, t = -0.10, p = 0.923) had no independent effect in the multivariable model."
    )
    res_3_4_b = (
        "In bivariate correlation analyses, baseline hand span correlated positively and significantly with dominant "
        "(right) handgrip strength gain (r = +0.47, r² = 0.219, p = 0.005; Figure 4A), and baseline arm span "
        "correlated positively and significantly with overall handgrip strength gain (r = +0.37, r² = 0.134, p = 0.033; "
        "Figure 4B) [Citation Needed - 24]."
    )
    add_body_para_with_markers(doc, res_3_4_a)
    add_body_para_with_markers(doc, res_3_4_b)

    tab4_headers_en = [
        "Predictor Variable",
        "B Coefficient",
        "Standard Error",
        "Standardized β",
        "t",
        "p"
    ]
    tab4_rows_en = [
        ["Intercept", "-6.62", "9.26", "—", "-0.71", "0.481"],
        ["Baseline Arm Span (cm)", "+0.09", "0.07", "+0.77", "+1.29", "0.207"],
        ["Baseline BMI (kg/m²)", "-0.15", "0.08", "-0.50", "-1.85", "0.075†"],
        ["Baseline Hand Span (cm)", "-0.04", "0.40", "-0.04", "-0.10", "0.923"],
        ["Mirwald PHV Maturity Offset (years)", "-0.04", "0.52", "-0.04", "-0.08", "0.939"],
    ]
    tab4_abbr_en = (
        "Model summary: R² = 0.280, Adjusted R² = 0.180, F(4, 29) = 2.813, p = 0.044*. B: Unstandardized regression "
        "coefficient; β: Standardized regression coefficient; BMI: Body mass index; cm: Centimeters; PHV: Peak height "
        "velocity; †p < 0.10."
    )
    add_clean_black_table(
        doc,
        "Table 4",
        "Ordinary Least Squares (OLS) Multiple Linear Regression Model Predicting Handgrip Strength Gains (N = 34)",
        tab4_headers_en,
        tab4_rows_en,
        [2.20, 0.90, 0.95, 0.85, 0.80, 0.80],
        "Abbreviations",
        tab4_abbr_en
    )

    fig4_title_en = (
        "Linear Regression Relationships and 95% Confidence Bands Between Morphological Leverage Traits and Strength Gains: "
        "(A) Baseline Hand Span vs. Dominant Handgrip Gain, (B) Baseline Arm Span vs. Overall Handgrip Gain"
    )
    fig4_abbr_en = (
        "cm: Centimeters; kg: Kilograms; r: Pearson correlation coefficient; r²: Coefficient of determination; *p < 0.05, **p < 0.01."
    )
    add_figure_block(doc, "Figure 4", fig4_title_en, FIG4_EN, "Abbreviations", fig4_abbr_en, width_in=6.3)

    doc.save(EN_DOCX)
    scrub_docx_metadata(EN_DOCX)
    print("Saved Q1_Master_Manuscript_EN.docx")

    md_text_en = f"""# {title_en}

## ABSTRACT

**Objective:** {ab_obj}

**Methods:** {ab_meth}

**Results:** {ab_res}

**Conclusion:** {ab_conc}

**Keywords:** {ab_kw}

## 2. METHODS

### 2.1. Study Design

{meth_2_1}

**Figure 1.** {fig1_title_en}
![Figure 1](Figure_1_Study_Design_EN.png)
**Abbreviations:** {fig1_abbr_en}

### 2.2. Participants and Statistical Power Statement

{meth_2_2_a}

{meth_2_2_b}

### 2.3. Supplemental Micro-Dosed Motor Skill Training Protocol

{meth_2_3_a}

{meth_2_3_b}

### 2.4. Data Collection Instruments and Testing Conditions

{meth_2_4_a}

{meth_2_4_b}

{meth_2_4_c}

{meth_2_4_d}

{meth_2_4_e}

### 2.5. Statistical Analysis

{meth_2_5_a}

{meth_2_5_b}

## 3. RESULTS

### 3.1. Baseline Characteristics and Age Group Comparison

{res_3_1}

**Table 1.** Baseline Anthropometric, Somatic Maturation, and Neuromotor Characteristics by Age Group (N = 34)

| Variable | Childhood Group (5–9 yrs, n = 17) | Adolescence Group (10–15 yrs, n = 17) | Full Cohort (N = 34) | t | p |
| :--- | :---: | :---: | :---: | :---: | :---: |
| Chronological Age (years) | 6.53 ± 0.87 | 11.59 ± 1.84 | 9.06 ± 2.93 | +10.24 | <0.001* |
| Standing Height (cm) | 119.23 ± 10.28 | 152.74 ± 12.50 | 135.98 ± 20.40 | +8.53 | <0.001* |
| Body Mass (kg) | 24.60 ± 8.08 | 62.25 ± 27.84 | 43.43 ± 27.80 | "+5.35" | <0.001* |
| BMI (kg/m²) | 16.93 ± 3.15 | 25.92 ± 8.66 | 21.42 ± 7.87 | +4.02 | <0.001* |
| Mirwald PHV Offset (years) | -5.43 ± 0.72 | -1.42 ± 1.99 | -3.42 ± 2.51 | +7.82 | <0.001* |
| Visual Reaction Time (s) | 38.92 ± 13.93 | 29.04 ± 13.76 | 33.98 ± 14.53 | -2.08 | 0.045* |
| Relative Grip Strength (kg/kg) | 0.29 ± 0.19 | 0.22 ± 0.10 | 0.25 ± 0.16 | -1.38 | 0.177 |
| Pro-Agility Shuttle Time (s) | 15.35 ± 4.23 | 13.18 ± 4.44 | 14.27 ± 4.41 | -1.46 | 0.155 |

**Abbreviations:** {tab1_abbr_en}

### 3.2. Performance Adaptations to the 12-Week Supplemental Intervention

{res_3_2_a}

{res_3_2_b}

**Table 2.** Pre- to Post-Intervention Neuromotor and Perceptual Adaptations Across the Cohort (N = 34)

| Outcome Variable | Pre-Test (Mean ± SD) | Post-Test (Mean ± SD) | Within-Group Δ (95% CI) | Δ (%) | t (p) | Cohen's d (ηp²) |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: |
| Visual Reaction Time (s) | 33.98 ± 14.53 | 29.55 ± 10.39 | -4.43 (-7.76, -1.10) | -13.0% | -2.71 (0.011*) | -0.35 (0.182) |
| Relative Grip Strength (kg/kg) | 0.25 ± 0.16 | 0.28 ± 0.16 | +0.02 (+0.002, +0.044) | +9.0% | +2.23 (0.033*) | 0.15 (0.131) |
| Right-Hand Grip Strength (kg) | 9.32 ± 5.96 | 11.45 ± 7.25 | +2.13 (+1.02, +3.24) | +22.9% | +3.91 (<0.001*) | 0.32 (0.317) |
| Left-Hand Grip Strength (kg) | 9.41 ± 6.09 | 10.10 ± 7.11 | +0.69 (-0.11, +1.48) | +7.3% | +1.76 (0.087†) | 0.10 (0.086) |
| Bilateral Asymmetry Index (%) | 16.76 ± 11.94 | 17.04 ± 11.66 | +0.28 (-4.90, +5.46) | +1.7% | +0.11 (0.914) | 0.02 (0.000) |
| Sayers Peak Power (Watts) | 516.5 ± 1275.9 | 576.4 ± 1293.9 | +59.9 (-6.5, +126.3) | +11.6% | +1.83 (0.076†) | 0.05 (0.092) |
| Pro-Agility Shuttle Time (s) | 14.27 ± 4.41 | 13.74 ± 3.46 | "-0.52 (-1.81, +0.76)" | -3.7% | -0.83 (0.413) | -0.13 (0.020) |
| Vertical Jump Height (cm) | 9.96 ± 5.26 | 9.54 ± 4.53 | -0.42 (-1.55, +0.71) | -4.2% | -0.76 (0.453) | -0.09 (0.017) |
| Sit-and-Reach Flexibility (cm) | 17.28 ± 6.67 | 16.43 ± 6.38 | -0.85 (-2.33, +0.62) | -4.9% | -1.18 (0.247) | -0.13 (0.040) |

**Abbreviations:** {tab2_abbr_en}

**Figure 2.** {fig2_title_en}
![Figure 2](Figure_2_Performance_Adaptations_EN.png)
**Abbreviations:** {fig2_abbr_en}

### 3.3. Smallest Worthwhile Change (SWC) Individual Responder Analysis

{res_3_3}

**Table 3.** Individual Responder Distribution Based on the Smallest Worthwhile Change (SWC = 0.20 × SD_pre) Threshold (N = 34)

| Performance Variable | SWC Threshold (0.20 × SD_pre) | Positive Responder n (%) | Stable / Non-Changer n (%) | Negative Responder n (%) |
| :--- | :---: | :---: | :---: | :---: |
| Visual Reaction Time | 2.91 s | 14 (41.2%) | 14 (41.2%) | 6 (17.6%) |
| Relative Handgrip Strength | 0.03 kg/kg | 10 (29.4%) | 23 (67.6%) | 1 (2.9%) |
| Pro-Agility Shuttle Time | 0.88 s | 10 (29.4%) | 17 (50.0%) | 7 (20.6%) |
| Sayers Peak Mechanical Power | 255.2 Watts | 6 (17.6%) | 28 (82.4%) | 0 (0.0%) |

**Abbreviations:** {tab3_abbr_en}

**Figure 3.** {fig3_title_en}
![Figure 3](Figure_3_SWC_Responder_Analysis_EN.png)
**Abbreviations:** {fig3_abbr_en}

### 3.4. Morphological and Maturational Regression Model of Strength Gains

{res_3_4_a}

{res_3_4_b}

**Table 4.** Ordinary Least Squares (OLS) Multiple Linear Regression Model Predicting Handgrip Strength Gains (N = 34)

| Predictor Variable | B Coefficient | Standard Error | Standardized β | t | p |
| :--- | :---: | :---: | :---: | :---: | :---: |
| Intercept | -6.62 | 9.26 | — | -0.71 | 0.481 |
| Baseline Arm Span (cm) | +0.09 | 0.07 | +0.77 | +1.29 | 0.207 |
| Baseline BMI (kg/m²) | -0.15 | 0.08 | -0.50 | -1.85 | 0.075† |
| Baseline Hand Span (cm) | -0.04 | 0.40 | -0.04 | -0.10 | 0.923 |
| Mirwald PHV Maturity Offset (years) | -0.04 | 0.52 | -0.04 | -0.08 | 0.939 |

**Abbreviations:** {tab4_abbr_en}

**Figure 4.** {fig4_title_en}
![Figure 4](Figure_4_Regression_Relationships_EN.png)
**Abbreviations:** {fig4_abbr_en}
"""
    with open(EN_MD, "w", encoding="utf-8") as f:
        f.write(clean_str(md_text_en))
    print("Saved Q1_Master_Manuscript_EN.md")


# =============================================================================
# 3. AYRI ATIF REHBERİ VE KAYNAKÇA RAPORU (TR & EN)
# =============================================================================
def build_citation_reports():
    rows_tr = [
        [
            "[Atıf Gerekli - 1]",
            "2.1. STROBE Raporlama Kılavuzu",
            "(von Elm ve ark., 2007)",
            "Gözlemsel ve retrospektif kohort çalışmalarının standart raporlama kılavuzu (STROBE Statement)."
        ],
        [
            "[Atıf Gerekli - 2]",
            "2.1. CERT Egzersiz Raporlama Şablonu",
            "(Slade ve ark., 2016)",
            "Egzersiz ve motor beceri müdahalelerinin (istasyon, süre, sıklık, bireysel ilerleme) standart raporlama şablonu."
        ],
        [
            "[Atıf Gerekli - 3]",
            "2.1. Nöromotor Öğrenme Etkisi ve Bazal Plato (≥6 Ay)",
            "(Škarabot ve ark., 2021; Pan ve ark., 2017)",
            "Erken dönem nöral adaptasyonları ve test alışma etkisini dışlamak için en az 6 ay düzenli eğitim geçmişi kriteri."
        ],
        [
            "[Atıf Gerekli - 4]",
            "2.2. DSM-5 Otizm Spektrum Bozukluğu Tanı Ölçütleri",
            "(American Psychiatric Association, 2013)",
            "Katılımcıların resmi klinik Otizm Spektrum Bozukluğu (OSB) tanı sınıflandırma referansı (DSM-5)."
        ],
        [
            "[Atıf Gerekli - 5]",
            "2.2. Helsinki Bildirgesi Araştırma Etiği İlkeleri",
            "(World Medical Association, 2013)",
            "Pediatrik insan araştırmalarında etik kurul onayı, ebeveyn yazılı onamı ve çocuk sözlü onayı standardı."
        ],
        [
            "[Atıf Gerekli - 6]",
            "2.2. Duyarlılık Güç Analizi (G*Power ve dz ≈ 0.49)",
            "(Faul ve ark., 2007; Cohen, 1988)",
            "N = 34 örneklem ile bağımlı örneklemler t-testinde (α = 0.05, %80 güç) saptanabilir minimum etki büyüklüğü hesabı."
        ],
        [
            "[Atıf Gerekli - 7]",
            "2.3. Otizmde Uyarlanmış Hareket ve Beceri Eğitimi",
            "(Healy ve ark., 2018; Ruggeri ve ark., 2020)",
            "Otizmli çocuk ve ergenlerde yapılandırılmış motor beceri antrenmanlarının nöromotor gelişim üzerindeki etkileri."
        ],
        [
            "[Atıf Gerekli - 8]",
            "2.3. Bireysel Meydan Okuma Noktası Çerçevesi",
            "(Guadagnoli ve Lee, 2004)",
            "Motor öğrenmede görev zorluğunun bireyin anlık kapasitesine göre ayarlanmasını açıklayan kuramsal çerçeve (Challenge Point Framework)."
        ],
        [
            "[Atıf Gerekli - 9]",
            "2.4. Sirkadiyen Ritim ve Çevresel Ölçüm Kontrolü",
            "(Atkinson ve Reilly, 1996; Chtourou ve Souissi, 2012)",
            "Reaksiyon süresi, kas kuvveti ve patlayıcı güç testlerinde gün içi sirkadiyen değişim ve ortam sıcaklığı kontrolü."
        ],
        [
            "[Atıf Gerekli - 10]",
            "2.4. ISAK Antropometri Protokolü ve VKİ Hesabı",
            "(World Health Organization, 2007; Nuttall, 2015)",
            "Boy, oturma boyu, vücut ağırlığı, kulaç ve el karış uzunluğu ölçüm standartları ile DSÖ çocukluk çağı VKİ (kg/m²) büyüme referansı."
        ],
        [
            "[Atıf Gerekli - 11]",
            "2.4. Mirwald Zirve Boy Hızı (PHV) Olgunlaşma Denklemi",
            "(Mirwald ve ark., 2002; Malina ve ark., 2015)",
            "Erkek çocuk ve ergenlerde oturma boyu, bacak uzunluğu, yaş, boy ve ağırlıktan biyolojik olgunlaşma offseti (yıl) hesaplama denklemi."
        ],
        [
            "[Atıf Gerekli - 12]",
            "2.4. Görsel Reaksiyon Süresi Ölçüm Protokolü",
            "(Magill ve Anderson, 2017; Laurenti ve ark., 2026)",
            "LED görsel uyaranlı basit reaksiyon süresi ölçümü ve otizmde algısal-motor tepki gecikmesi değerlendirmesi."
        ],
        [
            "[Atıf Gerekli - 13]",
            "2.4. Jamar Dinamometresi ve ASHT Ölçüm Standartları",
            "(Fess, 1992; Roberts ve ark., 2011)",
            "Amerikan El Terapistleri Derneği (ASHT) standart oturur pozisyonda (dirsek 90° fleksiyonda) el kavrama kuvveti ölçüm kılavuzu."
        ],
        [
            "[Atıf Gerekli - 14]",
            "2.4. Vücut Ağırlığına Normalize Rölatif Kuvvet Formülü",
            "(Jaric, 2002; Wind ve ark., 2010)",
            "Farklı vücut kütlelerine sahip pediatrik gruplarda mutlak kuvvet yerine rölatif kavrama kuvveti ({[(Sağ+Sol)/2]/Kilo}, kg/kg) kullanımı."
        ],
        [
            "[Atıf Gerekli - 15]",
            "2.4. Bilateral Fonksiyonel Asimetri İndeksi Formülü",
            "(Impellizzeri ve ark., 2007; Bishop ve ark., 2018)",
            "İki ekstremite arasındaki kuvvet dengesizliğini yüzdesel olarak hesaplayan ([|Sağ - Sol| / maks(Sağ, Sol)] × 100) standart denklem."
        ],
        [
            "[Atıf Gerekli - 16]",
            "2.4. Aktif Dikey Sıçrama ve Zirve Güç Kestirim Denklemleri",
            "(Duncan ve ark., 2008; Markovic ve ark., 2004)",
            "Karşıt hareketli dikey sıçrama yüksekliği ve vücut ağırlığından zirve mekanik güç (Sayers denklemi) kestirimi ve geçerliği."
        ],
        [
            "[Atıf Gerekli - 17]",
            "2.4. Pro-Agility (5-10-5) Çeviklik Testi Protokolü",
            "(Harman ve Garhammer, 2008; Stewart ve ark., 2014)",
            "9.14 metrelik (5-10-5 yarda) mekik parkurunda yön değiştirme hızı ve çeviklik ölçüm protokolü ve güvenirliği."
        ],
        [
            "[Atıf Gerekli - 18]",
            "2.4. Otur-Eriş (Sit-and-Reach) Esneklik Testi",
            "(Wells ve Dillon, 1952; Castro-Piñero ve ark., 2009)",
            "Çocuk ve ergenlerde hamstring ve lomber ekstansör esnekliğinin standart otur-eriş sehpası ile ölçümü."
        ],
        [
            "[Atıf Gerekli - 19]",
            "2.5. Shapiro-Wilk Normallik Testi",
            "(Shapiro ve Wilk, 1965; Razali ve Wah, 2011)",
            "N < 50 örneklem büyüklüklerinde parametrik test varsayımları için Shapiro-Wilk normallik testinin üstünlüğü."
        ],
        [
            "[Atıf Gerekli - 20]",
            "2.5. Etki Büyüklüğü (Cohen's d_av ve ηp²) Sınır Değerleri",
            "(Cohen, 1988; Lakens, 2013)",
            "Bağımlı örneklemlerde standartlaştırılmış etki büyüklüğü (Cohen's d_av) ve kısmi eta-kare (ηp²) hesaplama ve yorumlama ölçütleri."
        ],
        [
            "[Atıf Gerekli - 21]",
            "2.5. En Küçük Anlamlı Değişim (SWC = 0.20 × SS_ön)",
            "(Hopkins, 2004; Swinton ve ark., 2018)",
            "Spor bilimlerinde bireysel yanıtlayıcı (responder) sınıflandırması için SWC = 0.20 × Başlangıç SS eşik değeri yöntemi."
        ],
        [
            "[Atıf Gerekli - 22]",
            "2.5. Çoklu Doğrusal Regresyon (OLS) ve Korelasyon Analizi",
            "(Aiken ve West, 1991; Tabachnick ve Fidell, 2019)",
            "Sıradan En Küçük Kareler (OLS) çoklu doğrusal regresyon modellemesi, R², F testi ve standartlaştırılmış β katsayıları analizi."
        ],
        [
            "[Atıf Gerekli - 23]",
            "3.1. Bulgular: Yaş Grupları Reaksiyon Hızı ve Rölatif Kuvvet",
            "(Luna ve ark., 2004; Lloyd ve Oliver, 2012)",
            "Çocukluktan ergenliğe geçişte bilişsel işlemleme ve motor tepki hızının olgunlaşması ile rölatif kuvvetin gelişimsel karşılaştırması."
        ],
        [
            "[Atıf Gerekli - 24]",
            "3.4. Bulgular: El Karış ve Kulaç Uzunluğu Biyomekanik Kaldıraç Etkisi",
            "(Nicolay ve Walker, 2005; Hager-Ross ve Rösblad, 2002)",
            "El karış uzunluğu ve üst ekstremite antropometrik kaldıraç kollarının Jamar dinamometresi kavrama kuvveti gelişimi üzerindeki biyomekanik etkisi."
        ],
    ]

    rows_en = [
        [
            "[Citation Needed - 1]",
            "2.1. STROBE Reporting Guideline",
            "(von Elm et al., 2007)",
            "Standard reporting guideline for observational and retrospective cohort studies (STROBE Statement)."
        ],
        [
            "[Citation Needed - 2]",
            "2.1. CERT Exercise Reporting Template",
            "(Slade et al., 2016)",
            "Standardized reporting template for exercise and motor skill interventions (stations, dosage, progression)."
        ],
        [
            "[Citation Needed - 3]",
            "2.1. Neuromotor Familiarization & Baseline Plateau (≥6 Months)",
            "(Škarabot et al., 2021; Pan et al., 2017)",
            "Requirement of ≥6 months prior structured training to account for early neural adaptations and test familiarization."
        ],
        [
            "[Citation Needed - 4]",
            "2.2. DSM-5 Diagnostic Criteria for ASD",
            "(American Psychiatric Association, 2013)",
            "Standard clinical diagnostic classification for Autism Spectrum Disorder (DSM-5)."
        ],
        [
            "[Citation Needed - 5]",
            "2.2. Declaration of Helsinki Ethical Principles",
            "(World Medical Association, 2013)",
            "Ethical principles for medical and exercise research involving pediatric human subjects, guardian consent, and child assent."
        ],
        [
            "[Citation Needed - 6]",
            "2.2. Sensitivity Power Analysis (G*Power, dz ≈ 0.49)",
            "(Faul et al., 2007; Cohen, 1988)",
            "Calculation of minimum detectable standardized effect size for paired-samples t-tests (N = 34, α = 0.05, 80% power)."
        ],
        [
            "[Citation Needed - 7]",
            "2.3. Adapted Movement & Motor Skill Training in ASD",
            "(Healy et al., 2018; Ruggeri et al., 2020)",
            "Systematic evidence supporting structured motor skill interventions for improving motor competence in youth with ASD."
        ],
        [
            "[Citation Needed - 8]",
            "2.3. Challenge Point Framework",
            "(Guadagnoli & Lee, 2004)",
            "Theoretical framework regulating functional task difficulty according to individual learner skill level in motor learning."
        ],
        [
            "[Citation Needed - 9]",
            "2.4. Circadian Rhythm & Environmental Standardization",
            "(Atkinson & Reilly, 1996; Chtourou & Souissi, 2012)",
            "Control of diurnal variation (time-of-day ±30 min) and indoor temperature during reaction time and neuromuscular testing."
        ],
        [
            "[Citation Needed - 10]",
            "2.4. Anthropometric Standards & BMI Growth Reference",
            "(World Health Organization, 2007; Nuttall, 2015)",
            "Anthropometric measurement standards for height, weight, arm span, and hand span, and WHO BMI (kg/m²) growth reference for youth."
        ],
        [
            "[Citation Needed - 11]",
            "2.4. Mirwald Peak Height Velocity (PHV) Maturity Offset",
            "(Mirwald et al., 2002; Malina et al., 2015)",
            "Non-invasive sex-specific regression equation estimating somatic biological maturity offset (years from PHV) in boys."
        ],
        [
            "[Citation Needed - 12]",
            "2.4. Visual Reaction Time Testing Protocol",
            "(Magill & Anderson, 2017; Laurenti et al., 2026)",
            "Electronic LED simple visual response latency testing and perceptual-motor processing assessment in ASD."
        ],
        [
            "[Citation Needed - 13]",
            "2.4. Jamar Dynamometry & ASHT Testing Guidelines",
            "(Fess, 1992; Roberts et al., 2011)",
            "American Society of Hand Therapists (ASHT) standardized seated posture (elbow flexed at 90°) for isometric handgrip testing."
        ],
        [
            "[Citation Needed - 14]",
            "2.4. Body Mass-Normalized Relative Grip Strength",
            "(Jaric, 2002; Wind et al., 2010)",
            "Normalization of bilateral grip force to body mass ({[(Right+Left)/2]/Body Mass}, kg/kg) to control for body size heterogeneity."
        ],
        [
            "[Citation Needed - 15]",
            "2.4. Bilateral Functional Asymmetry Index Formula",
            "(Impellizzeri et al., 2007; Bishop et al., 2018)",
            "Validated equation quantifying inter-limb strength asymmetry ([|Right - Left| / max(Right, Left)] × 100)."
        ],
        [
            "[Citation Needed - 16]",
            "2.4. Countermovement Jump & Peak Power Prediction Equations",
            "(Duncan et al., 2008; Markovic et al., 2004)",
            "Cross-validation of peak power prediction equations (including Sayers equation) from vertical jump height and body mass."
        ],
        [
            "[Citation Needed - 17]",
            "2.4. Pro-Agility (5-10-5) Shuttle Test Protocol",
            "(Harman & Garhammer, 2008; Stewart et al., 2014)",
            "Standardized 9.14-m (5-10-5 yard) change-of-direction speed and agility assessment using electronic timing gates."
        ],
        [
            "[Citation Needed - 18]",
            "2.4. Sit-and-Reach Flexibility Assessment",
            "(Wells & Dillon, 1952; Castro-Piñero et al., 2009)",
            "Criterion validity and protocol for sit-and-reach hamstring and lumbar flexibility testing in youth."
        ],
        [
            "[Citation Needed - 19]",
            "2.5. Shapiro-Wilk Normality Test",
            "(Shapiro & Wilk, 1965; Razali & Wah, 2011)",
            "Statistical power and appropriateness of the Shapiro-Wilk test for verifying normality in sample sizes N < 50."
        ],
        [
            "[Citation Needed - 20]",
            "2.5. Effect Size Conventions (Cohen's d_av & ηp²)",
            "(Cohen, 1988; Lakens, 2013)",
            "Calculation and interpretation thresholds for standardized mean differences (Cohen's d_av) and partial eta-squared (ηp²)."
        ],
        [
            "[Citation Needed - 21]",
            "2.5. Smallest Worthwhile Change (SWC = 0.20 × SD_pre)",
            "(Hopkins, 2004; Swinton et al., 2018)",
            "Magnitude-based individual responder classification threshold (SWC = 0.20 × Baseline SD) in exercise science."
        ],
        [
            "[Citation Needed - 22]",
            "2.5. Ordinary Least Squares (OLS) Multiple Regression",
            "(Aiken & West, 1991; Tabachnick & Fidell, 2019)",
            "Multivariable linear regression modeling (R², Adjusted R², F-test, standardized β weights) and bivariate correlation."
        ],
        [
            "[Citation Needed - 23]",
            "3.1. Results: Developmental Reaction Speed & Relative Strength",
            "(Luna et al., 2004; Lloyd & Oliver, 2012)",
            "Maturation of cognitive processing and response speed from childhood to adolescence and relative strength scaling in youth."
        ],
        [
            "[Citation Needed - 24]",
            "3.4. Results: Biomechanical Leverage of Hand & Arm Span",
            "(Nicolay & Walker, 2005; Hager-Ross & Rösblad, 2002)",
            "Biomechanical influence of hand span and upper-extremity anthropometry on Jamar dynamometer grip force generation."
        ],
    ]

    references_apa7 = [
        "Aiken, L. S., & West, S. G. (1991). Multiple regression: Testing and interpreting interactions. Sage Publications.",
        "American Psychiatric Association. (2013). Diagnostic and statistical manual of mental disorders (5th ed.). American Psychiatric Publishing. https://doi.org/10.1176/appi.books.9780890425596",
        "Atkinson, G., & Reilly, T. (1996). Circadian variation in sports performance. Sports Medicine, 21(4), 292–312. https://doi.org/10.2165/00007256-199621040-00005",
        "Bishop, C., Read, P., Lake, J., Chavda, S., & Turner, A. (2018). Interlimb asymmetries: Understanding how to calculate differences from bilateral and unilateral tests. Strength and Conditioning Journal, 40(4), 1–6. https://doi.org/10.1519/SSC.0000000000000371",
        "Castro-Piñero, J., Chillón, P., Ortega, F. B., Montesinos, J. L., Sjöström, M., & Ruiz, J. R. (2009). Criterion-related validity of sit-and-reach and modified sit-and-reach test for estimating hamstring flexibility in children and adolescents aged 6–17 years. International Journal of Sports Medicine, 30(9), 658–662. https://doi.org/10.1055/s-0029-1224175",
        "Chtourou, H., & Souissi, N. (2012). The effect of training at a specific time of day: A review. Journal of Strength and Conditioning Research, 26(7), 1984–2005. https://doi.org/10.1519/JSC.0b013e31825770a7",
        "Cohen, J. (1988). Statistical power analysis for the behavioral sciences (2nd ed.). Lawrence Erlbaum Associates.",
        "Faul, F., Erdfelder, E., Lang, A.-G., & Buchner, A. (2007). G*Power 3: A flexible statistical power analysis program for the social, behavioral, and biomedical sciences. Behavior Research Methods, 39(2), 175–191. https://doi.org/10.3758/BF03193146",
        "Fess, E. E. (1992). Grip strength. In J. S. Casanova (Ed.), Clinical assessment recommendations (2nd ed., pp. 41–45). American Society of Hand Therapists.",
        "Guadagnoli, M. A., & Lee, T. D. (2004). Challenge point: A framework for conceptualizing the effects of various practice conditions in motor learning. Journal of Motor Behavior, 36(2), 212–224. https://doi.org/10.3200/JMBR.36.2.212-224",
        "Hager-Ross, C., & Rösblad, B. (2002). Norms for grip strength in children aged 4–16 years. Acta Paediatrica, 91(6), 617–625. https://doi.org/10.1111/j.1651-2227.2002.tb03290.x",
        "Harman, E., & Garhammer, J. (2008). Administration, scoring, and interpretation of selected tests. In T. R. Baechle & R. W. Earle (Eds.), Essentials of strength training and conditioning (3rd ed., pp. 249–292). Human Kinetics.",
        "Healy, S., Nacario, A., Braithwaite, R. E., & Hopper, C. (2018). The effect of physical activity interventions on youth with autism spectrum disorder: A meta-analysis. Autism Research, 11(6), 818–833. https://doi.org/10.1002/aur.1955",
        "Hopkins, W. G. (2004). How to interpret changes in an athletic performance test. Sportscience, 8, 1–7.",
        "Impellizzeri, F. M., Rampinini, E., Maffiuletti, N., & Marcora, S. M. (2007). A vertical jump force test for assessing bilateral strength asymmetry in athletes. Medicine & Science in Sports & Exercise, 39(11), 2044–2050. https://doi.org/10.1249/mss.0b013e31814fb55c",
        "Jaric, S. (2002). Muscle strength testing: Use of normalization for body size. Sports Medicine, 32(10), 615–631. https://doi.org/10.2165/00007256-200232100-00002",
        "Lakens, D. (2013). Calculating and reporting effect sizes to facilitate cumulative science: A practical primer for t-tests and ANOVAs. Frontiers in Psychology, 4, 863. https://doi.org/10.3389/fpsyg.2013.00863",
        "Laurenti, F., Presta, V., Compiani, M., Zobbi, G., Ilari, B., Picchi, M. P., Maré, E., Severini, F., Guarnieri, A., Mazzei, S., di Martino, O., Pozzi, G., Condello, G., & Gobbi, G. (2026). The effects of a sport-based training program on reaction time and fine motor coordination in children with autism spectrum disorder: A pilot study. Sports, 14(2), Article 80. https://doi.org/10.3390/sports14020080",
        "Lloyd, R. S., & Oliver, J. L. (2012). The youth physical development model: A new approach to long-term athletic development. Strength and Conditioning Journal, 34(3), 61–72. https://doi.org/10.1519/SSC.0b013e31825760ea",
        "Luna, B., Garver, K. E., Urban, T. A., Lazar, N. A., & Sweeney, J. A. (2004). Maturation of cognitive processes from late childhood to adulthood. Child Development, 75(5), 1357–1372. https://doi.org/10.1111/j.1467-8624.2004.00745.x",
        "Magill, R. A., & Anderson, D. I. (2017). Motor learning and control: Concepts and applications (11th ed.). McGraw-Hill Education.",
        "Malina, R. M., Rogol, A. D., Cumming, S. P., Coelho e Silva, M. J., & Figueredo, A. J. (2015). Biological maturation of youth athletes: Assessment and implications. British Journal of Sports Medicine, 49(13), 852–859. https://doi.org/10.1136/bjsports-2015-094623",
        "Markovic, G., Dizdar, D., Jukic, I., & Cardinale, M. (2004). Reliability and factorial validity of squat and countermovement jump tests. Journal of Strength and Conditioning Research, 18(3), 551–555. https://doi.org/10.1519/1533-4287(2004)18<551:RAFVOS>2.0.CO;2",
        "Mirwald, R. L., Baxter-Jones, A. D. G., Bailey, D. A., & Beunen, G. P. (2002). An assessment of maturity from anthropometric measurements. Medicine & Science in Sports & Exercise, 34(4), 689–694. https://doi.org/10.1097/00005768-200204000-00020",
        "Nicolay, C. W., & Walker, A. L. (2005). Grip strength and endurance: Influences of anthropometric variation, hand dominance, and gender. International Journal of Industrial Ergonomics, 35(7), 605–618. https://doi.org/10.1016/j.ergon.2005.01.007",
        "Nuttall, F. Q. (2015). Body Mass Index: Obesity, BMI, and health: A critical review. Nutrition Today, 50(3), 117–128. https://doi.org/10.1097/NT.0000000000000092",
        "Pan, C.-Y., Chu, C.-H., Tsai, C.-L., Sung, M.-C., Huang, C.-Y., & Ma, W.-Y. (2017). The impacts of physical activity intervention on physical and cognitive outcomes in children with autism spectrum disorder. Autism, 21(2), 190–202. https://doi.org/10.1177/1362361316633562",
        "Razali, N. M., & Wah, Y. B. (2011). Power comparisons of Shapiro-Wilk, Kolmogorov-Smirnov, Lilliefors and Anderson-Darling tests. Journal of Statistical Modeling and Analytics, 2(1), 21–33.",
        "Roberts, H. C., Denison, H. J., Martin, H. J., Patel, H. P., Syddall, H., Cooper, C., & Sayer, A. A. (2011). A review of the measurement of grip strength in clinical and epidemiological studies: Towards a standardised approach. Age and Ageing, 40(4), 423–429. https://doi.org/10.1093/ageing/afr051",
        "Ruggeri, A., Dancel, A., Johnson, R., & Sargent, B. (2020). The effect of motor and physical activity intervention on motor outcomes of children with autism spectrum disorder: A systematic review. Autism, 24(3), 544–568. https://doi.org/10.1177/1362361319885215",
        "Škarabot, J., Brownstein, C. G., Casolo, A., Del Vecchio, A., & Ansdell, P. (2021). The knowns and unknowns of neural adaptations to resistance training. European Journal of Applied Physiology, 121(3), 675–685. https://doi.org/10.1007/s00421-020-04567-3",
        "Duncan, M. J., Lyons, M., & Nevill, A. M. (2008). Evaluation of peak power prediction equations in male basketball players. European Journal of Applied Physiology, 104(4), 629–634. https://doi.org/10.1007/s00421-008-0814-1",
        "Shapiro, S. S., & Wilk, M. B. (1965). An analysis of variance test for normality (complete samples). Biometrika, 52(3–4), 591–611. https://doi.org/10.1093/biomet/52.3-4.591",
        "Slade, S. C., Dionne, C. E., Underwood, M., & Buchbinder, R. (2016). Consensus on Exercise Reporting Template (CERT): Explanation and elaboration statement. British Journal of Sports Medicine, 50(23), 1428–1437. https://doi.org/10.1136/bjsports-2016-096651",
        "Stewart, P. F., Turner, A. N., & Miller, S. C. (2014). Reliability, factorial validity, and interrelationships of five commonly used change of direction speed tests. Scandinavian Journal of Medicine & Science in Sports, 24(3), 500–506. https://doi.org/10.1111/sms.12019",
        "Swinton, P. A., Hemingway, B. S., Saunders, B., Gualano, B., & Dolan, E. (2018). A statistical framework to interpret individual response to intervention: Paving the way for personalized nutrition and exercise prescription. Frontiers in Nutrition, 5, 41. https://doi.org/10.3389/fnut.2018.00041",
        "Tabachnick, B. G., & Fidell, L. S. (2019). Using multivariate statistics (7th ed.). Pearson.",
        "von Elm, E., Altman, D. G., Egger, M., Pocock, S. J., Gøtzsche, P. C., & Vandenbroucke, J. P. (2007). The Strengthening the Reporting of Observational Studies in Epidemiology (STROBE) statement: Guidelines for reporting observational studies. The Lancet, 370(9596), 1453–1457. https://doi.org/10.1016/S0140-6736(07)61602-X",
        "Wells, K. F., & Dillon, E. K. (1952). The sit and reach—A test of back and leg flexibility. Research Quarterly, 23(1), 115–118. https://doi.org/10.1080/10671188.1952.10761965",
        "Wind, A. E., Takken, T., Helders, P. J. M., & Engelbert, R. H. H. (2010). Is grip strength a predictor for total muscle strength in healthy children, adolescents, and young adults? European Journal of Pediatrics, 169(3), 281–287. https://doi.org/10.1007/s00431-009-1010-4",
        "World Health Organization. (2007). Growth reference data for 5–19 years. Bulletin of the World Health Organization, 85(9), 660–667. https://www.who.int/tools/growth-reference-data-for-5to19-years",
        "World Medical Association. (2013). World Medical Association Declaration of Helsinki: Ethical principles for medical research involving human subjects. JAMA, 310(20), 2191–2194. https://doi.org/10.1001/jama.2013.281053",
    ]

    # --- Build Turkish Report DOCX & MD ---
    doc_tr = docx.Document()
    for section in doc_tr.sections:
        section.top_margin = Inches(0.9)
        section.bottom_margin = Inches(0.9)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.9)

    add_heading_para(doc_tr, "YÖNTEM VE BULGULAR BÖLÜMLERİ ATIF REHBERİ VE ÖNERİLEN KAYNAKÇA RAPORU", size_pt=13, space_before=0, space_after=8)

    p_intro = doc_tr.add_paragraph()
    p_intro.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_intro.paragraph_format.space_after = Pt(10)
    p_intro.paragraph_format.line_spacing = 1.35
    r_in = p_intro.add_run(
        "Bu rapor, Q1_Master_Makale_Taslagi_TR.docx belgesinin Yöntem ve Bulgular bölümleri içerisinde koyu kırmızı renkle "
        "işaretlenen [Atıf Gerekli - 1] ile [Atıf Gerekli - 24] arasındaki her bir araştırma kılavuzu, tanı ölçütü, "
        "veri toplama aracı, biyomekanik/antropometrik hesaplama denklemi ve istatistiksel analiz noktasına eklenmesi "
        "gereken birincil akademik kaynakları göstermektedir. Birinci bölümde numaralı eşleştirme tablosu, ikinci "
        "bölümde ise makalenin Kaynaklar (References) bölümüne doğrudan eklenebilecek APA 7 formatında tam künyeler yer almaktadır."
    )
    set_run_font(r_in, size_pt=11)

    add_heading_para(doc_tr, "1. Metin İçi İşaretlerin ([Atıf Gerekli - #]) Eşleştirme Tablosu", size_pt=12, space_before=8, space_after=6)

    headers_tr = [
        "İşaret No",
        "Alt Bölüm ve Konu",
        "Metin İçi APA 7 Gösterimi",
        "Atıf Gerekçesi ve Önerilen Kaynak"
    ]
    table_tr = doc_tr.add_table(rows=1 + len(rows_tr), cols=4)
    table_tr.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_black_apa_borders(table_tr)
    col_widths = [1.25, 1.65, 1.55, 2.25]

    for j, h_text in enumerate(headers_tr):
        cell = table_tr.rows[0].cells[j]
        cell.width = Inches(col_widths[j])
        set_cell_border_bottom(cell, sz='6')
        set_cell_margins(cell)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.1
        r = p.add_run(h_text)
        set_run_font(r, size_pt=9.5, bold=True)

    for i, r_data in enumerate(rows_tr):
        row = table_tr.rows[i + 1]
        for j, c_text in enumerate(r_data):
            cell = row.cells[j]
            cell.width = Inches(col_widths[j])
            set_cell_margins(cell)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.1
            r = p.add_run(c_text)
            if j == 0:
                set_run_font(r, size_pt=9.0, bold=True, color=RED_COLOR)
            elif j == 2:
                set_run_font(r, size_pt=9.0, bold=True, color=BLACK_COLOR)
            else:
                set_run_font(r, size_pt=9.0, bold=False, color=BLACK_COLOR)

    add_heading_para(doc_tr, "2. Kaynaklar Bölümüne Eklenebilecek Tam Kaynakça Listesi (APA 7)", size_pt=12, space_before=14, space_after=6)

    for ref_txt in references_apa7:
        p = doc_tr.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.left_indent = Inches(0.4)
        p.paragraph_format.first_line_indent = Inches(-0.4)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(5)
        p.paragraph_format.line_spacing = 1.25
        r = p.add_run(ref_txt)
        set_run_font(r, size_pt=10.5)

    doc_tr.save(REPORT_TR_DOCX)
    scrub_docx_metadata(REPORT_TR_DOCX)
    print("Saved Atif_Rehberi_ve_Kaynakca_Raporu_TR.docx")

    md_lines_tr = [
        "# YÖNTEM VE BULGULAR BÖLÜMLERİ ATIF REHBERİ VE ÖNERİLEN KAYNAKÇA RAPORU\n",
        "Bu rapor, `Q1_Master_Makale_Taslagi_TR.docx` belgesinin Yöntem ve Bulgular bölümleri içerisinde koyu kırmızı renkle işaretlenen `[Atıf Gerekli - 1]` ile `[Atıf Gerekli - 24]` arasındaki her bir araştırma kılavuzu, tanı ölçütü, veri toplama aracı, biyomekanik/antropometrik hesaplama denklemi ve istatistiksel analiz noktasına eklenmesi gereken birincil akademik kaynakları göstermektedir.\n",
        "## 1. Metin İçi İşaretlerin (`[Atıf Gerekli - #]`) Eşleştirme Tablosu\n",
        "| İşaret No | Alt Bölüm ve Konu | Metin İçi APA 7 Gösterimi | Atıf Gerekçesi ve Önerilen Kaynak |",
        "| :--- | :--- | :--- | :--- |"
    ]
    for r in rows_tr:
        md_lines_tr.append(f"| **{r[0]}** | {r[1]} | **{r[2]}** | {r[3]} |")
    md_lines_tr.append("\n## 2. Kaynaklar Bölümüne Eklenebilecek Tam Kaynakça Listesi (APA 7)\n")
    for ref in references_apa7:
        md_lines_tr.append(f"- {ref}")
    with open(REPORT_TR_MD, "w", encoding="utf-8") as f:
        f.write("\n".join(md_lines_tr) + "\n")
    print("Saved Atif_Rehberi_ve_Kaynakca_Raporu_TR.md")

    # --- Build English Report DOCX & MD ---
    doc_en = docx.Document()
    for section in doc_en.sections:
        section.top_margin = Inches(0.9)
        section.bottom_margin = Inches(0.9)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.9)

    add_heading_para(doc_en, "METHODS AND RESULTS CITATION GUIDE AND RECOMMENDED REFERENCE REPORT", size_pt=13, space_before=0, space_after=8)

    p_intro_en = doc_en.add_paragraph()
    p_intro_en.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_intro_en.paragraph_format.space_after = Pt(10)
    p_intro_en.paragraph_format.line_spacing = 1.35
    r_in_en = p_intro_en.add_run(
        "This report maps each crimson-highlighted marker ([Citation Needed - 1] through [Citation Needed - 24]) "
        "in the Methods and Results sections of Q1_Master_Manuscript_EN.docx to its corresponding reporting guideline, "
        "diagnostic standard, measurement instrument, biomechanical/anthropometric equation, and statistical procedure. "
        "Section 1 provides the numbered lookup table, and Section 2 provides the complete APA 7th Edition reference list."
    )
    set_run_font(r_in_en, size_pt=11)

    add_heading_para(doc_en, "1. In-Text Citation Marker ([Citation Needed - #]) Mapping Table", size_pt=12, space_before=8, space_after=6)

    headers_en = [
        "Marker No",
        "Subsection & Topic",
        "In-Text APA 7 Citation",
        "Methodological Justification & Primary Source"
    ]
    table_en = doc_en.add_table(rows=1 + len(rows_en), cols=4)
    table_en.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_black_apa_borders(table_en)

    for j, h_text in enumerate(headers_en):
        cell = table_en.rows[0].cells[j]
        cell.width = Inches(col_widths[j])
        set_cell_border_bottom(cell, sz='6')
        set_cell_margins(cell)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.1
        r = p.add_run(h_text)
        set_run_font(r, size_pt=9.5, bold=True)

    for i, r_data in enumerate(rows_en):
        row = table_en.rows[i + 1]
        for j, c_text in enumerate(r_data):
            cell = row.cells[j]
            cell.width = Inches(col_widths[j])
            set_cell_margins(cell)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.1
            r = p.add_run(c_text)
            if j == 0:
                set_run_font(r, size_pt=9.0, bold=True, color=RED_COLOR)
            elif j == 2:
                set_run_font(r, size_pt=9.0, bold=True, color=BLACK_COLOR)
            else:
                set_run_font(r, size_pt=9.0, bold=False, color=BLACK_COLOR)

    add_heading_para(doc_en, "2. Complete Reference List for Manuscript Inclusion (APA 7th Edition)", size_pt=12, space_before=14, space_after=6)

    for ref_txt in references_apa7:
        p = doc_en.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.left_indent = Inches(0.4)
        p.paragraph_format.first_line_indent = Inches(-0.4)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(5)
        p.paragraph_format.line_spacing = 1.25
        r = p.add_run(ref_txt)
        set_run_font(r, size_pt=10.5)

    doc_en.save(REPORT_EN_DOCX)
    scrub_docx_metadata(REPORT_EN_DOCX)
    print("Saved Citation_Guide_and_Reference_Report_EN.docx")

    md_lines_en = [
        "# METHODS AND RESULTS CITATION GUIDE AND RECOMMENDED REFERENCE REPORT\n",
        "This report maps each crimson-highlighted marker (`[Citation Needed - 1]` through `[Citation Needed - 24]`) in the Methods and Results sections of `Q1_Master_Manuscript_EN.docx` to its corresponding reporting guideline, diagnostic standard, measurement instrument, biomechanical/anthropometric equation, and statistical procedure.\n",
        "## 1. In-Text Citation Marker (`[Citation Needed - #]`) Mapping Table\n",
        "| Marker No | Subsection & Topic | In-Text APA 7 Citation | Methodological Justification & Primary Source |",
        "| :--- | :--- | :--- | :--- |"
    ]
    for r in rows_en:
        md_lines_en.append(f"| **{r[0]}** | {r[1]} | **{r[2]}** | {r[3]} |")
    md_lines_en.append("\n## 2. Complete Reference List for Manuscript Inclusion (APA 7th Edition)\n")
    for ref in references_apa7:
        md_lines_en.append(f"- {ref}")
    with open(REPORT_EN_MD, "w", encoding="utf-8") as f:
        f.write("\n".join(md_lines_en) + "\n")
    print("Saved Citation_Guide_and_Reference_Report_EN.md")


if __name__ == "__main__":
    build_tr_manuscript()
    build_en_manuscript()
    build_citation_reports()
