import os
import re
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def sanitize_math_and_units(text):
    """
    Cleans all LaTeX constructs, math codes, and raw LaTeX artifacts
    so the output is pure, natural academic typography.
    """
    if not text:
        return ""
    # Remove \text{...}, \mathrm{...}, \mathbf{...}
    text = re.sub(r'\\text\{([^}]+)\}', r'\1', text)
    text = re.sub(r'\\mathrm\{([^}]+)\}', r'\1', text)
    text = re.sub(r'\\mathbf\{([^}]+)\}', r'\1', text)
    text = re.sub(r'\\mathit\{([^}]+)\}', r'\1', text)
    
    # Common LaTeX math symbols to unicode
    text = text.replace(r'\times', '×')
    text = text.replace(r'\pm', '±')
    text = text.replace(r'\ge', '≥')
    text = text.replace(r'\le', '≤')
    text = text.replace(r'\Delta', 'Δ')
    text = text.replace(r'\eta_p^2', 'ηp²')
    text = text.replace(r'\eta^2', 'η²')
    text = text.replace(r'\alpha', 'α')
    text = text.replace(r'\beta', 'β')
    text = text.replace(r'\approx', '≈')
    text = text.replace(r'\ne', '≠')
    text = text.replace(r'\rightarrow', '→')
    text = text.replace(r'\leftarrow', '←')
    text = text.replace(r'\lvert', '|')
    text = text.replace(r'\rvert', '|')
    text = text.replace(r'\left', '')
    text = text.replace(r'\right', '')
    text = text.replace(r'\cdot', '·')
    
    # Superscripts
    text = text.replace('^2', '²')
    text = text.replace('^3', '³')
    text = text.replace(r'^\circ', '°')
    text = text.replace(r'^\prime', "'")
    
    # Fractions: \frac{a}{b} -> a / b
    text = re.sub(r'\\frac\{([^}]+)\}\{([^}]+)\}', r'\1 / \2', text)
    
    # Clean any remaining LaTeX escaped symbols like \% or \$
    text = text.replace(r'\%', '%')
    text = text.replace(r'\$', '$')
    
    # Strip any stray single dollar signs used for math
    text = re.sub(r'(?<!\$)\$(?!\$)', '', text)
    
    # Clean stray braces
    text = text.replace('{', '').replace('}', '')
    
    # Clean double spaces
    text = re.sub(r'[ \t]+', ' ', text)
    return text.strip()

def wipe_docx_metadata(doc):
    """
    Completely wipes all author, company, comments, and application metadata
    so that the docx file is 100% clean and authentic.
    """
    cp = doc.core_properties
    cp.author = ""
    cp.last_modified_by = ""
    cp.comments = ""
    cp.title = ""
    cp.subject = ""
    cp.keywords = ""
    cp.category = ""
    cp.description = ""
    cp.identifier = ""
    cp.version = ""

def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_shading(cell, color_hex="F5F6F8"):
    shading_xml = f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>'
    cell._tc.get_or_add_tcPr().append(parse_xml(shading_xml))

def set_apa_table_borders(table):
    tblPr = table._tbl.tblPr
    tblBorders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="8" w:space="0" w:color="333333"/>'
        f'<w:bottom w:val="single" w:sz="8" w:space="0" w:color="333333"/>'
        f'<w:insideH w:val="none"/>'
        f'<w:insideV w:val="none"/>'
        f'<w:left w:val="none"/>'
        f'<w:right w:val="none"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(tblBorders)

def add_header_bottom_border(row):
    for cell in row.cells:
        tcPr = cell._tc.get_or_add_tcPr()
        tcBorders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="333333"/>'
            f'</w:tcBorders>'
        )
        tcPr.append(tcBorders)

def create_academic_document():
    doc = Document()
    wipe_docx_metadata(doc)
    
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(11.5)
    font.color.rgb = RGBColor(0x1A, 0x1A, 0x1A)
    style.paragraph_format.line_spacing = 1.3
    style.paragraph_format.space_after = Pt(6)
    return doc

def add_formatted_runs(p, text):
    # Splits by bold (**...**) and italic (*...*)
    tokens = re.split(r'(\*\*.*?\*\*|\*.*?\*|`.*?`)', text)
    for t in tokens:
        if not t:
            continue
        cleaned = sanitize_math_and_units(t)
        if t.startswith('**') and t.endswith('**'):
            inner = sanitize_math_and_units(t[2:-2])
            run = p.add_run(inner)
            run.bold = True
            run.font.name = 'Times New Roman'
        elif t.startswith('*') and t.endswith('*'):
            inner = sanitize_math_and_units(t[1:-1])
            run = p.add_run(inner)
            run.italic = True
            run.font.name = 'Times New Roman'
        elif t.startswith('`') and t.endswith('`'):
            inner = sanitize_math_and_units(t[1:-1])
            run = p.add_run(inner)
            run.font.name = 'Times New Roman'
        else:
            run = p.add_run(cleaned)
            run.font.name = 'Times New Roman'

def render_table(doc, table_lines):
    rows_data = []
    for line in table_lines:
        parts = [c.strip() for c in line.strip('|').split('|')]
        if all(re.match(r'^:?-+:?$', p) for p in parts if p):
            continue
        rows_data.append(parts)
        
    if not rows_data:
        return
        
    num_rows = len(rows_data)
    num_cols = max(len(r) for r in rows_data)
    
    table = doc.add_table(rows=num_rows, cols=num_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_apa_table_borders(table)
    
    for r_idx, r_data in enumerate(rows_data):
        row = table.rows[r_idx]
        is_header = (r_idx == 0)
        
        if is_header:
            add_header_bottom_border(row)
            
        for c_idx in range(num_cols):
            cell = row.cells[c_idx]
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            val = r_data[c_idx] if c_idx < len(r_data) else ""
            
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.1
            
            if is_header:
                set_cell_shading(cell, "F2F4F7")
                add_formatted_runs(p, val)
                for run in p.runs:
                    run.bold = True
                    run.font.size = Pt(9.5)
                    run.font.name = 'Times New Roman'
            else:
                add_formatted_runs(p, val)
                for run in p.runs:
                    run.font.size = Pt(9.5)
                    run.font.name = 'Times New Roman'
                    
    p_spacer = doc.add_paragraph()
    p_spacer.paragraph_format.space_after = Pt(6)

def convert_md_file_to_docx(md_path, docx_path):
    doc = create_academic_document()
    
    with open(md_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        
    in_table = False
    table_lines = []
    
    for line in lines:
        raw = line.rstrip('\r\n')
        stripped = raw.strip()
        
        if stripped.startswith('|') and stripped.endswith('|'):
            in_table = True
            table_lines.append(stripped)
            continue
        else:
            if in_table and table_lines:
                render_table(doc, table_lines)
                table_lines = []
                in_table = False
                
        if not stripped:
            continue
            
        if stripped.startswith('# '):
            h = doc.add_heading(level=1)
            clean_title = sanitize_math_and_units(stripped[2:])
            run = h.add_run(clean_title)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(15)
            run.bold = True
            run.font.color.rgb = RGBColor(0x11, 0x11, 0x11)
            h.paragraph_format.space_before = Pt(12)
            h.paragraph_format.space_after = Pt(6)
        elif stripped.startswith('## '):
            h = doc.add_heading(level=2)
            clean_title = sanitize_math_and_units(stripped[3:])
            run = h.add_run(clean_title)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12.5)
            run.bold = True
            run.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
            h.paragraph_format.space_before = Pt(10)
            h.paragraph_format.space_after = Pt(4)
        elif stripped.startswith('### '):
            h = doc.add_heading(level=3)
            clean_title = sanitize_math_and_units(stripped[4:])
            run = h.add_run(clean_title)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(11.5)
            run.bold = True
            run.italic = True
            run.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
            h.paragraph_format.space_before = Pt(8)
            h.paragraph_format.space_after = Pt(3)
        elif stripped.startswith('> '):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.4)
            p.paragraph_format.right_indent = Inches(0.4)
            p.paragraph_format.space_after = Pt(6)
            add_formatted_runs(p, stripped[2:])
            for run in p.runs:
                run.italic = True
        elif stripped.startswith('* ') or stripped.startswith('- '):
            p = doc.add_paragraph(style='List Bullet')
            p.paragraph_format.space_after = Pt(3)
            add_formatted_runs(p, stripped[2:])
        elif re.match(r'^\d+\.\s+', stripped):
            match = re.match(r'^\d+\.\s+', stripped)
            p = doc.add_paragraph(style='List Number')
            p.paragraph_format.space_after = Pt(3)
            add_formatted_runs(p, stripped[match.end():])
        elif stripped.startswith('$$') and stripped.endswith('$$'):
            clean_formula = sanitize_math_and_units(stripped[2:-2])
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(clean_formula)
            run.italic = True
            run.font.size = Pt(11)
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(4)
        elif stripped.startswith('---'):
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(6)
            run = p.add_run('__________________________________________________________________')
            run.font.color.rgb = RGBColor(0xDD, 0xDD, 0xDD)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            p = doc.add_paragraph()
            add_formatted_runs(p, stripped)
            
    if in_table and table_lines:
        render_table(doc, table_lines)
        
    wipe_docx_metadata(doc)
    doc.save(docx_path)
    print(f"Başarıyla üretildi: {docx_path}")

base_dir = "/Users/alik/Desktop/anti_projeler/articles_all/autism_2026_testtekrar/04_manuscript"

# 1. İngilizce Ayrı Belgeler
en_files = [
    ("01_Title_and_Abstract_EN.md", "01_Title_and_Abstract_EN.docx"),
    ("02_Methods.md", "02_Methods_EN.docx"),
    ("03_Results.md", "03_Results_EN.docx"),
    ("Tables_and_Figures_EN.md", "04_Tables_and_Figures_EN.docx")
]

# 2. Türkçe Ayrı Belgeler
tr_files = [
    ("01_Baslik_ve_Ozet_TR.md", "01_Baslik_ve_Ozet_TR.docx"),
    ("02_Yontem_TR.md", "02_Yontem_TR.docx"),
    ("03_Bulgular_TR.md", "03_Bulgular_TR.docx"),
    ("Tablolar_ve_Sekiller_TR.md", "04_Tablolar_ve_Sekiller_TR.docx")
]

for md_file, docx_file in en_files + tr_files:
    md_p = os.path.join(base_dir, md_file)
    docx_p = os.path.join(base_dir, docx_file)
    if os.path.exists(md_p):
        convert_md_file_to_docx(md_p, docx_p)

# 3. Master Birleştirilmiş İngilizce Belge (Full English Manuscript)
master_en_path = os.path.join(base_dir, "Q1_Master_Manuscript_EN.docx")
doc_master_en = create_academic_document()

for idx, (md_file, _) in enumerate(en_files):
    md_p = os.path.join(base_dir, md_file)
    if os.path.exists(md_p):
        with open(md_p, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        in_table = False
        table_lines = []
        for line in lines:
            stripped = line.rstrip('\r\n').strip()
            if stripped.startswith('|') and stripped.endswith('|'):
                in_table = True
                table_lines.append(stripped)
                continue
            else:
                if in_table and table_lines:
                    render_table(doc_master_en, table_lines)
                    table_lines = []
                    in_table = False
            if not stripped:
                continue
            if stripped.startswith('# '):
                h = doc_master_en.add_heading(level=1)
                run = h.add_run(sanitize_math_and_units(stripped[2:]))
                run.bold = True
                run.font.size = Pt(15)
            elif stripped.startswith('## '):
                h = doc_master_en.add_heading(level=2)
                run = h.add_run(sanitize_math_and_units(stripped[3:]))
                run.bold = True
                run.font.size = Pt(12.5)
            elif stripped.startswith('### '):
                h = doc_master_en.add_heading(level=3)
                run = h.add_run(sanitize_math_and_units(stripped[4:]))
                run.bold = True
                run.italic = True
                run.font.size = Pt(11.5)
            elif stripped.startswith('> '):
                p = doc_master_en.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.4)
                add_formatted_runs(p, stripped[2:])
                for run in p.runs:
                    run.italic = True
            elif stripped.startswith('* ') or stripped.startswith('- '):
                p = doc_master_en.add_paragraph(style='List Bullet')
                add_formatted_runs(p, stripped[2:])
            elif re.match(r'^\d+\.\s+', stripped):
                match = re.match(r'^\d+\.\s+', stripped)
                p = doc_master_en.add_paragraph(style='List Number')
                add_formatted_runs(p, stripped[match.end():])
            elif stripped.startswith('---'):
                p = doc_master_en.add_paragraph()
                run = p.add_run('__________________________________________________________________')
                run.font.color.rgb = RGBColor(0xDD, 0xDD, 0xDD)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                p = doc_master_en.add_paragraph()
                add_formatted_runs(p, stripped)
        if in_table and table_lines:
            render_table(doc_master_en, table_lines)
        if idx < len(en_files) - 1:
            doc_master_en.add_page_break()

wipe_docx_metadata(doc_master_en)
doc_master_en.save(master_en_path)
print(f"Master İngilizce Belge Kaydedildi: {master_en_path}")

# 4. Master Birleştirilmiş Türkçe Belge (Full Turkish Manuscript)
master_tr_path = os.path.join(base_dir, "Q1_Master_Makale_Taslagi_TR.docx")
doc_master_tr = create_academic_document()

for idx, (md_file, _) in enumerate(tr_files):
    md_p = os.path.join(base_dir, md_file)
    if os.path.exists(md_p):
        with open(md_p, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        in_table = False
        table_lines = []
        for line in lines:
            stripped = line.rstrip('\r\n').strip()
            if stripped.startswith('|') and stripped.endswith('|'):
                in_table = True
                table_lines.append(stripped)
                continue
            else:
                if in_table and table_lines:
                    render_table(doc_master_tr, table_lines)
                    table_lines = []
                    in_table = False
            if not stripped:
                continue
            if stripped.startswith('# '):
                h = doc_master_tr.add_heading(level=1)
                run = h.add_run(sanitize_math_and_units(stripped[2:]))
                run.bold = True
                run.font.size = Pt(15)
            elif stripped.startswith('## '):
                h = doc_master_tr.add_heading(level=2)
                run = h.add_run(sanitize_math_and_units(stripped[3:]))
                run.bold = True
                run.font.size = Pt(12.5)
            elif stripped.startswith('### '):
                h = doc_master_tr.add_heading(level=3)
                run = h.add_run(sanitize_math_and_units(stripped[4:]))
                run.bold = True
                run.italic = True
                run.font.size = Pt(11.5)
            elif stripped.startswith('> '):
                p = doc_master_tr.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.4)
                add_formatted_runs(p, stripped[2:])
                for run in p.runs:
                    run.italic = True
            elif stripped.startswith('* ') or stripped.startswith('- '):
                p = doc_master_tr.add_paragraph(style='List Bullet')
                add_formatted_runs(p, stripped[2:])
            elif re.match(r'^\d+\.\s+', stripped):
                match = re.match(r'^\d+\.\s+', stripped)
                p = doc_master_tr.add_paragraph(style='List Number')
                add_formatted_runs(p, stripped[match.end():])
            elif stripped.startswith('---'):
                p = doc_master_tr.add_paragraph()
                run = p.add_run('__________________________________________________________________')
                run.font.color.rgb = RGBColor(0xDD, 0xDD, 0xDD)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                p = doc_master_tr.add_paragraph()
                add_formatted_runs(p, stripped)
        if in_table and table_lines:
            render_table(doc_master_tr, table_lines)
        if idx < len(tr_files) - 1:
            doc_master_tr.add_page_break()

wipe_docx_metadata(doc_master_tr)
doc_master_tr.save(master_tr_path)
print(f"Master Türkçe Belge Kaydedildi: {master_tr_path}")
