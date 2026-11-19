import os
import re
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_shading(cell, color_hex="F2F2F2"):
    shading_xml = f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>'
    cell._tc.get_or_add_tcPr().append(parse_xml(shading_xml))

def set_apa_table_borders(table):
    tblPr = table._tbl.tblPr
    tblBorders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
        f'<w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
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
            f'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            f'</w:tcBorders>'
        )
        tcPr.append(tcBorders)

def create_styled_document():
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(11)
    font.color.rgb = RGBColor(0x22, 0x22, 0x22)
    style.paragraph_format.line_spacing = 1.25
    style.paragraph_format.space_after = Pt(6)
    return doc

def add_inline_formatted_text(p, text):
    # Splits by bold (**...**) and inline math/code (`...` or $...$)
    tokens = re.split(r'(\*\*.*?\*\*|\*.*?\*|`.*?`|\$.*?\$)', text)
    for t in tokens:
        if not t:
            continue
        if t.startswith('**') and t.endswith('**'):
            run = p.add_run(t[2:-2])
            run.bold = True
        elif t.startswith('*') and t.endswith('*'):
            run = p.add_run(t[1:-1])
            run.italic = True
        elif t.startswith('`') and t.endswith('`'):
            run = p.add_run(t[1:-1])
            run.font.name = 'Consolas'
            run.font.size = Pt(10)
        elif t.startswith('$') and t.endswith('$'):
            run = p.add_run(t[1:-1])
            run.italic = True
        else:
            p.add_run(t)

def markdown_to_docx(md_path, docx_path, title_text=None):
    doc = create_styled_document()
    
    with open(md_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        
    in_table = False
    table_lines = []
    
    for line in lines:
        raw = line.rstrip('\r\n')
        stripped = raw.strip()
        
        # Check table
        if stripped.startswith('|') and stripped.endswith('|'):
            in_table = True
            table_lines.append(stripped)
            continue
        else:
            if in_table and table_lines:
                # Process table
                render_markdown_table(doc, table_lines)
                table_lines = []
                in_table = False
                
        if not stripped:
            continue
            
        if stripped.startswith('# '):
            h = doc.add_heading(level=1)
            run = h.add_run(stripped[2:])
            run.font.name = 'Times New Roman'
            run.font.size = Pt(16)
            run.bold = True
            run.font.color.rgb = RGBColor(0x11, 0x11, 0x11)
            h.paragraph_format.space_before = Pt(12)
            h.paragraph_format.space_after = Pt(6)
        elif stripped.startswith('## '):
            h = doc.add_heading(level=2)
            run = h.add_run(stripped[3:])
            run.font.name = 'Times New Roman'
            run.font.size = Pt(13)
            run.bold = True
            run.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
            h.paragraph_format.space_before = Pt(10)
            h.paragraph_format.space_after = Pt(4)
        elif stripped.startswith('### '):
            h = doc.add_heading(level=3)
            run = h.add_run(stripped[4:])
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
            add_inline_formatted_text(p, stripped[2:])
            for run in p.runs:
                run.italic = True
        elif stripped.startswith('* ') or stripped.startswith('- '):
            p = doc.add_paragraph(style='List Bullet')
            p.paragraph_format.space_after = Pt(3)
            add_inline_formatted_text(p, stripped[2:])
        elif re.match(r'^\d+\.\s+', stripped):
            match = re.match(r'^\d+\.\s+', stripped)
            p = doc.add_paragraph(style='List Number')
            p.paragraph_format.space_after = Pt(3)
            add_inline_formatted_text(p, stripped[match.end():])
        elif stripped.startswith('$$') and stripped.endswith('$$'):
            # Math block
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(stripped[2:-2].strip())
            run.italic = True
            run.font.size = Pt(11)
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(4)
        elif stripped.startswith('```'):
            continue
        elif stripped.startswith('---'):
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(6)
            run = p.add_run('__________________________________________________________________')
            run.font.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            p = doc.add_paragraph()
            add_inline_formatted_text(p, stripped)
            
    if in_table and table_lines:
        render_markdown_table(doc, table_lines)
        
    doc.save(docx_path)
    print(f"Kaydedildi: {docx_path}")

def render_markdown_table(doc, table_lines):
    rows_data = []
    for line in table_lines:
        parts = [c.strip() for c in line.strip('|').split('|')]
        # Skip separator row (e.g. :--- | :---:)
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
            p.paragraph_format.line_spacing = 1.05
            
            # Formatting
            if is_header:
                set_cell_shading(cell, "F4F5F7")
                add_inline_formatted_text(p, val)
                for run in p.runs:
                    run.bold = True
                    run.font.size = Pt(9.5)
            else:
                add_inline_formatted_text(p, val)
                for run in p.runs:
                    run.font.size = Pt(9.5)
                    
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

# Execute conversions
base_dir = "/Users/alik/Desktop/anti_projeler/articles_all/autism_2026_testtekrar/04_manuscript"

files = [
    ("01_Title_and_Abstract.md", "01_Title_and_Abstract.docx"),
    ("02_Methods.md", "02_Methods.docx"),
    ("03_Results.md", "03_Results.docx"),
    ("Tables_and_Figures.md", "Tables_and_Figures.docx")
]

for md_file, docx_file in files:
    md_p = os.path.join(base_dir, md_file)
    docx_p = os.path.join(base_dir, docx_file)
    if os.path.exists(md_p):
        markdown_to_docx(md_p, docx_p)

# Also generate a MASTER unified Word document!
master_docx_path = os.path.join(base_dir, "Q1_Makale_Master_Taslak_Tam_Metin.docx")
doc_master = create_styled_document()

# Add Title
h = doc_master.add_heading(level=1)
r = h.add_run("Q1 DÜZEYİNDE MAKALE MASTER TASLAĞI (TITLE, ABSTRACT, METHODS, RESULTS, TABLES)")
r.font.name = 'Times New Roman'
r.font.size = Pt(16)
r.bold = True

p_info = doc_master.add_paragraph()
r_sub = p_info.add_run("Otizm Spektrum Bozukluğu (OSB) Retrospektif Kohort Çalışması (2021 Ölçümleri, N=57, n=34 Tamamlayan)")
r_sub.italic = True

for md_file, _ in files:
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
                    render_markdown_table(doc_master, table_lines)
                    table_lines = []
                    in_table = False
            if not stripped:
                continue
            if stripped.startswith('# '):
                h = doc_master.add_heading(level=1)
                run = h.add_run(stripped[2:])
                run.bold = True
                run.font.size = Pt(15)
            elif stripped.startswith('## '):
                h = doc_master.add_heading(level=2)
                run = h.add_run(stripped[3:])
                run.bold = True
                run.font.size = Pt(13)
            elif stripped.startswith('### '):
                h = doc_master.add_heading(level=3)
                run = h.add_run(stripped[4:])
                run.bold = True
                run.italic = True
                run.font.size = Pt(11.5)
            elif stripped.startswith('> '):
                p = doc_master.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.4)
                add_inline_formatted_text(p, stripped[2:])
                for run in p.runs:
                    run.italic = True
            elif stripped.startswith('* ') or stripped.startswith('- '):
                p = doc_master.add_paragraph(style='List Bullet')
                add_inline_formatted_text(p, stripped[2:])
            elif re.match(r'^\d+\.\s+', stripped):
                match = re.match(r'^\d+\.\s+', stripped)
                p = doc_master.add_paragraph(style='List Number')
                add_inline_formatted_text(p, stripped[match.end():])
            elif stripped.startswith('$$') and stripped.endswith('$$'):
                p = doc_master.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(stripped[2:-2].strip())
                run.italic = True
            elif stripped.startswith('---'):
                p = doc_master.add_paragraph()
                run = p.add_run('__________________________________________________________________')
                run.font.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                p = doc_master.add_paragraph()
                add_inline_formatted_text(p, stripped)
        if in_table and table_lines:
            render_markdown_table(doc_master, table_lines)
        doc_master.add_page_break()

doc_master.save(master_docx_path)
print(f"Master Word Dokümanı Kaydedildi: {master_docx_path}")
