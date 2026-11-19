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
    if not text:
        return ""
    for ch in ['\u200b', '\u200c', '\u200d', '\ufeff', '\u00ad']:
        text = text.replace(ch, '')
    text = re.sub(r'\\text\{([^}]+)\}', r'\1', text)
    text = re.sub(r'\\mathrm\{([^}]+)\}', r'\1', text)
    text = re.sub(r'\\mathbf\{([^}]+)\}', r'\1', text)
    text = re.sub(r'\\mathit\{([^}]+)\}', r'\1', text)
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
    text = text.replace('^2', '²')
    text = text.replace('^3', '³')
    text = text.replace(r'^\circ', '°')
    text = text.replace(r'^\prime', "'")
    text = re.sub(r'\\frac\{([^}]+)\}\{([^}]+)\}', r'\1 / \2', text)
    text = text.replace(r'\%', '%')
    text = text.replace(r'\$', '$')
    text = re.sub(r'(?<!\$)\$(?!\$)', '', text)
    text = text.replace('{', '').replace('}', '')
    text = re.sub(r'[ \t]+', ' ', text)
    return text.strip()

def wipe_docx_metadata(doc):
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

def set_cell_margins(cell, top=60, bottom=60, left=80, right=80):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_apa_table_borders(table):
    # Strict APA: Top border (1 pt), Bottom border (1 pt), no vertical, no inside horizontal
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

def create_apa_document():
    doc = Document()
    wipe_docx_metadata(doc)
    
    # 1 inch (2.54 cm) margins on all sides
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    font.color.rgb = RGBColor(0x00, 0x00, 0x00) # Pure Black
    style.paragraph_format.line_spacing = 1.5   # Standard 1.5 line spacing
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(6)
    return doc

def add_formatted_runs(p, text):
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
            run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        elif t.startswith('*') and t.endswith('*'):
            inner = sanitize_math_and_units(t[1:-1])
            run = p.add_run(inner)
            run.italic = True
            run.font.name = 'Times New Roman'
            run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        elif t.startswith('`') and t.endswith('`'):
            inner = sanitize_math_and_units(t[1:-1])
            run = p.add_run(inner)
            run.font.name = 'Times New Roman'
            run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        else:
            run = p.add_run(cleaned)
            run.font.name = 'Times New Roman'
            run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

def render_apa_table(doc, table_lines):
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
            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
            val = r_data[c_idx] if c_idx < len(r_data) else ""
            
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.15
            
            # Align: First column left, other columns center
            if c_idx > 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                
            add_formatted_runs(p, val)
            for run in p.runs:
                run.font.name = 'Times New Roman'
                run.font.size = Pt(10) # 10pt in APA tables
                run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
                if is_header:
                    run.bold = True
                    
    p_spacer = doc.add_paragraph()
    p_spacer.paragraph_format.space_before = Pt(0)
    p_spacer.paragraph_format.space_after = Pt(6)

def insert_figure(doc, img_path, caption_text):
    if not os.path.exists(img_path):
        print(f"Warning: Image file not found: {img_path}")
        return
        
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(14)
    p_img.paragraph_format.space_after = Pt(4)
    r_img = p_img.add_run()
    # Fit nicely within 6.5 inch margins
    r_img.add_picture(img_path, width=Inches(6.2))
    
    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_before = Pt(2)
    p_cap.paragraph_format.space_after = Pt(14)
    r_cap = p_cap.add_run(sanitize_math_and_units(caption_text))
    r_cap.font.name = 'Times New Roman'
    r_cap.font.size = Pt(10)
    r_cap.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    r_cap.italic = True

def build_q1_document(md_path, docx_path, base_dir):
    doc = create_apa_document()
    
    with open(md_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        
    in_table = False
    table_lines = []
    
    i = 0
    n_lines = len(lines)
    
    while i < n_lines:
        raw = lines[i].rstrip('\r\n')
        stripped = raw.strip()
        
        # Table handling
        if stripped.startswith('|') and stripped.endswith('|'):
            in_table = True
            table_lines.append(stripped)
            i += 1
            continue
        else:
            if in_table and table_lines:
                render_apa_table(doc, table_lines)
                table_lines = []
                in_table = False
                
        if not stripped:
            i += 1
            continue
            
        # Markdown image detection: ![Caption](filename.png)
        img_match = re.match(r'^!\[(.*?)\]\((.*?)\)', stripped)
        if img_match:
            img_alt = img_match.group(1)
            img_rel = img_match.group(2)
            img_path = os.path.join(base_dir, img_rel) if not os.path.isabs(img_rel) else img_rel
            
            # Check if next line is caption like *Şekil X...* or *Figure X...*
            caption = img_alt
            if i + 1 < n_lines:
                next_line = lines[i+1].strip()
                if (next_line.startswith('*Şekil ') or next_line.startswith('*Figure ')) and next_line.endswith('*'):
                    caption = next_line[1:-1]
                    i += 1 # skip caption line since we consume it
            
            insert_figure(doc, img_path, caption)
            i += 1
            continue
            
        # Standalone italic figure caption (if image was processed or stand-alone)
        if (stripped.startswith('*Şekil ') or stripped.startswith('*Figure ')) and stripped.endswith('*'):
            # Only if not already consumed
            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_before = Pt(2)
            p_cap.paragraph_format.space_after = Pt(12)
            r_cap = p_cap.add_run(sanitize_math_and_units(stripped[1:-1]))
            r_cap.font.name = 'Times New Roman'
            r_cap.font.size = Pt(10)
            r_cap.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
            r_cap.italic = True
            i += 1
            continue

        # Table abbreviations line (9pt per anti-ai-writing-style Rule 2.2)
        if (stripped.startswith('*Kısaltmalar:') or stripped.startswith('*Abbreviations:')) and stripped.endswith('*'):
            p_abbr = doc.add_paragraph()
            p_abbr.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p_abbr.paragraph_format.space_before = Pt(0)
            p_abbr.paragraph_format.space_after = Pt(10)
            r_abbr = p_abbr.add_run(sanitize_math_and_units(stripped[1:-1]))
            r_abbr.font.name = 'Times New Roman'
            r_abbr.font.size = Pt(9)
            r_abbr.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
            r_abbr.italic = True
            i += 1
            continue

        if stripped.startswith('# '):
            # Main Title: 14pt, Bold, Centered, Black
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(12)
            clean_title = sanitize_math_and_units(stripped[2:])
            run = p.add_run(clean_title)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(14)
            run.bold = True
            run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        elif stripped.startswith('## '):
            # Section Heading 1 (e.g. METHODS / YÖNTEM): 12pt, Bold, Left, Black
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after = Pt(4)
            clean_title = sanitize_math_and_units(stripped[3:])
            run = p.add_run(clean_title.upper())
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
            run.bold = True
            run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        elif stripped.startswith('### '):
            # Subheading Level 2: 12pt, Bold, Left, Black
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(3)
            clean_title = sanitize_math_and_units(stripped[4:])
            
            # If it's Table title (e.g. Tablo 1. ... or Table 1. ...), format per APA
            if clean_title.startswith("Tablo ") or clean_title.startswith("Table "):
                run = p.add_run(clean_title)
                run.font.name = 'Times New Roman'
                run.font.size = Pt(11)
                run.bold = True
                run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
            else:
                run = p.add_run(clean_title)
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)
                run.bold = True
                run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        elif stripped.startswith('* ') or stripped.startswith('- '):
            p = doc.add_paragraph(style='List Bullet')
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(3)
            add_formatted_runs(p, stripped[2:])
        elif re.match(r'^\d+\.\s+', stripped):
            match = re.match(r'^\d+\.\s+', stripped)
            p = doc.add_paragraph(style='List Number')
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(3)
            add_formatted_runs(p, stripped[match.end():])
        elif stripped.startswith('---'):
            i += 1
            continue # Skip horizontal markdown lines for clean academic manuscript
        else:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(6)
            add_formatted_runs(p, stripped)
            
        i += 1
            
    if in_table and table_lines:
        render_apa_table(doc, table_lines)
        
    wipe_docx_metadata(doc)
    doc.save(docx_path)
    print(f"Q1 Formatında Word Belgesi Başarıyla Üretildi: {docx_path}")

if __name__ == '__main__':
    base_dir = "/Users/alik/Desktop/anti_projeler/articles_all/autism_2026_testtekrar/04_manuscript"
    
    md_tr = os.path.join(base_dir, "Q1_Master_Makale_Taslagi_TR.md")
    docx_tr = os.path.join(base_dir, "Q1_Master_Makale_Taslagi_TR.docx")
    
    md_en = os.path.join(base_dir, "Q1_Master_Manuscript_EN.md")
    docx_en = os.path.join(base_dir, "Q1_Master_Manuscript_EN.docx")
    
    build_q1_document(md_tr, docx_tr, base_dir)
    build_q1_document(md_en, docx_en, base_dir)
