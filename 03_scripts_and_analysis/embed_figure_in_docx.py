import os
import shutil
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

artifact_dir = "/Users/alik/.gemini/antigravity/brain/b070dc08-8fc0-46d0-9bc0-8e6ac510db31"
manuscript_dir = "/Users/alik/Desktop/anti_projeler/articles_all/autism_2026_testtekrar/04_manuscript"

img_tr = os.path.join(manuscript_dir, "Sekil_1_Arastirma_Tasarimi_TR.png")
img_en = os.path.join(manuscript_dir, "Figure_1_Study_Design_EN.png")

# Copy to artifact dir for direct embedding/viewing
shutil.copy(img_tr, os.path.join(artifact_dir, "Sekil_1_Arastirma_Tasarimi_TR.png"))
shutil.copy(img_en, os.path.join(artifact_dir, "Figure_1_Study_Design_EN.png"))
print("Görseller artifact dizinine kopyalandı.")

def wipe_metadata(doc):
    cp = doc.core_properties
    cp.author = ""
    cp.last_modified_by = ""
    cp.comments = ""
    cp.title = ""
    cp.subject = ""
    cp.keywords = ""
    cp.category = ""

def insert_figure_into_docx(docx_path, image_path, caption_text, target_heading="2.1."):
    doc = Document(docx_path)
    
    # Locate where to insert (after section 2.1 text or right before 2.2)
    insert_idx = None
    for idx, p in enumerate(doc.paragraphs):
        if "2.2." in p.text:
            insert_idx = idx
            break
            
    # If 2.2 found, insert right before 2.2
    if insert_idx is not None:
        p_target = doc.paragraphs[insert_idx]
        p_img = p_target.insert_paragraph_before()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(12)
        p_img.paragraph_format.space_after = Pt(4)
        run_img = p_img.add_run()
        run_img.add_picture(image_path, width=Inches(6.5))
        
        p_cap = p_target.insert_paragraph_before()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_after = Pt(14)
        run_cap = p_cap.add_run(caption_text)
        run_cap.font.name = 'Times New Roman'
        run_cap.font.size = Pt(10)
        run_cap.italic = True
        run_cap.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
    else:
        # Append to end
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(12)
        p_img.paragraph_format.space_after = Pt(4)
        run_img = p_img.add_run()
        run_img.add_picture(image_path, width=Inches(6.5))
        
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_after = Pt(14)
        run_cap = p_cap.add_run(caption_text)
        run_cap.font.name = 'Times New Roman'
        run_cap.font.size = Pt(10)
        run_cap.italic = True
        run_cap.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    wipe_metadata(doc)
    doc.save(docx_path)
    print(f"Görsel eklendi: {docx_path}")

cap_tr = "Şekil 1. Araştırma tasarımı, katılımcı uygunluk kriterleri, 14 haftalık deneysel protokol ve STROBE kohort akış şeması."
cap_en = "Figure 1. Study design, participant eligibility criteria, 14-week experimental timeline, and STROBE cohort flowchart."

insert_figure_into_docx(os.path.join(manuscript_dir, "02_Yontem_TR.docx"), img_tr, cap_tr)
insert_figure_into_docx(os.path.join(manuscript_dir, "02_Methods_EN.docx"), img_en, cap_en)
insert_figure_into_docx(os.path.join(manuscript_dir, "Q1_Master_Makale_Taslagi_TR.docx"), img_tr, cap_tr)
insert_figure_into_docx(os.path.join(manuscript_dir, "Q1_Master_Manuscript_EN.docx"), img_en, cap_en)
