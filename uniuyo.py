from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, ns
import re

# -------------------------------
# Load document
# -------------------------------
doc = Document("input.docx")

# -------------------------------
# Page setup
# -------------------------------
section = doc.sections[0]
section.left_margin = Inches(1.5)
section.right_margin = Inches(1)
section.top_margin = Inches(1)
section.bottom_margin = Inches(1)

# -------------------------------
# Page numbers (top right)
# -------------------------------
header = section.header
p = header.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run = p.add_run()
fldBegin = OxmlElement('w:fldChar')
fldBegin.set(ns.qn('w:fldCharType'), 'begin')
instr = OxmlElement('w:instrText')
instr.text = "PAGE"
fldEnd = OxmlElement('w:fldChar')
fldEnd.set(ns.qn('w:fldCharType'), 'end')
run._r.append(fldBegin)
run._r.append(instr)
run._r.append(fldEnd)

# -------------------------------
# Insert Table of Contents (TOC)
# -------------------------------
try:
    toc_paragraph = doc.paragraphs[0]
    toc_run = toc_paragraph.add_run()
    fldBegin = OxmlElement('w:fldChar')
    fldBegin.set(ns.qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.text = 'TOC \\o "1-3" \\h \\z \\u'
    fldEnd = OxmlElement('w:fldChar')
    fldEnd.set(ns.qn('w:fldCharType'), 'end')
    toc_run._r.append(fldBegin)
    toc_run._r.append(instrText)
    toc_run._r.append(fldEnd)
    toc_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
except Exception as e:
    print("⚠ TOC insertion skipped:", e)

# -------------------------------
# Track Figures/Tables
# -------------------------------
figure_count = 0
table_count = 0
figures = {}
tables = {}

# -------------------------------
# Format paragraphs, headings, in-text et al.
# -------------------------------
for para in doc.paragraphs:
    text = para.text.strip()
    if not text:
        continue

    # -----------------
    # Chapter headings
    if text.upper().startswith("CHAPTER"):
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.style = doc.styles['Heading 1']
        for run in para.runs:
            run.font.size = Pt(14)
            run.font.bold = True
            run.font.color.rgb = None
        continue

    # -----------------
    # ALL CAPS headings → Heading 2 (subhead, bold)
    elif text.isupper() and len(text.split()) <= 6:
        para.style = doc.styles['Heading 2']
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for run in para.runs:
            run.font.bold = True
            run.font.name = "Times New Roman"
            run.font.size = Pt(14)
            run.font.color.rgb = None
        continue

    # -----------------
    # Subheadings → Heading 3 (subhead, bold)
    elif text[0].isupper() and text[1:].islower():
        para.style = doc.styles['Heading 3']
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for run in para.runs:
            run.font.bold = True
            run.font.name = "Times New Roman"
            run.font.size = Pt(12)
            run.font.color.rgb = None

    # -----------------
    # Body formatting
    para.paragraph_format.first_line_indent = Inches(0.5)
    para.paragraph_format.line_spacing = 2
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # -----------------
    # Reconstruct paragraph text with italicized 'et al.'
    full_text = para.text
    para.text = ""
    last_index = 0
    for match in re.finditer(r"et al\.", full_text, flags=re.IGNORECASE):
        # Text before et al.
        before = full_text[last_index:match.start()]
        if before:
            run = para.add_run(before)
            run.font.name = "Times New Roman"
            run.font.size = Pt(12)
            run.font.color.rgb = None
        # Italicized et al.
        run = para.add_run(match.group())
        run.italic = True
        run.font.name = "Times New Roman"
        run.font.size = Pt(12)
        run.font.color.rgb = None
        last_index = match.end()
    # Remaining text
    after = full_text[last_index:]
    if after:
        run = para.add_run(after)
        run.font.name = "Times New Roman"
        run.font.size = Pt(12)
        run.font.color.rgb = None

    # -----------------
    # Figures
    if text.lower().startswith("figure"):
        figure_count += 1
        figures[text.lower()] = f"Figure {figure_count}"
        para.text = f"{figures[text.lower()]}: {text[6:].strip()}"

    # -----------------
    # Tables
    if text.lower().startswith("table"):
        table_count += 1
        tables[text.lower()] = f"Table {table_count}"
        para.text = f"{tables[text.lower()]}: {text[5:].strip()}"

# -------------------------------
# Replace in-text figure/table mentions
# -------------------------------
for para in doc.paragraphs:
    for key, val in figures.items():
        para.text = re.sub(r"\b" + re.escape(key) + r"\b", val, para.text, flags=re.IGNORECASE)
    for key, val in tables.items():
        para.text = re.sub(r"\b" + re.escape(key) + r"\b", val, para.text, flags=re.IGNORECASE)

# -------------------------------
# Format tables
# -------------------------------
for table in doc.tables:
    table.autofit = True
    for row in table.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                para.paragraph_format.line_spacing = 2
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                para.paragraph_format.first_line_indent = Inches(0.5)
                for run in para.runs:
                    run.font.name = "Times New Roman"
                    run.font.size = Pt(12)
                    run.font.color.rgb = None

# -------------------------------
# APA References
# -------------------------------
references_started = False
reference_paragraphs = []

for para in doc.paragraphs:
    if para.text.strip().upper() == "REFERENCES":
        references_started = True
        continue
    if references_started and para.text.strip() != "":
        reference_paragraphs.append(para)

if reference_paragraphs:
    reference_texts = [p.text for p in reference_paragraphs]
    reference_texts.sort(key=lambda x: x.split(',')[0].strip().lower())
    for p in reference_paragraphs:
        p.clear()
    for i, text in enumerate(reference_texts):
        para = reference_paragraphs[i]
        parts = text.split("et al.")
        for j, part in enumerate(parts):
            run = para.add_run(part)
            run.font.name = "Times New Roman"
            run.font.size = Pt(12)
            run.font.color.rgb = None
            if j < len(parts) - 1:
                etal_run = para.add_run("et al.")
                etal_run.italic = True
                etal_run.font.name = "Times New Roman"
                etal_run.font.size = Pt(12)
                etal_run.font.color.rgb = None
        para.paragraph_format.first_line_indent = Inches(-0.5)
        para.paragraph_format.left_indent = Inches(0.5)
        para.paragraph_format.line_spacing = 1  # single line
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

# -------------------------------
# List of Figures
# -------------------------------
if figures:
    doc.add_page_break()
    lof = doc.add_paragraph("LIST OF FIGURES")
    lof.alignment = WD_ALIGN_PARAGRAPH.CENTER
    lof.runs[0].font.bold = True
    lof.runs[0].font.size = Pt(14)
    lof.runs[0].font.color.rgb = None
    for val in figures.values():
        p = doc.add_paragraph(val)
        p.paragraph_format.left_indent = Inches(0.5)
        p.paragraph_format.first_line_indent = Inches(-0.5)
        p.paragraph_format.line_spacing = 2
        for run in p.runs:
            run.font.name = "Times New Roman"
            run.font.size = Pt(12)
            run.font.color.rgb = None

# -------------------------------
# List of Tables
# -------------------------------
if tables:
    doc.add_page_break()
    lot = doc.add_paragraph("LIST OF TABLES")
    lot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    lot.runs[0].font.bold = True
    lot.runs[0].font.size = Pt(14)
    lot.runs[0].font.color.rgb = None
    for val in tables.values():
        p = doc.add_paragraph(val)
        p.paragraph_format.left_indent = Inches(0.5)
        p.paragraph_format.first_line_indent = Inches(-0.5)
        p.paragraph_format.line_spacing = 2
        for run in p.runs:
            run.font.name = "Times New Roman"
            run.font.size = Pt(12)
            run.font.color.rgb = None

# -------------------------------
# Save final document
# -------------------------------
doc.save("formatted_uniuyo_project.docx")
print("✔ Formatting complete! Subheads bolded, et al italicized, references single-spaced.")
