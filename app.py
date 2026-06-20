import streamlit as st
import docx
from docx.shared import Pt, Inches
import pypdf
import re
import io

def split_mcq_options(line):
    matches = list(re.finditer(r'\b([A-D])\.\s', line))
    if not matches: return []
    options = []
    for i in range(len(matches)):
        start = matches[i].start()
        end = matches[i+1].start() if i + 1 < len(matches) else len(line)
        options.append(line[start:end].strip())
    return options

def add_options_grid(doc, options):
    if not options: return
    max_len = max(len(o) for o in options)
    num_cols = 4 if (len(options) == 4 and max_len < 20) else 2
    num_rows = (len(options) + num_cols - 1) // num_cols
    
    table = doc.add_table(rows=num_rows, cols=num_cols)
    table.autofit = False
    col_width = Inches(6.2 / num_cols)
    for row in table.rows:
        for cell in row.cells: cell.width = col_width
            
    for index, option in enumerate(options):
        row_idx = index // num_cols
        col_idx = index % num_cols
        cell = table.rows[row_idx].cells[col_idx]
        p = cell.paragraphs[0]
        p.paragraph_format.left_indent = Inches(0.1)
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(option)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        
    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = docx.oxml.OxmlElement('w:tcBorders')
            for b_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
                b = docx.oxml.OxmlElement(f'w:{b_name}')
                b.set(docx.oxml.ns.qn('w:val'), 'none')
                tcBorders.append(b)
            tcPr.append(tcBorders)

def process_paper(pdf_file, template_path, subject, grade, date, marks, exam_title):
    reader = pypdf.PdfReader(pdf_file)
    pdf_lines = []
    for page in reader.pages:
        text = page.extract_text()
        for line in text.split('\n'):
            if line.strip(): pdf_lines.append(line.strip())

    doc = docx.Document(template_path)
    
    # ADVANCED SEARCH & REPLACE ENGINE
    # This directly forces replacement regardless of Word's hidden formatting breaks
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                # Read the full combined plain text of the cell
                cell_text = "".join(p.text for p in cell.paragraphs)
                
                if "PLACEHOLDER_SUBJECT" in cell_text or "PLACEHOLDER_GRADE" in cell_text or "PLACEHOLDER_DATE" in cell_text or "PLACEHOLDER_MARKS" in cell_text or "PLACEHOLDER_EXAM" in cell_text:
                    # Clear out the broken segments completely
                    p = cell.paragraphs[0]
                    p.text = "" 
                    
                    # Determine what value to insert
                    if "PLACEHOLDER_SUBJECT" in cell_text:
                        text_to_write = f"SUBJECT : {subject}"
                    elif "PLACEHOLDER_GRADE" in cell_text:
                        text_to_write = f"GRADE : {grade}"
                    elif "PLACEHOLDER_DATE" in cell_text:
                        text_to_write = f"DATE : {date}"
                    elif "PLACEHOLDER_MARKS" in cell_text:
                        text_to_write = f"MARKS : {marks}"
                    elif "PLACEHOLDER_EXAM" in cell_text:
                        text_to_write = exam_title
                    
                    # Re-write the full text cleanly as a single segment
                    run = p.add_run(text_to_write)
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(11)
                    run.font.bold = True

    # Clear instructions text below the header
    while len(doc.paragraphs) > 0:
        p_to_remove = doc.paragraphs[-1]
        p_to_remove._element.getparent().remove(p_to_remove._element)

    current_options = []
    for line in pdf_lines:
        clean_line = line.strip()
        upper_line = clean_line.upper()
        
        if "SUBJECT:" in upper_line or "STANDARD:" in upper_line or "QUESTION PAPER" in upper_line or "MARKS:" in upper_line:
            continue
        if re.match(r'^\(\d+\)$', clean_line):
            continue

        has_explicit_options = bool(re.search(r'\b[A-D]\.\s', clean_line))
        is_broken_continuation = len(current_options) > 0 and not has_explicit_options and not re.match(r'^\d+\.', clean_line) and "SECTION -" not in upper_line and "►" not in clean_line

        if is_broken_continuation:
            current_options[-1] = current_options[-1] + " " + clean_line
            continue

        if not has_explicit_options and current_options:
            add_options_grid(doc, current_options)
            current_options = []

        if "SECTION -" in upper_line:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after = Pt(4)
            p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(clean_line)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(14)
            run.bold = True
        elif "►" in clean_line or "ANSWER THE FOLLOWING" in upper_line or "CHOOSE THE CORRECT" in upper_line:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(6)
            mark_match = re.search(r'\(\d{2}\)$', clean_line)
            if mark_match:
                mark_text = mark_match.group(0)
                instruction_text = clean_line
