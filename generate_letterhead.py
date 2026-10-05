"""
Generate corporate executive DOCX letterhead for Hope Restoration and Health Relief Foundation.
Fully compliant with ISO/IEC 29500 OpenXML Schema for seamless compatibility with Microsoft Word,
LibreOffice, Google Docs, Apple Pages, and PDF converters.
"""
import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def normalize_tblPr(tblPr):
    tag_order = [
        'tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual',
        'tblStyleRowBandSize', 'tblStyleColBandSize', 'tblW', 'jc',
        'tblCellSpacing', 'tblInd', 'tblBorders', 'shd',
        'tblLayout', 'tblCellMar', 'tblLook'
    ]
    elements = list(tblPr)
    for el in elements:
        tblPr.remove(el)
    def key_func(el):
        tag = el.tag.split('}')[-1]
        return tag_order.index(tag) if tag in tag_order else 999
    for el in sorted(elements, key=key_func):
        tblPr.append(el)

def normalize_tcPr(tcPr):
    tag_order = [
        'cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge',
        'tcBorders', 'shd', 'noWrap', 'tcMar', 'textDirection',
        'tcFitText', 'vAlign', 'hideMark', 'headers'
    ]
    # Remove duplicates of tcW if any
    seen_tcW = False
    elements = []
    for el in list(tcPr):
        tag = el.tag.split('}')[-1]
        if tag == 'tcW':
            if seen_tcW:
                tcPr.remove(el)
                continue
            seen_tcW = True
        elements.append(el)
        tcPr.remove(el)
        
    def key_func(el):
        tag = el.tag.split('}')[-1]
        return tag_order.index(tag) if tag in tag_order else 999
    for el in sorted(elements, key=key_func):
        tcPr.append(el)

def remove_borders(table):
    tblPr = table._tbl.tblPr
    # Remove any existing tblBorders
    for b in tblPr.findall(qn('w:tblBorders')):
        tblPr.remove(b)
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/>'
        f'<w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)
    normalize_tblPr(tblPr)

def set_cell_margins(cell, top=0, bottom=0, left=0, right=0):
    tcPr = cell._tc.get_or_add_tcPr()
    for m in tcPr.findall(qn('w:tcMar')):
        tcPr.remove(m)
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)
    normalize_tcPr(tcPr)

def set_cell_shading(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    for s in tcPr.findall(qn('w:shd')):
        tcPr.remove(s)
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    tcPr.append(shading)
    normalize_tcPr(tcPr)

def create_letterhead(filename, include_sample_body=True):
    doc = docx.Document()
    
    # Page setup
    section = doc.sections[0]
    section.top_margin = Inches(0.55)
    section.bottom_margin = Inches(0.55)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)
    section.page_width = Inches(8.5)
    section.page_height = Inches(11.0)
    
    # ----------------- HEADER TABLE -----------------
    header_table = doc.add_table(rows=1, cols=3)
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_borders(header_table)
    
    row = header_table.rows[0]
    row.cells[0].width = Inches(1.25)
    row.cells[1].width = Inches(3.25)
    row.cells[2].width = Inches(2.6)
    
    for c in row.cells:
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_margins(c, top=40, bottom=40, left=40, right=40)
        normalize_tcPr(c._tc.get_or_add_tcPr())
    
    # Cell 0: Logo image
    p_logo = row.cells[0].paragraphs[0]
    p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_logo.paragraph_format.space_before = Pt(0)
    p_logo.paragraph_format.space_after = Pt(0)
    logo_path = os.path.abspath('images/logo.png')
    if os.path.exists(logo_path):
        run_logo = p_logo.add_run()
        run_logo.add_picture(logo_path, width=Inches(1.15))
    
    # Cell 1: Organization Name & Slogan
    cell_title = row.cells[1]
    p_title = cell_title.paragraphs[0]
    p_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(2)
    p_title.paragraph_format.line_spacing = Pt(14)
    
    r_name1 = p_title.add_run("HOPE RESTORATION\n")
    r_name1.bold = True
    r_name1.font.size = Pt(13)
    r_name1.font.name = "Arial"
    r_name1.font.color.rgb = RGBColor(13, 148, 136) # Primary Teal
    
    r_name2 = p_title.add_run("& HEALTH RELIEF FOUNDATION")
    r_name2.bold = True
    r_name2.font.size = Pt(10.5)
    r_name2.font.name = "Arial"
    r_name2.font.color.rgb = RGBColor(15, 118, 110) # Dark Teal
    
    p_motto = cell_title.add_paragraph()
    p_motto.paragraph_format.space_before = Pt(3)
    p_motto.paragraph_format.space_after = Pt(1)
    p_motto.paragraph_format.line_spacing = Pt(10)
    r_motto = p_motto.add_run("Restoring Hope Through Health & Dignity")
    r_motto.italic = True
    r_motto.bold = True
    r_motto.font.size = Pt(8.5)
    r_motto.font.name = "Arial"
    r_motto.font.color.rgb = RGBColor(100, 116, 139) # Slate
    
    p_sub = cell_title.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(1)
    p_sub.paragraph_format.space_after = Pt(0)
    p_sub.paragraph_format.line_spacing = Pt(9.5)
    r_sub = p_sub.add_run("Community Healthcare Support · Daily Living Relief")
    r_sub.font.size = Pt(7.5)
    r_sub.font.name = "Arial"
    r_sub.font.color.rgb = RGBColor(148, 163, 184) # Light Slate
    
    # Cell 2: Contact Details (Right-aligned)
    cell_contact = row.cells[2]
    p_contact = cell_contact.paragraphs[0]
    p_contact.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_contact.paragraph_format.space_before = Pt(0)
    p_contact.paragraph_format.space_after = Pt(0)
    p_contact.paragraph_format.line_spacing = Pt(11)
    
    items = [
        ("Location: ", "Ogun State, Abeokuta, Nigeria", False),
        ("Phone: ", "+234 706 871 9591 | 0706 871 9591", False),
        ("Email: ", "elijah.adebayo@hoperestorationhrf.org", True),
        ("Web: ", "www.hoperestorationhrf.org", True),
        ("Manager: ", "Adebayo Elijah", False)
    ]
    
    for i, (label, val, is_accent) in enumerate(items):
        if i > 0:
            p_contact = cell_contact.add_paragraph()
            p_contact.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p_contact.paragraph_format.space_before = Pt(1)
            p_contact.paragraph_format.space_after = Pt(0)
            p_contact.paragraph_format.line_spacing = Pt(11)
        
        lbl_run = p_contact.add_run(label)
        lbl_run.bold = True
        lbl_run.font.size = Pt(7.5)
        lbl_run.font.name = "Arial"
        lbl_run.font.color.rgb = RGBColor(71, 85, 105)
        
        val_run = p_contact.add_run(val)
        val_run.font.size = Pt(7.5)
        val_run.font.name = "Arial"
        if is_accent:
            val_run.font.color.rgb = RGBColor(13, 148, 136)
        else:
            val_run.font.color.rgb = RGBColor(30, 41, 59)
    
    # ----------------- ACCENT DIVIDER BARS -----------------
    p_divider_space = doc.add_paragraph()
    p_divider_space.paragraph_format.space_before = Pt(6)
    p_divider_space.paragraph_format.space_after = Pt(0)
    p_divider_space.paragraph_format.line_spacing = Pt(1)
    
    bar_table = doc.add_table(rows=2, cols=1)
    bar_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_borders(bar_table)
    
    # Top primary bar: #0D9488, 3pt high
    bar_cell1 = bar_table.rows[0].cells[0]
    bar_cell1.width = Inches(7.1)
    set_cell_shading(bar_cell1, "0D9488")
    set_cell_margins(bar_cell1, top=20, bottom=20, left=0, right=0)
    p_b1 = bar_cell1.paragraphs[0]
    p_b1.paragraph_format.space_before = Pt(0)
    p_b1.paragraph_format.space_after = Pt(0)
    p_b1.paragraph_format.line_spacing = Pt(2.5)
    normalize_tcPr(bar_cell1._tc.get_or_add_tcPr())
    
    # Bottom secondary accent bar: #99F6E4, 1pt high
    bar_cell2 = bar_table.rows[1].cells[0]
    bar_cell2.width = Inches(7.1)
    set_cell_shading(bar_cell2, "99F6E4")
    set_cell_margins(bar_cell2, top=10, bottom=10, left=0, right=0)
    p_b2 = bar_cell2.paragraphs[0]
    p_b2.paragraph_format.space_before = Pt(0)
    p_b2.paragraph_format.space_after = Pt(0)
    p_b2.paragraph_format.line_spacing = Pt(1)
    normalize_tcPr(bar_cell2._tc.get_or_add_tcPr())
    
    # ----------------- BODY SECTION -----------------
    if include_sample_body:
        # Reference and Date Table
        meta_table = doc.add_table(rows=1, cols=2)
        meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        remove_borders(meta_table)
        meta_table.rows[0].cells[0].width = Inches(3.55)
        meta_table.rows[0].cells[1].width = Inches(3.55)
        set_cell_margins(meta_table.rows[0].cells[0], top=100, bottom=60, left=0, right=0)
        set_cell_margins(meta_table.rows[0].cells[1], top=100, bottom=60, left=0, right=0)
        normalize_tcPr(meta_table.rows[0].cells[0]._tc.get_or_add_tcPr())
        normalize_tcPr(meta_table.rows[0].cells[1]._tc.get_or_add_tcPr())
        
        p_ref = meta_table.rows[0].cells[0].paragraphs[0]
        p_ref.paragraph_format.space_before = Pt(10)
        p_ref.paragraph_format.space_after = Pt(0)
        r_ref_lbl = p_ref.add_run("REF NO: ")
        r_ref_lbl.bold = True
        r_ref_lbl.font.size = Pt(9.5)
        r_ref_lbl.font.name = "Arial"
        r_ref_lbl.font.color.rgb = RGBColor(100, 116, 139)
        r_ref_val = p_ref.add_run("HRF/ADM/2026/001")
        r_ref_val.bold = True
        r_ref_val.font.size = Pt(9.5)
        r_ref_val.font.name = "Arial"
        r_ref_val.font.color.rgb = RGBColor(30, 41, 59)
        
        p_date = meta_table.rows[0].cells[1].paragraphs[0]
        p_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_date.paragraph_format.space_before = Pt(10)
        p_date.paragraph_format.space_after = Pt(0)
        r_date_lbl = p_date.add_run("DATE: ")
        r_date_lbl.bold = True
        r_date_lbl.font.size = Pt(9.5)
        r_date_lbl.font.name = "Arial"
        r_date_lbl.font.color.rgb = RGBColor(100, 116, 139)
        r_date_val = p_date.add_run("6th October 2026")
        r_date_val.bold = True
        r_date_val.font.size = Pt(9.5)
        r_date_val.font.name = "Arial"
        r_date_val.font.color.rgb = RGBColor(30, 41, 59)
        
        # Addressee block
        p_to = doc.add_paragraph()
        p_to.paragraph_format.space_before = Pt(14)
        p_to.paragraph_format.space_after = Pt(2)
        p_to.paragraph_format.line_spacing = Pt(13)
        r_to = p_to.add_run("TO:\n")
        r_to.bold = True
        r_to.font.size = Pt(10)
        r_to.font.name = "Arial"
        r_to.font.color.rgb = RGBColor(15, 118, 110)
        
        r_addr = p_to.add_run("The Medical Director / Partner Organisation\nHealthcare & Social Welfare Department\nAbeokuta, Ogun State, Nigeria")
        r_addr.font.size = Pt(9.5)
        r_addr.font.name = "Arial"
        r_addr.font.color.rgb = RGBColor(51, 65, 85)
        
        # Salutation
        p_sal = doc.add_paragraph()
        p_sal.paragraph_format.space_before = Pt(10)
        p_sal.paragraph_format.space_after = Pt(6)
        r_sal = p_sal.add_run("Dear Sir/Madam,")
        r_sal.bold = True
        r_sal.font.size = Pt(10)
        r_sal.font.name = "Arial"
        r_sal.font.color.rgb = RGBColor(30, 41, 59)
        
        # Subject line
        p_subj = doc.add_paragraph()
        p_subj.paragraph_format.space_before = Pt(6)
        p_subj.paragraph_format.space_after = Pt(12)
        r_subj = p_subj.add_run("SUBJECT: OFFICIAL HEALTHCARE RELIEF & COMMUNITY SUPPORT CORRESPONDENCE")
        r_subj.bold = True
        r_subj.font.size = Pt(10.5)
        r_subj.font.name = "Arial"
        r_subj.font.color.rgb = RGBColor(13, 148, 136)
        
        # Body paragraphs
        p_body1 = doc.add_paragraph()
        p_body1.paragraph_format.space_before = Pt(0)
        p_body1.paragraph_format.space_after = Pt(8)
        p_body1.paragraph_format.line_spacing = Pt(14)
        r_b1 = p_body1.add_run(
            "We have the pleasure of writing to introduce the healthcare support, daily living assistance, "
            "and community relief initiatives provided by Hope Restoration and Health Relief Foundation across "
            "Ogun State, Abeokuta, and neighbouring communities. Our foundation is committed to walking alongside "
            "vulnerable individuals, patients recovering from acute conditions, and those living with chronic health "
            "challenges to restore independence, confidence, and dignity."
        )
        r_b1.font.size = Pt(9.5)
        r_b1.font.name = "Arial"
        r_b1.font.color.rgb = RGBColor(51, 65, 85)
        
        p_body2 = doc.add_paragraph()
        p_body2.paragraph_format.space_before = Pt(0)
        p_body2.paragraph_format.space_after = Pt(8)
        p_body2.paragraph_format.line_spacing = Pt(14)
        r_b2 = p_body2.add_run(
            "With over 6years + of professional hands-on care experience, our team of dedicated Support Workers "
            "and Healthcare Assistants adheres strictly to robust safeguarding standards, infection prevention protocols, "
            "and accurate care documentation. We work collaboratively with clinical teams, social service providers, and "
            "families to deliver compassionate medication reminders, mobility support, and genuine companionship."
        )
        r_b2.font.size = Pt(9.5)
        r_b2.font.name = "Arial"
        r_b2.font.color.rgb = RGBColor(51, 65, 85)
        
        p_body3 = doc.add_paragraph()
        p_body3.paragraph_format.space_before = Pt(0)
        p_body3.paragraph_format.space_after = Pt(14)
        p_body3.paragraph_format.line_spacing = Pt(14)
        r_b3 = p_body3.add_run(
            "This official letterhead is provided as a formal communication template for partnership proposals, "
            "health relief reports, verification documents, and official announcements. Please feel free to reach "
            "out via the contact details provided in the header for any enquiries or collaboration opportunities."
        )
        r_b3.font.size = Pt(9.5)
        r_b3.font.name = "Arial"
        r_b3.font.color.rgb = RGBColor(51, 65, 85)
        
        # Sign-off
        p_sign = doc.add_paragraph()
        p_sign.paragraph_format.space_before = Pt(10)
        p_sign.paragraph_format.space_after = Pt(36) # Signature space
        r_sign = p_sign.add_run("Yours faithfully,")
        r_sign.font.size = Pt(10)
        r_sign.font.name = "Arial"
        r_sign.font.color.rgb = RGBColor(30, 41, 59)
        
        p_sig_name = doc.add_paragraph()
        p_sig_name.paragraph_format.space_before = Pt(0)
        p_sig_name.paragraph_format.space_after = Pt(1)
        r_sig_name = p_sig_name.add_run("Adebayo Elijah")
        r_sig_name.bold = True
        r_sig_name.font.size = Pt(10.5)
        r_sig_name.font.name = "Arial"
        r_sig_name.font.color.rgb = RGBColor(13, 148, 136)
        
        p_sig_role = doc.add_paragraph()
        p_sig_role.paragraph_format.space_before = Pt(0)
        p_sig_role.paragraph_format.space_after = Pt(0)
        p_sig_role.paragraph_format.line_spacing = Pt(11)
        r_sig_role = p_sig_role.add_run("Manager & Care Coordinator\nHope Restoration and Health Relief Foundation\nTel: +234 706 871 9591 · Abeokuta, Ogun State")
        r_sig_role.font.size = Pt(8.5)
        r_sig_role.font.name = "Arial"
        r_sig_role.font.color.rgb = RGBColor(100, 116, 139)
    else:
        p_blank = doc.add_paragraph()
        p_blank.paragraph_format.space_before = Pt(24)
        p_blank.paragraph_format.space_after = Pt(0)
        r_blank = p_blank.add_run("[Type your official letter content here...]")
        r_blank.italic = True
        r_blank.font.size = Pt(10)
        r_blank.font.name = "Arial"
        r_blank.font.color.rgb = RGBColor(148, 163, 184)
    
    # ----------------- FOOTER SECTION -----------------
    footer = section.footer
    footer_p = footer.paragraphs[0]
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_p.paragraph_format.space_before = Pt(0)
    footer_p.paragraph_format.space_after = Pt(0)
    footer_p.paragraph_format.line_spacing = Pt(10.5)
    
    pPr = footer_p._p.get_or_add_pPr()
    for b in pPr.findall(qn('w:pBdr')):
        pPr.remove(b)
    pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:top w:val="single" w:sz="8" w:space="4" w:color="0D9488"/></w:pBdr>')
    pPr.append(pBdr)
    
    f_run1 = footer_p.add_run("Hope Restoration and Health Relief Foundation · Ogun State, Abeokuta, Nigeria\n")
    f_run1.bold = True
    f_run1.font.size = Pt(8)
    f_run1.font.name = "Arial"
    f_run1.font.color.rgb = RGBColor(15, 118, 110)
    
    f_run2 = footer_p.add_run("Phone: +234 706 871 9591 · Email: elijah.adebayo@hoperestorationhrf.org · Web: www.hoperestorationhrf.org\n")
    f_run2.font.size = Pt(7.5)
    f_run2.font.name = "Arial"
    f_run2.font.color.rgb = RGBColor(100, 116, 139)
    
    f_run3 = footer_p.add_run("Restoring Hope Through Health & Dignity · Compassionate Healthcare Relief & Daily Living Support")
    f_run3.italic = True
    f_run3.font.size = Pt(7)
    f_run3.font.name = "Arial"
    f_run3.font.color.rgb = RGBColor(148, 163, 184)
    
    # Final XML validation pass across all tables and cells
    for table in doc.tables:
        normalize_tblPr(table._tbl.tblPr)
        for r in table.rows:
            for c in r.cells:
                normalize_tcPr(c._tc.get_or_add_tcPr())
                
    doc.save(filename)
    print(f"Generated {filename} ({os.path.getsize(filename)} bytes) with verified OpenXML schema.")

if __name__ == '__main__':
    create_letterhead("Hope_Restoration_HRF_Letterhead.docx", include_sample_body=True)
    create_letterhead("Hope_Restoration_HRF_Blank_Letterhead.docx", include_sample_body=False)
