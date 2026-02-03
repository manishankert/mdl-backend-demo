# utils/docx_utils.py
from docx.shared import Pt, Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.table import Table


def shade_cell(cell, hex_fill="E7E6E6"):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)


def set_col_widths(table: Table, widths):
    for col_idx, w in enumerate(widths):
        for cell in table.columns[col_idx].cells:
            cell.width = w


def tight_paragraph(p: Paragraph):
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.0


def as_oxml(el):
    """Return underlying oxml element for Paragraph/Table/raw CT_* safely."""
    if hasattr(el, "_p"):   # Paragraph
        return el._p
    if hasattr(el, "_tbl"):  # Table
        return el._tbl
    if hasattr(el, "_element"):
        return el._element
    return el  # assume already oxml


def insert_after(anchor, new_block):
    """Insert new_block (Paragraph/Table or raw oxml) after anchor (Paragraph/Table or raw oxml)."""
    a = as_oxml(anchor)
    n = as_oxml(new_block)
    a.addnext(n)


def apply_grid_borders(tbl: Table):
    """Ensure visible borders regardless of style availability."""
    tbl_el = tbl._tbl
    tblPr = tbl_el.tblPr or tbl_el.get_or_add_tblPr()
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{side}")
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), "6")     # 0.5pt
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), "auto")
        borders.append(e)
    tblPr.append(borders)


def remove_paragraph(p):
    # safe remove of a docx paragraph
    p._element.getparent().remove(p._element)


def clear_runs(p: Paragraph):
    for r in list(p.runs):
        r._element.getparent().remove(r._element)


def para_text(p: Paragraph) -> str:
    return "".join(run.text for run in p.runs)


def rewrite_para_text(p, new_text: str):
    """Clear all runs in a paragraph and set to new_text."""
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    p.add_run(new_text)


def set_table_bold_borders(tbl, size=12, color="000000"):
    tblPr = tbl._tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl._tbl.insert(0, tblPr)

    tblBorders = tblPr.find(qn("w:tblBorders"))
    if tblBorders is None:
        tblBorders = OxmlElement("w:tblBorders")
        tblPr.append(tblBorders)

    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = tblBorders.find(qn(f"w:{edge}"))
        if border is None:
            border = OxmlElement(f"w:{edge}")
            tblBorders.append(border)

        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), str(size))   # "bold"
        border.set(qn("w:color"), color)


def set_table_cell_margins(tbl, top_in=0.06, bottom_in=0.06, left_in=0.06, right_in=0.06):
    """Adds internal cell padding for the whole table (Word: tblCellMar)."""
    def twips(inches: float) -> str:
        return str(int(round(inches * 1440)))

    tblPr = tbl._tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl._tbl.insert(0, tblPr)

    cellMar = tblPr.find(qn("w:tblCellMar"))
    if cellMar is None:
        cellMar = OxmlElement("w:tblCellMar")
        tblPr.append(cellMar)

    for side, val in (("top", top_in), ("bottom", bottom_in), ("left", left_in), ("right", right_in)):
        node = cellMar.find(qn(f"w:{side}"))
        if node is None:
            node = OxmlElement(f"w:{side}")
            cellMar.append(node)
        node.set(qn("w:w"), twips(val))
        node.set(qn("w:type"), "dxa")


def twips_from_inches(inches: float) -> int:
    return int(round(inches * 1440))


def set_table_preferred_width_and_indent(table, width_in=6.25, indent_in=0.05):
    # Disable autofit so Word respects widths
    table.autofit = False

    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)

    # Preferred table width
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"), str(twips_from_inches(width_in)))

    # Table indent from left
    tblInd = tblPr.find(qn("w:tblInd"))
    if tblInd is None:
        tblInd = OxmlElement("w:tblInd")
        tblPr.append(tblInd)
    tblInd.set(qn("w:type"), "dxa")
    tblInd.set(qn("w:w"), str(twips_from_inches(indent_in)))


def set_row_height_and_allow_break(row, height_in=0.48, allow_break_across_pages=True):
    tr = row._tr
    trPr = tr.get_or_add_trPr()

    trHeight = trPr.find(qn("w:trHeight"))
    if trHeight is None:
        trHeight = OxmlElement("w:trHeight")
        trPr.append(trHeight)

    trHeight.set(qn("w:val"), str(twips_from_inches(height_in)))
    trHeight.set(qn("w:hRule"), "atLeast")  # allows taller rows when needed

    cantSplit = trPr.find(qn("w:cantSplit"))
    if allow_break_across_pages:
        if cantSplit is not None:
            trPr.remove(cantSplit)
    else:
        if cantSplit is None:
            trPr.append(OxmlElement("w:cantSplit"))

    # Row height (exact)
    trHeight = trPr.find(qn("w:trHeight"))
    if trHeight is None:
        trHeight = OxmlElement("w:trHeight")
        trPr.append(trHeight)
    trHeight.set(qn("w:val"), str(twips_from_inches(height_in)))
    trHeight.set(qn("w:hRule"), "atLeast")

    # Allow row to break across pages:
    # Word uses <w:cantSplit/> to PREVENT breaking. So remove it if present.
    cantSplit = trPr.find(qn("w:cantSplit"))
    if allow_break_across_pages:
        if cantSplit is not None:
            trPr.remove(cantSplit)
    else:
        if cantSplit is None:
            cantSplit = OxmlElement("w:cantSplit")
            trPr.append(cantSplit)


def set_cell_preferred_width(cell, width_in: float):
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = tcPr.find(qn("w:tcW"))
    if tcW is None:
        tcW = OxmlElement("w:tcW")
        tcPr.append(tcW)
    tcW.set(qn("w:type"), "dxa")
    tcW.set(qn("w:w"), str(twips_from_inches(width_in)))


def set_table_column_widths(table, col_widths_in):
    # col_widths_in: list[float] length == number of cols
    for row in table.rows:
        for i, w in enumerate(col_widths_in):
            # python-docx visible width
            row.cells[i].width = Inches(w)
            # Word preferred width
            set_cell_preferred_width(row.cells[i], w)


def set_cell_paragraph_spacing_before(cell, before_pt: float):
    # Apply spacing-before to ALL paragraphs in a cell.
    for p in cell.paragraphs:
        p.paragraph_format.space_before = Pt(before_pt)


def apply_program_table_spacing(tbl):
    # Header row spacing-before = 3.8pt for all header cells
    header_row = tbl.rows[0]
    for cell in header_row.cells:
        set_cell_paragraph_spacing_before(cell, 3.8)

    # Subsequent rows: per-column spacing-before
    col_before_pts = [10.0, 0.0, 10.0, 3.8, 10.0]  # cols 1..5

    for r_i in range(1, len(tbl.rows)):
        row = tbl.rows[r_i]
        for c_i, before_pt in enumerate(col_before_pts):
            set_cell_paragraph_spacing_before(row.cells[c_i], before_pt)


def add_hyperlink(paragraph, text, url, bold=False, font_pt=12):
    """
    Add a clickable hyperlink to a paragraph.
    """
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")

    # Style like a normal hyperlink (blue + underline)
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0000FF")
    rPr.append(color)

    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rPr.append(u)

    if bold:
        b = OxmlElement("w:b")
        rPr.append(b)

    # FORCE FONT SIZE for hyperlink run
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(font_pt * 2)))
    rPr.append(sz)

    szCs = OxmlElement("w:szCs")
    szCs.set(qn("w:val"), str(int(font_pt * 2)))
    rPr.append(szCs)

    run.append(rPr)

    t = OxmlElement("w:t")
    t.text = text
    run.append(t)

    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def create_hyperlink_element(paragraph, url, text):
    """
    Add a hyperlink element to a paragraph.
    Returns the hyperlink OxmlElement (does NOT append it).
    """
    # Get relationship ID
    part = paragraph.part
    r_id = part.relate_to(url, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', is_external=True)

    # Create hyperlink element
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)

    # Create run for hyperlink text
    run = OxmlElement('w:r')

    # Run properties (blue + underline)
    rPr = OxmlElement('w:rPr')

    # Blue color
    color = OxmlElement('w:color')
    color.set(qn('w:val'), '0563C1')
    rPr.append(color)

    # Underline
    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')
    rPr.append(u)

    run.append(rPr)

    # Add text
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    run.append(t)

    hyperlink.append(run)

    return hyperlink
