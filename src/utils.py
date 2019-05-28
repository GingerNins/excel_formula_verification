from docx.text.run import Run
from docx.table import _Cell as Cell
from docx.styles.style import _ParagraphStyle as ParagraphStyle
from docx.oxml.ns import nsdecls, qn
from docx.oxml import parse_xml, OxmlElement


def add_outline_level(self: ParagraphStyle, level: int):
    """
    Adds a method to the _ParagraphFormat class that adds the outline level attribute to the document's xml.
    python-docx does not have this natively available, therefore a custom function is needed.
    :param self: _ParagraphFormat Class object
    :param level: Level at which the attribute is added, determines where item will reside in the document's TOC
    :return: N/A
    """
    assert isinstance(self, ParagraphStyle)
    outline_level = parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="{level}" />')
    self._element.get_or_add_pPr().append(outline_level)


# Note: Is this the appropriate place to put this?
ParagraphStyle.add_outline_level = add_outline_level


def add_field(self: Run, field: str):
    """
    Adds a Word Field to the Paragraph Run by modifying the document's xml
    This feature is not part of docx yet and thus a custom function is needed

    :param self: Location (run) in document to add Field
    :param field: Specific field-type to add
    """
    assert isinstance(self, Run)
    fld_char_begin = OxmlElement('w:fldChar')
    fld_char_begin.set(qn('w:fldCharType'), 'begin')
    instr_text = OxmlElement('w:instrText')
    instr_text.set(qn('xml:space'), 'preserve')
    instr_text.text = field

    fld_char_separate = OxmlElement('w:fldChar')
    fld_char_separate.set(qn('w:fldCharType'), 'separate')
    fld_char_text = OxmlElement('w:t')
    fld_char_text.text = 'Right-click to update field.'

    fld_char_end = OxmlElement('w:fldChar')
    fld_char_end.set(qn('w:fldCharType'), 'end')

    r_element = self._r
    r_element.append(fld_char_begin)
    r_element.append(instr_text)
    r_element.append(fld_char_separate)
    r_element.append(fld_char_text)
    r_element.append(fld_char_end)


Run.add_field = add_field


def shade_cell(self: Cell, color: str):
    """
    Shades the given cell with color.  openpyxl does not have this option natively.
    This method modifies the xml of the document.
    :param self: cell to add color to
    :param color: str representing color in hex format
    :return: n/a
    """
    assert isinstance(self, Cell)
    cell_color = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color}" />')
    self._tc.get_or_add_tcPr().append(cell_color)


Cell.shade_cell = shade_cell



