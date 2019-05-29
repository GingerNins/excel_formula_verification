from datetime import datetime
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from docx.styles.style import _ParagraphStyle as ParagraphStyle
from docx.table import _Cell as Cell
from docx.text.run import Run
from __init__ import add_method
from docx.text.paragraph import Paragraph


@add_method(ParagraphStyle)
def add_outline_level(self: ParagraphStyle, level: int):
    """
    Adds a method to the _ParagraphFormat class that adds the outline level attribute to the document's xml.
    python-docx does not have this natively available, therefore a custom function is needed.
    Method is added dynamically to the ParagraphStyle class.
    :param self: _ParagraphFormat Class object
    :param level: Level at which the attribute is added, determines where item will reside in the document's TOC
    :return: N/A
    """
    outline_level = parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="{level}" />')
    self._element.get_or_add_pPr().append(outline_level)


@add_method(Run)
def add_field(self: Run, field: str):
    """
    Adds a Word Field to the Paragraph Run by modifying the document's xml
    This feature is not part of python-docx yet and thus a custom function is needed.
    Method is added dynamically to the Run class.
    :param self: Location (run) in document to add Field
    :param field: Specific field-type to add
    """
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


@add_method(Paragraph)
def add_bottom_border(self: Paragraph):
    """
    Adds a horizontal rule to the bottom of a paragraph.  python-docx doesn't support natively, therefore a
    custom method is needed.  Method is added dynamically to the Paragraph class.
    :param self: paragraph to add border to
    :return: n/a
    """
    # TODO: update later so that val, sz, space and color can be customized
    p = self._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')

    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), 'orange')

    pBdr.append(bottom)
    p.append(pBdr)


@add_method(Cell)
def shade_cell(self: Cell, color: str):
    """
    Shades the given cell with color.  python-docx does not have this option natively.
    This method modifies the xml of the document.  Method is added dynamically to
    the Cell class.
    :param self: cell to add color to
    :param color: str representing color in hex format
    :return: n/a
    """
    cell_color = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color}" />')
    self._tc.get_or_add_tcPr().append(cell_color)





def create_filename(directory: str, description: str, initials: str = 'ELS', extension: str = 'docx') -> str:
    """
    Sets the name of the file based on the current timestamp and directory
    :param directory: Directory where document will be stored
    :param description: Description of the document
    :param initials: Initials of user
    :param extension: File type extension (should be docx or xlsx)
    :return: name of the document as string
    """
    if extension not in ['docx', 'xlsx']:
        raise ValueError("Extension must be docx or xlsx")
    now = datetime.now().strftime
    return f"{directory}{now('%Y%m%d')}_{initials}_{description}_{now('%H%M%S')}.{extension}"


