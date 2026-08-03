from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml.ns import qn


def _iter_block(parent_elem, parent_obj):
    for child in parent_elem.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, parent_obj)
        elif child.tag == qn("w:tbl"):
            table = Table(child, parent_obj)
            for row in table.rows:
                for cell in row.cells:
                    for p in _iter_block(cell._tc, cell):
                        yield p


def iterar_parrafos(doc):
    return _iter_block(doc.element.body, doc)