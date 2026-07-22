import re
import os
import copy
from lxml import etree
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"

PATRON = re.compile(
    r"^(Ilustraci\u00f3n|Tabla)\s+(\d+)\.\s*(.*)", re.IGNORECASE | re.DOTALL
)


def _make_run(rpr_elem, text, space_preserve=False):
    r = OxmlElement("w:r")
    if rpr_elem is not None:
        r.append(copy.deepcopy(rpr_elem))
    t = OxmlElement("w:t")
    if space_preserve:
        t.set(f"{{{XML_NS}}}space", "preserve")
    t.text = text
    r.append(t)
    return r


def _make_seq_field(etiqueta, numero, rpr_elem):
    def _run_fc(tipo):
        r = OxmlElement("w:r")
        if rpr_elem is not None:
            r.append(copy.deepcopy(rpr_elem))
        fc = OxmlElement("w:fldChar")
        fc.set(qn("w:fldCharType"), tipo)
        r.append(fc)
        return r

    return [
        _run_fc("begin"),
        _make_instr(rpr_elem, etiqueta),
        _run_fc("separate"),
        _make_run(rpr_elem, str(numero)),
        _run_fc("end"),
    ]


def _make_instr(rpr_elem, etiqueta):
    r = OxmlElement("w:r")
    if rpr_elem is not None:
        r.append(copy.deepcopy(rpr_elem))
    instr = OxmlElement("w:instrText")
    instr.set(f"{{{XML_NS}}}space", "preserve")
    instr.text = f" SEQ {etiqueta} \\* ARABIC "
    r.append(instr)
    return r


def _reescribir(p_elem, etiqueta, numero, texto_restante):
    rpr_elem = None
    primer_run = p_elem.find(qn("w:r"))
    if primer_run is not None:
        rpr_elem = primer_run.find(qn("w:rPr"))

    for child in list(p_elem):
        if child.tag != qn("w:pPr"):
            p_elem.remove(child)

    p_elem.append(_make_run(rpr_elem, f"{etiqueta} ", space_preserve=True))
    for elem in _make_seq_field(etiqueta, numero, rpr_elem):
        p_elem.append(elem)
    p_elem.append(_make_run(rpr_elem, ". ", space_preserve=True))
    if texto_restante:
        p_elem.append(_make_run(rpr_elem, texto_restante))


def procesar(ruta_entrada, callback=None):
    doc = Document(ruta_entrada)
    modificados = 0

    for p in doc.paragraphs:
        texto = "".join(r.text for r in p.runs).strip()
        m = PATRON.match(texto)
        if not m:
            continue

        tipo_raw = m.group(1)
        numero = int(m.group(2))
        resto = m.group(3).strip()

        tipo = "Ilustración" if "lustr" in tipo_raw.lower() else "Tabla"
        _reescribir(p._element, tipo, numero, resto)
        modificados += 1

        if callback and (numero <= 3 or numero % 100 == 0):
            callback(f"  OK {tipo} {numero}. {resto[:55]}")

    nombre_base, ext = os.path.splitext(ruta_entrada)
    nombre_base = nombre_base.replace("_numerado", "")
    ruta_salida = f"{nombre_base}_final{ext}"
    doc.save(ruta_salida)

    if callback:
        callback(f"{modificados} parrafos reescritos con campos SEQ.")
        callback(f"Archivo guardado: {ruta_salida}")

    return {
        "success": True,
        "modificados": modificados,
        "ruta_salida": ruta_salida,
    }
