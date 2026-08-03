import re
import os
import copy
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from core.parrafos import iterar_parrafos

XML_NS = "http://www.w3.org/XML/1998/namespace"
PATRON = re.compile(
    r"^(Ilustraci\u00f3n|Tabla)\s+\d+\.\s*(.*)", re.IGNORECASE | re.DOTALL
)
PATRON_HASH = re.compile(r"^(Ilustraci\u00f3n|Tabla)\s+#\.", re.IGNORECASE)


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

    r_instr = OxmlElement("w:r")
    if rpr_elem is not None:
        r_instr.append(copy.deepcopy(rpr_elem))
    instr = OxmlElement("w:instrText")
    instr.set(f"{{{XML_NS}}}space", "preserve")
    instr.text = f" SEQ {etiqueta} \\* ARABIC "
    r_instr.append(instr)

    return [
        _run_fc("begin"),
        r_instr,
        _run_fc("separate"),
        _make_run(rpr_elem, str(numero)),
        _run_fc("end"),
    ]


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


def renumerar(ruta_entrada, callback=None):
    doc = Document(ruta_entrada)
    contadores = {"Ilustración": 0, "Tabla": 0}
    modificados = 0
    advertencias = []

    for i, p in enumerate(iterar_parrafos(doc)):
        texto = "".join(r.text for r in p.runs).strip()

        if PATRON_HASH.match(texto):
            advertencias.append(
                f"  Parrafo ~{i+1} tiene '#.' sin procesar: {texto[:60]}"
            )
            continue

        m = PATRON.match(texto)
        if not m:
            continue

        tipo = "Ilustración" if "lustr" in m.group(1).lower() else "Tabla"
        resto = m.group(2).strip()

        contadores[tipo] += 1
        nuevo_num = contadores[tipo]

        _reescribir(p._element, tipo, nuevo_num, resto)
        modificados += 1

        if callback and (nuevo_num <= 3 or nuevo_num % 100 == 0):
            callback(f"  OK {tipo} {nuevo_num}. {resto[:55]}")

    if callback:
        callback(
            f"Ilustraciones renumeradas: {contadores['Ilustración']}"
        )
        callback(f"Tablas renumeradas: {contadores['Tabla']}")

    if advertencias:
        if callback:
            for w in advertencias:
                callback(f"  [ADVERTENCIA] {w}")

    nombre_base, ext = os.path.splitext(ruta_entrada)
    for sufijo in ["_final", "_numerado", "_renumerado"]:
        nombre_base = nombre_base.replace(sufijo, "")
    ruta_salida = f"{nombre_base}_renumerado{ext}"
    doc.save(ruta_salida)

    if callback:
        callback(f"Archivo guardado: {ruta_salida}")

    return {
        "success": True,
        "contadores": contadores,
        "modificados": modificados,
        "advertencias": advertencias,
        "ruta_salida": ruta_salida,
    }
