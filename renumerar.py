"""
renumerar.py — Recalcula toda la numeración de Ilustraciones y Tablas
======================================================================
Útil cuando:
  - Agregaste o eliminaste ilustraciones/tablas en el medio del documento
  - La numeración tiene saltos o duplicados
  - Quieres reiniciar la numeración desde 1

Lee el documento en orden, reasigna números correlativos (1, 2, 3...)
y reescribe los campos SEQ correctamente.

Funciona sobre documentos que ya tienen "Ilustración N." o "Tabla N."
(salida de los scripts anteriores). NO funciona sobre "#."

Uso:
    python renumerar.py documento_final.docx

Salida:
    documento_renumerado.docx
"""

import sys
import re
import os
import copy
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

XML_NS  = "http://www.w3.org/XML/1998/namespace"
PATRON  = re.compile(r'^(Ilustración|Tabla)\s+\d+\.\s*(.*)', re.IGNORECASE | re.DOTALL)
PATRON_HASH = re.compile(r'^(Ilustración|Tabla)\s+#\.', re.IGNORECASE)


def make_run(rpr_elem, text, space_preserve=False):
    r = OxmlElement("w:r")
    if rpr_elem is not None:
        r.append(copy.deepcopy(rpr_elem))
    t = OxmlElement("w:t")
    if space_preserve:
        t.set(f"{{{XML_NS}}}space", "preserve")
    t.text = text
    r.append(t)
    return r


def make_seq_field(etiqueta, numero, rpr_elem):
    def run_fc(tipo):
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
        run_fc("begin"),
        r_instr,
        run_fc("separate"),
        make_run(rpr_elem, str(numero)),
        run_fc("end"),
    ]


def reescribir_parrafo(p_elem, etiqueta, numero, texto_restante):
    rpr_elem = None
    primer_run = p_elem.find(qn("w:r"))
    if primer_run is not None:
        rpr_elem = primer_run.find(qn("w:rPr"))

    for child in list(p_elem):
        if child.tag != qn("w:pPr"):
            p_elem.remove(child)

    p_elem.append(make_run(rpr_elem, f"{etiqueta} ", space_preserve=True))
    for elem in make_seq_field(etiqueta, numero, rpr_elem):
        p_elem.append(elem)
    p_elem.append(make_run(rpr_elem, ". ", space_preserve=True))
    if texto_restante:
        p_elem.append(make_run(rpr_elem, texto_restante))


def renumerar(ruta_entrada):
    doc = Document(ruta_entrada)
    contadores = {"Ilustración": 0, "Tabla": 0}
    modificados = 0
    advertencias = []

    print(f"\nProcesando: {ruta_entrada}")

    for i, p in enumerate(doc.paragraphs):
        texto = "".join(r.text for r in p.runs).strip()

        if PATRON_HASH.match(texto):
            advertencias.append(f"  Párrafo ~{i+1} tiene '#.' sin procesar: {texto[:60]}")
            continue

        m = PATRON.match(texto)
        if not m:
            continue

        tipo  = "Ilustración" if "lustr" in m.group(1).lower() else "Tabla"
        resto = m.group(2).strip()

        contadores[tipo] += 1
        nuevo_num = contadores[tipo]

        reescribir_parrafo(p._element, tipo, nuevo_num, resto)
        modificados += 1

        if nuevo_num <= 3 or nuevo_num % 100 == 0:
            print(f"  OK {tipo} {nuevo_num}. {resto[:55]}")

    print(f"\nResultado:")
    print(f"  Ilustraciones renumeradas: {contadores['Ilustración']}")
    print(f"  Tablas renumeradas:        {contadores['Tabla']}")
    print(f"  Total modificados:         {modificados}")

    if advertencias:
        print(f"\n  ADVERTENCIAS ({len(advertencias)}):")
        for a in advertencias:
            print(a)

    nombre_base, ext = os.path.splitext(ruta_entrada)
    for sufijo in ["_final", "_numerado", "_renumerado"]:
        nombre_base = nombre_base.replace(sufijo, "")
    ruta_salida = f"{nombre_base}_renumerado{ext}"
    doc.save(ruta_salida)

    print(f"\nArchivo guardado: {ruta_salida}")
    print("Recuerda abrir en Word y presionar Ctrl+A -> F9 para actualizar campos.\n")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python renumerar.py documento_final.docx")
        sys.exit(1)
    ruta = sys.argv[1]
    if not os.path.exists(ruta):
        print(f"Archivo no encontrado: {ruta}")
        sys.exit(1)
    renumerar(ruta)
