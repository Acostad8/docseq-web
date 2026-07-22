"""
Inserta campos SEQ de Word en párrafos de Ilustración/Tabla
============================================================
Toma el documento_numerado.docx (ya con estilos aplicados) y reescribe
cada párrafo para que tenga la estructura interna de campo SEQ que Word
necesita para generar la Tabla de ilustraciones y Tabla de tablas.

La estructura que genera es equivalente a cuando Word inserta un título
desde Referencias > Insertar título:

    Ilustración { SEQ Ilustración \* ARABIC } . Texto de la leyenda

Instalación:
    pip install python-docx lxml

Uso:
    python insertar_campos_seq.py documento_numerado.docx

Salida:
    documento_final.docx
"""

import sys
import re
import os
import copy
from lxml import etree
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Namespace XML base
W   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"

PATRON = re.compile(r'^(Ilustraci\u00f3n|Tabla)\s+(\d+)\.\s*(.*)', re.IGNORECASE | re.DOTALL)


def make_run(rpr_elem, text, space_preserve=False):
    """Crea un <w:r> con formato copiado de rpr_elem (puede ser None)."""
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
    """
    Construye la secuencia de elementos XML para un campo SEQ:
        <w:r><w:fldChar begin/></w:r>
        <w:r><w:instrText> SEQ Ilustración \* ARABIC </w:instrText></w:r>
        <w:r><w:fldChar separate/></w:r>
        <w:r><w:t>N</w:t></w:r>
        <w:r><w:fldChar end/></w:r>
    """
    elementos = []

    def run_with_fldChar(tipo):
        r = OxmlElement("w:r")
        if rpr_elem is not None:
            r.append(copy.deepcopy(rpr_elem))
        fc = OxmlElement("w:fldChar")
        fc.set(qn("w:fldCharType"), tipo)
        r.append(fc)
        return r

    elementos.append(run_with_fldChar("begin"))

    r_instr = OxmlElement("w:r")
    if rpr_elem is not None:
        r_instr.append(copy.deepcopy(rpr_elem))
    instr = OxmlElement("w:instrText")
    instr.set(f"{{{XML_NS}}}space", "preserve")
    instr.text = f" SEQ {etiqueta} \\* ARABIC "
    r_instr.append(instr)
    elementos.append(r_instr)

    elementos.append(run_with_fldChar("separate"))
    elementos.append(make_run(rpr_elem, str(numero)))
    elementos.append(run_with_fldChar("end"))

    return elementos


def reescribir_parrafo(p_elem, etiqueta, numero, texto_restante):
    """
    Reescribe el contenido XML del párrafo con estructura:
        "Ilustración " [SEQ] ". texto_restante"
    Preserva <w:pPr> y hereda formato del primer run.
    """
    rpr_elem = None
    primer_run = p_elem.find(qn("w:r"))
    if primer_run is not None:
        rpr_elem = primer_run.find(qn("w:rPr"))

    # Eliminar todos los hijos excepto pPr
    for child in list(p_elem):
        if child.tag != qn("w:pPr"):
            p_elem.remove(child)

    p_elem.append(make_run(rpr_elem, f"{etiqueta} ", space_preserve=True))
    for elem in make_seq_field(etiqueta, numero, rpr_elem):
        p_elem.append(elem)
    p_elem.append(make_run(rpr_elem, ". ", space_preserve=True))
    if texto_restante:
        p_elem.append(make_run(rpr_elem, texto_restante))


def procesar(ruta_entrada):
    doc = Document(ruta_entrada)
    modificados = 0

    print(f"\nProcesando: {ruta_entrada}")

    for p in doc.paragraphs:
        texto = "".join(r.text for r in p.runs).strip()
        m = PATRON.match(texto)
        if not m:
            continue

        tipo_raw = m.group(1)
        numero   = int(m.group(2))
        resto    = m.group(3).strip()

        tipo = "Ilustración" if "lustr" in tipo_raw.lower() else "Tabla"
        reescribir_parrafo(p._element, tipo, numero, resto)
        modificados += 1

        if numero <= 3 or numero % 100 == 0:
            print(f"  OK {tipo} {numero}. {resto[:55]}")

    nombre_base, ext = os.path.splitext(ruta_entrada)
    nombre_base = nombre_base.replace("_numerado", "")
    ruta_salida = f"{nombre_base}_final{ext}"
    doc.save(ruta_salida)

    print(f"\n{modificados} parrafos reescritos con campos SEQ.")
    print(f"Archivo guardado: {ruta_salida}")
    print("""
Pasos en Word:
  1. Abre _final.docx
  2. Ctrl+A  luego  F9   (actualiza todos los campos del documento)
  3. Referencias -> Insertar tabla de ilustraciones
     Etiqueta de titulo: "Ilustración"  -> Aceptar
  4. Repite con etiqueta "Tabla"
""")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python insertar_campos_seq.py documento_numerado.docx")
        sys.exit(1)
    ruta = sys.argv[1]
    if not os.path.exists(ruta):
        print(f"Archivo no encontrado: {ruta}")
        sys.exit(1)
    procesar(ruta)
