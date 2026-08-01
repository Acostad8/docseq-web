import re
import os
from docx import Document

PREFIJOS = {
    "Ilustración": "Título de ilustración",
    "Tabla": "Título de tabla",
}
PATRON = re.compile(r"^(Ilustraci\u00f3n|Tabla)\s+#\.\s*", re.IGNORECASE)


def _texto_completo(paragraph):
    return "".join(run.text for run in paragraph.runs)


def _aplicar_estilo(paragraph, style_name, doc):
    try:
        paragraph.style = doc.styles[style_name]
    except KeyError:
        try:
            base = doc.styles["Caption"]
        except KeyError:
            base = doc.styles["Normal"]
        new_style = doc.styles.add_style(style_name, base.type)
        new_style.base_style = base
        paragraph.style = new_style


def _reemplazar_texto(paragraph, prefijo_nuevo):
    texto = _texto_completo(paragraph)
    match = PATRON.match(texto)
    if not match:
        return False
    prefijo_viejo = match.group(0)
    texto_nuevo = prefijo_nuevo + texto[len(prefijo_viejo):]
    runs = paragraph.runs
    if not runs:
        return False
    runs[0].text = texto_nuevo
    for run in runs[1:]:
        run.text = ""
    return True


def procesar(ruta_entrada, callback=None):
    doc = Document(ruta_entrada)
    contadores = {k: 0 for k in PREFIJOS}
    pendientes = []

    if callback:
        callback("Escaneando documento...")

    for paragraph in doc.paragraphs:
        texto = _texto_completo(paragraph)
        m = PATRON.match(texto)
        if m:
            tipo = next(k for k in PREFIJOS if k.lower() == m.group(1).lower())
            contadores[tipo] += 1
            paragraph._nuevo_prefijo = f"{tipo} {contadores[tipo]}. "
            paragraph._tipo = tipo
            pendientes.append(paragraph)

    if not pendientes:
        return {
            "success": False,
            "error": "No se encontraron patrones 'Ilustración #.' o 'Tabla #.'",
            "contadores": contadores,
            "ruta_salida": None,
        }

    if callback:
        for tipo, cuenta in contadores.items():
            callback(f"  {tipo}s: {cuenta}")
        callback("Aplicando numeración y estilos...")

    errores = 0
    for paragraph in pendientes:
        try:
            _reemplazar_texto(paragraph, paragraph._nuevo_prefijo)
            _aplicar_estilo(paragraph, PREFIJOS[paragraph._tipo], doc)
        except Exception as e:
            errores += 1
            if callback:
                callback(f"  [ERROR] {e}")

    nombre_base, ext = os.path.splitext(ruta_entrada)
    ruta_salida = f"{nombre_base}_numerado{ext}"
    doc.save(ruta_salida)

    if callback:
        callback(f"Archivo guardado: {ruta_salida}")

    return {
        "success": True,
        "contadores": contadores,
        "modificados": len(pendientes),
        "errores": errores,
        "ruta_salida": ruta_salida,
    }
