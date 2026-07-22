"""
Numerador automático de Ilustraciones y Tablas para documentos Word (.docx)
===========================================================================
Reemplaza:
    "Ilustración #. ..." → "Ilustración 1. ..."
    "Tabla #. ..."       → "Tabla 1. ..."

Y aplica los estilos Word "Título de ilustración" / "Título de tabla"
para que puedas generar la Tabla de ilustraciones y Tabla de tablas
desde Referencias > Insertar tabla de ilustraciones.

Instalación:
    pip install python-docx

Uso:
    python numerar_ilustraciones_tablas.py documento.docx

Salida:
    documento_numerado.docx  (el original NO se modifica)
"""

import sys
import re
from copy import deepcopy
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

# ── Configuración ──────────────────────────────────────────────────────────────
PREFIJOS = {
    "Ilustración": "Título de ilustración",   # estilo Word en español
    "Tabla":       "Título de tabla",         # estilo Word en español
}
PATRON = re.compile(r'^(Ilustración|Tabla)\s+#\.\s*', re.IGNORECASE)
# ───────────────────────────────────────────────────────────────────────────────


def get_paragraph_full_text(paragraph):
    """Devuelve el texto completo de un párrafo concatenando todos los runs."""
    return "".join(run.text for run in paragraph.runs)


def set_paragraph_style(paragraph, style_name, doc):
    """
    Aplica un estilo al párrafo. Si el estilo no existe en el documento,
    lo crea basado en 'Caption' (estilo base de leyendas en Word).
    """
    try:
        paragraph.style = doc.styles[style_name]
    except KeyError:
        # El estilo no existe → crearlo basado en Caption
        try:
            base = doc.styles["Caption"]
        except KeyError:
            base = doc.styles["Normal"]
        new_style = doc.styles.add_style(style_name, base.type)
        new_style.base_style = base
        paragraph.style = new_style
        print(f"  [INFO] Estilo '{style_name}' creado (basado en Caption).")


def replace_text_preserving_format(paragraph, nuevo_texto):
    """
    Reemplaza el texto del primer run con el nuevo prefijo numerado
    y preserva el resto del contenido (formato, texto siguiente).
    
    Estrategia: modifica el texto del primer run para cambiar solo
    la parte "Ilustración #." → "Ilustración N." manteniendo el resto.
    """
    texto_completo = get_paragraph_full_text(paragraph)
    match = PATRON.match(texto_completo)
    if not match:
        return False

    prefijo_viejo = match.group(0)        # "Ilustración #. "
    prefijo_nuevo = paragraph._nuevo_prefijo  # "Ilustración 42. "

    # Reconstruir texto reemplazando solo el prefijo al inicio
    texto_nuevo_completo = prefijo_nuevo + texto_completo[len(prefijo_viejo):]

    # Vaciar todos los runs y poner el texto en el primero
    runs = paragraph.runs
    if not runs:
        return False

    # Poner todo el texto en el primer run
    runs[0].text = texto_nuevo_completo
    # Limpiar el resto de runs
    for run in runs[1:]:
        run.text = ""

    return True


def procesar_documento(ruta_entrada):
    doc = Document(ruta_entrada)

    contadores = {prefijo: 0 for prefijo in PREFIJOS}
    parrafos_a_modificar = []

    # ── Paso 1: identificar y contar ──────────────────────────────────────────
    print("\n🔍 Escaneando documento...")
    for i, paragraph in enumerate(doc.paragraphs):
        texto = get_paragraph_full_text(paragraph)
        match = PATRON.match(texto)
        if match:
            tipo = next(
                k for k in PREFIJOS
                if k.lower() == match.group(1).lower()
            )
            contadores[tipo] += 1
            nuevo_prefijo = f"{tipo} {contadores[tipo]}. "
            paragraph._nuevo_prefijo = nuevo_prefijo
            paragraph._tipo = tipo
            parrafos_a_modificar.append(paragraph)

    print(f"\n📊 Encontrados:")
    for tipo, cuenta in contadores.items():
        print(f"   • {tipo}s: {cuenta}")

    if not parrafos_a_modificar:
        print("\n⚠️  No se encontró ningún párrafo con el patrón 'Ilustración #.' o 'Tabla #.'")
        print("   Verifica que el texto en tu documento sea exactamente así.")
        return

    # ── Paso 2: aplicar cambios ───────────────────────────────────────────────
    print("\n✏️  Aplicando numeración y estilos...")
    errores = 0
    for paragraph in parrafos_a_modificar:
        try:
            # 1. Reemplazar texto
            replace_text_preserving_format(paragraph)
            # 2. Aplicar estilo de Word
            estilo = PREFIJOS[paragraph._tipo]
            set_paragraph_style(paragraph, estilo, doc)
        except Exception as e:
            errores += 1
            print(f"  [ERROR] Párrafo '{get_paragraph_full_text(paragraph)[:60]}': {e}")

    # ── Paso 3: guardar ───────────────────────────────────────────────────────
    nombre_base, ext = os.path.splitext(ruta_entrada)
    ruta_salida = f"{nombre_base}_numerado{ext}"
    doc.save(ruta_salida)

    print(f"\n✅ Listo. Archivo guardado en:")
    print(f"   {ruta_salida}")
    if errores:
        print(f"\n⚠️  {errores} párrafo(s) con errores (revisa la consola).")
    print("\n📌 Próximos pasos en Word:")
    print("   1. Abre el archivo _numerado.docx")
    print("   2. Ve al lugar donde quieres la Tabla de ilustraciones")
    print("   3. Referencias > Insertar tabla de ilustraciones")
    print("      → Etiqueta de título: 'Ilustración'")
    print("   4. Repite para la Tabla de tablas con etiqueta 'Tabla'")


# ── Bug fix: la función replace_text_preserving_format necesita el atributo ──
def replace_text_preserving_format(paragraph):
    texto_completo = get_paragraph_full_text(paragraph)
    match = PATRON.match(texto_completo)
    if not match:
        return False

    prefijo_viejo = match.group(0)
    prefijo_nuevo = paragraph._nuevo_prefijo
    texto_nuevo_completo = prefijo_nuevo + texto_completo[len(prefijo_viejo):]

    runs = paragraph.runs
    if not runs:
        return False

    runs[0].text = texto_nuevo_completo
    for run in runs[1:]:
        run.text = ""

    return True


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python numerar_ilustraciones_tablas.py ruta/al/documento.docx")
        sys.exit(1)

    ruta = sys.argv[1]
    if not os.path.exists(ruta):
        print(f"❌ Archivo no encontrado: {ruta}")
        sys.exit(1)

    procesar_documento(ruta)
