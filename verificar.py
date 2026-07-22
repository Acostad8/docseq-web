"""
verificar.py — Verifica la numeración de Ilustraciones y Tablas
================================================================
Detecta:
  - Saltos en la numeración (ej: va de 5 a 7, falta el 6)
  - Duplicados (mismo número dos veces)
  - Párrafos con el patrón antiguo "#." que no fueron procesados
  - Totales encontrados

Uso:
    python verificar.py documento.docx
"""

import sys
import re
import os
from docx import Document

PATRON_NUMERADO = re.compile(r'^(Ilustración|Tabla)\s+(\d+)\.\s*(.*)', re.IGNORECASE | re.DOTALL)
PATRON_SIN_NUM  = re.compile(r'^(Ilustración|Tabla)\s+#\.', re.IGNORECASE)


def verificar(ruta):
    doc = Document(ruta)

    registros = {"Ilustración": [], "Tabla": []}
    huerfanos = []

    for i, p in enumerate(doc.paragraphs):
        texto = "".join(r.text for r in p.runs).strip()

        if PATRON_SIN_NUM.match(texto):
            huerfanos.append((i + 1, texto[:80]))
            continue

        m = PATRON_NUMERADO.match(texto)
        if m:
            tipo  = "Ilustración" if "lustr" in m.group(1).lower() else "Tabla"
            num   = int(m.group(2))
            resto = m.group(3).strip()[:60]
            registros[tipo].append((num, resto, i + 1))

    print("\n" + "=" * 60)
    print("REPORTE DE VERIFICACIÓN")
    print("=" * 60)

    errores_totales = 0

    for tipo, lista in registros.items():
        print(f"\n── {tipo}s ──────────────────────────────────────")
        print(f"   Total encontradas: {len(lista)}")

        if not lista:
            continue

        numeros = [n for n, _, _ in lista]

        # Duplicados
        vistos = {}
        for num, texto, linea in lista:
            vistos.setdefault(num, []).append((texto, linea))
        duplicados = {n: v for n, v in vistos.items() if len(v) > 1}

        if duplicados:
            print(f"\n   ⚠  DUPLICADOS ({len(duplicados)}):")
            for num, ocurrencias in duplicados.items():
                print(f"      #{num} aparece {len(ocurrencias)} veces:")
                for texto, linea in ocurrencias:
                    print(f"         Párrafo ~{linea}: {texto}")
            errores_totales += len(duplicados)
        else:
            print("   ✔  Sin duplicados")

        # Saltos
        esperado = list(range(1, len(lista) + 1))
        saltos = []
        prev = 0
        for num, texto, linea in sorted(lista, key=lambda x: x[0]):
            if num != prev + 1:
                saltos.append((prev + 1, num - 1, linea))
            prev = num

        if saltos:
            print(f"\n   ⚠  SALTOS EN NUMERACIÓN ({len(saltos)}):")
            for desde, hasta, linea in saltos:
                rango = f"{desde}" if desde == hasta else f"{desde}–{hasta}"
                print(f"      Faltan números {rango} (detectado cerca del párrafo ~{linea})")
            errores_totales += len(saltos)
        else:
            print("   ✔  Numeración continua sin saltos")

        # Primeros y últimos
        ordenados = sorted(lista, key=lambda x: x[0])
        print(f"\n   Primeros 3:")
        for num, texto, _ in ordenados[:3]:
            print(f"      {tipo} {num}. {texto}")
        print(f"   Últimos 3:")
        for num, texto, _ in ordenados[-3:]:
            print(f"      {tipo} {num}. {texto}")

    # Huérfanos (sin numerar)
    print(f"\n── Sin numerar (patrón '#.' encontrado) ────────────")
    if huerfanos:
        print(f"   ⚠  {len(huerfanos)} párrafo(s) sin procesar:")
        for linea, texto in huerfanos[:10]:
            print(f"      Párrafo ~{linea}: {texto}")
        if len(huerfanos) > 10:
            print(f"      ... y {len(huerfanos) - 10} más")
        errores_totales += len(huerfanos)
    else:
        print("   ✔  Ninguno — todos fueron procesados")

    print("\n" + "=" * 60)
    if errores_totales == 0:
        print("✅ Todo correcto. El documento está listo.")
    else:
        print(f"⚠  Se encontraron {errores_totales} problema(s). Revisa los detalles arriba.")
    print("=" * 60 + "\n")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python verificar.py documento.docx")
        sys.exit(1)
    ruta = sys.argv[1]
    if not os.path.exists(ruta):
        print(f"Archivo no encontrado: {ruta}")
        sys.exit(1)
    verificar(ruta)
