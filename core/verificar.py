import re
from docx import Document

from core.parrafos import iterar_parrafos

PATRON_NUMERADO = re.compile(
    r"^(Ilustraci\u00f3n|Tabla)\s+(\d+)\.\s*(.*)", re.IGNORECASE | re.DOTALL
)
PATRON_SIN_NUM = re.compile(r"^(Ilustraci\u00f3n|Tabla)\s+#\.", re.IGNORECASE)


def verificar(ruta_entrada, callback=None):
    doc = Document(ruta_entrada)

    registros = {"Ilustración": [], "Tabla": []}
    huerfanos = []

    for i, p in enumerate(iterar_parrafos(doc)):
        texto = "".join(r.text for r in p.runs).strip()

        if PATRON_SIN_NUM.match(texto):
            huerfanos.append((i + 1, texto[:80]))
            continue

        m = PATRON_NUMERADO.match(texto)
        if m:
            tipo = "Ilustración" if "lustr" in m.group(1).lower() else "Tabla"
            num = int(m.group(2))
            resto = m.group(3).strip()[:60]
            registros[tipo].append((num, resto, i + 1))

    errores_totales = 0
    resultado = {}

    for tipo, lista in registros.items():
        info = {"total": len(lista), "duplicados": [], "saltos": []}
        resultado[tipo] = info

        if callback:
            callback(f"\n--- {tipo}s ---")
            callback(f"  Total encontradas: {len(lista)}")

        if not lista:
            continue

        if callback:
            callback(f"  Primeros:")
            for num, texto, _ in sorted(lista, key=lambda x: x[0])[:3]:
                callback(f"    {tipo} {num}. {texto}")
            callback(f"  Ultimos:")
            for num, texto, _ in sorted(lista, key=lambda x: x[0])[-3:]:
                callback(f"    {tipo} {num}. {texto}")

        vistos = {}
        for num, texto, linea in lista:
            vistos.setdefault(num, []).append((texto, linea))
        duplicados = {n: v for n, v in vistos.items() if len(v) > 1}

        if duplicados:
            info["duplicados"] = [
                {"numero": n, "ocurrencias": len(ocs), "detalles": ocs}
                for n, ocs in duplicados.items()
            ]
            errores_totales += len(duplicados)
            if callback:
                callback(f"  [ERROR] {len(duplicados)} numero(s) duplicados")
                for n, ocs in duplicados.items():
                    callback(f"    #{n} aparece {len(ocs)} veces")
        else:
            if callback:
                callback(f"  Sin duplicados")

        prev = 0
        saltos = []
        for num, texto, linea in sorted(lista, key=lambda x: x[0]):
            if num > prev + 1:
                saltos.append((prev + 1, num - 1, linea))
            if num > prev:
                prev = num

        if saltos:
            info["saltos"] = saltos
            errores_totales += len(saltos)
            if callback:
                callback(f"  [ERROR] {len(saltos)} salto(s) en la numeracion")
                for desde, hasta, linea in saltos:
                    rango = f"{desde}" if desde == hasta else f"{desde}-{hasta}"
                    callback(f"    Faltan numeros {rango} (parrafo ~{linea})")
        else:
            if callback:
                callback(f"  Numeracion continua sin saltos")

    resultado["huerfanos"] = [
        {"linea": linea, "texto": texto} for linea, texto in huerfanos
    ]
    if huerfanos:
        errores_totales += len(huerfanos)
        if callback:
            callback(f"\n--- Sin numerar (patron '#.') ---")
            callback(f"  [ERROR] {len(huerfanos)} parrafo(s) sin procesar")
            for linea, texto in huerfanos[:10]:
                callback(f"    Parrafo ~{linea}: {texto}")
    else:
        if callback:
            callback(f"\n  Todos los parrafos fueron procesados")

    if callback:
        if errores_totales == 0:
            callback("  Todo correcto. El documento esta listo.")
        else:
            callback(f"  Se encontraron {errores_totales} problema(s).")

    return {
        "success": errores_totales == 0,
        "errores_totales": errores_totales,
        "detalle": resultado,
    }
