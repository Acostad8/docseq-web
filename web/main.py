import os
import tempfile
import shutil
from pathlib import Path
from contextlib import asynccontextmanager
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
import sys
import time

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import core.numerar
import core.insertar_seq
import core.renumerar
import core.verificar

UPLOAD_DIR = Path(tempfile.gettempdir()) / "docseq_web"
UPLOAD_DIR.mkdir(exist_ok=True)

logs = {}


@asynccontextmanager
async def lifespan(app: FastAPI):
    now = time.time()
    for d in UPLOAD_DIR.iterdir():
        if d.is_dir():
            try:
                if now - d.stat().st_mtime > 3600:
                    shutil.rmtree(d, ignore_errors=True)
            except:
                pass
    stale = [sid for sid, msgs in list(logs.items()) if not (UPLOAD_DIR / sid).exists()]
    for sid in stale:
        logs.pop(sid, None)
    yield


app = FastAPI(title="DocSeq Web", lifespan=lifespan)

app.mount("/static", StaticFiles(directory=Path(__file__).parent / "static"), name="static")


@app.get("/", response_class=HTMLResponse)
async def index():
    index_path = Path(__file__).parent / "static" / "index.html"
    return index_path.read_text(encoding="utf-8")


def _make_callback(session_id: str):
    def cb(msg: str):
        logs.setdefault(session_id, []).append(msg)
    return cb


@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    if not (file.filename or "").lower().endswith(".docx"):
        raise HTTPException(400, "Solo archivos .docx")
    session_id = os.urandom(8).hex()
    session_dir = UPLOAD_DIR / session_id
    session_dir.mkdir(parents=True, exist_ok=True)
    dest = session_dir / "original.docx"
    with open(dest, "wb") as f:
        content = await file.read()
        f.write(content)
    logs[session_id] = [f"Archivo subido: {file.filename}"]
    return {"session_id": session_id, "filename": file.filename}


@app.post("/api/procesar")
async def procesar(session_id: str = Form(...), accion: str = Form(...)):
    session_dir = UPLOAD_DIR / session_id
    original = session_dir / "original.docx"
    if not original.exists():
        raise HTTPException(400, "Sesión no encontrada")

    ruta_base = str(original)
    logs[session_id] = []
    cb = _make_callback(session_id)

    try:
        if accion == "numerar":
            cb("Paso 1 — Numerando ilustraciones y tablas...")
            res = core.numerar.procesar(ruta_base, callback=cb)
            if res["ruta_salida"]:
                shutil.copy(res["ruta_salida"], session_dir / "resultado.docx")
            cb("")
            if res.get("success"):
                c = res["contadores"]
                cb("  Resumen:")
                for tipo, cuenta in c.items():
                    cb(f"    {tipo}s:  {cuenta}")
                cb(f"    Modificados:  {res['modificados']}")
            return {"success": res["success"], "resultado": res, "logs": logs[session_id]}

        elif accion == "insertar_seq":
            cb("Paso 2 — Insertando campos SEQ...")
            entrada_seq = session_dir / "resultado.docx"
            if not entrada_seq.exists():
                entrada_seq = original
            res = core.insertar_seq.procesar(str(entrada_seq), callback=cb)
            if res.get("ruta_salida"):
                shutil.copy(res["ruta_salida"], session_dir / "resultado.docx")
            return {"success": res["success"], "resultado": res, "logs": logs[session_id]}

        elif accion == "renumerar":
            cb("Renumerando...")
            entrada_ren = session_dir / "resultado.docx"
            if not entrada_ren.exists():
                entrada_ren = original
            res = core.renumerar.renumerar(str(entrada_ren), callback=cb)
            if res.get("ruta_salida"):
                shutil.copy(res["ruta_salida"], session_dir / "resultado.docx")
            return {"success": res["success"], "resultado": res, "logs": logs[session_id]}

        elif accion == "verificar":
            cb("Verificando...")
            res = core.verificar.verificar(ruta_base, callback=cb)
            return {"success": res["success"], "resultado": res, "logs": logs[session_id]}

        elif accion == "todo":
            cb("═" * 48)
            cb("Proceso completo:  Paso 1  →  Paso 2")
            cb("")
            cb("▶ Paso 1 — Numerar")
            res1 = core.numerar.procesar(ruta_base, callback=cb)
            if not res1.get("success"):
                cb("")
                cb("Paso 1 falló. Se cancela el proceso.")
                c = res1["contadores"]
                if any(v > 0 for v in c.values()):
                    cb("  Resumen parcial:")
                    for tipo, cuenta in c.items():
                        cb(f"    {tipo}s:  {cuenta}")
                return {"success": False, "resultado": res1, "logs": logs[session_id]}

            c = res1["contadores"]
            cb("")
            cb("  Resumen:")
            for tipo, cuenta in c.items():
                cb(f"    {tipo}s:  {cuenta}")
            cb(f"    Modificados:  {res1['modificados']}")
            cb("")
            cb("▶ Paso 2 — Insertar SEQ")
            ruta1 = res1["ruta_salida"]
            res2 = core.insertar_seq.procesar(ruta1, callback=cb)
            if res2.get("ruta_salida"):
                shutil.copy(res2["ruta_salida"], session_dir / "resultado.docx")
            cb("")
            cb("Proceso completo finalizado.")
            cb("Abre el archivo en Word y presiona Ctrl+A → F9.")
            cb("═" * 48)
            return {"success": True, "resultado": res2, "logs": logs[session_id]}

        else:
            raise HTTPException(400, f"Acción desconocida: {accion}")

    except Exception as e:
        cb(f"ERROR: {str(e)}")
        return {"success": False, "error": str(e), "logs": logs[session_id]}


@app.get("/api/download/{session_id}")
async def download(session_id: str):
    session_dir = UPLOAD_DIR / session_id
    resultado = session_dir / "resultado.docx"
    if not resultado.exists():
        raise HTTPException(404, "No hay archivo disponible")
    return FileResponse(
        str(resultado),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename="documento_procesado.docx"
    )


@app.get("/api/logs/{session_id}")
async def get_logs(session_id: str):
    return {"logs": logs.get(session_id, [])}



