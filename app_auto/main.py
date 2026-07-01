import os
import shutil
import asyncio
from contextlib import asynccontextmanager
from concurrent.futures import ThreadPoolExecutor

from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware

from logic import (
    procesar_carpeta,
    read_pdf_text,
    es_albaran,
    extract_invoice_with_agent,
    csv_to_matrix,
    clasificar_factura,
    mover_pdf,
    guardar_historial,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")

_fallback = os.path.join(BASE_DIR, "facturas")
FACTURAS_DIR = os.getenv("FACTURAS_DIR", _fallback).strip('"').strip("'")
CARPETA_ENTRADA = os.path.join(FACTURAS_DIR, "entrada")

# Intervalo de vigilancia en segundos (configurable en .env)
INTERVALO = int(os.getenv("INTERVALO_VIGILANCIA", "30"))

_executor = ThreadPoolExecutor(max_workers=1)
_procesando = False


async def vigilar_entrada():
    global _procesando
    while True:
        await asyncio.sleep(INTERVALO)
        if _procesando:
            continue
        if not os.path.exists(CARPETA_ENTRADA):
            continue
        pdfs = [f for f in os.listdir(CARPETA_ENTRADA) if f.lower().endswith(".pdf")]
        if not pdfs:
            continue
        print(f"[Vigilante] {len(pdfs)} PDF(s) detectados — procesando...")
        _procesando = True
        try:
            loop = asyncio.get_event_loop()
            await loop.run_in_executor(_executor, procesar_carpeta)
        finally:
            _procesando = False


@asynccontextmanager
async def lifespan(_app: FastAPI):
    tarea = asyncio.create_task(vigilar_entrada())
    print(f"[Vigilante] Activo — comprobando cada {INTERVALO}s")
    yield
    tarea.cancel()


# =========================================================
# APP
# =========================================================

app = FastAPI(title="Procesador Automático de Facturas", lifespan=lifespan)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# =========================================================
# HOME
# =========================================================

@app.get("/")
def home():
    return FileResponse(os.path.join(TEMPLATES_DIR, "index.html"))

# =========================================================
# PROCESAR CARPETA (manual desde la UI)
# =========================================================

@app.post("/procesar")
def procesar():
    resultados = procesar_carpeta()
    return JSONResponse({"resultados": resultados})

# =========================================================
# UPLOAD PDF (llamado desde Power Automate)
# =========================================================

@app.post("/upload-pdf")
async def upload_pdf(file: UploadFile = File(...)):

    os.makedirs(CARPETA_ENTRADA, exist_ok=True)
    pdf_path = os.path.join(CARPETA_ENTRADA, file.filename)

    with open(pdf_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    try:
        text = read_pdf_text(pdf_path)

        if es_albaran(text):
            mover_pdf(pdf_path, "albaran")
            return JSONResponse({"archivo": file.filename, "estado": "albaran"})

        if len(text) < 80:
            mover_pdf(pdf_path, "imagen")
            return JSONResponse({"archivo": file.filename, "estado": "imagen"})

        result_csv = extract_invoice_with_agent(
            file_name=file.filename,
            invoice_text=text,
        )

        tabla = csv_to_matrix(result_csv)

        if len(tabla) < 2:
            mover_pdf(pdf_path, "error")
            return JSONResponse({"archivo": file.filename, "estado": "error", "detalle": "respuesta vacía del modelo"})

        fila = tabla[1]
        tipo = clasificar_factura(fila)
        mover_pdf(pdf_path, tipo)
        guardar_historial(result_csv, "auto")

        return JSONResponse({"archivo": file.filename, "estado": tipo})

    except Exception as e:
        try:
            mover_pdf(pdf_path, "error")
        except Exception:
            pass
        return JSONResponse(
            {"archivo": file.filename, "estado": "error", "detalle": str(e)},
            status_code=500,
        )

# =========================================================
# ESTADISTICAS
# =========================================================

@app.get("/estadisticas")
def estadisticas():

    carpetas = {
        "entrada":         os.path.join(FACTURAS_DIR, "entrada"),
        "procesadas":      os.path.join(FACTURAS_DIR, "procesadas"),
        "imagenes":        os.path.join(FACTURAS_DIR, "imagenes"),
        "error":           os.path.join(FACTURAS_DIR, "error"),
        "examinadas":      os.path.join(FACTURAS_DIR, "examinadas"),
        "manual":          os.path.join(FACTURAS_DIR, "corregir_manualmente"),
        "albaran":         os.path.join(FACTURAS_DIR, "albaran"),
        "reenviar_pedido": os.path.join(FACTURAS_DIR, "reenviar_falta_pedidocliente"),
    }

    datos = {}
    for nombre, ruta in carpetas.items():
        if not os.path.exists(ruta):
            datos[nombre] = 0
        else:
            datos[nombre] = len([f for f in os.listdir(ruta) if f.lower().endswith(".pdf")])

    return JSONResponse(datos)
