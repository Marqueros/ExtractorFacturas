import os

from fastapi import FastAPI
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware

from logic import procesar_carpeta

app = FastAPI(title="Procesador Automático de Facturas")

# =========================================================
# RUTAS
# =========================================================

BASE_DIR = os.path.dirname(
    os.path.abspath(__file__)
)


TEMPLATES_DIR = os.path.join(
    BASE_DIR,
    "templates"
)

# =========================================================
# CORS
# =========================================================

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

    return FileResponse(
        os.path.join(
            TEMPLATES_DIR,
            "index.html"
        )
    )

# =========================================================
# PROCESAR CARPETA
# =========================================================

@app.post("/procesar")
def procesar():

    resultados = procesar_carpeta()

    return JSONResponse({
        "resultados": resultados
    })

# =========================================================
# ESTADISTICAS
# =========================================================

@app.get("/estadisticas")
def estadisticas():

    carpetas = {
        "entrada":    "facturas/entrada",
        "procesadas": "facturas/procesadas",
        "imagenes":   "facturas/imagenes",
        "error":      "facturas/error",
        "examinadas": "facturas/examinadas",
        "manual":     "facturas/corregir_manualmente",
        "albaran":    "facturas/albaran",
    }

    datos = {}

    for nombre, ruta in carpetas.items():

        if not os.path.exists(ruta):

            datos[nombre] = 0

        else:

            datos[nombre] = len(
                [
                    f
                    for f in os.listdir(ruta)
                    if f.lower().endswith(".pdf")
                ]
            )

    return JSONResponse(datos)

@app.get("/")
def home():

    print("ESTOY EN APP_AUTO")

    return FileResponse(
        os.path.join(
            TEMPLATES_DIR,
            "index.html"
        )
    )