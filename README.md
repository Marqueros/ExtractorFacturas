# Extractor de Facturas

Aplicación web que extrae automáticamente los datos estructurados de facturas en PDF usando IA (GPT-4.1). Permite subir una o varias facturas, visualizar los datos extraídos en tabla, corregirlos manualmente y exportarlos a Excel.

---

## Funcionalidades

- **Extracción automática** de datos de facturas PDF mediante LLM (OpenAI GPT-4.1 o compatible)
- **Soporte multiidioma**: detecta facturas en español, inglés, portugués, francés, italiano, alemán, neerlandés, chino, griego y otros
- **Tabla interactiva** con los datos extraídos, editable desde el navegador
- **Exportación a Excel** con formato y estilos aplicados
- **Historial diario** en Excel, acumulado por fecha con marca de hora
- **Corrección manual**: los cambios realizados en la UI se guardan en el historial y en un fichero de facturas corregidas
- **Clasificación automática de PDFs** en carpetas según resultado de la extracción:
  - `facturas/procesadas` — todos los campos obligatorios extraídos correctamente
  - `facturas/corregir_manualmente` — faltan campos obligatorios
  - `facturas/imagenes` — PDF escaneado sin texto extraíble
  - `facturas/examinadas` — factura ya revisada y corregida por un usuario.

---

## Campos extraídos

| Campo | Descripción |
|---|---|
| Archivo | Nombre del archivo PDF |
| BaseImp | Base imponible |
| BaseIRPF | Base IRPF (si aplica) |
| Buyer | CIF/NIF/VAT del comprador |
| Empresa | Razón social del comprador |
| FEscaneo | Fecha y hora de procesamiento |
| FFactura | Fecha de la factura |
| FOperacion | Fecha de operación |
| ImporIVA | Importe del IVA |
| Moneda | Moneda (EUR, USD, GBP…) |
| NombreProveedor | Razón social del proveedor |
| NumeroFactura | Número de factura |
| PedidoCliente | Número de pedido del cliente |
| Proveedor | CIF/NIF/VAT del proveedor |
| TipoIVA | Tipo de IVA principal |
| TipoIVA2 | Segundo tipo de IVA |
| TipoIVA3 | Tercer tipo de IVA |
| TotalFact | Importe total de la factura |

---

## Requisitos

- Python 3.10+
- API key de OpenAI (o endpoint compatible)

---

## Instalación

```bash
# Clonar el repositorio
git clone <url-del-repo>
cd Extraer-texto-facturas

# Crear entorno virtual
python -m venv .venv
.venv\Scripts\activate      # Windows
# source .venv/bin/activate  # Linux/Mac

# Instalar dependencias
pip install -r requirements.txt
```

---

## Configuración

Crear un fichero `.env` en la raíz del proyecto:

```env
OPENAI_API_KEY=sk-...
OPENAI_MODEL=gpt-4.1
# Opcional: para usar un endpoint alternativo (Azure OpenAI, proxy, etc.)
# OPENAI_BASE_URL=https://...
```

---

## Uso

```bash
uvicorn app.main:app --reload
```

La aplicación estará disponible en `http://localhost:8000`.

### Flujo de uso

1. Abre el navegador en `http://localhost:8000`
2. Sube una o varias facturas PDF
3. Los datos extraídos aparecen en tabla
4. Edita cualquier celda para corregir un valor
5. Descarga el resultado en Excel

---

## Estructura del proyecto

```
├── app/
│   ├── main.py           # API FastAPI (endpoints)
│   └── logic.py          # Lógica: extracción PDF, llamada al LLM, Excel, historial
├── templates/
│   └── index.html        # Interfaz web
├── static/
│   └── img/              # Recursos estáticos
├── facturas/
│   ├── entrada/          # PDFs recién subidos
│   ├── procesadas/       # Extracción completa
│   ├── corregir_manualmente/  # Campos incompletos
│   ├── imagenes/         # PDFs sin texto extraíble
│   └── examinadas/       # Revisadas por usuario
├── historial/
│   ├── historial_facturas_usuario_YYYY-MM-DD.xlsx
│   └── facturas_corregidas.xlsx
├── requirements.txt
└── .env                  # Variables de entorno (no subir al repositorio)
```

---

## API endpoints

| Método | Ruta | Descripción |
|---|---|---|
| `GET` | `/` | Interfaz web |
| `POST` | `/extraer` | Extrae datos de uno o varios PDFs y guarda historial |
| `POST` | `/extraer-excel` | Extrae datos y devuelve fichero Excel para descarga |
| `GET` | `/historial` | Descarga el historial Excel del día actual |
| `GET` | `/historial-json` | Devuelve el historial del día en formato JSON |
| `POST` | `/guardar-correccion` | Guarda una corrección manual sobre una factura ya procesada |

---

## Dependencias principales

| Paquete | Uso |
|---|---|
| `fastapi` | Framework web |
| `uvicorn` | Servidor ASGI |
| `pdfplumber` | Extracción de texto de PDFs |
| `openai` | Cliente OpenAI / Azure OpenAI |
| `openpyxl` | Generación y lectura de ficheros Excel |
| `python-dotenv` | Carga de variables de entorno |
| `python-multipart` | Subida de ficheros |
