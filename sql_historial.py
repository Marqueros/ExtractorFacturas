"""
Guarda en SQL Server las facturas clasificadas como "examinada"
(extracción correcta), tanto del flujo automático como del manual.

Es un complemento del historial Excel, no un sustituto: si la conexión
a SQL Server falla o no está configurada, se registra un aviso y el
flujo de procesamiento de facturas continúa con normalidad.
"""

import os

from dotenv import load_dotenv

_ROOT = os.path.dirname(os.path.abspath(__file__))
load_dotenv(dotenv_path=os.path.join(_ROOT, ".env"))

SQL_SERVER = os.getenv("SQL_SERVER", "").strip()
SQL_DATABASE = os.getenv("SQL_DATABASE", "").strip()
SQL_DRIVER = os.getenv("SQL_DRIVER", "ODBC Driver 17 for SQL Server").strip()

TABLA = "FacturasExaminadas"

COLUMNAS = [
    "Archivo",
    "BaseImp",
    "BaseIRPF",
    "Buyer",
    "Empresa",
    "FEscaneo",
    "FFactura",
    "FOperacion",
    "ImporIVA",
    "Moneda",
    "NombreProveedor",
    "NumeroFactura",
    "PedidoCliente",
    "Proveedor",
    "TipoIVA",
    "TipoIVA2",
    "TipoIVA3",
    "TotalFact",
]


def _conectar():
    import pyodbc

    if not SQL_SERVER or not SQL_DATABASE:
        raise RuntimeError(
            "SQL_SERVER / SQL_DATABASE no están configurados en .env; "
            "no se puede guardar en SQL Server."
        )

    conn_str = (
        f"DRIVER={{{SQL_DRIVER}}};"
        f"SERVER={SQL_SERVER};"
        f"DATABASE={SQL_DATABASE};"
        "Trusted_Connection=yes;"
    )

    return pyodbc.connect(conn_str, timeout=5)


def _crear_tabla_si_no_existe(cursor):
    columnas_sql = ",\n".join(f"[{c}] NVARCHAR(255) NULL" for c in COLUMNAS)

    cursor.execute(f"""
        IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = '{TABLA}')
        CREATE TABLE {TABLA} (
            Id INT IDENTITY(1,1) PRIMARY KEY,
            {columnas_sql},
            Origen NVARCHAR(20) NULL,
            FechaInsercion DATETIME NOT NULL DEFAULT GETDATE(),
            CONSTRAINT UQ_{TABLA}_Archivo UNIQUE (Archivo)
        )
    """)


def guardar_factura_examinada_sql(fila, origen):
    """
    Inserta o actualiza (según 'Archivo') una factura "examinada" en SQL Server.

    fila:   lista de valores en el mismo orden que EXPECTED_HEADERS
            (Archivo, BaseImp, ..., TotalFact).
    origen: "auto" o "manual".

    No propaga excepciones: un fallo de SQL Server no debe romper el
    procesamiento de facturas ni el guardado del historial Excel.
    """
    try:
        datos = dict(zip(COLUMNAS, (list(fila) + ["-"] * len(COLUMNAS))[:len(COLUMNAS)]))

        with _conectar() as conn:
            cursor = conn.cursor()
            _crear_tabla_si_no_existe(cursor)
            conn.commit()

            columnas_sin_archivo = [c for c in COLUMNAS if c != "Archivo"]
            set_clause = ", ".join(f"[{c}] = ?" for c in columnas_sin_archivo)
            valores_update = [datos[c] for c in columnas_sin_archivo]

            cursor.execute(
                f"UPDATE {TABLA} SET {set_clause}, Origen = ?, FechaInsercion = GETDATE() "
                f"WHERE Archivo = ?",
                (*valores_update, origen, datos["Archivo"]),
            )

            if cursor.rowcount == 0:
                columnas_insert = ", ".join(f"[{c}]" for c in COLUMNAS)
                placeholders = ", ".join("?" for _ in COLUMNAS)
                cursor.execute(
                    f"INSERT INTO {TABLA} ({columnas_insert}, Origen) "
                    f"VALUES ({placeholders}, ?)",
                    (*[datos[c] for c in COLUMNAS], origen),
                )

            conn.commit()

    except Exception as e:
        print(f"AVISO: no se pudo guardar la factura en SQL Server ({origen}): {e}")
