import pandas as pd
from datetime import datetime
from pathlib import Path

# ===================================================================
# CONFIGURACIÓN
# ===================================================================
CARPETA_ARCHIVOS = r"\\AZEUNLIHP001\Teamcore\Archive"
CARPETA_SALIDA   = r"C:\Users\HernaGab1\OneDrive - Diageo\Desktop\Auto_1\Resumen_Crudo"

CANTIDAD_ULTIMOS_ARCHIVOS = 15

# Crea carpeta de salida si no existe
Path(CARPETA_SALIDA).mkdir(parents=True, exist_ok=True)

# Nombre del archivo con fecha de hoy
hoy = datetime.now().strftime("%Y%m%d")
ARCHIVO_SALIDA = Path(CARPETA_SALIDA) / f"Resumen_Data_Cruda_{hoy}.xlsx"

# ===================================================================
# 1. Últimos N archivos
# ===================================================================
print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] Iniciando → {ARCHIVO_SALIDA.name}")

archivos_txt = list(Path(CARPETA_ARCHIVOS).glob("*.txt"))
if not archivos_txt:
    print("No hay archivos .txt")
    exit()

archivos_txt.sort(key=lambda x: x.stat().st_ctime, reverse=True)
ultimos = archivos_txt[:CANTIDAD_ULTIMOS_ARCHIVOS]

print(f"  Procesando los {len(ultimos)} archivos más recientes:")
for a in ultimos:
    print(f"    → {a.name}")

# ===================================================================
# 2. Leer y procesar (CON customer_sap)
# ===================================================================
lista_resumen = []

for archivo in ultimos:
    print(f"  Leyendo → {archivo.name}")
    try:
        df = pd.read_csv(archivo, sep="|", encoding="utf-8", dtype=str,
                         engine="python", on_bad_lines="skip")
    except Exception as e:
        print(f"    Error leyendo archivo: {e}")
        continue

    # Columnas obligatorias + customer_sap
    cols_req = ["dimtiempo_date", "vts_piezas", "vts_importe_real", "customer_sap"]
    if not all(col in df.columns for col in cols_req):
        print(f"    Faltan columnas (incluyendo customer_sap) → saltado")
        continue

    # Tomamos el código SAP (constante por archivo)
    codigo_sap = df["customer_sap"].dropna()
    if codigo_sap.empty:
        codigo_sap = "SIN_CODIGO"
    else:
        codigo_sap = codigo_sap.iloc[0].strip()

    # Nos quedamos solo con las columnas numéricas
    df = df[["dimtiempo_date", "vts_piezas", "vts_importe_real"]].copy()

    # Limpieza de fechas
    df["dimtiempo_date"] = df["dimtiempo_date"].astype(str).str.strip().str.zfill(8)
    df = df[df["dimtiempo_date"].str.match(r"^\d{8}$")]
    if df.empty:
        print(f"    No hay fechas válidas → saltado")
        continue

    # Conversión numérica
    df["vts_piezas"] = pd.to_numeric(df["vts_piezas"], errors="coerce").fillna(0)
    df["vts_importe_real"] = pd.to_numeric(df["vts_importe_real"], errors="coerce").fillna(0)
    df["AñoMes"] = df["dimtiempo_date"].str[:6]

    # Fechas min/max del archivo
    fecha_min = df["dimtiempo_date"].min()
    fecha_max = df["dimtiempo_date"].max()

    # Distribuidor y País del nombre del archivo
    partes = archivo.name.upper().split("_")
    distribuidor = partes[0] if len(partes) > 0 else "DESCONOCIDO"
    pais = partes[1] if len(partes) > 1 else "XX"

    # Resumen por AñoMes
    temp = df.groupby("AñoMes").agg(
        Botellas_Total=("vts_piezas", "sum"),
        Importe_Real_Total=("vts_importe_real", "sum")
    ).reset_index()

    # Añadimos metadatos
    temp["Distribuidor"] = distribuidor
    temp["País"] = pais
    temp["Código_Distribuidor"] = codigo_sap
    temp["Fecha_Min"] = fecha_min
    temp["Fecha_Max"] = fecha_max
    temp["Archivo_Procesado"] = archivo.name

    lista_resumen.append(temp)

if not lista_resumen:
    print("No se generaron datos")
    exit()

# ===================================================================
# 3. Resumen consolidado
# ===================================================================
print("  Generando resumen final...")
final = pd.concat(lista_resumen, ignore_index=True)

resumen = final.groupby(["Código_Distribuidor", "Distribuidor", "País", "AñoMes"], as_index=False).agg(
    Botellas_Total=("Botellas_Total", "sum"),
    Importe_Real_Total=("Importe_Real_Total", "sum"),
    Fecha_Min=("Fecha_Min", "min"),
    Fecha_Max=("Fecha_Max", "max")
)

resumen = resumen.sort_values(["País", "AñoMes", "Distribuidor"], ascending=[True, False, True])
resumen["Importe_Real_Total"] = resumen["Importe_Real_Total"].round(2)

# ===================================================================
# 4. Guardar Excel
# ===================================================================
with pd.ExcelWriter(ARCHIVO_SALIDA, engine="openpyxl") as writer:
    resumen.to_excel(writer, sheet_name="Resumen_Crudo", index=False)

print(f"\n¡PERFECTO! Ahora incluye Código_Distribuidor")
print(f"   Archivo → {ARCHIVO_SALIDA.name}")
print(f"   Registros → {len(resumen)}")
print(f"   Hora → {datetime.now():%Y-%m-%d %H:%M:%S}")

# Abrir carpeta
import os
os.startfile(CARPETA_SALIDA)
