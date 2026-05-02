import pandas as pd
from datetime import datetime
from pathlib import Path
import os

# ===================================================================
# CONFIGURACIÓN
# ===================================================================
CARPETA_ARCHIVOS = r"XXXXXXXXXXXXX"
CARPETA_SALIDA   = r"C:XXXXXXXXXXXX"

CANTIDAD_ULTIMOS_ARCHIVOS = 15

# Crea carpeta de salida si no existe
Path(CARPETA_SALIDA).mkdir(parents=True, exist_ok=True)

# Nombre del archivo con fecha de hoy
hoy = datetime.now().strftime("%Y%m%d")
ARCHIVO_SALIDA = Path(CARPETA_SALIDA) / f"Productos_por_mapear_{hoy}.xlsx"

# ===================================================================
# 1. Tomar los últimos 14 archivos por fecha de creación
# ===================================================================
print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] Iniciando Mapeo de Productos → {ARCHIVO_SALIDA.name}")

archivos_txt = list(Path(CARPETA_ARCHIVOS).glob("*.txt"))
if not archivos_txt:
    print("No se encontraron archivos .txt")
    exit()

archivos_txt.sort(key=lambda x: x.stat().st_mtime, reverse=True)  # más reciente primero
ultimos = archivos_txt[:CANTIDAD_ULTIMOS_ARCHIVOS]

print(f"  Analizando los {len(ultimos)} archivos más recientes:")
for a in ultimos:
    print(f"    → {a.name}")

# ===================================================================
# 2. Procesar cada archivo buscando productos SIN MAPEAR
# ===================================================================
lista_productos_pendientes = []

for archivo in ultimos:
    print(f"  Revisando → {archivo.name}")
    try:
        df = pd.read_csv(
            archivo, sep="|", encoding="utf-8", dtype=str,
            engine="python", on_bad_lines="skip"
        )
    except Exception as e:
        print(f"    Error leyendo archivo: {e}")
        continue

    # Columnas obligatorias para este reporte
    columnas_necesarias = [
        "customer_sap",
        "dimtiendas_cadenanombre",
        "dimtiendas_formato",
        "dimproductos_upc",
        "dimproductos_clvCliente",
        "vts_piezas"                      # ← agregada para sumar cantidades
    ]

    if not all(col in df.columns for col in columnas_necesarias):
        print(f"    Faltan columnas clave (incluyendo vts_piezas) → saltado")
        continue

    # Filtrar productos SIN MAPEAR
    condicion_sin_mapear = (
        (df["dimproductos_clvCliente"].isnull()) |
        (df["dimproductos_clvCliente"].str.strip() == "") |
        (df["dimproductos_clvCliente"].str.upper() == "SIN IDENTIFICAR") |
        (df["dimproductos_clvCliente"] == "0")
    )

    pendientes = df[condicion_sin_mapear][columnas_necesarias].copy()

    if pendientes.empty:
        print(f"    No hay productos sin mapear en este archivo")
        continue

    # Convertir ventas a numérico (errores → 0)
    pendientes["vts_piezas"] = pd.to_numeric(pendientes["vts_piezas"], errors="coerce").fillna(0)

    # Extraer Distribuidor y País del nombre del archivo
    partes = archivo.name.upper().split("_")
    distribuidor_nombre = partes[0] if len(partes) > 0 else "DESCONOCIDO"
    pais = partes[1] if len(partes) > 1 else "XX"

    # Añadir columnas de trazabilidad
    pendientes["Archivo_Origen"] = archivo.name
    pendientes["Distribuidor_Nombre_Archivo"] = distribuidor_nombre
    pendientes["País_Archivo"] = pais
    pendientes["Fecha_Procesado"] = datetime.now().strftime("%Y-%m-%d %H:%M")

    lista_productos_pendientes.append(pendientes)

# ===================================================================
# 3. Si no hay nada → generar Excel vacío con encabezados
# ===================================================================
if not lista_productos_pendientes:
    print("No se encontraron productos sin mapear en ningún archivo.")
    columnas_vacias = [
        "customer_sap", "dimtiendas_cadenanombre", "dimtiendas_formato",
        "dimproductos_upc", "dimproductos_clvCliente", "Total_Piezas_SIN_MAPEAR",
        "Archivo_Origen", "Fecha_Procesado"
    ]
    df_vacio = pd.DataFrame(columns=columnas_vacias)
    with pd.ExcelWriter(ARCHIVO_SALIDA, engine="openpyxl") as writer:
        df_vacio.to_excel(writer, sheet_name="Por_Mapear", index=False)
    print(f"   Archivo vacío generado: {ARCHIVO_SALIDA.name}")
else:
    # ===================================================================
    # 4. Consolidar y agrupar/sumar
    # ===================================================================
    print(f"  Consolidando {sum(len(df) for df in lista_productos_pendientes)} filas sin mapear...")
    final = pd.concat(lista_productos_pendientes, ignore_index=True)

    # Asegurar que vts_piezas sea numérico
    final["vts_piezas"] = pd.to_numeric(final["vts_piezas"], errors="coerce").fillna(0)

    # Agrupar por producto único y sumar piezas
    agregacion = final.groupby(["customer_sap", "dimproductos_upc"], as_index=False).agg({
        "vts_piezas": "sum",
        "dimtiendas_cadenanombre": "first",
        "dimtiendas_formato": "first",
        "dimproductos_clvCliente": "first",
        "Archivo_Origen": lambda x: ", ".join(x.unique()),  # todos los archivos donde apareció
        "Fecha_Procesado": "max",                           # la fecha más reciente
        "Distribuidor_Nombre_Archivo": "first",
        "País_Archivo": "first"
    })

    # Renombrar columna de suma
    agregacion = agregacion.rename(columns={"vts_piezas": "Total_Piezas_SIN_MAPEAR"})

    # Orden final
    agregacion = agregacion.sort_values(["dimtiendas_formato", "dimtiendas_cadenanombre", "dimproductos_upc"])

    # ===================================================================
    # 5. Columnas finales para el Excel
    # ===================================================================
    columnas_finales = [
        "customer_sap",                  # Código SAP distribuidor
        "dimtiendas_cadenanombre",       # Nombre cadena
        "dimtiendas_formato",            # País / Formato
        "dimproductos_upc",              # UPC del distribuidor
        "dimproductos_clvCliente",       # Código Diageo (vacío o SIN IDENTIFICAR)
        "Total_Piezas_SIN_MAPEAR",       # ← la suma que querías
        "Archivo_Origen",
        "Fecha_Procesado"
    ]

    final_ordenado = agregacion[columnas_finales]

    # Guardar con formato
    with pd.ExcelWriter(ARCHIVO_SALIDA, engine="openpyxl") as writer:
        final_ordenado.to_excel(writer, sheet_name="Por_Mapear", index=False)

        # Autoajustar columnas
        worksheet = writer.sheets["Por_Mapear"]
        for i, col in enumerate(final_ordenado.columns, 1):
            max_len = max(final_ordenado[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.column_dimensions[worksheet.cell(row=1, column=i).column_letter].width = min(max_len, 50)

    print(f"\n¡ÉXITO! {len(final_ordenado)} productos listos para mapear")
    print(f"   Total piezas sin mapear: {final_ordenado['Total_Piezas_SIN_MAPEAR'].sum():,.0f}")
    print(f"   Archivo generado → {ARCHIVO_SALIDA.name}")
    print(f"   Hora → {datetime.now():%Y-%m-%d %H:%M:%S}")

# Abrir la carpeta al finalizar
os.startfile(CARPETA_SALIDA)
