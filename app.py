import pandas as pd
import numpy as np

# ==============================================================================
# 1. CARGA Y LIMPIEZA DE DATOS DESDE EXCEL
# ==============================================================================

def cargar_y_limpiar_movimientos(file_path="MOVIMIENTOS 2026.xlsx", sheet_name="Detalle1"):
    """
    Carga el archivo de movimientos contables y ajusta los encabezados dinámicamente.
    """
    try:
        # Carga del Excel omitiendo encabezados iniciales de reporte si existen
        df_raw = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # Localizar la fila que contiene los nombres de las columnas reales
        header_row_idx = None
        for idx, row in df_raw.iterrows():
            if "PROVEEDORCLIENTE" in row.values or "Fecha y Hora contable" in row.values:
                header_row_idx = idx
                break
        
        if header_row_idx is not None:
            df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=header_row_idx + 1)
        else:
            df = df_raw.copy()

        # Normalizar nombres de columnas
        df.columns = df.columns.str.strip().str.upper()
        
        # Limpieza de montos numéricos
        columnas_numericas = ['SUBTOTAL', 'CARGO', 'ABONO', 'SALDO']
        for col in columnas_numericas:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)

        # Formato de fecha
        if 'FECHA Y HORA CONTABLE' in df.columns:
            df['FECHA Y HORA CONTABLE'] = pd.to_datetime(df['FECHA Y HORA CONTABLE'], errors='coerce')

        return df

    except Exception as e:
        print(f"Nota: No se pudo cargar el archivo local directamente ({e}). Utilizando dataset estructurado de respaldo...")
        return crear_dataset_ejemplo()

# ==============================================================================
# 2. DATASET DE RESPALDO (ESTRUCTURA COMPLETA COMPATIBLE)
# ==============================================================================

def crear_dataset_ejemplo():
    data = [
        {
            "FECHA Y HORA CONTABLE": "2026-09-30 14:48:39",
            "DESCRIPCIÓN": "ABONO TRANSFERENCIA SPEI",
            "PROVEEDORCLIENTE": "BBVA",
            "SUBTOTAL": 150000.0,
            "CARGO": 0.0,
            "ABONO": 150000.0,
            "SALDO": 181260.09,
            "AREA": "RANCHO",
            "APLICADO EN": "FINANCIERO",
            "SUBCUENTA": "LA PERLA",
            "CONCEPTO": "TRANSFERENCIA ENTRE CUENTAS 012680001265451814"
        },
        {
            "FECHA Y HORA CONTABLE": "2026-09-28 14:37:46",
            "DESCRIPCIÓN": "ABONO TRANSFERENCIA SPEI",
            "PROVEEDORCLIENTE": "BBVA",
            "SUBTOTAL": 100000.0,
            "CARGO": 0.0,
            "ABONO": 100000.0,
            "SALDO": 126905.09,
            "AREA": "RANCHO",
            "APLICADO EN": "FINANCIERO",
            "SUBCUENTA": "LA PERLA",
            "CONCEPTO": "TRANSFERENCIA ENTRE CUENTAS 012680001265451814"
        },
        {
            "FECHA Y HORA CONTABLE": "2026-01-06 10:00:00",
            "DESCRIPCIÓN": "CARGO FLETE DE HORTALIZAS",
            "PROVEEDORCLIENTE": "JUANA SILVIA HERNANDEZ GONZALEZ",
            "SUBTOTAL": 5459.03,
            "CARGO": 6332.48,
            "ABONO": 0.0,
            "SALDO": 0.0,
            "AREA": "LA PERLA",
            "APLICADO EN": "BROCOLI",
            "SUBCUENTA": "FLETE",
            "CONCEPTO": "FLETE DE BROCOLI DE CAMPO A PLANTA"
        },
        {
            "FECHA Y HORA CONTABLE": "2026-01-07 11:30:00",
            "DESCRIPCIÓN": "COMPRA DIESEL MAQUINARIA",
            "PROVEEDORCLIENTE": "GASOLINERA LA PERLA SA DE CV",
            "SUBTOTAL": 7211.64,
            "CARGO": 8365.50,
            "ABONO": 0.0,
            "SALDO": 1628962.20,
            "AREA": "LA PERLA",
            "APLICADO EN": "EQUIPO",
            "SUBCUENTA": "DIESEL",
            "CONCEPTO": "SUMINISTRO DIESEL TRACTORES"
        }
    ]
    df = pd.DataFrame(data)
    df['FECHA Y HORA CONTABLE'] = pd.to_datetime(df['FECHA Y HORA CONTABLE'])
    return df

# ==============================================================================
# 3. FUNCIONES DE ANÁLISIS Y REPORTING FINANCIERO
# ==============================================================================

def generar_resumen_por_area(df):
    """Calcula total de cargos, abonos y flujo neto por Área."""
    resumen = df.groupby('AREA')[['CARGO', 'ABONO']].sum().reset_index()
    resumen['FLUJO_NETO'] = resumen['ABONO'] - resumen['CARGO']
    return resumen

def generar_resumen_por_aplicacion(df):
    """Calcula desgloses agrupados por Aplicado En y Subcuenta."""
    resumen = df.groupby(['APLICADO EN', 'SUBCUENTA'])[['CARGO', 'ABONO']].sum().reset_index()
    resumen['NETO'] = resumen['ABONO'] - resumen['CARGO']
    return resumen

# ==============================================================================
# 4. EJECUCIÓN DEL SCRIPT
# ==============================================================================

if __name__ == "__main__":
    # Cargar datos
    df_movimientos = cargar_y_limpiar_movimientos()

    print("=== VISTA PREVIA DE LOS DATOS PROCESADOS ===")
    print(df_movimientos[['FECHA Y HORA CONTABLE', 'PROVEEDORCLIENTE', 'CARGO', 'ABONO', 'AREA', 'SUBCUENTA']].head())

    print("\n=== RESUMEN POR ÁREA ===")
    resumen_area = generar_resumen_por_area(df_movimientos)
    print(resumen_area.to_string(index=False))

    print("\n=== RESUMEN POR APLICADO EN / SUBCUENTA ===")
    resumen_app = generar_resumen_por_aplicacion(df_movimientos)
    print(resumen_app.to_string(index=False))
