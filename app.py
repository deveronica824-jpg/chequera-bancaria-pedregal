import pandas as pd
import numpy as np

def cargar_y_limpiar_movimientos(archivo_path):
    """
    Carga y consolida las hojas de un libro de Excel financiero o reporte
    detectando automáticamente la fila de encabezados.
    """
    xls = pd.ExcelFile(archivo_path)
    hojas_procesadas = []

    for nombre_hoja in xls.sheet_names:
        df_raw = pd.read_excel(archivo_path, sheet_name=nombre_hoja)
        
        # Buscar la fila que contiene los nombres de las columnas principales
        header_idx = None
        for idx, row in df_raw.iterrows():
            valores_fila = [str(val).upper().strip() for val in row.values if pd.notna(val)]
            if any("PROVEEDOR" in v or "CARGO" in v or "ABONO" in v for v in valores_fila):
                header_idx = idx
                break
        
        if header_idx is not None:
            # Reasignar encabezados y limpiar filas superiores
            df_hoja = pd.read_excel(archivo_path, sheet_name=nombre_hoja, skiprows=header_idx + 1)
            # Normalizar nombres de columnas
            df_hoja.columns = [str(col).strip().upper().replace(" ", "_") for col in df_raw.iloc[header_idx]]
            df_hoja['HOJA_ORIGEN'] = nombre_hoja
            hojas_procesadas.append(df_hoja)

    if not hojas_procesadas:
        raise ValueError("No se encontraron tablas con formato válido en el archivo.")

    # Concatenar todas las pestañas procesadas
    df_consolidado = pd.concat(hojas_procesadas, ignore_index=True)

    # Renombrar columnas comunes para estandarizar
    mapeo_columnas = {
        'FECHA_Y_HORA_CONTABLE': 'FECHA',
        'PROVEEDORCLIENTE': 'PROVEEDOR_CLIENTE',
        'PROVEEDOR_/_CLIENTE': 'PROVEEDOR_CLIENTE',
        'APLICADO_EN': 'APLICADO_EN'
    }
    df_consolidado.rename(columns=mapeo_columnas, inplace=True)

    # Filtrar filas vacías o de totales que no contienen movimientos
    cols_clave = [col for col in ['CARGO', 'ABONO', 'SALDO'] if col in df_consolidado.columns]
    df_consolidado.dropna(subset=cols_clave, how='all', inplace=True)

    # Convertir y limpiar columnas numéricas
    for col in ['SUBTOTAL', 'CARGO', 'ABONO', 'SALDO']:
        if col in df_consolidado.columns:
            df_consolidado[col] = pd.to_numeric(
                df_consolidado[col].astype(str).str.replace('$', '', regex=False).str.replace(',', '', regex=False),
                errors='coerce'
            ).fillna(0.0)

    # Convertir columna de Fecha
    if 'FECHA' in df_consolidado.columns:
        df_consolidado['FECHA'] = pd.to_datetime(df_consolidado['FECHA'], errors='coerce')

    return df_consolidado


def generar_reporte_financiero(df):
    """
    Genera resúmenes ejecutivos a partir del dataframe consolidado.
    """
    reportes = {}

    # 1. Resumen por Área (RANCHO, LA PERLA, CASAS, etc.)
    if 'AREA' in df.columns:
        reportes['resumen_area'] = df.groupby('AREA').agg(
            Total_Cargos=('CARGO', 'sum'),
            Total_Abonos=('ABONO', 'sum'),
            Total_Movimientos=('CARGO', 'count')
        ).reset_index()
        reportes['resumen_area']['Flujo_Neto'] = reportes['resumen_area']['Total_Abonos'] - reportes['resumen_area']['Total_Cargos']

    # 2. Resumen por Subcuenta y Aplicado En
    cols_agrupacion = [c for c in ['AREA', 'APLICADO_EN', 'SUBCUENTA'] if c in df.columns]
    if cols_agrupacion:
        reportes['desglose_subcuenta'] = df.groupby(cols_agrupacion).agg(
            Total_Cargos=('CARGO', 'sum'),
            Total_Abonos=('ABONO', 'sum')
        ).reset_index()

    # 3. Top Proveedores / Clientes por volumen de egresos
    if 'PROVEEDOR_CLIENTE' in df.columns:
        reportes['top_proveedores'] = df.groupby('PROVEEDOR_CLIENTE').agg(
            Total_Pagado=('CARGO', 'sum'),
            Total_Cobrado=('ABONO', 'sum')
        ).sort_values(by='Total_Pagado', ascending=False).head(10).reset_index()

    return reportes


# ==========================================
# EJECUCIÓN PRINCIPAL
# ==========================================
if __name__ == "__main__":
    ruta_archivo = 'MOVIMIENTOS 2026.xlsx'  # Reemplazar por la ruta de tu archivo
    
    print("Cargando y procesando datos...")
    df_movimientos = cargar_y_limpiar_movimientos(ruta_archivo)
    
    print(f"\n¡Procesamiento exitoso! Total de movimientos válidos: {len(df_movimientos)}")
    print("\n--- VISTA PREVIA DE LOS DATOS ---")
    print(df_movimientos[['FECHA', 'PROVEEDOR_CLIENTE', 'CARGO', 'ABONO', 'AREA', 'SUBCUENTA']].head())

    # Generar reportes
    reportes = generar_reporte_financiero(df_movimientos)

    if 'resumen_area' in reportes:
        print("\n--- RESUMEN POR ÁREA ---")
        print(reportes['resumen_area'].to_string(index=False))

    # Guardar resultado consolidado a Excel
    archivo_salida = 'MOVIMIENTOS_CONSOLIDADOS_2026.xlsx'
    with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
        df_movimientos.to_excel(writer, sheet_name='Movimientos_Limpios', index=False)
        for nombre_reporte, df_rep in reportes.items():
            df_rep.to_excel(writer, sheet_name=nombre_reporte[:31], index=False)
            
    print(f"\nReporte completo exportado correctamente a: {archivo_salida}")
