import pandas as pd

# 1. Carga de datos desde tu archivo de Excel o PDF procesado
# Asegúrate de colocar la ruta correcta a tu archivo
archivo_path = 'MOVIMIENTOS 2026.xlsx'

try:
    # Cargar los datos omitiendo filas de encabezados adicionales
    df = pd.read_excel(archivo_path, sheet_name='Detalle1', skiprows=2)
    
    # Asignar y limpiar nombres de columnas
    df.columns = [
        'Fecha_Hora', 'Descripcion', 'Proveedor_Cliente', 'Subtotal', 
        'Cargo', 'Abono', 'Saldo', 'Area', 'Aplicado_En', 'Subcuenta', 'Concepto'
    ]
    
    # Rellenar valores nulos en importes numéricos con 0
    df['Cargo'] = df['Cargo'].fillna(0.0)
    df['Abono'] = df['Abono'].fillna(0.0)
    df['Subtotal'] = df['Subtotal'].fillna(0.0)
    
    # Formatear la fecha
    df['Fecha_Hora'] = pd.to_datetime(df['Fecha_Hora'])

except Exception as e:
    print(f"Error al cargar el archivo directamente: {e}")
    # Estructura fallback alternativa en caso de carga limpia manual
    df = pd.DataFrame([
        {"Fecha_Hora": "2026-09-30 14:48:39", "Proveedor_Cliente": "BBVA", "Cargo": 0.0, "Abono": 150000.0, "Saldo": 181260.09, "Area": "RANCHO", "Aplicado_En": "FINANCIERO", "Subcuenta": "LA PERLA"},
        {"Fecha_Hora": "2026-09-28 14:37:46", "Proveedor_Cliente": "BBVA", "Cargo": 0.0, "Abono": 100000.0, "Saldo": 126905.09, "Area": "RANCHO", "Aplicado_En": "FINANCIERO", "Subcuenta": "LA PERLA"},
        {"Fecha_Hora": "2026-09-02 12:58:22", "Proveedor_Cliente": "BBVA", "Cargo": 0.0, "Abono": 92701.0, "Saldo": 138035.09, "Area": "RANCHO", "Aplicado_En": "FINANCIERO", "Subcuenta": "LA PERLA"},
        {"Fecha_Hora": "2026-09-02 13:06:58", "Proveedor_Cliente": "BBVA", "Cargo": 0.0, "Abono": 79679.0, "Saldo": 217714.09, "Area": "RANCHO", "Aplicado_En": "FINANCIERO", "Subcuenta": "LA PERLA"},
        {"Fecha_Hora": "2026-09-24 14:24:57", "Proveedor_Cliente": "BBVA", "Cargo": 0.0, "Abono": 100000.0, "Saldo": 141105.92, "Area": "RANCHO", "Aplicado_En": "FINANCIERO", "Subcuenta": "LA PERLA"},
        {"Fecha_Hora": "2026-09-08 15:25:47", "Proveedor_Cliente": "BBVA", "Cargo": 0.0, "Abono": 66202.6, "Saldo": 94676.48, "Area": "RANCHO", "Aplicado_En": "FINANCIERO", "Subcuenta": "LA PERLA"}
    ])

# 2. Funciones de Análisis Financiero

def resumen_general(dataframe):
    """Genera un resumen global de Cargos, Abonos y Flujo Neto."""
    total_cargos = dataframe['Cargo'].sum()
    total_abonos = dataframe['Abono'].sum()
    flujo_neto = total_abonos - total_cargos
    
    return pd.Series({
        'Total Cargos': total_cargos,
        'Total Abonos': total_abonos,
        'Flujo Neto': flujo_neto
    })

def desglose_por_columna(dataframe, columna='Area'):
    """Agrupa los movimientos por la columna deseada (Area, Aplicado_En, Subcuenta, etc.)."""
    agrupado = dataframe.groupby(columna)[['Cargo', 'Abono']].sum().reset_index()
    agrupado['Balance_Neto'] = agrupado['Abono'] - agrupado['Cargo']
    return agrupado.sort_values(by='Balance_Neto', ascending=False)

# 3. Ejecución de consultas e impresión de resultados

print("=== VISTA PREVIA DE LOS DATOS ===")
print(df.head(10))

print("\n=== RESUMEN GLOBAL DE MOVIMIENTOS ===")
print(resumen_general(df))

print("\n=== BALANCE AGRUPADO POR ÁREA ===")
print(desglose_por_columna(df, columna='Area'))

print("\n=== BALANCE AGRUPADO POR SUBCUENTA ===")
print(desglose_por_columna(df, columna='Subcuenta'))

# 4. Exportar reporte limpio a un nuevo archivo Excel
# df.to_excel("MOVIMIENTOS_2026_PROCESADO.xlsx", index=False)
