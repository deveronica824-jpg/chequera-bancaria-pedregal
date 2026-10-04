import streamlit as st
import pandas as pd
import numpy as np

# Configuración de página de Streamlit
st.set_page_config(page_title="Chequera Bancaria Pedregal", layout="wide")

st.title("📊 Chequera Bancaria Pedregal")
st.write("Sube tu archivo de movimientos en Excel para generar el análisis financiero.")

@st.cache_data
def cargar_y_limpiar_movimientos(archivo):
    """
    Carga y consolida las hojas de un libro de Excel financiero o reporte,
    detectando automáticamente la fila de encabezados. Funciona con st.file_uploader o rutas locales.
    """
    xls = pd.ExcelFile(archivo)
    hojas_procesadas = []

    for nombre_hoja in xls.sheet_names:
        df_raw = pd.read_excel(xls, sheet_name=nombre_hoja)
        
        # Buscar la fila que contiene los nombres de las columnas principales
        header_idx = None
        for idx, row in df_raw.iterrows():
            valores_fila = [str(val).upper().strip() for val in row.values if pd.notna(val)]
            if any("PROVEEDOR" in v or "CARGO" in v or "ABONO" in v for v in valores_fila):
                header_idx = idx
                break
        
        if header_idx is not None:
            # Reasignar encabezados y limpiar filas superiores
            df_hoja = pd.read_excel(xls, sheet_name=nombre_hoja, skiprows=header_idx + 1)
            # Normalizar nombres de columnas
            df_hoja.columns = [str(col).strip().upper().replace(" ", "_") for col in df_raw.iloc[header_idx]]
            df_hoja['HOJA_ORIGEN'] = nombre_hoja
            hojas_procesadas.append(df_hoja)

    if not hojas_procesadas:
        st.error("No se encontraron tablas con formato válido en el archivo subido.")
        return pd.DataFrame()

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

    # Filtrar filas vacías o de totales
    cols_clave = [col for col in ['CARGO', 'ABONO', 'SALDO'] if col in df_consolidado.columns]
    if cols_clave:
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

    # 1. Resumen por Área
    if 'AREA' in df.columns:
        resumen_area = df.groupby('AREA').agg(
            Total_Cargos=('CARGO', 'sum'),
            Total_Abonos=('ABONO', 'sum'),
            Total_Movimientos=('CARGO', 'count')
        ).reset_index()
        resumen_area['Flujo_Neto'] = resumen_area['Total_Abonos'] - resumen_area['Total_Cargos']
        reportes['resumen_area'] = resumen_area

    # 2. Resumen por Subcuenta y Aplicado En
    cols_agrupacion = [c for c in ['AREA', 'APLICADO_EN', 'SUBCUENTA'] if c in df.columns]
    if cols_agrupacion:
        reportes['desglose_subcuenta'] = df.groupby(cols_agrupacion).agg(
            Total_Cargos=('CARGO', 'sum'),
            Total_Abonos=('ABONO', 'sum')
        ).reset_index()

    # 3. Top Proveedores / Clientes
    if 'PROVEEDOR_CLIENTE' in df.columns:
        reportes['top_proveedores'] = df.groupby('PROVEEDOR_CLIENTE').agg(
            Total_Pagado=('CARGO', 'sum'),
            Total_Cobrado=('ABONO', 'sum')
        ).sort_values(by='Total_Pagado', ascending=False).head(10).reset_index()

    return reportes


# Selector de archivo en la interfaz de Streamlit
archivo_subido = st.file_uploader("Carga aquí tu archivo de Excel (.xlsx / .xls)", type=["xlsx", "xls"])

if archivo_subido is not None:
    with st.spinner("Procesando datos del archivo..."):
        df_movimientos = cargar_y_limpiar_movimientos(archivo_subido)
        
    if not df_movimientos.empty:
        st.success(f"¡Procesamiento exitoso! Total de registros analizados: {len(df_movimientos)}")

        # Métricas principales
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Cargos (Egresos)", f"${df_movimientos['CARGO'].sum():,.2f}")
        col2.metric("Total Abonos (Ingresos)", f"${df_movimientos['ABONO'].sum():,.2f}")
        flujo_neto = df_movimientos['ABONO'].sum() - df_movimientos['CARGO'].sum()
        col3.metric("Flujo Neto", f"${flujo_neto:,.2f}")

        # Pestañas con los resultados
        tab1, tab2, tab3 = st.tabs(["📋 Detalle de Movimientos", "🏢 Resumen por Área", "🔝 Top Proveedores"])

        with tab1:
            st.subheader("Tabla Consolidada de Movimientos")
            st.dataframe(df_movimientos, use_container_width=True)

        reportes = generar_reporte_financiero(df_movimientos)

        with tab2:
            st.subheader("Resumen por Área Financial")
            if 'resumen_area' in reportes:
                st.dataframe(reportes['resumen_area'], use_container_width=True)
            else:
                st.info("No se encontró la columna 'AREA' para generar el reporte.")

        with tab3:
            st.subheader("Top Proveedores / Clientes")
            if 'top_proveedores' in reportes:
                st.dataframe(reportes['top_proveedores'], use_container_width=True)
            else:
                st.info("No se encontró la columna 'PROVEEDOR_CLIENTE'.")
else:
    st.info("👆 Sube un archivo Excel arriba para desplegar los reportes.")
