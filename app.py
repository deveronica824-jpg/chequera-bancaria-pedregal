import sqlite3
import pandas as pd
import streamlit as st
from datetime import datetime

# ==========================================
# 1. CONFIGURACIÓN Y BASE DE DATOS
# ==========================================
DB_NAME = "gestion_chequeras.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    
    # Cuentas Bancarias
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS cuentas_bancarias (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT UNIQUE NOT NULL
        )
    ''')
    
    # Categorías / Centros de Costo
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS ReglasClasificacion (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            patron_busqueda TEXT UNIQUE NOT NULL,
            proveedor TEXT,
            cuenta TEXT,
            subcuenta TEXT,
            centro_costo TEXT
        )
    ''')

    # Movimientos Bancarios
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS movimientos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cuenta_bancaria TEXT NOT NULL,
            fecha DATE NOT NULL,
            concepto TEXT NOT NULL,
            proveedor TEXT,
            cuenta TEXT,
            subcuenta TEXT,
            centro_costo TEXT,
            cargo REAL DEFAULT 0.0,
            abono REAL DEFAULT 0.0,
            referencia TEXT
        )
    ''')
    conn.commit()
    conn.close()

init_db()

# ==========================================
# 2. FUNCIONES DE LÓGICA DE NEGOCIO
# ==========================================
def obtener_conexion():
    return sqlite3.connect(DB_NAME)

def clasificar_automatico(concepto):
    """Aplica reglas predefinidas o historial para sugerir clasificación."""
    conn = obtener_conexion()
    cursor = conn.cursor()
    cursor.execute("SELECT patron_busqueda, proveedor, cuenta, subcuenta, centro_costo FROM ReglasClasificacion")
    reglas = cursor.fetchall()
    conn.close()

    for patron, prov, cta, subcta, cc in reglas:
        if patron.lower() in concepto.lower():
            return prov, cta, subcta, cc
    return "Sin Clasificar", "General", "General", "General"

def cargar_excel_inicial(file_path):
    """Carga y consolida hojas de Enero a Octubre de un libro Excel."""
    excel_file = pd.ExcelFile(file_path)
    hojas_objetivo = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre']
    
    conn = obtener_conexion()
    cursor = conn.cursor()
    
    # Asegurar cuenta base por defecto
    cursor.execute("INSERT OR IGNORE INTO cuentas_bancarias (nombre) VALUES ('Cuenta Principal')")
    
    total_registros = 0
    for sheet in excel_file.sheet_names:
        if sheet.strip().lower() in hojas_objetivo:
            df = pd.read_excel(file_path, sheet_name=sheet)
            # Limpieza y mapeo flexible de columnas
            df.columns = [str(c).strip().lower() for c in df.columns]
            
            for _, row in df.iterrows():
                concepto = str(row.get('concepto', row.get('descripcion', 'Movimiento')))
                prov, cta, subcta, cc = clasificar_automatico(concepto)
                
                fecha = row.get('fecha', datetime.now().strftime('%Y-%m-%d'))
                cargo = float(row.get('cargo', row.get('egreso', 0.0)) or 0.0)
                abono = float(row.get('abono', row.get('ingreso', 0.0)) or 0.0)
                ref = str(row.get('referencia', ''))

                cursor.execute('''
                    INSERT INTO movimientos 
                    (cuenta_bancaria, fecha, concepto, proveedor, cuenta, subcuenta, centro_costo, cargo, abono, referencia)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', ('Cuenta Principal', str(fecha)[:10], concepto, prov, cta, subcta, cc, cargo, abono, ref))
                total_registros += 1

    conn.commit()
    conn.close()
    return total_registros

# ==========================================
# 3. INTERFAZ GRÁFICA DE USUARIO (STREAMLIT)
# ==========================================
st.set_page_config(page_title="Gestión de Chequeras y Centros de Costo", layout="wide")
st.title("🏦 Sistema de Gestión Multicuenta y Control de Costos")

# Menú Principal
opcion = st.sidebar.selectbox("Módulo principal", [
    "Resumen Mensual & Dashboard",
    "Reporte por Proveedor",
    "Captura / Edición Directa",
    "Carga de Archivos (Excel/PDF)",
    "Configuración de Reglas"
])

conn = obtener_conexion()

# ------------------------------------------
# MÓDULO: RESUMEN MENSUAL & DASHBOARD
# ------------------------------------------
if opcion == "Resumen Mensual & Dashboard":
    st.header("📊 Resumen Mensual y Estado de Cuentas")
    
    cuentas = pd.read_sql("SELECT nombre FROM cuentas_bancarias", conn)['nombre'].tolist()
    cuenta_sel = st.selectbox("Seleccionar Cuenta Bancaria", ["Todas"] + cuentas)
    
    query = "SELECT * FROM movimientos"
    if cuenta_sel != "Todas":
        query += f" WHERE cuenta_bancaria = '{cuenta_sel}'"
        
    df = pd.read_sql(query, conn)
    
    if not df.empty:
        df['fecha'] = pd.to_datetime(df['fecha'])
        df['Mes-Año'] = df['fecha'].dt.to_period('M')
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Egresos (Cargos)", f"${df['cargo'].sum():,.2f}")
        col2.metric("Total Ingresos (Abonos)", f"${df['abono'].sum():,.2f}")
        col3.metric("Balance Net", f"${(df['abono'].sum() - df['cargo'].sum()):,.2f}")
        
        st.subheader("Resumen Agrupado por Mes")
        resumen_mes = df.groupby('Mes-Año')[['cargo', 'abono']].sum()
        st.bar_chart(resumen_mes)
        st.dataframe(resumen_mes, use_container_width=True)
        
        st.subheader("Detalle por Centros de Costo")
        resumen_cc = df.groupby('centro_costo')[['cargo', 'abono']].sum()
        st.dataframe(resumen_cc, use_container_width=True)
    else:
        st.info("No hay datos registrados en la base de datos.")

# ------------------------------------------
# MÓDULO: REPORTE POR PROVEEDOR
# ------------------------------------------
elif opcion == "Reporte por Proveedor":
    st.header("🔍 Consulta y Reporte por Proveedor")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        prov_list = pd.read_sql("SELECT DISTINCT proveedor FROM movimientos", conn)['proveedor'].dropna().tolist()
        prov_sel = st.selectbox("Seleccionar Proveedor", prov_list if prov_list else ["Sin Datos"])
    with col2:
        fecha_inicio = st.date_input("Fecha Inicio", datetime(2026, 1, 1))
    with col3:
        fecha_fin = st.date_input("Fecha Fin", datetime(2026, 12, 31))
        
    if st.button("Generar Consulta"):
        q = f"""
            SELECT fecha, cuenta_bancaria, concepto, cuenta, subcuenta, centro_costo, cargo, abono, referencia 
            FROM movimientos 
            WHERE proveedor = '{prov_sel}' 
            AND fecha BETWEEN '{fecha_inicio}' AND '{fecha_fin}'
            ORDER BY fecha ASC
        """
        df_prov = pd.read_sql(q, conn)
        
        st.write(f"### Movimientos de: **{prov_sel}**")
        st.dataframe(df_prov, use_container_width=True)
        
        col_t1, col_t2 = st.columns(2)
        col_t1.metric("Total Pagado (Cargo)", f"${df_prov['cargo'].sum():,.2f}")
        col_t2.metric("Total Abonado", f"${df_prov['abono'].sum():,.2f}")
        
        # Botón para descargar reporte listo para imprimir/exportar
        csv = df_prov.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📄 Exportar Reporte a CSV / Excel",
            data=csv,
            file_name=f"Reporte_Proveedor_{prov_sel}.csv",
            mime="text/csv"
        )

# ------------------------------------------
# MÓDULO: CAPTURA Y EDICIÓN DIRECTA
# ------------------------------------------
elif opcion == "Captura / Edición Directa":
    st.header("✏️ Captura Directa y Modificación de Movimientos")
    
    tab1, tab2 = st.tabs(["➕ Capturar Nuevo Movimiento", "📝 Modificar / Eliminar Existentes"])
    
    with tab1:
        with st.form("form_captura"):
            cuentas = pd.read_sql("SELECT nombre FROM cuentas_bancarias", conn)['nombre'].tolist()
            c_cta_ban = st.selectbox("Cuenta Bancaria", cuentas if cuentas else ["Cuenta Principal"])
            c_fecha = st.date_input("Fecha", datetime.now())
            c_concepto = st.text_input("Concepto / Descripción")
            c_prov = st.text_input("Proveedor")
            c_cta = st.text_input("Cuenta Contable")
            c_subcta = st.text_input("Subcuenta")
            c_cc = st.text_input("Centro de Costo")
            c_cargo = st.number_input("Cargo (Egreso)", min_value=0.0, format="%.2f")
            c_abono = st.number_input("Abono (Ingreso)", min_value=0.0, format="%.2f")
            c_ref = st.text_input("Referencia")
            
            if st.form_submit_button("Guardar Movimiento"):
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO movimientos (cuenta_bancaria, fecha, concepto, proveedor, cuenta, subcuenta, centro_costo, cargo, abono, referencia)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (c_cta_ban, c_fecha, c_concepto, c_prov, c_cta, c_subcta, c_cc, c_cargo, c_abono, c_ref))
                conn.commit()
                st.success("¡Movimiento guardado exitosamente!")

    with tab2:
        df_edit = pd.read_sql("SELECT * FROM movimientos ORDER BY id DESC LIMIT 100", conn)
        edited_df = st.data_editor(df_edit, num_rows="dynamic", key="editor_movs")
        if st.button("Guardar Cambios en la Base de Datos"):
            edited_df.to_sql("movimientos", conn, if_exists="replace", index=False)
            st.success("¡Base de datos actualizada con éxito!")

# ------------------------------------------
# MÓDULO: CARGA DE ARCHIVOS
# ------------------------------------------
elif opcion == "Carga de Archivos (Excel/PDF)":
    st.header("📥 Ingesta de Archivos Excel")
    st.write("Sube el archivo Excel conteniendo las pestañas requeridas (`enero`, `febrero`, ..., `octubre`).")
    
    uploaded_file = st.file_uploader("Seleccionar archivo Excel", type=["xlsx", "xls"])
    if uploaded_file is not None:
        if st.button("Procesar Carga Inicial"):
            with open("temp_upload.xlsx", "wb") as f:
                f.write(uploaded_file.getbuffer())
            registros = cargar_excel_inicial("temp_upload.xlsx")
            st.success(f"Se han procesado e importado con éxito {registros} registros a la base de datos.")

# ------------------------------------------
# MÓDULO: CONFIGURACIÓN DE REGLAS
# ------------------------------------------
elif opcion == "Configuración de Reglas":
    st.header("⚙️ Reglas de Clasificación Automática")
    st.write("Define reglas para asignar automáticamente Cuenta, Subcuenta y Centro de Costos según palabras clave del concepto.")
    
    with st.form("form_regla"):
        patron = st.text_input("Palabra o patrón a buscar en el concepto (Ej. 'GASOLINA', 'NOMINA')")
        prov = st.text_input("Asignar Proveedor")
        cta = st.text_input("Asignar Cuenta")
        subcta = st.text_input("Asignar Subcuenta")
        cc = st.text_input("Asignar Centro de Costo")
        
        if st.form_submit_button("Guardar Regla"):
            cursor = conn.cursor()
            cursor.execute("""
                INSERT OR REPLACE INTO ReglasClasificacion (patron_busqueda, proveedor, cuenta, subcuenta, centro_costo)
                VALUES (?, ?, ?, ?, ?)
            """, (patron, prov, cta, subcta, cc))
            conn.commit()
            st.success("Regla guardada.")

    reglas_df = pd.read_sql("SELECT * FROM ReglasClasificacion", conn)
    st.dataframe(reglas_df, use_container_width=True)

conn.close()
