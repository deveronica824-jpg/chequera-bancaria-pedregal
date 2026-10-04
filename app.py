import sqlite3
import pandas as pd
import streamlit as st
from datetime import datetime

DB_NAME = "gestion_chequeras.db"

# ==========================================
# 1. CONEXIÓN Y CREACIÓN DE TABLAS SEGURO
# ==========================================
def obtener_conexion():
    return sqlite3.connect(DB_NAME, check_same_thread=False)

def init_db():
    try:
        conn = obtener_conexion()
        cursor = conn.cursor()
        
        # Tabla de Cuentas Bancarias
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS cuentas_bancarias (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT UNIQUE NOT NULL
            )
        ''')
        
        # Tabla de Movimientos Bancarios
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS movimientos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cuenta_bancaria TEXT NOT NULL,
                fecha TEXT NOT NULL,
                concepto TEXT,
                proveedor TEXT,
                cuenta TEXT,
                subcuenta TEXT,
                centro_costo TEXT,
                cargo REAL DEFAULT 0.0,
                abono REAL DEFAULT 0.0,
                referencia TEXT
            )
        ''')
        
        # Insertar cuenta base si no existe
        cursor.execute("INSERT OR IGNORE INTO cuentas_bancarias (nombre) VALUES ('Cuenta Principal')")
        
        conn.commit()
        conn.close()
    except Exception as e:
        st.error(f"Error al inicializar la base de datos: {e}")

# Inicializar DB en cada arranque de la app
init_db()

# ==========================================
# 2. INTERFAZ Y NAVEGACIÓN
# ==========================================
st.set_page_config(page_title="Gestión de Chequeras y Costos", layout="wide")
st.title("🏦 Sistema de Gestión Multicuenta y Control de Costos")

opcion = st.sidebar.selectbox("Módulo Principal", [
    "📊 Resumen Mensual & Dashboard",
    "🔍 Reporte por Proveedor",
    "✏️ Captura y Edición Directa",
    "📥 Carga de Datos (Excel)",
    "➕ Gestión de Cuentas Bancarias"
])

conn = obtener_conexion()

# ------------------------------------------
# MÓDULO 1: RESUMEN MENSUAL Y DASHBOARD
# ------------------------------------------
if opcion == "📊 Resumen Mensual & Dashboard":
    st.header("📊 Resumen Mensual y Estado de Cuentas")
    
    try:
        cuentas_df = pd.read_sql("SELECT nombre FROM cuentas_bancarias", conn)
        cuentas = cuentas_df['nombre'].tolist() if not cuentas_df.empty else ["Cuenta Principal"]
    except Exception:
        cuentas = ["Cuenta Principal"]
        
    cuenta_sel = st.selectbox("Seleccionar Cuenta Bancaria", ["Todas"] + cuentas)
    
    query = "SELECT * FROM movimientos"
    if cuenta_sel != "Todas":
        query += f" WHERE cuenta_bancaria = '{cuenta_sel}'"
        
    df = pd.read_sql(query, conn)
    
    if not df.empty:
        df['cargo'] = pd.to_numeric(df['cargo'], errors='coerce').fillna(0)
        df['abono'] = pd.to_numeric(df['abono'], errors='coerce').fillna(0)
        df['fecha_dt'] = pd.to_datetime(df['fecha'], errors='coerce')
        df['Mes-Año'] = df['fecha_dt'].dt.to_period('M').astype(str)
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Egresos (Cargos)", f"${df['cargo'].sum():,.2f}")
        col2.metric("Total Ingresos (Abonos)", f"${df['abono'].sum():,.2f}")
        col3.metric("Balance Neto", f"${(df['abono'].sum() - df['cargo'].sum()):,.2f}")
        
        st.markdown("---")
        st.subheader("Resumen Agrupado por Mes")
        resumen_mes = df.groupby('Mes-Año')[['cargo', 'abono']].sum()
        st.bar_chart(resumen_mes)
        st.dataframe(resumen_mes.style.format("${:,.2f}"), use_container_width=True)
        
        st.subheader("Detalle por Centros de Costo")
        resumen_cc = df.groupby('centro_costo')[['cargo', 'abono']].sum()
        st.dataframe(resumen_cc.style.format("${:,.2f}"), use_container_width=True)
    else:
        st.info("La base de datos está vacía. Por favor dirígete al módulo 'Carga de Datos (Excel)' para importar tu archivo.")

# ------------------------------------------
# MÓDULO 2: REPORTE POR PROVEEDOR
# ------------------------------------------
elif opcion == "🔍 Reporte por Proveedor":
    st.header("🔍 Consulta y Reporte por Proveedor")
    
    df_provs = pd.read_sql("SELECT DISTINCT proveedor FROM movimientos WHERE proveedor IS NOT NULL AND proveedor != '' AND proveedor != 'nan'", conn)
    prov_list = sorted(df_provs['proveedor'].tolist()) if not df_provs.empty else []
    
    if prov_list:
        col1, col2, col3 = st.columns(3)
        with col1:
            prov_sel = st.selectbox("Seleccionar Proveedor", prov_list)
        with col2:
            fecha_inicio = st.date_input("Fecha Inicio", datetime(2026, 1, 1))
        with col3:
            fecha_fin = st.date_input("Fecha Fin", datetime(2026, 10, 31))
            
        if st.button("Generar Reporte"):
            q = f"""
                SELECT fecha, cuenta_bancaria, concepto, cuenta, subcuenta, centro_costo, cargo, abono, referencia 
                FROM movimientos 
                WHERE proveedor = '{prov_sel}' 
                AND fecha BETWEEN '{fecha_inicio}' AND '{fecha_fin}'
                ORDER BY fecha ASC
            """
            df_res = pd.read_sql(q, conn)
            df_res['cargo'] = pd.to_numeric(df_res['cargo'], errors='coerce').fillna(0)
            df_res['abono'] = pd.to_numeric(df_res['abono'], errors='coerce').fillna(0)
            
            st.write(f"### Movimientos registrados de: **{prov_sel}**")
            st.dataframe(df_res, use_container_width=True)
            
            c1, c2 = st.columns(2)
            c1.metric("Total Pagado (Cargos)", f"${df_res['cargo'].sum():,.2f}")
            c2.metric("Total Abonado", f"${df_res['abono'].sum():,.2f}")
            
            csv = df_res.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📄 Exportar / Imprimir Reporte (CSV)",
                data=csv,
                file_name=f"Reporte_Proveedor_{prov_sel}.csv",
                mime="text/csv"
            )
    else:
        st.warning("No hay proveedores registrados aún en la base de datos.")

# ------------------------------------------
# MÓDULO 3: CAPTURA Y EDICIÓN DIRECTA
# ------------------------------------------
elif opcion == "✏️ Captura y Edición Directa":
    st.header("✏️ Captura Directa y Modificación de Movimientos")
    
    tab1, tab2 = st.tabs(["➕ Registrar Nuevo Movimiento", "📝 Modificar Base de Datos"])
    
    with tab1:
        with st.form("form_captura"):
            cuentas_df = pd.read_sql("SELECT nombre FROM cuentas_bancarias", conn)
            cuentas = cuentas_df['nombre'].tolist() if not cuentas_df.empty else ["Cuenta Principal"]
            
            c_cta_ban = st.selectbox("Cuenta Bancaria", cuentas)
            c_fecha = st.date_input("Fecha", datetime.now())
            c_concepto = st.text_input("Concepto / Descripción")
            c_prov = st.text_input("Proveedor")
            c_cta = st.text_input("Cuenta Contable")
            c_subcta = st.text_input("Subcuenta")
            c_cc = st.text_input("Centro de Costo")
            c_cargo = st.number_input("Cargo (Egreso)", min_value=0.0, format="%.2f")
            c_abono = st.number_input("Abono (Ingreso)", min_value=0.0, format="%.2f")
            c_ref = st.text_input("Referencia")
            
            if st.form_submit_button("Guardar en Base de Datos"):
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO movimientos (cuenta_bancaria, fecha, concepto, proveedor, cuenta, subcuenta, centro_costo, cargo, abono, referencia)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (c_cta_ban, str(c_fecha), c_concepto, c_prov, c_cta, c_subcta, c_cc, c_cargo, c_abono, c_ref))
                conn.commit()
                st.success("¡Movimiento registrado correctamente!")

    with tab2:
        df_edit = pd.read_sql("SELECT * FROM movimientos ORDER BY id DESC LIMIT 200", conn)
        if not df_edit.empty:
            edited_df = st.data_editor(df_edit, num_rows="dynamic", key="editor_movs")
            if st.button("Guardar Cambios"):
                edited_df.to_sql("movimientos", conn, if_exists="replace", index=False)
                st.success("¡Base de datos actualizada correctamente!")
        else:
            st.info("No hay movimientos para editar.")

# ------------------------------------------
# MÓDULO 4: CARGA DE DATOS EXCEL
# ------------------------------------------
elif opcion == "📥 Carga de Datos (Excel)":
    st.header("📥 Ingesta e Importación de Archivo Excel")
    st.write("Carga tu archivo Excel (`MOVIMIENTOS 2026.xlsx`) para poblar la base de datos con las hojas de enero a octubre.")
    
    uploaded_file = st.file_uploader("Seleccionar archivo Excel", type=["xlsx", "xls"])
    if uploaded_file is not None:
        if st.button("Procesar y Cargar a Base de Datos"):
            try:
                excel = pd.ExcelFile(uploaded_file)
                hojas = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre']
                cursor = conn.cursor()
                total = 0
                
                for sheet in excel.sheet_names:
                    if sheet.strip().lower() in hojas:
                        df_sheet = pd.read_excel(uploaded_file, sheet_name=sheet)
                        df_sheet.columns = [str(c).strip().lower() for c in df_sheet.columns]
                        
                        for _, row in df_sheet.iterrows():
                            concepto = str(row.get('concepto', row.get('descripcion', '')))
                            if pd.isna(concepto) or concepto.strip() in ['', 'nan']:
                                continue
                            
                            fecha = str(row.get('fecha', '2026-01-01'))[:10]
                            prov = str(row.get('proveedor', 'Sin Proveedor'))
                            cta = str(row.get('cuenta', 'General'))
                            subcta = str(row.get('subcuenta', 'General'))
                            cc = str(row.get('centro de costo', row.get('centro_costo', 'General')))
                            
                            cargo_raw = str(row.get('cargo', row.get('egreso', 0))).replace('$', '').replace(',', '').strip()
                            abono_raw = str(row.get('abono', row.get('ingreso', 0))).replace('$', '').replace(',', '').strip()
                            
                            cargo = float(cargo_raw) if cargo_raw not in ['', 'nan', 'None'] else 0.0
                            abono = float(abono_raw) if abono_raw not in ['', 'nan', 'None'] else 0.0
                            ref = str(row.get('referencia', '')) if not pd.isna(row.get('referencia')) else ''

                            cursor.execute('''
                                INSERT INTO movimientos 
                                (cuenta_bancaria, fecha, concepto, proveedor, cuenta, subcuenta, centro_costo, cargo, abono, referencia)
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            ''', ('Cuenta Principal', fecha, concepto, prov, cta, subcta, cc, cargo, abono, ref))
                            total += 1
                            
                conn.commit()
                st.success(f"¡Se han importado exitosamente {total} registros a la base de datos!")
            except Exception as e:
                st.error(f"Error procesando el archivo: {e}")

# ------------------------------------------
# MÓDULO 5: GESTIÓN MULTICUENTA
# ------------------------------------------
elif opcion == "➕ Gestión de Cuentas Bancarias":
    st.header("⚙️ Configuración Multicuenta")
    
    with st.form("form_cuenta"):
        nueva_cuenta = st.text_input("Nombre de la nueva cuenta bancaria / chequera")
        if st.form_submit_button("Agregar Cuenta"):
            if nueva_cuenta:
                cursor = conn.cursor()
                cursor.execute("INSERT OR IGNORE INTO cuentas_bancarias (nombre) VALUES (?)", (nueva_cuenta,))
                conn.commit()
                st.success(f"Cuenta '{nueva_cuenta}' agregada exitosamente.")
                
    st.subheader("Cuentas Registradas")
    df_ctas = pd.read_sql("SELECT * FROM cuentas_bancarias", conn)
    st.dataframe(df_ctas, use_container_width=True)

conn.close()
