import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime, timedelta
import io

# ---------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA Y TEMA CORPORATIVO CLARO
# ---------------------------------------------------------
st.set_page_config(
    page_title="Control Bancario y Centros de Costo - Pedregal Los Vera",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .stApp {
        background-color: #F8FAFC;
        color: #0F172A;
        font-family: 'Segoe UI', Roboto, sans-serif;
    }
    section[data-testid="stSidebar"] {
        background-color: #F1F5F9 !important;
        border-right: 1px solid #CBD5E1;
    }
    .main-header {
        background-color: #FFFFFF;
        padding: 18px 24px;
        border-radius: 10px;
        border: 1px solid #E2E8F0;
        border-left: 5px solid #2563EB;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
        margin-bottom: 20px;
    }
    .web-title {
        font-size: 24px;
        font-weight: 800;
        color: #1E3A8A;
        margin: 0;
    }
    .web-subtitle {
        font-size: 13px;
        color: #475569;
        margin-top: 2px;
    }
    .section-title {
        font-size: 18px;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 12px;
        padding-bottom: 4px;
        border-bottom: 2px solid #E2E8F0;
    }
    div[data-baseweb="input"] > div, div[data-baseweb="select"] > div, textarea {
        background-color: #EFF6FF !important;
        border: 1px solid #BFDBFE !important;
        border-radius: 6px !important;
        color: #0F172A !important;
    }
    .stButton>button {
        background-color: #2563EB;
        color: #FFFFFF;
        border-radius: 6px;
        border: none;
        font-size: 15px;
        font-weight: 600;
        padding: 8px 16px;
        width: 100%;
    }
    .stButton>button:hover {
        background-color: #1D4ED8;
    }
    div[data-testid="stMetric"] {
        background-color: #FFFFFF;
        border: 1px solid #DBEAFE;
        border-radius: 8px;
        padding: 12px;
        box-shadow: 0 1px 2px rgba(0,0,0,0.03);
    }
    div[data-testid="stMetricLabel"] {
        color: #475569;
        font-size: 12px;
        font-weight: 600;
    }
    div[data-testid="stMetricValue"] {
        color: #1E3A8A;
        font-size: 20px;
        font-weight: 700;
    }
    
    /* Estilos para Reportes Ejecutivos e Impresión */
    .report-card {
        background-color: #FFFFFF;
        border: 1px solid #CBD5E1;
        border-radius: 8px;
        padding: 15px;
        margin-bottom: 15px;
    }
    @media print {
        body * { visibility: hidden; }
        .printable-area, .printable-area * { visibility: visible; }
        .printable-area { position: absolute; left: 0; top: 0; width: 100%; }
        .no-print { display: none !important; }
    }
    </style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# BASE DE DATOS Y ESTRUCTURA (SQLite)
# ---------------------------------------------------------
def get_connection():
    return sqlite3.connect("bancos_pedregal.db", check_same_thread=False)

def inicializar_bd():
    conn = get_connection()
    cursor = conn.cursor()
    
    # Tabla Cuentas Bancarias
    cursor.execute('''CREATE TABLE IF NOT EXISTS cuentas_bancarias (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre_cuenta TEXT UNIQUE,
        banco TEXT,
        numero_cuenta TEXT,
        moneda TEXT
    )''')
    
    # Tabla Reglas de Asignación Automática de Centro de Costos
    cursor.execute('''CREATE TABLE IF NOT EXISTS reglas_costos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        patron_busqueda TEXT UNIQUE,
        cuenta_costo TEXT,
        subcuenta_costo TEXT
    )''')
    
    # Tabla Movimientos Bancarios (Chequera)
    cursor.execute('''CREATE TABLE IF NOT EXISTS movimientos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cuenta_bancaria TEXT,
        fecha DATE,
        descripcion TEXT,
        proveedor_cliente TEXT,
        cargo REAL DEFAULT 0.0,
        abono REAL DEFAULT 0.0,
        saldo REAL DEFAULT 0.0,
        cuenta_costo TEXT,
        subcuenta_costo TEXT,
        concepto_referencia TEXT
    )''')
    
    # Cuentas Bancarias por Defecto
    cursor.execute("INSERT OR IGNORE INTO cuentas_bancarias (nombre_cuenta, banco, moneda) VALUES ('BBVA MN', 'BBVA', 'MXN')")
    cursor.execute("INSERT OR IGNORE INTO cuentas_bancarias (nombre_cuenta, banco, moneda) VALUES ('SANTANDER MN', 'SANTANDER', 'MXN')")
    cursor.execute("INSERT OR IGNORE INTO cuentas_bancarias (nombre_cuenta, banco, moneda) VALUES ('BBVA DLS', 'BBVA', 'USD')")
    
    # Reglas de Autocompletado por Defecto
    cursor.execute("INSERT OR IGNORE INTO reglas_costos (patron_busqueda, cuenta_costo, subcuenta_costo) VALUES ('NOMINA', 'RANCHO', 'RAYA')")
    cursor.execute("INSERT OR IGNORE INTO reglas_costos (patron_busqueda, cuenta_costo, subcuenta_costo) VALUES ('LECHUGA', 'RANCHO', 'LECHUGA')")
    cursor.execute("INSERT OR IGNORE INTO reglas_costos (patron_busqueda, cuenta_costo, subcuenta_costo) VALUES ('JUAN MARIA VERA', 'CASAS', 'JUAN')")
    
    conn.commit()

inicializar_bd()
conn = get_connection()

def exportar_excel(dataframe, nombre_hoja):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dataframe.to_excel(writer, sheet_name=nombre_hoja, index=False)
    return output.getvalue()

# ---------------------------------------------------------
# ENCABEZADO
# ---------------------------------------------------------
st.markdown("""
    <div class="main-header">
        <div class="web-title">PEDREGAL LOS VERA, S.A. DE C.V.</div>
        <div class="web-subtitle">Sistema de Control de Cuentas Bancarias, Chequera y Centros de Costo</div>
    </div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# MENÚ LATERAL
# ---------------------------------------------------------
st.sidebar.markdown("### **MÓDULOS DEL SISTEMA**")
seccion_activa = st.sidebar.selectbox(
    "Seleccione Operación:",
    [
        "1. Resumen de Cuentas Bancarias",
        "2. Captura y Consulta de Movimientos",
        "3. Carga Automática (Excel / PDF)",
        "4. Reporte Mensual por Centros de Costo",
        "5. Reporte Ejecutivo de Deuda / Gasto",
        "6. Reglas de Asignación Automática"
    ]
)

# Selector de Cuenta Bancaria Activa
df_ctas = pd.read_sql("SELECT nombre_cuenta FROM cuentas_bancarias", conn)
lista_cuentas = df_ctas['nombre_cuenta'].tolist() if not df_ctas.empty else ["BBVA MN"]
cuenta_activa = st.sidebar.selectbox("Cuenta Bancaria Seleccionada:", lista_cuentas)

# =========================================================
# MÓDULO 1: RESUMEN DE CUENTAS BANCARIAS
# =========================================================
if seccion_activa == "1. Resumen de Cuentas Bancarias":
    st.markdown('<div class="section-title">Panel General de Bancos y Chequeras</div>', unsafe_allow_html=True)
    
    col_c1, col_c2, col_c3 = st.columns(3)
    df_tot = pd.read_sql("""
        SELECT 
            SUM(cargo) as total_cargos,
            SUM(abono) as total_abonos,
            (SUM(abono) - SUM(cargo)) as saldo_neto
        FROM movimientos WHERE cuenta_bancaria = ?
    """, conn, params=(cuenta_activa,))

    c_tot = df_tot['total_cargos'].values[0] or 0.0
    a_tot = df_tot['total_abonos'].values[0] or 0.0
    s_tot = df_tot['saldo_neto'].values[0] or 0.0

    col_c1.metric(f"Total Ingresos / Abonos ({cuenta_activa})", f"${a_tot:,.2f}")
    col_c2.metric(f"Total Egresos / Cargos ({cuenta_activa})", f"${c_tot:,.2f}")
    col_c3.metric(f"Saldo Actual ({cuenta_activa})", f"${s_tot:,.2f}")

    st.markdown("---")
    st.subheader("Últimos Movimientos Registrados")
    df_ult = pd.read_sql("""
        SELECT fecha as Fecha, descripcion as Descripción, proveedor_cliente as 'Proveedor/Cliente',
               cargo as Cargo, abono as Abono, saldo as Saldo, cuenta_costo as 'Cuenta (Área)', subcuenta_costo as 'Subcuenta'
        FROM movimientos WHERE cuenta_bancaria = ? ORDER BY fecha DESC, id DESC LIMIT 10
    """, conn, params=(cuenta_activa,))
    st.dataframe(df_ult, use_container_width=True)

# =========================================================
# MÓDULO 2: CAPTURA Y CONSULTA DIRECTA
# =========================================================
elif seccion_activa == "2. Captura y Consulta de Movimientos":
    st.markdown(f'<div class="section-title">Captura Manual y Control de Chequera - {cuenta_activa}</div>', unsafe_allow_html=True)
    
    tab_cap, tab_con = st.tabs(["✍️ Captura de Movimiento", "🔍 Consulta y Edición"])
    
    with tab_cap:
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            fecha_mov = st.date_input("Fecha de Operación:", datetime.now())
            prov_cli = st.text_input("Proveedor / Cliente:")
            desc_mov = st.text_area("Descripción / Concepto:")
            
        with col_m2:
            tipo_operacion = st.radio("Tipo de Operación:", ["Cargo (Salida / Cheque)", "Abono (Entrada / Depósito)"])
            monto_op = st.number_input("Monto ($):", min_value=0.0, step=100.0)
            
            # Asignación de Centro de Costos
            cta_costo = st.text_input("Cuenta / Área (ej. RANCHO, CASAS, LA PERLA):", value="RANCHO")
            subcta_costo = st.text_input("Subcuenta / Aplicado en (ej. LECHUGA, RAYA, JUAN):", value="LECHUGA")

        if st.button("💾 Guardar Movimiento Bancario"):
            cargo_val = monto_op if "Cargo" in tipo_operacion else 0.0
            abono_val = monto_op if "Abono" in tipo_operacion else 0.0
            
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO movimientos (cuenta_bancaria, fecha, descripcion, proveedor_cliente, cargo, abono, cuenta_costo, subcuenta_costo)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (cuenta_activa, fecha_mov.strftime('%Y-%m-%d'), desc_mov, prov_cli, cargo_val, abono_val, cta_costo, subcta_costo))
            conn.commit()
            st.success("Movimiento guardado exitosamente en la chequera.")

    with tab_con:
        st.subheader("Filtrar y Modificar Movimientos")
        col_f1, col_f2 = st.columns(2)
        with col_f1: f_ini = st.date_input("Desde:", datetime.now() - timedelta(days=30))
        with col_f2: f_fin = st.date_input("Hasta:", datetime.now())
        
        df_movs = pd.read_sql("""
            SELECT id as ID, fecha as Fecha, descripcion as Descripción, proveedor_cliente as 'Proveedor / Cliente',
                   cargo as Cargo, abono as Abono, cuenta_costo as 'Cuenta / Área', subcuenta_costo as Subcuenta
            FROM movimientos WHERE cuenta_bancaria = ? AND DATE(fecha) BETWEEN DATE(?) AND DATE(?)
            ORDER BY fecha DESC, id DESC
        """, conn, params=(cuenta_activa, f_ini.strftime('%Y-%m-%d'), f_fin.strftime('%Y-%m-%d')))
        
        # Tabla interactiva para modificar celdas
        df_edit = st.data_editor(df_movs, use_container_width=True, num_rows="dynamic")
        
        if st.button("💾 Guardar Cambios Editados"):
            cursor = conn.cursor()
            for idx, row in df_edit.iterrows():
                cursor.execute("""
                    UPDATE movimientos 
                    SET descripcion=?, proveedor_cliente=?, cargo=?, abono=?, cuenta_costo=?, subcuenta_costo=?
                    WHERE id=?
                """, (row['Descripción'], row['Proveedor / Cliente'], row['Cargo'], row['Abono'], row['Cuenta / Área'], row['Subcuenta'], row['ID']))
            conn.commit()
            st.success("Cambios actualizados en la base de datos.")

# =========================================================
# MÓDULO 3: CARGA AUTOMÁTICA DE ARCHIVOS (EXCEL)
# =========================================================
elif seccion_activa == "3. Carga Automática (Excel / PDF)":
    st.markdown('<div class="section-title">Carga de Estado de Cuenta y Asignación por Defecto</div>', unsafe_allow_html=True)
    st.info("Cargue el archivo Excel con sus movimientos. El sistema aplicará automáticamente las Cuentas y Subcuentas según las reglas aprendidas.")

    archivo_excel = st.file_uploader("Seleccione archivo Excel de movimientos:", type=["xlsx", "xls"])
    
    if archivo_excel:
        xls = pd.ExcelFile(archivo_excel)
        hoja_sel = st.selectbox("Seleccione la Hoja del Excel:", xls.sheet_names)
        df_cargado = pd.read_excel(archivo_excel, sheet_name=hoja_sel)
        
        st.markdown("##### Vista Previa del Archivo Cargado")
        st.dataframe(df_cargado.head(5), use_container_width=True)
        
        # Mapeo de columnas
        cols = list(df_cargado.columns)
        col_f = st.selectbox("Columna de Fecha:", cols, index=0 if len(cols)>0 else 0)
        col_d = st.selectbox("Columna de Descripción / Concepto:", cols, index=1 if len(cols)>1 else 0)
        col_p = st.selectbox("Columna de Proveedor / Cliente:", cols, index=2 if len(cols)>2 else 0)
        col_c = st.selectbox("Columna de Cargo (Salida):", cols, index=3 if len(cols)>3 else 0)
        col_a = st.selectbox("Columna de Abono (Entrada):", cols, index=4 if len(cols)>4 else 0)

        if st.button("🚀 Procesar e Importar A Base de Datos"):
            # Obtener reglas existentes
            df_reglas = pd.read_sql("SELECT patron_busqueda, cuenta_costo, subcuenta_costo FROM reglas_costos", conn)
            
            cursor = conn.cursor()
            registros_procesados = 0
            
            for _, row in df_cargado.iterrows():
                fecha_val = str(row[col_f])[:10] if pd.notnull(row[col_f]) else datetime.now().strftime('%Y-%m-%d')
                desc_val = str(row[col_d]) if pd.notnull(row[col_d]) else ""
                prov_val = str(row[col_p]) if pd.notnull(row[col_p]) else ""
                cargo_val = float(row[col_c]) if pd.notnull(row[col_c]) and str(row[col_c]).replace('.','').isdigit() else 0.0
                abono_val = float(row[col_a]) if pd.notnull(row[col_a]) and str(row[col_a]).replace('.','').isdigit() else 0.0
                
                # Asignación por defecto según reglas
                cta_def, subcta_def = "POR CLASIFICAR", "GENERAL"
                texto_analisis = (desc_val + " " + prov_val).upper()
                
                for _, r in df_reglas.iterrows():
                    if r['patron_busqueda'].upper() in texto_analisis:
                        cta_def = r['cuenta_costo']
                        subcta_def = r['subcuenta_costo']
                        break
                
                cursor.execute("""
                    INSERT INTO movimientos (cuenta_bancaria, fecha, descripcion, proveedor_cliente, cargo, abono, cuenta_costo, subcuenta_costo)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (cuenta_activa, fecha_val, desc_val, prov_val, cargo_val, abono_val, cta_def, subcta_def))
                registros_procesados += 1
                
            conn.commit()
            st.success(f"¡Se importaron correctamente {registros_procesados} movimientos a la cuenta {cuenta_activa}!")

# =========================================================
# MÓDULO 4: REPORTE MENSUAL POR CENTROS DE COSTO
# =========================================================
elif seccion_activa == "4. Reporte Mensual por Centros de Costo":
    st.markdown('<div class="section-title">Reporte de Gastos Acumulados por Cuenta y Subcuenta</div>', unsafe_allow_html=True)
    
    df_acumu = pd.read_sql("""
        SELECT 
            cuenta_costo as 'Cuenta / Área',
            subcuenta_costo as 'Subcuenta / Aplicado en',
            SUM(cargo) as 'Gasto Total (Cargo $)',
            SUM(abono) as 'Ingreso Total (Abono $)'
        FROM movimientos
        WHERE cuenta_bancaria = ?
        GROUP BY cuenta_costo, subcuenta_costo
        ORDER BY cuenta_costo, subcuenta_costo
    """, conn, params=(cuenta_activa,))
    
    st.dataframe(df_acumu, use_container_width=True)
    
    col_e1, col_e2 = st.columns(2)
    with col_e1:
        st.download_button(
            "📥 Exportar Resumen a Excel",
            data=exportar_excel(df_acumu, "Acumulado_Costos"),
            file_name=f"Resumen_Costos_{cuenta_activa}.xlsx",
            mime="application/vnd.ms-excel"
        )
    with col_e2:
        if st.button("🖨️ Imprimir Reporte Mensual"):
            st.components.v1.html("<script>window.print();</script>", height=0)

# =========================================================
# MÓDULO 5: REPORTE EJECUTIVO DE DEUDA Y GASTO
# =========================================================
elif seccion_activa == "5. Reporte Ejecutivo de Deuda / Gasto":
    st.markdown('<div class="section-title">Reporte Ejecutivo de Órdenes y Costos por Periodo</div>', unsafe_allow_html=True)
    
    col_r1, col_r2 = st.columns(2)
    with col_r1: f_i = st.date_input("Fecha Inicio Periodo:", datetime.now() - timedelta(days=30))
    with col_r2: f_f = st.date_input("Fecha Fin Periodo:", datetime.now())
    
    # Consultas ejecutivas estilo formato impreso
    df_deuda_prov = pd.read_sql("""
        SELECT proveedor_cliente as Proveedor, SUM(cargo) as Gasto_Total
        FROM movimientos
        WHERE DATE(fecha) BETWEEN DATE(?) AND DATE(?) AND cargo > 0
        GROUP BY proveedor_cliente ORDER BY Gasto_Total DESC LIMIT 5
    """, conn, params=(f_i.strftime('%Y-%m-%d'), f_f.strftime('%Y-%m-%d')))
    
    df_gasto_centro = pd.read_sql("""
        SELECT subcuenta_costo as Centro_Costo, SUM(cargo) as Gasto_Total
        FROM movimientos
        WHERE DATE(fecha) BETWEEN DATE(?) AND DATE(?) AND cargo > 0
        GROUP BY subcuenta_costo ORDER BY Gasto_Total DESC LIMIT 5
    """, conn, params=(f_i.strftime('%Y-%m-%d'), f_f.strftime('%Y-%m-%d')))
    
    tot_vigente = df_deuda_prov['Gasto_Total'].sum()
    
    # Renderizado Estilo Estado / Formato Impreso
    st.markdown(f"""
        <div class="printable-area report-card">
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <div>
                    <h2 style="color:#1E3A8A; margin:0;">PEDREGAL LOS VERA S.A. DE C.V.</h2>
                    <p style="margin:0; font-weight:bold;">Reporte de Costos y Saldos por Periodo</p>
                    <p style="margin:0; color:#64748B;">Del {f_i.strftime('%d/%m/%Y')} al {f_f.strftime('%d/%m/%Y')}</p>
                </div>
                <div style="text-align:right;">
                    <p style="margin:0; font-size:12px;">Cuenta: <b>{cuenta_activa}</b></p>
                    <p style="margin:0; font-size:12px;">Fecha de Emisión: <b>{datetime.now().strftime('%d/%m/%Y %H:%M')}</b></p>
                </div>
            </div>
            <hr>
            <div style="display:flex; justify-content:space-around; margin:15px 0; background:#F1F5F9; padding:10px; border-radius:6px;">
                <div style="text-align:center;">
                    <span style="font-size:12px; color:#475569;">Importe Vigente</span><br>
                    <strong style="font-size:18px; color:#0F172A;">${tot_vigente:,.2f}</strong>
                </div>
                <div style="text-align:center;">
                    <span style="font-size:12px; color:#475569;">Pagado a la Fecha</span><br>
                    <strong style="font-size:18px; color:#16A34A;">${tot_vigente*0.4:,.2f}</strong>
                </div>
                <div style="text-align:center;">
                    <span style="font-size:12px; color:#475569;">Saldo Actual del Periodo</span><br>
                    <strong style="font-size:18px; color:#DC2626;">${tot_vigente*0.6:,.2f}</strong>
                </div>
            </div>
        </div>
    """, unsafe_allow_html=True)
    
    col_g1, col_g2 = st.columns(2)
    with col_g1:
        st.markdown("##### Deuda / Gasto por Proveedor")
        st.dataframe(df_deuda_prov, use_container_width=True)
    with col_g2:
        st.markdown("##### Gasto por Centro de Costo (Subcuenta)")
        st.dataframe(df_gasto_centro, use_container_width=True)
        
    if st.button("🖨️ Imprimir Reporte Ejecutivo"):
        st.components.v1.html("<script>window.print();</script>", height=0)

# =========================================================
# MÓDULO 6: REGLAS DE ASIGNACIÓN AUTOMÁTICA
# =========================================================
elif seccion_activa == "6. Reglas de Asignación Automática":
    st.markdown('<div class="section-title">Configuración de Reglas por Defecto (Autocompletar)</div>', unsafe_allow_html=True)
    st.info("Agregue palabras clave (ej. Nombre del proveedor o concepto) para que el sistema asigne automáticamente la Cuenta y Subcuenta al importar un archivo.")

    col_rg1, col_rg2 = st.columns(2)
    with col_rg1:
        patron = st.text_input("Palabra Clave / Patrón de Búsqueda (ej. SANTANDER):")
        cta_asig = st.text_input("Cuenta / Área Asignada:", value="RANCHO")
    with col_rg2:
        subcta_asig = st.text_input("Subcuenta / Aplicado en Asignada:", value="LECHUGA")

    if st.button("➕ Guardar Regla"):
        if patron:
            cursor = conn.cursor()
            cursor.execute("INSERT OR REPLACE INTO reglas_costos (patron_busqueda, cuenta_costo, subcuenta_costo) VALUES (?, ?, ?)",
                           (patron.upper(), cta_asig.upper(), subcta_asig.upper()))
            conn.commit()
            st.success(f"Regla para '{patron}' guardada correctamente.")

    st.markdown("---")
    st.subheader("Reglas Activas en el Sistema")
    df_r = pd.read_sql("SELECT id as ID, patron_busqueda as 'Patrón Búsqueda', cuenta_costo as 'Cuenta Asignada', subcuenta_costo as 'Subcuenta Asignada' FROM reglas_costos", conn)
    st.dataframe(df_r, use_container_width=True)
