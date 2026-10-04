import sqlite3
import pandas as pd
import re

DB_NAME = "gestion_chequeras.db"

def limpiar_monto(valor):
    if pd.isna(valor) or valor is None:
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip()
    texto_limpio = re.sub(r'[^\d.-]', '', texto)
    try:
        return float(texto_limpio) if texto_limpio != '' else 0.0
    except ValueError:
        return 0.0

def inicializar_y_poblar_db(excel_path="MOVIMIENTOS 2026.xlsx"):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    # Crear tablas
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS cuentas_bancarias (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT UNIQUE NOT NULL
        )
    ''')

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

    cursor.execute("INSERT OR IGNORE INTO cuentas_bancarias (nombre) VALUES ('Cuenta Principal')")

    # Leer las pestañas de Enero a Octubre
    excel_file = pd.ExcelFile(excel_path)
    hojas = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre']

    total_registros = 0
    for sheet in excel_file.sheet_names:
        if sheet.strip().lower() in hojas:
            df = pd.read_excel(excel_path, sheet_name=sheet)
            df.columns = [str(c).strip().lower() for c in df.columns]

            for _, row in df.iterrows():
                concepto = str(row.get('concepto', row.get('descripcion', '')))
                if pd.isna(concepto) or concepto.strip() == '' or concepto == 'nan':
                    continue

                fecha = str(row.get('fecha', '2026-01-01'))[:10]
                proveedor = str(row.get('proveedor', 'Sin Proveedor'))
                cuenta = str(row.get('cuenta', 'General'))
                subcuenta = str(row.get('subcuenta', 'General'))
                centro_costo = str(row.get('centro de costo', row.get('centro_costo', 'General')))
                cargo = limpiar_monto(row.get('cargo', row.get('egreso', 0.0)))
                abono = limpiar_monto(row.get('abono', row.get('ingreso', 0.0)))
                ref = str(row.get('referencia', '')) if not pd.isna(row.get('referencia')) else ''

                cursor.execute('''
                    INSERT INTO movimientos 
                    (cuenta_bancaria, fecha, concepto, proveedor, cuenta, subcuenta, centro_costo, cargo, abono, referencia)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', ('Cuenta Principal', fecha, concepto, proveedor, cuenta, subcuenta, centro_costo, cargo, abono, ref))
                total_registros += 1

    conn.commit()
    conn.close()
    print(f"Base de datos precargada con éxito. Total de movimientos guardados: {total_registros}")

if __name__ == "__main__":
    inicializar_y_poblar_db()
