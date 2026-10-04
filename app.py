import sqlite3
import pandas as pd
import streamlit as st
import re
from datetime import datetime

# ==========================================
# FUNCIÓN AUXILIAR DE LIMPIEZA NUMÉRICA
# ==========================================
def limpiar_monto(valor):
    """Convierte de forma segura textos, vacíos o formatos de moneda a float."""
    if pd.isna(valor) or valor is None:
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    
    # Si es texto, limpiar símbolos de moneda, comas y espacios
    texto = str(valor).strip()
    texto_limpio = re.sub(r'[^\d.-]', '', texto)
    
    try:
        return float(texto_limpio) if texto_limpio != '' else 0.0
    except ValueError:
        return 0.0

# ==========================================
# LÓGICA DE INGESTA MEJORADA
# ==========================================
def cargar_excel_inicial(uploaded_file):
    """Carga y consolida hojas de Enero a Octubre de un libro Excel cargado."""
    excel_file = pd.ExcelFile(uploaded_file)
    hojas_objetivo = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre']
    
    conn = sqlite3.connect("gestion_chequeras.db")
    cursor = conn.cursor()
    
    cursor.execute("INSERT OR IGNORE INTO cuentas_bancarias (nombre) VALUES ('Cuenta Principal')")
    
    total_registros = 0
    for sheet in excel_file.sheet_names:
        if sheet.strip().lower() in hojas_objetivo:
            df = pd.read_excel(excel_file, sheet_name=sheet)
            
            # Normalizar nombres de columnas (minúsculas y sin espacios extra)
            df.columns = [str(c).strip().lower() for c in df.columns]
            
            for _, row in df.iterrows():
                concepto = str(row.get('concepto', row.get('descripcion', 'Movimiento')))
                if pd.isna(concepto) or concepto.strip() == '':
                    continue  # Omitir filas totalmente vacías
                
                prov, cta, subcta, cc = clasificar_automatico(concepto)
                
                fecha = row.get('fecha', datetime.now().strftime('%Y-%m-%d'))
                if pd.isna(fecha):
                    fecha = datetime.now().strftime('%Y-%m-%d')
                
                # Conversión segura de valores numéricos
                val_cargo = row.get('cargo', row.get('egreso', 0.0))
                val_abono = row.get('abono', row.get('ingreso', 0.0))
                
                cargo = limpiar_monto(val_cargo)
                abono = limpiar_monto(val_abono)
                ref = str(row.get('referencia', '')) if not pd.isna(row.get('referencia')) else ''

                cursor.execute('''
                    INSERT INTO movimientos 
                    (cuenta_bancaria, fecha, concepto, proveedor, cuenta, subcuenta, centro_costo, cargo, abono, referencia)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', ('Cuenta Principal', str(fecha)[:10], concepto, prov, cta, subcta, cc, cargo, abono, ref))
                total_registros += 1

    conn.commit()
    conn.close()
    return total_registros
