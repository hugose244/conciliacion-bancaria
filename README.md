# conciliacion-bancaria
import pandas as pd
from fuzzywuzzy import fuzz

# Función para cargar datos desde archivos
def cargar_datos(transacciones_path, contabilidad_path):
    transacciones = pd.read_excel(transacciones_path)
    contabilidad = pd.read_excel(contabilidad_path)
    return transacciones, contabilidad

# Función para determinar similitud entre conceptos
def comparar_conceptos(concepto1, concepto2, threshold=80):
    if pd.isna(concepto1) or pd.isna(concepto2):
        return False
    return fuzz.ratio(concepto1.lower(), concepto2.lower()) >= threshold

# Función para clasificar diferencias
def clasificar_diferencias(row):
    if pd.isna(row['Debe_Banco']) and pd.isna(row['Haber_Banco']):
        return 'Partida Conciliatoria (+)' if row['Debe_Contabilidad'] > 0 else 'Partida Conciliatoria (-)'
    if pd.isna(row['Debe_Contabilidad']) and pd.isna(row['Haber_Contabilidad']):
        return 'Partida de Ajuste (+)' if row['Debe_Banco'] > 0 else 'Partida de Ajuste (-)'
    if row['Diferencia_Monto'] != 0 or row['Diferencia_Concepto'] == 'Sí':
        return 'Partida de Ajuste'
    return 'Conciliado'

# Función para realizar la conciliación bancaria con comparaciones avanzadas
def conciliar_bancos(transacciones, contabilidad):
    conciliacion = transacciones.merge(contabilidad, on='Referencia', how='outer', indicator=True, suffixes=('_Banco', '_Contabilidad'))
    
    conciliacion['Diferencia_Monto'] = (
        conciliacion[['Debe_Banco', 'Haber_Banco']].sum(axis=1, skipna=True) - 
        conciliacion[['Debe_Contabilidad', 'Haber_Contabilidad']].sum(axis=1, skipna=True)
    )
    
    conciliacion['Diferencia_Concepto'] = conciliacion.apply(
        lambda row: 'Sí' if comparar_conceptos(row['Concepto_Banco'], row['Concepto_Contabilidad']) is False else 'No',
        axis=1
    )
    
    conciliacion['Estado'] = conciliacion.apply(clasificar_diferencias, axis=1)
    return conciliacion.drop(columns=['_merge'])

# Función para generar el reporte en Excel
def generar_reporte(conciliacion, output_path):
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        conciliacion.to_excel(writer, sheet_name='Conciliacion', index=False)
    print(f"✅ Reporte generado: {output_path}")
