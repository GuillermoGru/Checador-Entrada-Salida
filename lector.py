# lector.py

# ------Importaciones------
# Traemos pandas para leer datos y herramientas de openpyxl para dibujar en el Excel
import pandas as pd
from pathlib import Path
from openpyxl.styles import Border, Side, Alignment, Font

# ------Lectura de Archivos Nuevos------

def cargar_eventos(ruta_archivo: Path) -> pd.DataFrame:
    # Lee el Excel de la checadora normal saltándose la primera fila para tomar bien los títulos
    try:
        df = pd.read_excel(ruta_archivo, header=1)
        if 'Género' in df.columns:
            df['Género'] = df['Género'].fillna('No especificado')
        return df
    except Exception as e:
        print(f"[Error] Fallo al leer Eventos '{ruta_archivo}': {e}")
        return None

def cargar_eventos_veladores(ruta_archivo: Path) -> pd.DataFrame:
    # Hace exactamente lo mismo pero exclusivo para el archivo de los veladores
    try:
        df = pd.read_excel(ruta_archivo, header=1)
        return df
    except Exception as e:
        print(f"[Error] Fallo al leer Veladores '{ruta_archivo}': {e}")
        return None

# ------Lectura de la Base de Datos------

def cargar_base_datos(ruta_archivo: Path) -> dict:
    # Abre A2026.xlsx y carga todas sus pestañas en la memoria al mismo tiempo
    try:
        dfs = pd.read_excel(ruta_archivo, sheet_name=None, header=1)
        return dfs
    except Exception as e:
        print(f"[Error] Fallo al leer Base de Datos '{ruta_archivo}': {e}")
        return None

# ------Escritura y Diseño Visual en Excel------

def actualizar_hoja_bd(ruta_bd: Path, nombre_hoja: str, df_nuevo: pd.DataFrame):
    try:
        # Importamos openpyxl explícitamente para borrar filas previas
        import openpyxl
        wb = openpyxl.load_workbook(ruta_bd)
        if nombre_hoja in wb.sheetnames:
            ws = wb[nombre_hoja]
            # Borramos todas las filas de datos
            if ws.max_row >= 3:
                ws.delete_rows(3, ws.max_row - 2)
        wb.save(ruta_bd)
        wb.close()
        
        # Pegamos la información limpia
        with pd.ExcelWriter(ruta_bd, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            df_nuevo.to_excel(writer, sheet_name=nombre_hoja, index=False, header=False, startrow=2)
            worksheet = writer.sheets[nombre_hoja]
            
            borde_delgado = Border(
                left=Side(style='thin'), right=Side(style='thin'), 
                top=Side(style='thin'), bottom=Side(style='thin')
            )
            fuente_estandar = Font(name='Arial', size=11, bold=False)
            alineacion_izq = Alignment(horizontal='left', vertical='center')
            
            maxima_fila = 2 + len(df_nuevo)
            maxima_columna = len(df_nuevo.columns)
            
            for fila in worksheet.iter_rows(min_row=3, max_row=maxima_fila, min_col=1, max_col=maxima_columna):
                for celda in fila:
                    celda.border = borde_delgado
                    celda.font = fuente_estandar
                    celda.alignment = alineacion_izq
    except Exception as e:
        print(f"[Error] No se pudo actualizar la hoja '{nombre_hoja}': {e}")