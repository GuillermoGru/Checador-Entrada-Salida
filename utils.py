# utils.py

# ------Importaciones------
import pandas as pd
from datetime import datetime

# ------Catálogo de Empleados------

def actualizar_catalogo_empleados(df_existente: pd.DataFrame, df_eventos: pd.DataFrame, df_veladores: pd.DataFrame) -> pd.DataFrame:
    # Agarra las listas viejas y nuevas, las mezcla, y elimina a los repetidos 
    # para que en la hoja 'Empleados' siempre haya uno de cada uno
    lista_df = []
    
    if df_existente is not None and not df_existente.empty:
        lista_df.append(df_existente)
        
    if df_eventos is not None and not df_eventos.empty:
        cols_emp = ['ID de Empleado', 'Nombre', 'Apellido', 'Género', 'ID de Departamento', 'Departamento', 'ID del Cargo', 'Cargo']
        cols_presentes = [c for c in cols_emp if c in df_eventos.columns]
        lista_df.append(df_eventos[cols_presentes].copy())
        
    if df_veladores is not None and not df_veladores.empty:
        cols_vel = ['ID de Empleado', 'Nombre', 'Apellido', 'Departamento']
        cols_presentes = [c for c in cols_vel if c in df_veladores.columns]
        df_vel = df_veladores[cols_presentes].copy()
        df_vel['Cargo'] = 'Velador' 
        lista_df.append(df_vel)
        
    if lista_df:
        df_final = pd.concat(lista_df, ignore_index=True)
        # Nos quedamos con el primero que encontramos y borramos el resto
        df_final = df_final.drop_duplicates(subset=['ID de Empleado'], keep='first')
        df_final = df_final.sort_values(by='ID de Empleado')
        return df_final
    return None

# ------Lógica de Asistencia General------

def procesar_marcas_asistencia(df_eventos: pd.DataFrame, margen_minutos: int = 60) -> pd.DataFrame:
    # La máquina checadora solo da una lista larga de horas. Esta función las agrupa en parejas de "Entrada y Salida"
    if df_eventos is None or df_eventos.empty:
        return pd.DataFrame()
        
    cols_necesarias = ['ID de Empleado', 'Fecha', 'Hora']
    if not all(col in df_eventos.columns for col in cols_necesarias):
        return pd.DataFrame()
        
    df_temp = df_eventos[cols_necesarias].copy()
    df_temp['Fecha_str'] = df_temp['Fecha'].astype(str).str.split(' ').str[0]
    df_temp['Hora_str'] = df_temp['Hora'].astype(str)
    df_temp['FechaHora'] = pd.to_datetime(df_temp['Fecha_str'] + ' ' + df_temp['Hora_str'], errors='coerce')
    df_temp = df_temp.dropna(subset=['FechaHora'])
    df_temp = df_temp.sort_values(by=['ID de Empleado', 'FechaHora'])
    
    margen_segundos = margen_minutos * 60
    
    resultados = []
    # Revisamos día por día para cada empleado normal
    for (empleado_id, fecha_str), grupo in df_temp.groupby(['ID de Empleado', 'Fecha_str']):
        marcas = grupo['FechaHora'].tolist()
        entrada_actual = None
        
        for hora in marcas:
            if entrada_actual is None:
                entrada_actual = hora
            else:
                diff = hora - entrada_actual
                
                # Si pasó muy poco tiempo entre checadas, significa que fue un accidente (-1)
                if diff.total_seconds() < margen_segundos:
                    resultados.append({
                        'ID de Empleado': empleado_id,
                        'Fecha': fecha_str,
                        'Hora entrada': hora.strftime('%H:%M'),
                        'Hora salida': None,
                        'Diferencia horas': -1
                    })
                else:
                    # Hacen buena pareja, calculamos las horas que trabajó
                    horas = int(diff.total_seconds() // 3600)
                    minutos = int((diff.total_seconds() % 3600) // 60)
                    resultados.append({
                        'ID de Empleado': empleado_id,
                        'Fecha': fecha_str,
                        'Hora entrada': entrada_actual.strftime('%H:%M'),
                        'Hora salida': hora.strftime('%H:%M'),
                        'Diferencia horas': f"{horas:02d}:{minutos:02d}"
                    })
                    entrada_actual = None 
                    
        # Si se acabó el día y no marcó salida, es una huérfana (0)
        if entrada_actual is not None:
            resultados.append({
                'ID de Empleado': empleado_id,
                'Fecha': fecha_str,
                'Hora entrada': entrada_actual.strftime('%H:%M'),
                'Hora salida': None,
                'Diferencia horas': 0
            })
            
    return pd.DataFrame(resultados)

# ------Limpieza y Guardado General------

def actualizar_datos_historicos(df_existente: pd.DataFrame, df_eventos_nuevos: pd.DataFrame, margen_minutos: int = 60) -> pd.DataFrame:
    # Une los pares nuevos con el historial de A2026.xlsx y borra basuritas o duplicados
    df_procesados = procesar_marcas_asistencia(df_eventos_nuevos, margen_minutos)
    
    if df_procesados.empty:
        return df_existente if df_existente is not None else None
        
    lista_df = []
    if df_existente is not None and not df_existente.empty:
        lista_df.append(df_existente)
        
    lista_df.append(df_procesados)
    
    if lista_df:
        df_final = pd.concat(lista_df, ignore_index=True)
        
        # Sistema automático: si un registro antes era 0 (huérfano) pero hoy ya le encontramos pareja, borramos el 0 viejo
        df_final['clave_temp'] = df_final['ID de Empleado'].astype(str) + df_final['Fecha'].astype(str) + df_final['Hora entrada'].astype(str)
        claves_con_pareja = df_final.loc[df_final['Diferencia horas'].astype(str).str.contains(':'), 'clave_temp'].unique()
        es_huerfano = df_final['Diferencia horas'].astype(str) == '0'
        tiene_pareja = df_final['clave_temp'].isin(claves_con_pareja)
        
        df_final = df_final[~(es_huerfano & tiene_pareja)]
        df_final = df_final.drop(columns=['clave_temp'])
        
        # Limpieza estándar anti-clonaciones
        df_final = df_final.drop_duplicates(subset=['ID de Empleado', 'Fecha', 'Hora entrada', 'Diferencia horas'], keep='last')
        df_final = df_final.sort_values(by=['ID de Empleado', 'Fecha', 'Hora entrada'])
        return df_final
        
    return None

# ------Lógica de Asistencia Veladores------

def procesar_marcas_veladores(df_veladores: pd.DataFrame, margen_minutos: int = 60) -> pd.DataFrame:
    # Los veladores trabajan de noche, así que no podemos agrupar sus horas en el mismo día.
    # Aquí los ordenamos cronológicamente en el tiempo de forma libre
    if df_veladores is None or df_veladores.empty:
        return pd.DataFrame()
        
    cols_necesarias = ['ID de Empleado', 'Fecha', 'Hora']
    if not all(col in df_veladores.columns for col in cols_necesarias):
        return pd.DataFrame()
        
    df_temp = df_veladores[cols_necesarias].copy()
    df_temp['Fecha_str'] = df_temp['Fecha'].astype(str).str.split(' ').str[0]
    df_temp['Hora_str'] = df_temp['Hora'].astype(str)
    df_temp['FechaHora'] = pd.to_datetime(df_temp['Fecha_str'] + ' ' + df_temp['Hora_str'], errors='coerce')
    df_temp = df_temp.dropna(subset=['FechaHora'])
    
    df_temp = df_temp.sort_values(by=['ID de Empleado', 'FechaHora'])
    
    margen_segundos = margen_minutos * 60
    resultados = []
    
    for empleado_id, grupo in df_temp.groupby('ID de Empleado'):
        marcas = grupo['FechaHora'].tolist()
        entrada_actual = None
        
        for hora in marcas:
            if entrada_actual is None:
                entrada_actual = hora
            else:
                diff = hora - entrada_actual
                
                # Leemos qué día de la semana fue. Los fines de semana los veladores hacen turnos de hasta 26 horas
                if entrada_actual.weekday() in [5, 6]:
                    max_turno_segundos = 26 * 3600
                else:
                    max_turno_segundos = 14 * 3600
                
                if diff.total_seconds() < margen_segundos:
                    # Doble checada por accidente
                    resultados.append({
                        'ID de Empleado': empleado_id,
                        'Fecha': hora.strftime('%Y-%m-%d'),
                        'Hora entrada': hora.strftime('%H:%M'),
                        'Hora salida': None,
                        'Diferencia horas': -1
                    })
                elif diff.total_seconds() > max_turno_segundos:
                    # El velador nunca marcó salida, la marcamos huérfana
                    resultados.append({
                        'ID de Empleado': empleado_id,
                        'Fecha': entrada_actual.strftime('%Y-%m-%d'),
                        'Hora entrada': entrada_actual.strftime('%H:%M'),
                        'Hora salida': None,
                        'Diferencia horas': 0
                    })
                    entrada_actual = hora # La checada actual se vuelve el inicio de su nuevo día
                else:
                    # Emparejado con éxito (incluso si cruzó la medianoche)
                    horas = int(diff.total_seconds() // 3600)
                    minutos = int((diff.total_seconds() % 3600) // 60)
                    resultados.append({
                        'ID de Empleado': empleado_id,
                        'Fecha': entrada_actual.strftime('%Y-%m-%d'),
                        'Hora entrada': entrada_actual.strftime('%H:%M'),
                        'Hora salida': hora.strftime('%H:%M'),         
                        'Diferencia horas': f"{horas:02d}:{minutos:02d}"
                    })
                    entrada_actual = None 
                    
        if entrada_actual is not None:
            resultados.append({
                'ID de Empleado': empleado_id,
                'Fecha': entrada_actual.strftime('%Y-%m-%d'),
                'Hora entrada': entrada_actual.strftime('%H:%M'),
                'Hora salida': None,
                'Diferencia horas': 0
            })
            
    return pd.DataFrame(resultados)

# ------Limpieza y Guardado Veladores------

def actualizar_datos_veladores(df_existente: pd.DataFrame, df_veladores_nuevos: pd.DataFrame, margen_minutos: int = 60) -> pd.DataFrame:
    # Igual que con los empleados, mezcla lo nuevo con el historial y limpia huérfanos falsos
    df_procesados = procesar_marcas_veladores(df_veladores_nuevos, margen_minutos)
    
    if df_procesados.empty:
        return df_existente if df_existente is not None else None
        
    lista_df = []
    if df_existente is not None and not df_existente.empty:
        lista_df.append(df_existente)
        
    lista_df.append(df_procesados)
    
    if lista_df:
        df_final = pd.concat(lista_df, ignore_index=True)
        
        df_final['clave_temp'] = df_final['ID de Empleado'].astype(str) + df_final['Fecha'].astype(str) + df_final['Hora entrada'].astype(str)
        claves_con_pareja = df_final.loc[df_final['Diferencia horas'].astype(str).str.contains(':'), 'clave_temp'].unique()
        es_huerfano = df_final['Diferencia horas'].astype(str) == '0'
        tiene_pareja = df_final['clave_temp'].isin(claves_con_pareja)
        
        df_final = df_final[~(es_huerfano & tiene_pareja)]
        df_final = df_final.drop(columns=['clave_temp'])
        
        df_final = df_final.drop_duplicates(subset=['ID de Empleado', 'Fecha', 'Hora entrada', 'Diferencia horas'], keep='last')
        df_final = df_final.sort_values(by=['ID de Empleado', 'Fecha', 'Hora entrada'])
        return df_final
        
    return None