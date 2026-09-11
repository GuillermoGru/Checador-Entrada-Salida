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

def actualizar_datos_historicos(df_existente: pd.DataFrame, df_eventos_nuevos: pd.DataFrame, df_ignorar: pd.DataFrame, margen_minutos: int = 60) -> pd.DataFrame:
    if df_eventos_nuevos is None or df_eventos_nuevos.empty:
        return df_existente if df_existente is not None else None
        
    # Aplicar filtro de Ignorar a la data cruda
    if df_ignorar is not None and not df_ignorar.empty and 'ID de Empleado' in df_ignorar.columns:
        id_ign = df_ignorar['ID de Empleado'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        f_ign = df_ignorar['Fecha'].astype(str).str[:10].str.strip()
        h_ign = df_ignorar['Hora'].astype(str).str[:5].str.strip()
        claves_ign = id_ign + "_" + f_ign + "_" + h_ign
        
        id_ev = df_eventos_nuevos['ID de Empleado'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        f_ev = df_eventos_nuevos['Fecha'].astype(str).str[:10].str.strip()
        h_ev = df_eventos_nuevos['Hora'].astype(str).str[:5].str.strip()
        claves_ev = id_ev + "_" + f_ev + "_" + h_ev
        
        df_eventos_nuevos = df_eventos_nuevos[~claves_ev.isin(claves_ign)].copy()
        
    if df_eventos_nuevos.empty:
        return df_existente if df_existente is not None else None

    # Definir el rango de fechas de los eventos leídos
    fechas_ev = pd.to_datetime(df_eventos_nuevos['Fecha'].astype(str).str.split(' ').str[0], errors='coerce')
    min_fecha = fechas_ev.min()
    max_fecha = fechas_ev.max()

    # Procesar parejas de horas
    df_procesados = procesar_marcas_asistencia(df_eventos_nuevos, margen_minutos)
    
    lista_df = []
    if df_existente is not None and not df_existente.empty:
        # 4. Eliminar registros del historial viejo que caigan en el periodo que estamos reconstruyendo
        df_existente_temp = df_existente.copy()
        df_existente_temp['Fecha_dt'] = pd.to_datetime(df_existente_temp['Fecha'], errors='coerce')
        # Conservamos solo lo que esté antes del inicio o después del final de este archivo de eventos
        mask_fuera_rango = (df_existente_temp['Fecha_dt'] < min_fecha) | (df_existente_temp['Fecha_dt'] > max_fecha) | df_existente_temp['Fecha_dt'].isna()
        df_existente_filtrado = df_existente_temp[mask_fuera_rango].drop(columns=['Fecha_dt'])
        lista_df.append(df_existente_filtrado)
        
    if not df_procesados.empty:
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

# ------Lógica de Asistencia Veladores------

def procesar_marcas_veladores(df_veladores: pd.DataFrame, margen_minutos: int = 60, max_lv: int = 14, max_sab: int = 26, max_dom: int = 26) -> pd.DataFrame:
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
                
                # Leemos qué día de la semana fue para aplicar el máximo correspondiente
                dia_semana = entrada_actual.weekday()
                if dia_semana <= 4: # Lunes a Viernes
                    max_turno_segundos = max_lv * 3600
                elif dia_semana == 5: # Sábado
                    max_turno_segundos = max_sab * 3600
                elif dia_semana == 6: # Domingo
                    max_turno_segundos = max_dom * 3600
                
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

def actualizar_datos_veladores(df_existente: pd.DataFrame, df_veladores_nuevos: pd.DataFrame, df_ignorar: pd.DataFrame, margen_minutos: int = 60, max_lv: int = 14, max_sab: int = 26, max_dom: int = 26) -> pd.DataFrame:
    if df_veladores_nuevos is None or df_veladores_nuevos.empty:
        return df_existente if df_existente is not None else None
        
    if df_ignorar is not None and not df_ignorar.empty and 'ID de Empleado' in df_ignorar.columns:
        id_ign = df_ignorar['ID de Empleado'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        f_ign = df_ignorar['Fecha'].astype(str).str[:10].str.strip()
        h_ign = df_ignorar['Hora'].astype(str).str[:5].str.strip()
        claves_ign = id_ign + "_" + f_ign + "_" + h_ign
        
        id_ev = df_veladores_nuevos['ID de Empleado'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        f_ev = df_veladores_nuevos['Fecha'].astype(str).str[:10].str.strip()
        h_ev = df_veladores_nuevos['Hora'].astype(str).str[:5].str.strip()
        claves_ev = id_ev + "_" + f_ev + "_" + h_ev
        
        df_veladores_nuevos = df_veladores_nuevos[~claves_ev.isin(claves_ign)].copy()
        
    if df_veladores_nuevos.empty:
        return df_existente if df_existente is not None else None

    fechas_ev = pd.to_datetime(df_veladores_nuevos['Fecha'].astype(str).str.split(' ').str[0], errors='coerce')
    min_fecha = fechas_ev.min()
    max_fecha = fechas_ev.max()

    df_procesados = procesar_marcas_veladores(df_veladores_nuevos, margen_minutos, max_lv, max_sab, max_dom)
    
    lista_df = []
    if df_existente is not None and not df_existente.empty:
        df_existente_temp = df_existente.copy()
        df_existente_temp['Fecha_dt'] = pd.to_datetime(df_existente_temp['Fecha'], errors='coerce')
        mask_fuera_rango = (df_existente_temp['Fecha_dt'] < min_fecha) | (df_existente_temp['Fecha_dt'] > max_fecha) | df_existente_temp['Fecha_dt'].isna()
        df_existente_filtrado = df_existente_temp[mask_fuera_rango].drop(columns=['Fecha_dt'])
        lista_df.append(df_existente_filtrado)
        
    if not df_procesados.empty:
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