# main.py

# ------Importaciones------
import argparse
from pathlib import Path
import pandas as pd
import lector
import utils
import reportes

# ------Herramientas del Menú------

def pedir_fecha(tipo="inicio"):
    # Le pide al usuario el día, mes y año y verifica que sea una fecha que de verdad exista en el calendario
    while True:
        print(f"\n[Fecha de {tipo}] (Escribe 'c' para cancelar y volver al menú)")
        dia = input("Día (DD): ")
        if dia.lower() == 'c': return None
        mes = input("Mes (MM): ")
        if mes.lower() == 'c': return None
        anio = input("Año (YYYY): ")
        if anio.lower() == 'c': return None
        
        try:
            fecha_str = f"{int(dia):02d}/{int(mes):02d}/{int(anio)}"
            pd.to_datetime(fecha_str, format="%d/%m/%Y")
            return fecha_str
        except ValueError:
            print("  [!] Fecha inválida. Por favor, verifica e intenta de nuevo.")

# ------Configuración Inicial------

def main():
    # Prepara el programa para recibir los archivos desde la consola negra (terminal)
    parser = argparse.ArgumentParser(description="Sistema de procesamiento de reportes de asistencia.")
    parser.add_argument("-bd", "--base_datos", type=Path, required=True, help="Ruta de la base de datos principal")
    parser.add_argument("-e", "--eventos", type=Path, help="Archivo de Eventos generales")
    parser.add_argument("-v", "--veladores", type=Path, help="Archivo de Eventos de Veladores")
    args = parser.parse_args()

# ------Carga de Información------

    # Abrimos la libreta principal (A2026.xlsx)
    db_sheets = lector.cargar_base_datos(args.base_datos)
    if db_sheets is None:
        return

    # Revisamos si el usuario nos dio un archivo nuevo de empleados para procesar
    df_eventos = None
    if args.eventos:
        df = lector.cargar_eventos(args.eventos)
        if df is not None and not df.empty:
            df_eventos = df

    # Revisamos si el usuario nos dio un archivo nuevo de veladores para procesar
    df_veladores = None
    if args.veladores:
        df = lector.cargar_eventos_veladores(args.veladores)
        if df is not None and not df.empty:
            df_veladores = df

# ------Actualización de Catálogos e Historiales------

    # 1. Actualizamos la lista de empleados para que nadie falte
    df_emp_actual = db_sheets.get('Empleados') 
    df_empleados_nuevo = utils.actualizar_catalogo_empleados(df_emp_actual, df_eventos, df_veladores)

    if df_empleados_nuevo is not None:
        lector.actualizar_hoja_bd(args.base_datos, 'Empleados', df_empleados_nuevo)

    # 2. Vamos a leer las Variables y la pestaña Ignorar
    df_datos_actual = db_sheets.get('Datos') 
    margen_minutos = 60
    max_lv = 14
    max_sab = 26
    max_dom = 26
    
    try:
        df_vars = pd.read_excel(args.base_datos, sheet_name='Variables')
        def obtener_variable(nombre, valor_por_defecto):
            fila = df_vars[df_vars['Variable'].astype(str).str.contains(nombre, case=False, na=False)]
            if not fila.empty:
                valor_extraido = pd.to_numeric(fila['Valor'].iloc[0], errors='coerce')
                if pd.notna(valor_extraido): return int(valor_extraido)
            return valor_por_defecto

        margen_minutos = obtener_variable('Margen', 60)
        max_lv = obtener_variable('Max Horas L-V', 14)
        max_sab = obtener_variable('Max Horas Sabado', 26)
        max_dom = obtener_variable('Max Horas Domingo', 26)
    except:
        pass
        
    # Leemos la pestaña Ignorar
    try:
        df_ignorar = pd.read_excel(args.base_datos, sheet_name='Ignorar', header=1)
    except:
        df_ignorar = pd.DataFrame()

    # 3. Procesamos y guardamos la asistencia normal
    df_datos_nuevo = utils.actualizar_datos_historicos(df_datos_actual, df_eventos, df_ignorar, margen_minutos)
    if df_datos_nuevo is not None:
        lector.actualizar_hoja_bd(args.base_datos, 'Datos', df_datos_nuevo)

    # 4. Procesamos y guardamos la asistencia de los veladores
    df_datos_veladores_actual = db_sheets.get('DatosVeladores')
    df_datos_veladores_nuevo = utils.actualizar_datos_veladores(
        df_datos_veladores_actual, df_veladores, df_ignorar, margen_minutos, max_lv, max_sab, max_dom
    )
    if df_datos_veladores_nuevo is not None:
        lector.actualizar_hoja_bd(args.base_datos, 'DatosVeladores', df_datos_veladores_nuevo)

# ------Menú Interactivo------

    # Bucle infinito que mantiene vivo el programa hasta que el usuario decida salir
    while True:
        print("\n" + "="*40)
        print("          MENÚ DE REPORTES")
        print("="*40)
        print("1. Generar reporte de empleados")
        print("2. Generar reporte de veladores")
        print("0. Salir")
        
        opcion = input("\nSelecciona una opción: ")
        
        if opcion == '0':
            print("Saliendo del sistema...")
            break
            
        elif opcion in ['1', '2']:
            tipo = "EMPLEADOS" if opcion == '1' else "VELADORES"
            print(f"\n--- REPORTE DE {tipo} ---")
            
            # Pedimos las fechas de corte
            f_inicio = pedir_fecha("inicio")
            if not f_inicio: continue
            
            f_fin = pedir_fecha("fin")
            if not f_fin: continue
            
            # Pedimos cómo quieren que se divida el reporte
            print("\n[Tipo de Periodo]")
            print("1. Semanal\n2. Quincenal\n3. Mensual")
            op_periodo = input("Elige el periodo (o 'c' para cancelar): ")
            if op_periodo.lower() == 'c': continue
            
            periodos = {'1': 'Semanal', '2': 'Quincenal', '3': 'Mensual'}
            if op_periodo not in periodos:
                print("  [!] Opción inválida.")
                continue
                
            periodo = periodos[op_periodo]
            print(f"\n[*] Generando reporte...")
            
            # Mandamos a hacer el Excel según lo que eligieron
            if opcion == '1':
                archivo = reportes.generar_reporte_empleados(args.base_datos, f_inicio, f_fin, periodo)
            else:
                archivo = reportes.generar_reporte_veladores(args.base_datos, f_inicio, f_fin, periodo)
            
            if archivo:
                print(f"  -> [+] Éxito: Reporte guardado como '{archivo}'")
        else:
            print("  [!] Opción no válida.")

if __name__ == "__main__":
    main()