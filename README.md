# Guía de Ejecución: Sistema de Procesamiento de Reportes de Asistencia

Este documento explica paso a paso cómo preparar, instalar y ejecutar el sistema de procesamiento de asistencia desde cero en un entorno Windows.

## 1. Requisitos Previos

Antes de ejecutar el programa, asegúrate de tener instalado el entorno adecuado en tu computadora con Windows.

### Instalar Python
1. Descarga el instalador de Python desde la página oficial (https://www.python.org/downloads/windows/).
2. Ejecuta el instalador. **MUY IMPORTANTE**: En la primera pantalla de instalación, marca la casilla que dice **"Add Python to PATH"** (o "Agregar Python al PATH").
3. Haz clic en "Install Now" y espera a que termine.

### Instalar Dependencias
El sistema utiliza herramientas especializadas para leer datos y dibujar en Excel. Necesitas instalar estas librerías.
1. Presiona la tecla `Windows + R`, escribe `cmd` y presiona Enter para abrir la consola de comandos.
2. Escribe el siguiente comando y presiona Enter:
   ```cmd
   pip install pandas openpyxl
   ```
3. Espera a que se descarguen e instalen las librerías.

## 2. Preparación de los Archivos

El sistema está compuesto por cuatro archivos de código principales que deben estar en la misma carpeta:
* `main.py`: Es el archivo principal que controla el menú interactivo y la configuración inicial
* `lector.py`: Se encarga de la lectura de archivos nuevos y la escritura/diseño visual en la base de datos principal
* `utils.py`: Contiene la lógica para procesar marcas de asistencia, agrupar pares de entrada/salida y actualizar catálogos
* `reportes.py`: Es el motor encargado de las matemáticas de tiempo, gestión de permisos y generación de los reportes finales en Excel

### Archivos Excel Necesarios
Debes tener listos los siguientes archivos en formato `.xlsx`:
1. **Base de Datos Principal (ej. `A2026.xlsx`)**: Este archivo es **obligatorio**.
2. **Archivo de Eventos (Opcional)**: El Excel generado por la máquina checadora normal.
3. **Archivo de Veladores (Opcional)**: El Excel generado para el turno de veladores.

## 3. Ejecución del Programa Paso a Paso

1. Abre la consola de comandos (`cmd`) o PowerShell.
2. Navega hasta la carpeta donde guardaste los archivos de código usando el comando `cd`. Por ejemplo:
   ```cmd
   cd C:\Users\TuUsuario\Documentos\SistemaAsistencia
   ```
3. Ejecuta el programa llamando a `main.py`. Debes pasarle las rutas de los archivos usando argumentos de la terminal. 

**Comando Básico (Solo Base de Datos):**
```cmd
python main.py -bd A2026.xlsx
```
*(Nota: El argumento `-bd` es obligatorio).*

**Comando Completo (Con eventos nuevos y veladores):**
```cmd
python main.py -bd A2026.xlsx -e ReporteEventos.xlsx -v ReporteVeladores.xlsx
```
* `-e` : Procesa un archivo nuevo de checadas generales.
* `-v` : Procesa un archivo nuevo exclusivo de veladores.

## 4. Uso del Menú Interactivo

Una vez ejecutado el comando, el programa actualizará automáticamente los historiales y te mostrará el siguiente menú interactivo:

```text
========================================
          MENÚ DE REPORTES
========================================
1. Generar reporte de empleados
2. Generar reporte de veladores
0. Salir
```

### Generar un Reporte
Si eliges la opción `1` o `2`, el programa te pedirá lo siguiente:
1. **Fecha de inicio y fin**: Se te pedirá el Día, Mes y Año. Deben ser fechas válidas; de lo contrario, el sistema te pedirá que intentes de nuevo.
2. **Tipo de Periodo**: Podrás elegir cómo dividir el reporte:
   * `1` para Semanal.
   * `2` para Quincenal.
   * `3` para Mensual.

