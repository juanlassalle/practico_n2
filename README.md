# practico_n2
Práctico numero 2 de la unidad 3 - Introducción al análisis de datos

# Guía de inicio: Python, Visual Studio Code y análisis de datos con Pandas

El presente documento describe los pasos necesarios para configurar un entorno de
trabajo con Python en Visual Studio Code, crear un entorno virtual e importar un
archivo de Excel para su análisis mediante la biblioteca Pandas.


---

## 1. Requisitos para comenzar a trabajar con Python en Visual Studio Code

Para desarrollar y ejecutar programas en Python dentro de Visual Studio Code se
requieren los siguientes elementos:

| Elemento | Descripción |
|---|---|
| **Python** | Intérprete instalado en el sistema, preferentemente una versión 3.10 o superior. |
| **Visual Studio Code** | Editor de código disponible para Windows, Linux y macOS. |
| **Extensión Python (Microsoft)** | Aporta resaltado de sintaxis, autocompletado, depuración y gestión de intérpretes. |
| **Terminal integrada** | Permite ejecutar comandos, instalar bibliotecas y correr scripts sin salir del editor. |
| **pip** | Gestor de paquetes incluido con Python, utilizado para instalar bibliotecas externas. |

### 1.1. Verificación de la instalación

Una vez instalado Python, se comprueba su disponibilidad desde la terminal
integrada de Visual Studio Code (`Ctrl` + `` ` ``):

```bash
python --version
pip --version
```

En sistemas Linux o macOS, el comando correspondiente suele ser `python3` y
`pip3`. En Windows, también se puede utilizar el lanzador `py`.

Si el sistema no reconoce el comando, se debe verificar que Python se haya
instalado con la opción **Add Python to PATH** activada.

### 1.2. Instalación de la extensión

1. Se abre Visual Studio Code.
2. Se accede al panel **Extensions** (`Ctrl` + `Shift` + `X`).
3. Se busca **Python**, publicada por Microsoft, y se instala.
4. Se recarga la ventana cuando el editor lo solicita.

### 1.3. Estructura recomendada del proyecto

Se recomienda organizar los archivos en una carpeta específica, manteniendo el
script y el archivo de datos en la misma ubicación:

```text
Practico_N2/
├── .venv/                  # Entorno virtual (no se comparte ni se versiona)
├── Registros_Socios.xlsx   # Archivo de datos de origen
├── practico.py             # Script de Python
└── README.md               # Documentación del proyecto
```

---

## 2. Creación y activación de un entorno virtual en Visual Studio Code

Un entorno virtual permite aislar las bibliotecas de cada proyecto. De esta
manera, se evita que las instalaciones de un proyecto afecten a otros o a la
instalación global de Python.

### 2.1. Creación del entorno

Se abre la terminal integrada y se ejecuta el siguiente comando dentro de la
carpeta del proyecto:

```bash
python -m venv .venv
```

En Linux o macOS se utiliza:

```bash
python3 -m venv .venv
```

Como resultado, se genera una carpeta denominada `.venv` que contiene una copia
independiente del intérprete y su propio directorio de paquetes.

### 2.2. Activación del entorno

El comando de activación depende del sistema operativo y de la terminal empleada.

**Windows — PowerShell:**

```powershell
.venv\Scripts\Activate.ps1
```

**Windows — Símbolo del sistema (CMD):**

```bat
.venv\Scripts\activate.bat
```

**Linux / macOS:**

```bash
source .venv/bin/activate
```

Cuando el entorno se encuentra activo, el indicador `(.venv)` aparece al inicio
de la línea de comandos.

> **Nota (PowerShell):** si el sistema bloquea la ejecución de scripts, se puede
> utilizar la terminal CMD dentro de Visual Studio Code, o bien ejecutar el
> entorno con el lanzador `py -m venv .venv` y activarlo mediante `activate.bat`.

### 2.3. Selección del intérprete en Visual Studio Code

Para que el editor utilice el entorno creado, se siguen estos pasos:

1. Se abre la paleta de comandos con `Ctrl` + `Shift` + `P`.
2. Se escribe **Python: Select Interpreter**.
3. Se selecciona la entrada que corresponde a `.venv\Scripts\python.exe`.

A partir de ese momento, las ejecuciones y el autocompletado del editor utilizan
las bibliotecas instaladas dentro del entorno virtual.

### 2.4. Desactivación del entorno

Para salir del entorno virtual se ejecuta:

```bash
deactivate
```

---

## 3. Trabajo con Pandas y lectura de archivos Excel

### 3.1. Instalación de las bibliotecas necesarias

Con el entorno virtual activado, se instalan Pandas y el motor de lectura de
archivos `.xlsx`:

```bash
python -m pip install --upgrade pip
pip install pandas openpyxl
```

- **pandas:** proporciona la estructura `DataFrame` y las herramientas de
  limpieza y transformación de datos.
- **openpyxl:** biblioteca requerida por Pandas para leer y escribir archivos
  Excel con extensión `.xlsx`. Si no se encuentra instalada, se produce el error
  `Missing optional dependency 'openpyxl'`.

La verificación de las versiones se realiza mediante:

```bash
python -c "import pandas; print(pandas.__version__)"
```

### 3.2. Importación de un archivo Excel como DataFrame

Un **DataFrame** es una tabla bidimensional compuesta por filas y columnas.
Para crearlo a partir de un archivo Excel se utiliza la función
`pd.read_excel()`.

El siguiente ejemplo emplea una ruta calculada a partir de la ubicación del
propio script, lo que permite que el proyecto funcione aunque se traslade de
carpeta:

```python
import os
import pandas as pd

# Ubicación del archivo, en la misma carpeta que el script
# Lectura del Excel: se omite la primera fila, que contiene un título
datos = pd.read_excel(ruta_archivo, skiprows=1)

#En caso contrario esto
ruta_carpeta = os.path.dirname(os.path.abspath(__file__))
ruta_archivo = os.path.join(ruta_carpeta, 'Registro_Socios.xlsx')

# Lectura del Excel: se omite la primera fila, que contiene un título
datos = pd.read_excel(ruta_archivo, skiprows=1)

# Visualización de las primeras filas
print(datos.head())
```

### 3.3. Parámetros habituales de lectura

| Parámetro | Función |
|---|---|
| `skiprows=1` | Ignora la primera fila del archivo. Se utiliza cuando el Excel contiene un título antes de los encabezados. |
| `sheet_name='Hoja1'` | Indica la hoja que se desea leer. |
| `usecols` | Limita la lectura a columnas específicas. |
| `dtype` | Define el tipo de dato con el que se carga cada columna. |
| `na_values` | Permite indicar valores adicionales que deben interpretarse como faltantes. |

### 3.4. Configuración de la visualización

Para evitar que la consola(evitar usar esto) oculte columnas mediante el indicador `...`, se
pueden ajustar las siguientes opciones:

```python
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', 50)
```

### 3.5. Inspección inicial del dataset

Antes de aplicar cualquier transformación, se recomienda releva la estructura y
la calidad de los datos:

```python
print(datos.shape)              # Cantidad de filas y columnas
print(datos.columns)            # Nombres de las columnas
print(datos.info())             # Tipos de dato y valores presentes
print(datos.isnull().sum())     # Valores faltantes por columna
print(datos.duplicated().sum()) # Filas completamente duplicadas
```

Estos comandos permiten detectar valores ausentes, registros repetidos,
inconsistencias de formato y errores de tipeo que requieren tratamiento.

### 3.6. Exportación de resultados

Una vez procesados los datos, el resultado puede guardarse en nuevos formatos:

```python
# Exportación a Excel(no lo recomiendo)
datos.to_excel('Registros_Socios_LIMPIO.xlsx', index=False)

# Exportación a HTML para su visualización en un navegador
datos.to_html('Reporte_Socios.html', index=False)
```

---

## 4. Problemas frecuentes y soluciones

| Problema | Causa habitual | Solución |
|---|---|---|
| `FileNotFoundError` | El archivo Excel no se encuentra en la carpeta de trabajo. | Verificar el nombre exacto, la extensión real del archivo y la ruta utilizada. |
| `Missing optional dependency 'openpyxl'` | Falta el motor de lectura de Excel. | Ejecutar `pip install openpyxl` dentro del entorno virtual. |
| El módulo `pandas` no se encuentra | El paquete fue instalado fuera del entorno activo. | Activar `.venv` y repetir la instalación. |
| Encabezados incorrectos | El Excel contiene un título en la primera fila. | Utilizar `skiprows=1` en `pd.read_excel()`. |
| Columnas ocultas con `...` | Límites de visualización de Pandas. | Ajustar `display.max_columns` y `display.width`. |
| Teléfonos en notación científica | La columna se cargó como número decimal. | Convertirla a entero y luego a cadena de texto, o leerla directamente como `dtype=str`. |

---

## 5. Flujo de trabajo recomendado

1. Se crea la carpeta del proyecto y se abre en Visual Studio Code.
2. Se genera y se activa el entorno virtual `.venv`.
3. Se selecciona el intérprete del entorno dentro del editor.
4. Se instalan `pandas` y `openpyxl`.
5. Se copia el archivo Excel junto al script.
6. Se carga la información mediante `pd.read_excel()`.
7. Se realiza la inspección inicial y se relevan los problemas de calidad.
8. Se aplican las limpiezas y transformaciones correspondientes.
9. Se repite la inspección para validar los resultados.
10. Se exporta el dataset final y se documenta el procedimiento.

---

## 6. Comandos de resumen

```bash
# Crear el entorno virtual
python -m venv .venv

# Activarlo (Windows - CMD)
.venv\Scripts\activate.bat

# Instalar las dependencias
pip install pandas openpyxl

# Ejecutar el script
python practico.py

# Generar el archivo de dependencias (opcional)
pip freeze > requirements.txt

# Restaurar el entorno en otra máquina (opcional)
pip install -r requirements.txt
```

La correcta aplicación de estos pasos garantiza un entorno reproducible y
permite trabajar con archivos Excel de forma consistente y documentada.
