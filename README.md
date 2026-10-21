# Dashboard BCL - reporte local

[Auto_Reporte_de_Competencia.py](Auto_Reporte_de_Competencia.py) es la version
para Python local del [notebook original](Auto_Reporte_de_Competencia.ipynb).
El notebook se conserva sin cambios.

## Preparacion en VS Code (Windows)

Requiere Python 3.10 o superior y la extension Python de Microsoft para VS Code.
Desde la terminal PowerShell, con esta carpeta abierta:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

En `Python: Select Interpreter`, selecciona
`.\.venv\Scripts\python.exe`. No necesitas Google Colab.

## Ejecutar paso a paso por celdas (recomendado)

1. Instala las extensiones **Python** y **Jupyter** de Microsoft en VS Code.
   `ipykernel`, incluido en las dependencias, permite usar Python local.
2. Abre esta carpeta como proyecto y abre `Auto_Reporte_de_Competencia.py`.
3. Pulsa **Run Cell / Ejecutar celda** encima de `# %% 1. Preparacion e
   importaciones`. Selecciona el kernel de `.venv` si VS Code lo solicita.
4. Ejecuta las siguientes celdas en orden, con **Run Cell** o **Shift + Enter**.
   No uses **Run Python File** ni **Run All** si quieres avanzar manualmente.

Las nueve celdas son: preparacion, catalogos y salida, HR, Admetricks, OOH,
union, base lista, checkpoint y auditoria, y limpieza final con exportacion.
Los selectores de Excel aparecen al ejecutar la celda de cada fuente, no todos
al principio. Puedes inspeccionar `base_prueba`, `base_prueba_admetrics`,
`base_prueba_ooh` y `df_base` en la ventana interactiva despues de sus etapas.

Ejecuta cada celda una sola vez y detente si falla. Para empezar otro reporte,
reinicia el kernel y vuelve a ejecutar desde la primera celda: algunas etapas
renombran o filtran los DataFrames y no deben repetirse sobre datos ya limpiados.
La celda 2 crea una carpeta nueva para esa sesion. En modo interactivo, la ruta
predeterminada depende del directorio de trabajo del kernel si `__file__` no
esta definido; verifica la ruta `Resultados:` que imprime.

Si aun esta abierto un selector de la ejecucion completa anterior, cancelalo
antes de empezar por celdas. No ejecutes simultaneamente ambos modos.

## Ejecutar el reporte completo

```powershell
.\.venv\Scripts\python.exe .\Auto_Reporte_de_Competencia.py
```

Tambien puedes abrir el script y usar **Run Python File in Terminal**.
Se abriran cuatro selectores de archivos, en este orden:

1. Un Excel de catalogos, con las hojas originales `Cat_Marca`, `Cat_Master`,
   `Cat_CIA`, `Cat_Medio`, `Cat_Grupo`, `Cat_Canal`, `Cat_Q`, `Cat_Descuento`,
   `Cat_Digital` y `Cat_OOH`. Puede tener cualquier nombre.
2. Tres Excel de HR.
3. Cuatro Excel de Admetricks.
4. Tres Excel de OOH.

Selecciona todos los archivos de cada grupo en su ventana usando Ctrl o Shift.
Se admiten `.xlsx` y `.xlsm`. Se conservan las columnas requeridas del notebook.
Cancelar una seleccion, repetir un archivo dentro del grupo o seleccionar una
cantidad incorrecta detiene la ejecucion con un error visible.

Para ejecutar sin ventanas, proporciona todas las rutas por argumentos:

```powershell
.\.venv\Scripts\python.exe .\Auto_Reporte_de_Competencia.py `
  --catalogos "C:\datos\Catalogos.xlsx" `
  --hr "C:\datos\HR1.xlsx" "C:\datos\HR2.xlsx" "C:\datos\HR3.xlsx" `
  --admetricks "C:\datos\AD1.xlsx" "C:\datos\AD2.xlsx" "C:\datos\AD3.xlsx" "C:\datos\AD4.xlsx" `
  --ooh "C:\datos\OOH1.xlsx" "C:\datos\OOH2.xlsx" "C:\datos\OOH3.xlsx" `
  --salida "C:\reportes\BCL"
```

## Resultados y comportamiento

Cada ejecucion crea una subcarpeta nueva con fecha y hora dentro de `salidas`
(junto al script), o dentro de la carpeta indicada con `--salida`.
No cambia el directorio de trabajo ni sobrescribe resultados anteriores.

Se guardan `BASE_HR.xlsx`, `BASE_Admetricks.xlsx`, `BASE_OOH.xlsx`,
`Bases_Unidas.xlsx`, `BASE_LISTA.xlsx`, `checkpoint_base.csv` y
`Base_2026_Q1.csv`. Ya no hay cargas ni descargas desde el navegador.
Los errores de lectura, hojas o columnas faltantes se muestran en la terminal;
una ejecucion fallida puede dejar resultados parciales en su subcarpeta.

Se conservan los cruces, descuentos y filtros del notebook, incluidos los filtros
de 2025, abril de 2026 y las versiones de Villa del Sur. El periodo y el nombre
`Base_2026_Q1.csv` **no se actualizan automaticamente**.
La instalacion de paquetes se realiza antes de ejecutar, no dentro del reporte.
La celda opcional de borrado masivo **no se incorpora** al script: los archivos
de entrada y los resultados permanecen en disco.
