# %% 1. Preparacion e importaciones
"""Version local del reporte de competencia. Ejecutar con Python 3.10 o superior."""

from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path
from typing import Sequence

import numpy as np
import pandas as pd


def validar_archivos(
    rutas: Sequence[Path], cantidad: int, grupo: str
) -> list[Path]:
    if len(rutas) != cantidad:
        raise ValueError(
            f"{grupo}: se necesitan {cantidad} archivos; se recibieron {len(rutas)}."
        )
    archivos = [ruta.expanduser().resolve() for ruta in rutas]
    if len(set(archivos)) != cantidad:
        raise ValueError(f"{grupo}: no selecciones el mismo archivo mas de una vez.")
    for archivo in archivos:
        if not archivo.is_file():
            raise FileNotFoundError(f"{grupo}: no existe el archivo {archivo}")
        if archivo.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ValueError(f"{grupo}: se requiere Excel .xlsx o .xlsm: {archivo}")
    return archivos


def seleccionar_archivos(grupo: str, cantidad: int) -> list[Path]:
    import tkinter as tk
    from tkinter import filedialog

    ventana = tk.Tk()
    ventana.withdraw()
    try:
        seleccion = filedialog.askopenfilenames(
            parent=ventana,
            title=f"{grupo}: selecciona {cantidad} archivo(s) Excel",
            filetypes=[("Archivos Excel", "*.xlsx *.xlsm")],
        )
    finally:
        ventana.destroy()
    if not seleccion:
        raise RuntimeError(f"Seleccion cancelada: {grupo}. No se proceso el reporte.")
    return validar_archivos([Path(ruta) for ruta in seleccion], cantidad, grupo)


def leer_argumentos(argv: Sequence[str] | None = None) -> argparse.Namespace:
    carpeta_proyecto = Path(__file__).resolve().parent if '__file__' in globals() else Path.cwd()
    parser = argparse.ArgumentParser(
        description=(
            "Genera el reporte local de competencia. Si omites un grupo de "
            "archivos, se abre un selector para ese grupo."
        )
    )
    parser.add_argument("--catalogos", type=Path, help="Archivo de catalogos Excel")
    parser.add_argument("--hr", type=Path, nargs=3, help="Los tres archivos HR")
    parser.add_argument(
        "--admetricks", type=Path, nargs=4, help="Los cuatro archivos Admetricks"
    )
    parser.add_argument("--ooh", type=Path, nargs=3, help="Los tres archivos OOH")
    parser.add_argument(
        "--salida",
        type=Path,
        default=carpeta_proyecto / "salidas",
        help="Carpeta donde crear una subcarpeta nueva con los resultados",
    )
    return parser.parse_args(argv)


# %% 2. Catalogos y carpeta de resultados
if __name__ == "__main__":
    # La ventana interactiva de VS Code no usa los argumentos de su kernel.
    args = leer_argumentos([] if "get_ipython" in globals() else None)
    archivo_catalogos = (
        validar_archivos([args.catalogos], 1, "Catalogos")[0]
        if args.catalogos is not None
        else seleccionar_archivos("Catalogos", 1)[0]
    )
    carpeta_salida = args.salida.expanduser().resolve() / datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    carpeta_salida.mkdir(parents=True, exist_ok=False)
    print(f"Resultados: {carpeta_salida}")
    # Leer las hojas del DataFrame catalogos
    cat_marca = pd.read_excel(archivo_catalogos, sheet_name='Cat_Marca')
    cat_master = pd.read_excel(archivo_catalogos, sheet_name='Cat_Master')
    cat_cia = pd.read_excel(archivo_catalogos, sheet_name='Cat_CIA')
    cat_medio = pd.read_excel(archivo_catalogos, sheet_name='Cat_Medio')
    cat_grupo = pd.read_excel(archivo_catalogos, sheet_name='Cat_Grupo')
    cat_canal = pd.read_excel(archivo_catalogos, sheet_name='Cat_Canal')
    cat_q = pd.read_excel(archivo_catalogos, sheet_name='Cat_Q')
    cat_descuento = pd.read_excel(archivo_catalogos, sheet_name='Cat_Descuento')
    cat_digital = pd.read_excel(archivo_catalogos, sheet_name='Cat_Digital')
    cat_ooh = pd.read_excel(archivo_catalogos, sheet_name='Cat_OOH')


# %% 3. HR
if __name__ == "__main__":
    archivos_hr = (
        validar_archivos(args.hr, 3, "HR")
        if args.hr is not None
        else seleccionar_archivos("HR", 3)
    )
    # Las columnas comunes y deseadas que mencionaste, ahora incluyendo "Week"
    columnas_deseadas = [
     'Month', 'Week', 'Anunciante', 'Marca', 'Producto', 'Submarca', 'Version',
     'Channel', 'Tipo de Spot', 'Tipo de Medio', 'Cost', 'Sector', 'Subsector',
     'Rtg% A13-54', 'Rtg# A13-54'
    ]

    # Leer los archivos locales de HR.
    base1_df = pd.read_excel(archivos_hr[0])
    base2_df = pd.read_excel(archivos_hr[1])
    base3_df = pd.read_excel(archivos_hr[2])

    # Concatenar los tres DataFrames, seleccionando solo las columnas deseadas.
    base_prueba = pd.concat([
     base1_df[columnas_deseadas],
     base2_df[columnas_deseadas],
     base3_df[columnas_deseadas]
    ], ignore_index=True)

    # Convertir la columna "Cuartil" de la hoja "Cat_Q" a tipo str
    cat_q['Mes'] = cat_q['Mes'].astype(str)

    # Procesar las columnas "Month" y "Week"
    base_prueba['Month'] = base_prueba['Month'].astype(str)
    base_prueba['Week'] = base_prueba['Week'].astype(str)

    # Crear la columna "Año" tomando los primeros 4 dígitos de "Month"
    base_prueba['Año'] = base_prueba['Month'].str[:4]

    # Crear la columna "Mes" tomando solo los últimos dos dígitos de "Month"
    base_prueba['Mes'] = base_prueba['Month'].str[-2:]

    # Crear la columna "Semana" tomando solo los últimos dos dígitos de "Week"
    base_prueba['Semana'] = base_prueba['Week'].str[-2:]

    # Crear la columna "Concat" concatenando "Marca", "Producto" y "Submarca" con un espacio después de cada dato
    base_prueba['Concat'] = base_prueba['Marca'] + ' ' + base_prueba['Producto'] + ' ' + base_prueba['Submarca']

    # ---- PROCESO 1: Comparación con "Cat_Marca" ----
    base_prueba['Marca*'] = 'Sin coincidencia'
    base_prueba['Subproducto'] = 'Sin coincidencia'
    base_prueba['Caracteristica'] = 'Sin coincidencia'

    for i, row in base_prueba.iterrows():
     concat_value = row['Concat']

     # Buscar coincidencias exactas en la hoja "Cat_Marca"
     match_marca = cat_marca[cat_marca['Concat'] == concat_value]

     if not match_marca.empty:
         # Si hay coincidencia, copiar los valores de las columnas correspondientes
         base_prueba.at[i, 'Marca*'] = match_marca['Marca2'].iloc[0]
         base_prueba.at[i, 'Subproducto'] = match_marca['Subproducto'].iloc[0]
         base_prueba.at[i, 'Caracteristica'] = match_marca['Caracteristica'].iloc[0]

    # ---- PROCESO 2: Comparación con "Cat_Master" ----
    base_prueba['Megamarca'] = 'Sin coincidencia'
    base_prueba['Categoria'] = 'Sin coincidencia'
    base_prueba['Tipo'] = 'Sin coincidencia'

    for i, row in base_prueba.iterrows():
     marca_value = row['Marca*']

     # Buscar coincidencias exactas en la hoja "Cat_Master"
     match_master = cat_master[cat_master['Marca2'] == marca_value]

     if not match_master.empty:
         # Si hay coincidencia, copiar los valores de las columnas correspondientes
         base_prueba.at[i, 'Megamarca'] = match_master['Megamarca'].iloc[0]
         base_prueba.at[i, 'Categoria'] = match_master['Categoria'].iloc[0]
         base_prueba.at[i, 'Tipo'] = match_master['Tipo'].iloc[0]

    # ---- PROCESO 3: Comparación con "Cat_CIA" ----
    base_prueba['CIA'] = 'Sin coincidencia'
    base_prueba['Open_CIA'] = 'Sin coincidencia'

    for i, row in base_prueba.iterrows():
     megamarca_value = row['Megamarca']

     # Buscar coincidencias exactas en la hoja "Cat_CIA"
     match_cia = cat_cia[cat_cia['Megamarca'] == megamarca_value]

     if not match_cia.empty:
         # Si hay coincidencia, copiar los valores de las columnas correspondientes
         base_prueba.at[i, 'CIA'] = match_cia['CIA'].iloc[0]
         base_prueba.at[i, 'Open_CIA'] = match_cia['Open_CIA'].iloc[0]

    # ---- PROCESO 4: Comparaciones simultáneas con "Cat_Medio", "Cat_Grupo", "Cat_Canal", "Cat_Q" ----
    base_prueba['Medio'] = 'Sin coincidencia'
    base_prueba['Grupo'] = 'Sin coincidencia'
    base_prueba['Cadena'] = 'Sin coincidencia'
    base_prueba['Cuartil'] = 'Sin coincidencia'

    # Convertir la columna "Mes" en ambas tablas a tipo int para trabajar en formato general
    base_prueba['Mes'] = pd.to_numeric(base_prueba['Mes'], errors='coerce')  # Convertir a números enteros
    cat_q['Mes'] = pd.to_numeric(cat_q['Mes'], errors='coerce')  # Convertir también en cat_q

    # Iterar sobre las filas del DataFrame base_prueba
    for i, row in base_prueba.iterrows():
     # Valores de las columnas a comparar
     channel_value = row['Channel']
     categoria_value = row['Categoria']
     cadena_value = row['Channel']
     mes_value = row['Mes']  # La columna 'Mes' ahora está en formato numérico

     # Comparar con "Cat_Medio"
     match_medio = cat_medio[cat_medio['Cadena_origen'] == channel_value]
     if not match_medio.empty:
         base_prueba.at[i, 'Medio'] = match_medio['Medio'].iloc[0]

     # Comparar con "Cat_Grupo"
     match_grupo = cat_grupo[cat_grupo['Categoria'] == categoria_value]
     if not match_grupo.empty:
         base_prueba.at[i, 'Grupo'] = match_grupo['Grupo'].iloc[0]

     # Comparar con "Cat_Canal"
     match_canal = cat_canal[cat_canal['Canal'] == cadena_value]
     if not match_canal.empty:
         base_prueba.at[i, 'Cadena'] = match_canal['Cadena'].iloc[0]

     # Comparar con la columna "Mes" de "Cat_Q"
     match_q = cat_q[cat_q['Mes'] == mes_value]
     if not match_q.empty:
         base_prueba.at[i, 'Cuartil'] = match_q['Número de Q'].iloc[0]


    # ---- PROCESO 5: Comparación con "Cat_Descuento" ----
    # Crear la columna "Concatenado" concatenando "Año", "Medio", "Cadena" y "Cuartil" sin espacios
    base_prueba['Concatenado'] = base_prueba['Año'] + base_prueba['Medio'] + base_prueba['Cadena'] + base_prueba['Cuartil']

    # Inicializar la columna "Descuento"
    base_prueba['Descuento'] = 'Sin coincidencia'

    for i, row in base_prueba.iterrows():
     concatenado_value = row['Concatenado']

     # Buscar coincidencias exactas en la hoja "Cat_Descuento"
     match_descuento = cat_descuento[cat_descuento['Llave'] == concatenado_value]

     if not match_descuento.empty:
         # Si hay coincidencia, copiar el valor de la columna "Descuento Promedio"
         base_prueba.at[i, 'Descuento'] = match_descuento['Descuento Promedio'].iloc[0]

    # ---- PROCESO 6: Calcular "Coste ajustado" ----
    # Convertir "Cost" a valores numéricos
    base_prueba['Cost'] = base_prueba['Cost'].replace(r'[$,]', '', regex=True).astype(float)

    # Convertir "Descuento" a valores numéricos (ya está en formato decimal, no es porcentaje)
    base_prueba['Descuento'] = pd.to_numeric(base_prueba['Descuento'], errors='coerce').fillna(0)

    # Calcular el "Coste ajustado"
    base_prueba['Coste ajustado'] = base_prueba['Cost'] * (1 - base_prueba['Descuento'])


    # Mostrar las primeras filas del DataFrame resultante
    print(base_prueba.head())

    # Exportar el DataFrame a un archivo Excel llamado "BASE_LISTA.xlsx"
    base_prueba.to_excel(carpeta_salida / 'BASE_HR.xlsx', index=False)

    base_prueba.info()

# %% 4. Admetricks
if __name__ == "__main__":
    archivos_admetrics = (
        validar_archivos(args.admetricks, 4, "Admetricks")
        if args.admetricks is not None
        else seleccionar_archivos("Admetricks", 4)
    )
    # Las columnas comunes y deseadas específicas para los archivos de Admetrics
    columnas_comunes_2 = [
    'Fecha', 'Industria', 'Marca', 'Anunciante', 'Nombre de campaña',
    'Landing page de la campaña', 'Etiquetas de campaña', 'Sitio web',
    'Secciones del sitio web', 'Editor', 'Formato', 'Tamaño de aviso',
    'Duración (Video)', 'Omitible (Video)', 'Posición', 'Advertisement',
    'Screenshot', 'Países', 'Dispositivos', 'Hospedado por', 'Vendido por',
    'Impacto', 'Impresiones', 'Valorización (Local)', 'Valorización (Dólares)',
    'Ads Count', 'Web Report'
    ]

    # Leer los archivos locales de Admetricks.
    admetrics1 = pd.read_excel(archivos_admetrics[0])
    admetrics2 = pd.read_excel(archivos_admetrics[1])
    admetrics3 = pd.read_excel(archivos_admetrics[2])
    admetrics4 = pd.read_excel(archivos_admetrics[3])

    # Concatenar los cuatro DataFrames, seleccionando solo las columnas deseadas.
    # Advertencia: Si alguna columna de 'columnas_comunes_2' no existe en alguno de
    # los archivos, este código generará un error (KeyError).
    base_prueba_admetrics = pd.concat([
    admetrics1[columnas_comunes_2],
    admetrics2[columnas_comunes_2],
    admetrics3[columnas_comunes_2],
    admetrics4[columnas_comunes_2]
    ], ignore_index=True)
    # Filtrar el DataFrame base_prueba_admetrics eliminando las filas donde "Marca" coincida con "Eliminados" en cat_digital
    base_prueba_admetrics = base_prueba_admetrics[~base_prueba_admetrics['Marca'].isin(cat_digital['Eliminados'])]

    # Crear la columna "Año" extrayendo el año de la columna "Fecha"
    base_prueba_admetrics['Año'] = base_prueba_admetrics['Fecha'].dt.year

    # Crear la columna "Mes" extrayendo el mes de la columna "Fecha"
    base_prueba_admetrics['Mes'] = base_prueba_admetrics['Fecha'].dt.month

    # ---- PROCESO 1: Comparación con "Cat_Digital" ----
    # Asegurarse de que la columna "Marca" comience con mayúsculas
    base_prueba_admetrics['Marca'] = base_prueba_admetrics['Marca'].str.title()

    # Inicializar las columnas "Marca*" y "Subproducto" en el DataFrame base_prueba_admetrics
    base_prueba_admetrics['Marca*'] = 'Sin coincidencia'
    base_prueba_admetrics['Subproducto'] = 'Sin coincidencia'

    # Iterar sobre las filas del DataFrame base_prueba_admetrics
    for i, row in base_prueba_admetrics.iterrows():
     marca_value = row['Marca']  # Valor de la columna "Marca" en la fila actual

     # Buscar coincidencias exactas en la hoja "Cat_Digital"
     match_digital = cat_digital[cat_digital['Concat'] == marca_value]

     if not match_digital.empty:
         # Si hay coincidencia, copiar los valores de las columnas "Marca2" y "Subproducto" de "Cat_Digital"
         base_prueba_admetrics.at[i, 'Marca*'] = match_digital['Marca2'].iloc[0]
         base_prueba_admetrics.at[i, 'Subproducto'] = match_digital['Subproducto'].iloc[0]

    # ---- PROCESO 2: Comparación de "Marca*" con "Cat_Digital" ----
    # Inicializar la columna "Subproducto" en el DataFrame base_prueba_admetrics
    base_prueba_admetrics['Subproducto'] = 'Sin coincidencia'

    # Iterar sobre las filas del DataFrame base_prueba_admetrics
    for i, row in base_prueba_admetrics.iterrows():
     marca_estrella_value = row['Marca*']  # Valor de la columna "Marca*" en la fila actual

     # Buscar coincidencias exactas en la hoja "Cat_Digital"
     match_digital = cat_digital[cat_digital['Marca2'] == marca_estrella_value]

     if not match_digital.empty:
         # Si hay coincidencia, copiar el valor de la columna "Subproducto" de "Cat_Digital"
         base_prueba_admetrics.at[i, 'Subproducto'] = match_digital['Subproducto'].iloc[0]


    # ---- PROCESO 2: Comparación de "Subproducto" con "Cat_Master" ----
    # Inicializar las nuevas columnas en el DataFrame base_prueba_admetrics
    base_prueba_admetrics['Megamarca'] = 'Sin coincidencia'
    base_prueba_admetrics['Categoria'] = 'Sin coincidencia'
    base_prueba_admetrics['Tipo'] = 'Sin coincidencia'
    base_prueba_admetrics['Caracteristica'] = 'Sin coincidencia'

    # Iterar sobre las filas del DataFrame base_prueba_admetrics
    for i, row in base_prueba_admetrics.iterrows():
     subproducto_value = row['Subproducto']  # Valor de la columna "Subproducto" en la fila actual

     # Buscar coincidencias exactas en la hoja "Cat_Master"
     match_master = cat_master[cat_master['Subproducto'] == subproducto_value]

     if not match_master.empty:
         # Si hay coincidencia, copiar los valores de las columnas correspondientes de "Cat_Master"
         base_prueba_admetrics.at[i, 'Megamarca'] = match_master['Megamarca'].iloc[0]
         base_prueba_admetrics.at[i, 'Categoria'] = match_master['Categoria'].iloc[0]
         base_prueba_admetrics.at[i, 'Tipo'] = match_master['Tipo'].iloc[0]
         base_prueba_admetrics.at[i, 'Caracteristica'] = match_master['Caracteristica'].iloc[0]


    # ---- PROCESO 3: Comparación de "Megamarca" con "Cat_CIA" ----
    # Inicializar las nuevas columnas en el DataFrame base_prueba_admetrics
    base_prueba_admetrics['CIA'] = 'Sin coincidencia'
    base_prueba_admetrics['Open_CIA'] = 'Sin coincidencia'

    # Iterar sobre las filas del DataFrame base_prueba_admetrics
    for i, row in base_prueba_admetrics.iterrows():
     megamarca_value = row['Megamarca']  # Valor de la columna "Megamarca" en la fila actual

     # Buscar coincidencias exactas en la hoja "Cat_CIA"
     match_cia = cat_cia[cat_cia['Megamarca'] == megamarca_value]

     if not match_cia.empty:
         # Si hay coincidencia, copiar los valores de las columnas correspondientes de "Cat_CIA"
         base_prueba_admetrics.at[i, 'CIA'] = match_cia['CIA'].iloc[0]
         base_prueba_admetrics.at[i, 'Open_CIA'] = match_cia['Open_CIA'].iloc[0]


    # ---- PROCESO 4: Agregar columnas "Medio", "Descuento", "Cadena", "Coste ajustado" y "Concat" ----
    # Crear la columna "Medio" y asignar el valor "Digital" a todos los datos
    base_prueba_admetrics['Medio'] = 'Digital'

    # Crear la columna "Descuento" y asignar el valor 0 a todos los datos
    base_prueba_admetrics['Descuento'] = 0

    # Crear la columna "Cadena" y asignar el valor "Otros" a todos los datos
    base_prueba_admetrics['Cadena'] = 'Otros'

    # Crear la columna "Coste ajustado" copiando los valores de la columna "Valorización (Local)"
    base_prueba_admetrics['Coste ajustado'] = base_prueba_admetrics['Valorización (Local)']

    # Crear la columna "Concat" copiando los valores de la columna "Marca"
    base_prueba_admetrics['Concat'] = base_prueba_admetrics['Marca']


    # ---- PROCESO 5: Comparación de "Categoria" con "Cat_Grupo" ----
    # Inicializar la columna "Grupo" en el DataFrame base_prueba_admetrics
    base_prueba_admetrics['Grupo'] = 'Sin coincidencia'

    # Iterar sobre las filas del DataFrame base_prueba_admetrics
    for i, row in base_prueba_admetrics.iterrows():
     categoria_value = row['Categoria']  # Valor de la columna "Categoria" en la fila actual

     # Buscar coincidencias exactas en la hoja "Cat_Grupo"
     match_grupo = cat_grupo[cat_grupo['Categoria'] == categoria_value]

     if not match_grupo.empty:
         # Si hay coincidencia, copiar el valor de la columna "Grupo" de "Cat_Grupo"
         base_prueba_admetrics.at[i, 'Grupo'] = match_grupo['Grupo'].iloc[0]


    # ---- PROCESO 6: Comparación de "Mes" con "Cat_Q" ----
    # Convertir la columna "Mes" en ambas tablas a tipo int para formato general
    base_prueba_admetrics['Mes'] = pd.to_numeric(base_prueba_admetrics['Mes'], errors='coerce')  # Convertir a números enteros
    cat_q['Mes'] = pd.to_numeric(cat_q['Mes'], errors='coerce')  # Convertir también en cat_q

    # Inicializar la columna "Cuartil" en el DataFrame base_prueba_ooh
    base_prueba_admetrics['Cuartil'] = 'Sin coincidencia'

    # Iterar sobre las filas del DataFrame base_prueba_ooh
    for i, row in base_prueba_admetrics.iterrows():
     mes_value = row['Mes']  # Valor de la columna "Mes" en la fila actual (ya en formato numérico)

     # Buscar coincidencias exactas en la hoja "Cat_Q"
     match_q = cat_q[cat_q['Mes'] == mes_value]

     if not match_q.empty:
         # Si hay coincidencia, copiar el valor de la columna "Número de Q" de "Cat_Q" a "Cuartil"
         base_prueba_admetrics.at[i, 'Cuartil'] = match_q['Número de Q'].iloc[0]


    # ---- PROCESO 7: Crear la columna "Concatenado" ----
    # Concatenar las columnas "Año", "Medio", "Cadena" y "Cuartil" sin espacios
    base_prueba_admetrics['Concatenado'] = (
     base_prueba_admetrics['Año'].astype(str) +
     base_prueba_admetrics['Medio'].astype(str) +
     base_prueba_admetrics['Cadena'].astype(str) +
     base_prueba_admetrics['Cuartil'].astype(str)
    )


    # Mostrar las primeras filas del DataFrame resultante
    print(base_prueba_admetrics.head())

    # Exportar el DataFrame a un archivo Excel llamado "BASE_LISTA.xlsx"
    base_prueba_admetrics.to_excel(carpeta_salida / 'BASE_Admetricks.xlsx', index=False)

# %% 5. OOH
if __name__ == "__main__":
    archivos_ooh = (
        validar_archivos(args.ooh, 3, "OOH")
        if args.ooh is not None
        else seleccionar_archivos("OOH", 3)
    )
    # Las columnas comunes y deseadas específicas para los archivos de OOH
    columnas_ooh_deseadas = [
    'Fecha', 'Sector', 'SubSector', 'Categoria', 'Anunciante', 'Marca',
    'ProductoTag', 'Tag', 'Agencia', 'Direccion', 'Ciudad', 'Url',
    'TipoPublicidad', 'Lat', 'Lon', 'Proveedor', 'Tamano', 'Inserciones', 'Costo'
    ]

    # Leer los archivos locales de OOH.
    ooh1 = pd.read_excel(archivos_ooh[0])
    ooh2 = pd.read_excel(archivos_ooh[1])
    ooh3 = pd.read_excel(archivos_ooh[2])

    # Concatenar los tres DataFrames, seleccionando solo las columnas deseadas.
    # Advertencia: Si alguna columna de 'columnas_ooh_deseadas' no existe en alguno
    # de los archivos, este código generará un error (KeyError).
    base_prueba_ooh = pd.concat([
     ooh1[columnas_ooh_deseadas],
     ooh2[columnas_ooh_deseadas],
     ooh3[columnas_ooh_deseadas]
    ], ignore_index=True)  # Usamos ignore_index=True para reiniciar el índice en el DataFrame combinado

    # Filtrar el DataFrame base_prueba_ooh eliminando las filas donde "Marca" coincida con "Eliminados" en cat_ooh
    base_prueba_ooh = base_prueba_ooh[~base_prueba_ooh['Marca'].isin(cat_ooh['Eliminados'])]


    # Asegurarse de que la columna "Fecha" esté en formato datetime
    base_prueba_ooh['Fecha'] = pd.to_datetime(base_prueba_ooh['Fecha'], errors='coerce')

    # Crear la columna "Año" extrayendo el año de la columna "Fecha"
    base_prueba_ooh['Año'] = base_prueba_ooh['Fecha'].dt.year

    # Crear la columna "Mes" extrayendo el mes de la columna "Fecha"
    base_prueba_ooh['Mes'] = base_prueba_ooh['Fecha'].dt.month

    # Crear la columna "Concat" concatenando "Categoria", "Marca" y "ProductoTag" con un espacio entre cada dato
    base_prueba_ooh['Concat'] = (
     base_prueba_ooh['Categoria'].astype(str) + ' ' +
     base_prueba_ooh['Marca'].astype(str) + ' ' +
     base_prueba_ooh['ProductoTag'].astype(str)
    )

    # ---- PROCESO 1: Comparación de "Concat" con "Cat_OOH" ----
    # Inicializar las nuevas columnas en el DataFrame base_prueba_ooh
    base_prueba_ooh['Marca*'] = 'Sin coincidencia'
    base_prueba_ooh['Subproducto'] = 'Sin coincidencia'
    base_prueba_ooh['Caracteristica'] = 'Sin coincidencia'

    # Iterar sobre las filas del DataFrame base_prueba_ooh
    for i, row in base_prueba_ooh.iterrows():
     concat_value = row['Concat']  # Valor de la columna "Concat" en la fila actual

     # Buscar coincidencias exactas en la hoja "Cat_OOH"
     match_ooh = cat_ooh[cat_ooh['Concat2'] == concat_value]

     if not match_ooh.empty:
         # Si hay coincidencia, copiar los valores de las columnas correspondientes de "Cat_OOH"
         base_prueba_ooh.at[i, 'Marca*'] = match_ooh['Marca'].iloc[0]
         base_prueba_ooh.at[i, 'Subproducto'] = match_ooh['Subproducto'].iloc[0]
         base_prueba_ooh.at[i, 'Caracteristica'] = match_ooh['Caracteristica'].iloc[0]

    # ---- PROCESO 2: Comparación de "Subproducto" con "Cat_Master" ----
    # Inicializar las nuevas columnas en el DataFrame base_prueba_ooh
    base_prueba_ooh['Megamarca'] = 'Sin coincidencia'
    base_prueba_ooh['Categoria'] = 'Sin coincidencia'
    base_prueba_ooh['Tipo'] = 'Sin coincidencia'

    # Iterar sobre las filas del DataFrame base_prueba_ooh
    for i, row in base_prueba_ooh.iterrows():
     subproducto_value = row['Subproducto']  # Valor de la columna "Subproducto" en la fila actual

     # Buscar coincidencias exactas en la hoja "Cat_Master"
     match_master = cat_master[cat_master['Subproducto'] == subproducto_value]

     if not match_master.empty:
         # Si hay coincidencia, copiar los valores de las columnas correspondientes de "Cat_Master"
         base_prueba_ooh.at[i, 'Megamarca'] = match_master['Megamarca'].iloc[0]
         base_prueba_ooh.at[i, 'Categoria'] = match_master['Categoria'].iloc[0]
         base_prueba_ooh.at[i, 'Tipo'] = match_master['Tipo'].iloc[0]


    # ---- PROCESO 3: Comparación de "Megamarca" con "Cat_CIA" ----
    # Inicializar las nuevas columnas en el DataFrame base_prueba_ooh
    base_prueba_ooh['CIA'] = 'Sin coincidencia'
    base_prueba_ooh['Open_CIA'] = 'Sin coincidencia'

    # Iterar sobre las filas del DataFrame base_prueba_ooh
    for i, row in base_prueba_ooh.iterrows():
     megamarca_value = row['Megamarca']  # Valor de la columna "Megamarca" en la fila actual

     # Buscar coincidencias exactas en la hoja "Cat_CIA"
     match_cia = cat_cia[cat_cia['Megamarca'] == megamarca_value]

     if not match_cia.empty:
         # Si hay coincidencia, copiar los valores de las columnas correspondientes de "Cat_CIA"
         base_prueba_ooh.at[i, 'CIA'] = match_cia['CIA'].iloc[0]
         base_prueba_ooh.at[i, 'Open_CIA'] = match_cia['Open_CIA'].iloc[0]

    # ---- PROCESO 4: Crear las columnas "Medio" y "Cadena" ----
    # Crear la columna "Medio" y asignar el valor "OOH" a todas las filas
    base_prueba_ooh['Medio'] = 'OOH'

    # Crear la columna "Cadena" y asignar el valor "Otros" a todas las filas
    base_prueba_ooh['Cadena'] = 'Otros'

    # ---- PROCESO 5: Comparación de "Categoria" con "Cat_Grupo" ----
    # Inicializar la columna "Grupo" en el DataFrame base_prueba_ooh
    base_prueba_ooh['Grupo'] = 'Sin coincidencia'

    # Iterar sobre las filas del DataFrame base_prueba_ooh
    for i, row in base_prueba_ooh.iterrows():
     categoria_value = row['Categoria']  # Valor de la columna "Categoria" en la fila actual

     # Buscar coincidencias exactas en la hoja "Cat_Grupo"
     match_grupo = cat_grupo[cat_grupo['Categoria'] == categoria_value]

     if not match_grupo.empty:
         # Si hay coincidencia, copiar el valor de la columna "Grupo" de "Cat_Grupo"
         base_prueba_ooh.at[i, 'Grupo'] = match_grupo['Grupo'].iloc[0]

    # ---- PROCESO 6: Comparación de "Mes" con "Cat_Q" ----
    # Convertir la columna "Mes" en ambas tablas a tipo int para formato general
    base_prueba_ooh['Mes'] = pd.to_numeric(base_prueba_ooh['Mes'], errors='coerce')  # Convertir a números enteros
    cat_q['Mes'] = pd.to_numeric(cat_q['Mes'], errors='coerce')  # Convertir también en cat_q

    # Inicializar la columna "Cuartil" en el DataFrame base_prueba_ooh
    base_prueba_ooh['Cuartil'] = 'Sin coincidencia'

    # Iterar sobre las filas del DataFrame base_prueba_ooh
    for i, row in base_prueba_ooh.iterrows():
     mes_value = row['Mes']  # Valor de la columna "Mes" en la fila actual (ya en formato numérico)

     # Buscar coincidencias exactas en la hoja "Cat_Q"
     match_q = cat_q[cat_q['Mes'] == mes_value]

     if not match_q.empty:
         # Si hay coincidencia, copiar el valor de la columna "Número de Q" de "Cat_Q" a "Cuartil"
         base_prueba_ooh.at[i, 'Cuartil'] = match_q['Número de Q'].iloc[0]

    # ---- PROCESO 7: Crear la columna "Concatenado" ----
    # Concatenar las columnas "Año", "Medio", "Cadena" y "Cuartil" sin espacios
    base_prueba_ooh['Concatenado'] = (
     base_prueba_ooh['Año'].astype(str) +
     base_prueba_ooh['Medio'].astype(str) +
     base_prueba_ooh['Cadena'].astype(str) +
     base_prueba_ooh['Cuartil'].astype(str)
    )

    #






    # ---- PROCESO 8: Crear las columnas "Descuento" y "Coste ajustado" ----
    # Crear la columna "Descuento" con el valor 0.72 para todas las filas
    base_prueba_ooh['Descuento'] = 0.1001








    # Crear la columna "Coste ajustado" calculando "Costo * (1 - Descuento)"
    base_prueba_ooh['Coste ajustado'] = base_prueba_ooh['Costo'] * (1 - base_prueba_ooh['Descuento'])


    # Mostrar las primeras filas del DataFrame resultante
    print(base_prueba_ooh.head())

    # Exportar el DataFrame a un archivo Excel llamado "BASE_LISTA.xlsx"
    base_prueba_ooh.to_excel(carpeta_salida / 'BASE_OOH.xlsx', index=False)

# %% 6. Union de las bases
if __name__ == "__main__":
    # Columnas extra ya existentes en las bases y que se deben mantener
    columnas_extra = [
     'Año', 'Mes', 'Semana', 'Concat', 'Marca*', 'Subproducto', 'Caracteristica',
     'Megamarca', 'Categoria', 'Tipo', 'CIA', 'Open_CIA', 'Medio',
     'Grupo', 'Cadena', 'Cuartil', 'Concatenado', 'Descuento', 'Coste ajustado'
    ]

    # ===============================
    # BASE_PRUEBA (HR)
    # ===============================
    cols_base_prueba_original = [
     'Month', 'Week', 'Anunciante', 'Marca', 'Producto', 'Submarca', 'Version',
     'Channel', 'Tipo de Spot', 'Tipo de Medio', 'Cost', 'Rtg% A13-54'
    ]

    rename_base_prueba = {
     'Month': 'Month',
     'Anunciante': 'Anunciante_base',
     'Marca': 'Marca_base',
     'Producto': 'Producto_base',
     'Submarca': 'Submarca_base',
     'Version': 'Version_base',
     'Channel': 'Channel_base',
     'Tipo de Spot': 'Tipo de Spot_base',
     'Tipo de Medio': 'Tipo de Medio_base',
     'Cost': 'Cost_base',
     'Rtg% A13-54': 'Rtg% P13-54 [MEM] [HO]'
    }

    # Mantener las columnas solicitadas + las extras que ya estaban
    cols_a_mantener_hr = [col for col in cols_base_prueba_original + columnas_extra if col in base_prueba.columns]
    base_prueba = base_prueba[cols_a_mantener_hr].rename(columns=rename_base_prueba)

    # ===============================
    # BASE_PRUEBA_ADMETRICS
    # ===============================
    cols_base_prueba_admetrics_original = [
     'Fecha', 'Anunciante', 'Marca', 'Nombre de campaña', 'Etiquetas de campaña', 'Valorización (Local)'
    ]

    rename_base_prueba_admetrics = {
     'Fecha': 'Month',
     'Anunciante': 'Anunciante_base',
     'Marca': 'Marca_base',
     'Nombre de campaña': 'Producto_base',
     'Etiquetas de campaña': 'Version_base',
     'Valorización (Local)': 'Cost_base'
    }

    cols_a_mantener_admetrics = [col for col in cols_base_prueba_admetrics_original + columnas_extra if col in base_prueba_admetrics.columns]
    base_prueba_admetrics = base_prueba_admetrics[cols_a_mantener_admetrics].rename(columns=rename_base_prueba_admetrics)

    # Agregar columnas requeridas y llenarlas con valores
    base_prueba_admetrics['Submarca_base'] = "Otros"
    base_prueba_admetrics['Channel_base'] = "Otros"
    base_prueba_admetrics['Tipo de Spot_base'] = "Otros"
    base_prueba_admetrics['Tipo de Medio_base'] = "Digital"


    # Agregar la columna "Semana" con el valor "No se considera" si no existe
    if 'Semana' not in base_prueba_admetrics.columns:
     base_prueba_admetrics['Semana'] = "No se considera"

    # ===============================
    # BASE_PRUEBA_OOH
    # ===============================
    cols_base_prueba_ooh_original = [
     'Fecha', 'Anunciante', 'Marca', 'ProductoTag', 'Tag','Costo'
    ]

    rename_base_prueba_ooh = {
     'Fecha': 'Month',
     'Anunciante': 'Anunciante_base',
     'Marca': 'Marca_base',
     'ProductoTag': 'Producto_base',
     'Tag': 'Version_base',
     'Costo': 'Cost_base',
    }



    cols_a_mantener_ooh = [col for col in cols_base_prueba_ooh_original + columnas_extra if col in base_prueba_ooh.columns]
    base_prueba_ooh = base_prueba_ooh[cols_a_mantener_ooh].rename(columns=rename_base_prueba_ooh)

    # Agregar columnas requeridas y llenarlas con valores
    base_prueba_ooh['Submarca_base'] = "Otros"
    base_prueba_ooh['Channel_base'] = "Otros"
    base_prueba_ooh['Tipo de Spot_base'] = "Otros"
    base_prueba_ooh['Tipo de Medio_base'] = "OOH"

    # Agregar la columna "Semana" con el valor "No se considera" si no existe
    if 'Semana' not in base_prueba_ooh.columns:
     base_prueba_ooh['Semana'] = "No se considera"


    # ===============================
    # Comprobación rápida
    # ===============================
    print("BASE HR columnas:", base_prueba.columns.tolist())
    print("BASE ADMETRICS columnas:", base_prueba_admetrics.columns.tolist())
    print("BASE OOH columnas:", base_prueba_ooh.columns.tolist())

    # UNIR TODAS LAS BASES
    # ===============================
    unidas = pd.concat([base_prueba, base_prueba_admetrics, base_prueba_ooh], ignore_index=True)

    # Orden final de columnas que quieres
    orden_columnas = [
     'Month', 'Anunciante_base', 'Marca_base', 'Producto_base', 'Submarca_base',
     'Version_base', 'Channel_base', 'Tipo de Spot_base', 'Tipo de Medio_base',
     'Cost_base', 'Rtg% P13-54 [MEM] [HO]', 'Año', 'Mes', 'Semana', 'Concat', 'Marca*',
     'Subproducto', 'Caracteristica', 'Megamarca', 'Categoria', 'Tipo', 'CIA',
     'Open_CIA', 'Medio', 'Grupo', 'Cadena', 'Cuartil', 'Concatenado',
     'Descuento', 'Coste ajustado'
    ]

    # Filtrar y reordenar columnas según el orden definido (solo las que existen)
    orden_final = [col for col in orden_columnas if col in unidas.columns]
    unidas = unidas[orden_final]


    # EXPORTAR UN SOLO EXCEL CON UNA HOJA

    nombre_excel = carpeta_salida / "Bases_Unidas.xlsx"
    unidas.to_excel(nombre_excel, index=False)

    print(f"Archivo '{nombre_excel}' creado con todas las bases unidas y columnas ordenadas.")

    unidas.info()

# %% 7. Base lista
if __name__ == "__main__":
    # ---- CREAR BASE_LISTA ----
    # Seleccionar las columnas comunes, incluyendo la excepción 'Rtg% P13-54 [MEM] [HO]'
    columnas_comunes = [
     'Año', 'Mes', 'Semana', 'Rtg% P13-54 [MEM] [HO]', 'Concat', 'Marca*', 'Subproducto',
     'Megamarca', 'Categoria', 'Tipo', 'Caracteristica', 'CIA', 'Open_CIA',
     'Medio', 'Grupo', 'Cadena', 'Cuartil', 'Concatenado', 'Descuento', 'Coste ajustado', 'Version_base'
    ]

    # Crear un DataFrame temporal para cada base
    base_ooh_temp = base_prueba_ooh.copy()
    base_hr_temp = base_prueba.copy()
    base_admetrics_temp = base_prueba_admetrics.copy()

    # Añadir las columnas 'Rtg% P13-54 [MEM] [HO]' y 'Semana' a las bases que no las tienen
    if 'Rtg% P13-54 [MEM] [HO]' not in base_prueba_ooh.columns:
     base_ooh_temp['Rtg% P13-54 [MEM] [HO]'] = "No se considera"

    if 'Rtg% P13-54 [MEM] [HO]' not in base_prueba_admetrics.columns:
     base_admetrics_temp['Rtg% P13-54 [MEM] [HO]'] = "No se considera"

    if 'Semana' not in base_prueba_ooh.columns:
     base_ooh_temp['Semana'] = "No se considera"

    if 'Semana' not in base_prueba_admetrics.columns:
     base_admetrics_temp['Semana'] = "No se considera"


    # Concatenar las tres bases usando las columnas comunes
    BASE_LISTA = pd.concat([
     base_ooh_temp[columnas_comunes],
     base_hr_temp[columnas_comunes],
     base_admetrics_temp[columnas_comunes]
    ], ignore_index=True)

     # ---- Crear las nuevas columnas antes de exportar ----

    # Crear la columna "Csoto_master" dividiendo "Coste ajustado" entre 1,000,000
    BASE_LISTA['Costo_master'] = BASE_LISTA['Coste ajustado'] / 1_000_000

    # Crear la columna "Brand" copiando literalmente los valores de "Marca*"
    BASE_LISTA['Brand'] = BASE_LISTA['Marca*']

    # Crear la columna "Mes_" con el nombre del mes en texto basado en la columna Mes
    mes_dict = {
     1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril', 5: 'Mayo',
     6: 'Junio', 7: 'Julio', 8: 'Agosto', 9: 'Septiembre', 10: 'Octubre',
     11: 'Noviembre', 12: 'Diciembre'
    }
    BASE_LISTA['Mes_'] = BASE_LISTA['Mes'].map(mes_dict)

    # Crear la columna "Fecha_" combinando "Mes" y "Año", iniciando siempre el día con "1"
    BASE_LISTA['Fecha_'] = pd.to_datetime(
     dict(year=BASE_LISTA['Año'], month=BASE_LISTA['Mes'], day=1)
    )


    # ---- Reordenar las columnas de BASE_LISTA ----
    # Especificamos el orden deseado de las columnas
    column_order = [
     'Rtg% P13-54 [MEM] [HO]', 'Año', 'Mes','Semana', 'Concat', 'Marca*',
     'Subproducto', 'Megamarca', 'Categoria', 'Tipo', 'Caracteristica',
     'CIA', 'CIA_abierta', 'Medio', 'Grupo', 'Cadena', 'Cuartil',
     'Concatenado', 'Descuento', 'Costo ajustado', 'Costo_master',
     'Brand', 'Mes_', 'Fecha_'
    ]


    # Exportar el DataFrame a un archivo Excel llamado "BASE_LISTA.xlsx"
    BASE_LISTA.to_excel(carpeta_salida / 'BASE_LISTA.xlsx', index=False)

    BASE_LISTA.info()

# %% 8. Checkpoint y auditoria
if __name__ == "__main__":
    # --- PARA MEJORAR RENDIMIENTO ---

    # 1. Creamos la copia
    df_base = BASE_LISTA.copy()

    # 2. Exportamos a CSV temporal (para limpiar metadatos de Excel)
    df_base.to_csv((carpeta_salida / "checkpoint_base.csv"), index=False, encoding='utf-8-sig')

    # 3. Lo volvemos a leer (esto acelera todos los procesos de limpieza que siguen)
    df_base = pd.read_csv((carpeta_salida / "checkpoint_base.csv"))

    print("El código ahora correrá mucho más rápido.")

    # ==========================================
    # AUDITORÍA DE FILAS (Houston, aquí los números)
    # ==========================================

    df_base['Año'] = pd.to_numeric(df_base['Año'], errors='coerce')
    df_base['Mes'] = pd.to_numeric(df_base['Mes'], errors='coerce')
    # 1. Definimos los "Culpables"
    mask_marcas = df_base['Marca*'].str.strip().isin(['Sin coincidencia', 'No_se_considera', 'No se considera'])
    mask_2025 = (df_base['Año'] == 2025)
    mask_abril = (df_base['Año'] == 2026) & (df_base['Mes'] == 4)

    # 2. Conteos individuales
    n_marcas = mask_marcas.sum()
    n_2025 = mask_2025.sum()
    n_abril = mask_abril.sum()

    # 3. Conteo de la Unión (Evitando contar doble si una fila es de 2025 Y es 'Sin coincidencia')
    # Esto representa el TOTAL real que el código está eliminando
    mask_total_eliminados = mask_marcas | mask_2025 | mask_abril
    n_total_eliminados = mask_total_eliminados.sum()

    # 4. Versiones que hay que borrar (son de Argentina)
    mask_villa = df_base['Version_base'].isin([
        '(EXT) HECHA CON AGUA MINERAL VILLA DEL SUR Y JUGO DE FRUTAS',
        'HECHA CON AGUA MINERAL VILLA DEL SUR Y JUGO DE FRUTAS'
    ])
    n_eliminados_villa = mask_villa.sum()

    # Final

    mask_total_final = mask_total_eliminados | mask_villa
    n_total_final_eliminados = mask_total_final.sum()

    # 5. Resultados
    print("--- REPORTE DE FILAS ELIMINADAS ---")
    print(f"Total inicial en df_base: {len(df_base):,}")
    print(f"(-) Por Marcas (Sin coinc./No cons.): {n_marcas:,}")
    print(f"(-) Por Año 2025:                    {n_2025:,}")
    print(f"(-) Por Abril 2026:                  {n_abril:,}")
    print(f"(-) Villa del Sur:                   {n_eliminados_villa:,}")
    print("--------------------------------------")
    print(f"(=) TOTAL ÚNICO A ELIMINAR:          {n_total_final_eliminados:,}")
    print(f"(=) RESULTADO FINAL ESPERADO:        {len(df_base) - n_total_final_eliminados:,}")

    # ==========================================
# %% 9. Limpieza final y exportacion
if __name__ == "__main__":
    # 1. DEFINICIONES (Los Diccionarios)
    # ==========================================
    # Definimos primero todas las "máscaras" o mapeos para que Python los conozca


    df_base = df_base.rename(columns={
        "Marca*": "Marca",
        "Rtg% P13-54 [MEM] [HO]": "RTG"
    })

    # Ahora, tus limpiezas ortográficas se aplicarán sobre la NUEVA columna 'Marca'
    # (la que antes era Marca*)


    submarca_jumex = { "Jumex Fresh": "Jumex", "Ami Frut": "Jumex", "Bida": "Jumex" }
    submarca_DelValle = { "Del Valle Frut": "Del Valle", "Del Valle Junior": "Del Valle" }
    submarca_fabs = { 'Amstel Ultra': 'Amstel', 'Amstel Ultra Seltzer': 'Amstel', 'Michelob Ultra': 'Michelob' }

    limpieza_marcas = {
        "AGA de manzana": "AGA de Manzana", "Aguas frescas": "Aguas Frescas",
        "Del valle ": "Del Valle", "Del valle Fizz": "Del Valle Fizz",
        "Del valle Junior": "Del Valle Junior", "Fresca fusión": "Fresca Fusión",
        "Peñafiel mineral": "Peñafiel Mineral", "Peñafiel sabores": "Peñafiel Sabores",
        "Red bull": "Red Bull", "Schwepps 0.0 ": "Schwepps 0.0",
        "Schwepps mineral ": "Schwepps Mineral", "Skarch ": "Skarch",
        "Topochico mineral ": "Topochico Mineral", "Del valle": "Del Valle"
    }

    mapeo_cat = { "FSD": "Flavors", "Rehydration": "ORS", "Coke": "Colas" }
    mapeo_brand_extra = { "Del Valle Junior": "Del Valle", "Del Valle Frut": "Del Valle", "Néctasis": "Boing" }
    mapeo_subproducto = { "Hidratación con Causa": "Jumex Hydrolit", "Vitamins": "Suerox" }

    # Listas para la lógica de CIA
    tccc = ["Bacardi & Coca-Cola", "Topo Chico HS", "Jack&Coke"]
    heineken = ["Amstel Ultra", "Amstel", "Amstel Ultra Seltzer", "Heineken"]
    modelo = ["Michelin", "Michelob Ultra"]
    cuervo = ["Caribe Cooler", "New Mix", "Tequila New Mix", "Seagram's HS"]
    fabs_list = ["CERVEZAS", "COCKTAILES/MEZCLADORES", "COCTELES/MEZCLADORES",
                 "COMPARTIDA BEBIDAS ALCOHOLICAS", "HARD SELTZERS", "TEQUILAS"]


    # ==========================================
    # 3. EJECUCIÓN DE LA LIMPIEZA
    # ==========================================
    df_base = df_base.rename(columns={"Rtg% P13-54 [MEM] [HO]": "RTG"})
    df_base["RTG"] = df_base["RTG"].replace("No se considera", 0).fillna(0).astype(float)
    # Creamos el objeto Datetime (esto mantiene la naturaleza del dato)
    df_base['Date_'] = pd.to_datetime(dict(year=df_base['Año'], month=df_base['Mes'], day=1))
    # Convertimos a Periodo de día (esto elimina la micro-precisión de nanosegundos)
    # pero lo mantenemos como objeto temporal, no como string.
    df_base['Date_'] = df_base['Date_'].dt.date


    df_base['Marca'] = df_base['Marca'].replace(limpieza_marcas)


    # Usamos str.strip() por seguridad para quitar espacios en blanco
    df_base = df_base[df_base['Marca'].str.strip() != 'Sin coincidencia']

    # También quitamos el que tenías de "No_se_considera" por si las moscas
    df_base = df_base[df_base['Marca'].str.strip() != 'No_se_considera']



    # Limpieza de Semana y Week_dummy
    df_base['Semana'] = df_base['Semana'].fillna(0).replace("No se considera", 0).astype(float)
    df_base["Week_dummy"] = np.where(df_base["Semana"] != 0, 1, 0)

    # Mapeos de Brand (Aquí ya están definidos arriba, no habrá error)
    df_base["Brand"] = df_base["Brand"].replace(submarca_jumex)
    df_base["Brand"] = df_base["Brand"].replace(submarca_DelValle)
    df_base["Brand"] = df_base["Brand"].replace(submarca_fabs)
    df_base["Brand"] = df_base["Brand"].replace(mapeo_brand_extra)

    # Limpieza ortográfica y categorías
    df_base['Marca'] = df_base['Marca'].replace(limpieza_marcas)
    df_base.loc[df_base["Categoria"].isin(fabs_list), "Categoria"] = "FABs"
    df_base["Categoria"] = df_base["Categoria"].replace(mapeo_cat)
    df_base["Subproducto"] = df_base["Subproducto"].replace(mapeo_subproducto)

    # Lógica Vectorizada de CIA
    condiciones = [
        df_base['Subproducto'].isin(tccc),
        df_base['Subproducto'].isin(heineken),
        df_base['Subproducto'].isin(modelo),
        df_base['Subproducto'].isin(cuervo),
        (df_base['CIA'].notnull()) & (df_base['CIA'].astype(str).str.strip() != "")
    ]
    resultados = ["TCCC", "Heineken México", "Grupo Modelo", "Casa Cuervo", df_base['CIA']]
    df_base['CIA'] = np.select(condiciones, resultados, default=df_base['CIA'])

    # Crear columnas que el Dashboard viejo espera para no romper visuales
    df_base["Categoria_"] = ""
    df_base["CIA2"] = ""

    column_order = [
        'RTG', 'Año', 'Mes', 'Semana', 'Concat', 'Marca', 'Subproducto',
        'Megamarca', 'Categoria', 'Tipo', 'Caracteristica', 'CIA', 'Open_CIA',
        'Medio', 'Grupo', 'Cadena', 'Cuartil', 'Concatenado', 'Descuento',
        'Coste ajustado', 'Costo_master', 'Date_', 'Brand', 'Mes_',
        'Week_dummy', 'Categoria_', 'CIA2', 'Version_base'
    ]

    # Reordenamos las columnas.
    df_base = df_base[column_order]

    df_base['Año'] = pd.to_numeric(df_base['Año'], errors='coerce')
    df_base['Mes'] = pd.to_numeric(df_base['Mes'], errors='coerce')

    # Filtro 1: Borrar todo lo que sea 2025
    df_base = df_base[df_base['Año'] != 2025]

    # Borramos las versiones de argentina
    versiones = df_base[df_base['Version_base'].isin([
        '(EXT) HECHA CON AGUA MINERAL VILLA DEL SUR Y JUGO DE FRUTAS',
        'HECHA CON AGUA MINERAL VILLA DEL SUR Y JUGO DE FRUTAS'
    ])].index

    df_base.drop(versiones, inplace=True)

    # Filtro 2: Borrar lo que sea Abril (4) del 2026
    # Usamos una máscara para mantener todo lo que NO cumpla esa condición
    df_base = df_base[~((df_base['Año'] == 2026) & (df_base['Mes'] == 4))]



    # Limpiando los tipos de datos

    # --- LIMPIEZA DE COLUMNAS DE DINERO ---
    # Quitamos símbolos de pesos y comas para que puedan ser números reales
    for col in ['Coste ajustado', 'Costo_master']:
        # Convertimos a string por si acaso, limpiamos y pasamos a float
        df_base[col] = df_base[col].astype(str).replace(r'[$,]', '', regex=True)
        df_base[col] = pd.to_numeric(df_base[col], errors='coerce').fillna(0)

    # --- LIMPIEZA ESTÉTICA DE SEMANA ---
    # Pasamos de 1.0 a 1 (entero)
    df_base['Semana'] = df_base['Semana'].astype(int)


    # Orden de columnas

    # --- REORDENAR: RTG AL INICIO ---

    # 1. Creamos la lista de columnas poniendo 'RTG' primero
    # y sumamos el resto de las columnas (quitando 'RTG' de su posición original)
    columnas = ['RTG'] + [col for col in df_base.columns if col != 'RTG']

    # 2. Reindexamos el DataFrame
    df_base = df_base[columnas]


    # --- REORDENAR Y LIMPIEZA FINAL DE COLUMNAS ---
    # 1. Definimos las columnas que queremos REALMENTE en el CSV (las 27 originales)
    columnas_finales = [
        'RTG', 'Año', 'Mes', 'Semana', 'Concat', 'Marca', 'Subproducto',
        'Megamarca', 'Categoria', 'Tipo', 'Caracteristica', 'CIA', 'Open_CIA',
        'Medio', 'Grupo', 'Cadena', 'Cuartil', 'Concatenado', 'Descuento',
        'Coste ajustado', 'Costo_master', 'Date_', 'Brand', 'Mes_',
        'Week_dummy', 'Categoria_', 'CIA2'
    ]

    # 2. Reindexamos eliminando 'Version_base' de un plumazo
    df_base = df_base[columnas_finales]

    # ==========================================
    # 4. EXPORTACIÓN A CSV
    # ==========================================
    nombre_csv = carpeta_salida / "Base_2026_Q1.csv"
    # IMPORTANTE: No uses sep=',' si usas encoding='utf-8-sig',
    # con dejarlo por defecto Excel lo leerá perfecto.
    df_base.to_csv(nombre_csv, index=False, encoding='utf-8-sig')

    print(f"Archivo {nombre_csv} generado con {len(df_base):,} filas.")

    df_base.info()
    print(f"Reporte terminado. Resultados guardados en: {carpeta_salida}")
