import streamlit as st
import pandas as pd
from io import BytesIO
import os

def procesar_archivo(file, nombre_archivo):
    df = pd.read_excel(file)
    columnas_deseadas = ["Identificación", "NUI", "Factura", "Centro de costo", "Saldo Factura", "Mes de Cobro"]
    
    # Verificar si las columnas existen en el archivo
    columnas_presentes = [col for col in columnas_deseadas if col in df.columns]
    df_filtrado = df[columnas_presentes]

    # Eliminar guiones en "NUI" y "Factura"
    if "NUI" in df_filtrado.columns:
        df_filtrado["NUI"] = df_filtrado["NUI"].astype(str).str.replace("-", "", regex=True)
    if "Factura" in df_filtrado.columns:
        df_filtrado["Factura"] = df_filtrado["Factura"].astype(str).str.replace("-", "", regex=True)

    # Convertir "Centro de costo" a mayúsculas
    if "Centro de costo" in df_filtrado.columns:
        df_filtrado["Centro de costo"] = df_filtrado["Centro de costo"].astype(str).str.upper()

    # Reemplazar valores nulos con "NA"
    df_filtrado.fillna("NA", inplace=True)

    # Filtrar filas donde "Factura" esté vacía
    if "Factura" in df_filtrado.columns:
        df_filtrado = df_filtrado[df_filtrado["Factura"] != "NA"]

    # Procesar "Mes de Cobro" si existe
    if "Mes de Cobro" in df_filtrado.columns:
        df_filtrado["Mes de Cobro"] = df_filtrado["Mes de Cobro"].astype(str)
        df_filtrado[["mes", "año"]] = df_filtrado["Mes de Cobro"].str.split(" ", expand=True).fillna("")

        meses_dict = {
            "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
            "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12
        }

        df_filtrado["mes"] = df_filtrado["mes"].str.lower().map(meses_dict)
        df_filtrado["año"] = pd.to_numeric(df_filtrado["año"], errors='coerce')

    # Agregar columna con el nombre del archivo
    df_filtrado.insert(0, "nombre_archivo", nombre_archivo)

    return df_filtrado

def generar_xlsx(df):
    output = BytesIO()
    df.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)
    return output

def generar_csv(df):
    output = BytesIO()
    df.to_csv(output, index=False, encoding='utf-8')
    output.seek(0)
    return output

# Configuración de la página
st.set_page_config(page_title="Captura de datos por columna - Cartera", page_icon="📂", layout="centered")
st.title("📂 Captura de datos por columna - Cartera")

st.markdown("Sube un archivo Excel, extrae columnas específicas y descarga el CSV resultante.")

archivo = st.file_uploader("Cargar archivo Excel", type=["xlsx"])

if archivo is not None:
    nombre_archivo = archivo.name  # Guardamos el nombre del archivo
    df_filtrado = procesar_archivo(archivo, nombre_archivo)
    st.success("Archivo procesado correctamente.")
    st.dataframe(df_filtrado)  # Muestra la tabla con las columnas seleccionadas
    
    xlsx = generar_xlsx(df_filtrado)
    nombre_salida_xlsx = os.path.splitext(archivo.name)[0] + ".xlsx"
    st.download_button(label="📥 Descargar Excel", data=xlsx, file_name=nombre_salida_xlsx, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    csv = generar_csv(df_filtrado)
    nombre_salida_csv = os.path.splitext(archivo.name)[0] + ".csv"
    st.download_button(label="📥 Descargar CSV", data=csv, file_name=nombre_salida_csv, mime="text/csv")