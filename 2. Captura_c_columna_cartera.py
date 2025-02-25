import streamlit as st
import pandas as pd
from io import BytesIO
import os

def procesar_archivo(file):
    df = pd.read_excel(file)
    columnas_deseadas = ["Identificación", "NUI", "Factura", "Centro de costo", "Saldo Factura", "Mes de Cobro"] 
    df_filtrado = df[columnas_deseadas]
    
    # Eliminar guiones de las columnas nfacturasiigo y nui
    df_filtrado["NUI"] = df_filtrado["NUI"].astype(str).str.replace("-", "", regex=True)
    df_filtrado["Factura"] = df_filtrado["Factura"].astype(str).str.replace("-", "", regex=True)
    
    
    # Convertir las columnas address y localidad a mayúsculas
    df_filtrado["Centro de costo"] = df_filtrado["Centro de costo"].astype(str).str.upper()

    # Reemplazar valores vacíos o nulos con "NA" (excepto en "factura")
    df_filtrado.fillna("NA", inplace=True)

    # Eliminar filas donde "factura" esté vacía
    df_filtrado = df_filtrado[df_filtrado["factura"].notna() & (df_filtrado["factura"] != "NA") & (df_filtrado["factura"].str.strip() != "")]

    # Convertir "Mes de Cobro" en mes (numérico) y año
    if "Mes de Cobro" in df_filtrado.columns:
        df_filtrado["Mes de Cobro"] = df_filtrado["Mes de Cobro"].astype(str)  # Asegurar que es texto
        df_filtrado[["mes", "año"]] = df_filtrado["Mes de Cobro"].str.split(" ", expand=True)

        # Diccionario de meses para conversión a número
        meses_dict = {
            "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
            "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12
        }
        
        # Convertir nombre del mes a número
        df_filtrado["mes"] = df_filtrado["mes"].str.lower().map(meses_dict)
        
        # Convertir año a número
        df_filtrado["año"] = pd.to_numeric(df_filtrado["año"], errors='coerce')

    # Agregar una nueva columna al inicio con el nombre del archivo
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
    df_filtrado = procesar_archivo(archivo)
    st.success("Archivo procesado correctamente.")
    st.dataframe(df_filtrado)  # Muestra la tabla con las columnas seleccionadas
    
    xlsx = generar_xlsx(df_filtrado)
    nombre_salida = os.path.splitext(archivo.name)[0] + ".xlsx"
    st.download_button(label="📥 Descargar Excel", data=xlsx, file_name=nombre_salida, mime="text/xlsx")

    csv = generar_csv(df_filtrado)
    nombre_salida = os.path.splitext(archivo.name)[0] + ".csv"
    st.download_button(label="📥 Descargar CSV", data=csv, file_name=nombre_salida, mime="text/csv")
