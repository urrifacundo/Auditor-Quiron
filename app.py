import streamlit as st
import pandas as pd
from auditor import ejecutar_auditoria
import os

st.title("Sistema de Auditoría Operativa - Guía Quirón")
st.write("Subí el archivo Excel de denuncias de los operadores para generar el reporte de control automatizado.")

archivo_subido = st.file_uploader("Elegir archivo Excel", type=["xlsx", "xls"])

if archivo_subido is not None:
    # Guardamos temporalmente el archivo subido
    with open("temp_entrada.xlsx", "wb") as f:
        f.write(archivo_subido.getbuffer())
        
    st.info("Procesando y auditando registros...")
    
    # Ejecutamos la auditoría
    ejecutar_auditoria("temp_entrada.xlsx", "temp_salida.xlsx")
    
    st.success("¡Auditoría finalizada con éxito!")
    
    with open("temp_salida.xlsx", "rb") as f:
        st.download_button(
            label="Descargar Reporte Visual Auditado",
            data=f,
            file_name="reporte_auditoria_visual.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheet.sheet"
        )
