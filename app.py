import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
import re
import copy
import io

# --- DICCIONARIOS INTELIGENTES ---
CONECTORES = {"DE LA", "DEL", "DE", "SAN", "SANTA", "VON", "VAN", "LOS", "LAS"}
NOMBRES_COMUNES = {"JUAN", "CARLOS", "LUIS", "JOSE", "JOSÉ", "MARIA", "MARÍA", "ANDRES", "ANDRÉS", "DAVID", "ALEJANDRO", "ARTURO", "FREDY", "EDWIN", "GERARDO", "WILSON", "EDISON", "JHON", "JORGE", "ALEXANDER", "JULIAN", "DIEGO", "DANIEL", "MIGUEL", "ANGEL", "HERNANDO", "GUILLERMO", "GUSTAVO", "JAIME", "ALBERTO", "HECTOR", "JAIRO", "CESAR", "JULIO", "TITO", "ERNESTO", "NELSON", "CONRADO", "ALIRIO", "LIBANIEL", "EDER", "DIANA", "LUCIA"}
APELLIDOS_COMUNES = {"GOMEZ", "GÓMEZ", "ZAPATA", "PEREZ", "PÉREZ", "OSORIO", "VERA", "BETANCUR", "MORALES", "GALEANO", "ESPINOSA", "GUARIN", "CELIS", "RAMIREZ", "RAMÍREZ", "HERNANDEZ", "HERNÁNDEZ", "TORO", "OCAMPO", "ARROYAVE", "ARANGO", "MUÑOZ", "LONDOÑO", "AGUIRRE", "AMESQUITA", "MARIN", "MARÍN", "BETANCURT", "TOBON", "TOBÓN", "GARCIA", "GARCÍA", "MEJIA", "MEJÍA", "ARANZAZU", "OSPINA", "SANCHEZ", "SÁNCHEZ", "GAVIRIA", "CANO", "RUIZ", "BARRERA", "GALLO", "RAMOS", "GRAJALES", "GRISALES", "BOTERO", "CASTRO", "BARRETO", "ZAMBRANO", "BUITRAGO", "OBANDO", "GALLEGO", "MESA", "ARIAS", "CASTAÑO", "HERRERA", "MOLINA", "CARDONA", "PARRA", "TASCON", "TASCÓN", "DIAZ", "DÍAZ"}

def agrupar_conectores(tokens):
    resultado = []
    i = 0
    while i < len(tokens):
        palabra = tokens[i].upper()
        if i + 2 < len(tokens) and f"{palabra} {tokens[i+1].upper()}" in {"DE LA", "DE LOS", "DE LAS"}:
            resultado.append(f"{tokens[i]} {tokens[i+1]} {tokens[i+2]}")
            i += 3
        elif i + 1 < len(tokens) and palabra in CONECTORES:
            resultado.append(f"{tokens[i]} {tokens[i+1]}")
            i += 2
        elif i + 1 < len(tokens) and f"{palabra} {tokens[i+1].upper()}" in {"DE JESUS", "DEL CARMEN"}:
            resultado.append(f"{tokens[i]} {tokens[i+1]}")
            i += 2
        else:
            resultado.append(tokens[i])
            i += 1
    return resultado

def analizar_nombre(texto, opciones):
    texto = str(texto).strip()
    es_caso_b = False
    
    if re.match(r"^\(B\)", texto, re.IGNORECASE):
        es_caso_b = True
        texto = re.sub(r"^\(B\)\s*", "", texto, flags=re.IGNORECASE).strip()
    elif re.match(r"^\(A\)", texto, re.IGNORECASE):
        es_caso_b = False
        texto = re.sub(r"^\(A\)\s*", "", texto, flags=re.IGNORECASE).strip()
    else:
        es_caso_b = (opciones["orden_default"] == "NombresPrimero")
        
    tokens = agrupar_conectores(texto.split())
    p_ape, s_ape, p_nom, s_nom = "", "", "", ""
    
    if len(tokens) == 1:
        if es_caso_b: p_nom = tokens[0]
        else: p_ape = tokens[0]
        
    elif len(tokens) == 2:
        if es_caso_b: p_nom, p_ape = tokens[0], tokens[1]
        else: p_ape, p_nom = tokens[0], tokens[1]
        
    elif len(tokens) == 3:
        if es_caso_b:
            p_nom = tokens[0]
            if tokens[1].upper() in APELLIDOS_COMUNES and tokens[2].upper() in APELLIDOS_COMUNES:
                p_ape, s_ape = tokens[1], tokens[2]
            else:
                s_nom, p_ape = tokens[1], tokens[2]
        else:
            p_ape = tokens[0]
            if tokens[1].upper() in NOMBRES_COMUNES or tokens[2].upper() in {"DE JESUS", "DEL CARMEN"}:
                p_nom, s_nom = tokens[1], tokens[2]
            elif tokens[1].upper() in APELLIDOS_COMUNES:
                s_ape, p_nom = tokens[1], tokens[2]
            else:
                if opciones["tres_palabras_default"] == "2A_1N":
                    s_ape, p_nom = tokens[1], tokens[2]
                else:
                    p_nom, s_nom = tokens[1], tokens[2]
                    
    elif len(tokens) >= 4:
        if es_caso_b:
            p_nom, s_nom, p_ape = tokens[0], tokens[1], tokens[2]
            s_ape = " ".join(tokens[3:])
        else:
            p_ape, s_ape, p_nom = tokens[0], tokens[1], tokens[2]
            s_nom = " ".join(tokens[3:])
            
    return p_ape, s_ape, p_nom, s_nom

def copiar_estilo(origen, destino):
    if origen.has_style:
        destino.font = copy.copy(origen.font)
        destino.border = copy.copy(origen.border)
        destino.fill = copy.copy(origen.fill)
        destino.alignment = copy.copy(origen.alignment)

# --- INTERFAZ WEB STREAMLIT ---
st.set_page_config(page_title="Procesador de Nombres", layout="centered")
st.title("Separador Inteligente de Nombres")
st.write("Sube un archivo de Excel para dividir la columna de nombres en Primer/Segundo Apellido y Nombre, conservando el formato y los colores originales del documento.")

st.header("1. Configuración")
col1, col2 = st.columns(2)
with col1:
    orden_val = st.radio("Orden por defecto (sin A ni B):", 
                         options=["ApellidosPrimero", "NombresPrimero"],
                         format_func=lambda x: "Apellidos primero" if x == "ApellidosPrimero" else "Nombres primero")
with col2:
    tres_pal_val = st.radio("Manejo de 3 palabras desconocidas:", 
                            options=["2A_1N", "1A_2N"],
                            format_func=lambda x: "2 Apellidos, 1 Nombre" if x == "2A_1N" else "1 Apellido, 2 Nombres")

opciones = {"orden_default": orden_val, "tres_palabras_default": tres_pal_val}

st.header("2. Subir Archivo")
archivo_subido = st.file_uploader("Selecciona el archivo Excel (.xlsx)", type=["xlsx"])

if archivo_subido is not None:
    with st.spinner('Procesando el archivo sin perder la estética...'):
        try:
            wb = openpyxl.load_workbook(archivo_subido)
            sheet = wb.active
            
            fila_header = -1
            col_nombres = -1
            
            for r in range(1, min(30, sheet.max_row + 1)):
                for c in range(1, sheet.max_column + 1):
                    val = str(sheet.cell(row=r, column=c).value).upper()
                    if ("NOMBRE" in val or "APELLIDO" in val) and val not in ["PRIMER NOMBRE", "SEGUNDO NOMBRE", "PRIMER APELLIDO", "SEGUNDO APELLIDO"]:
                        fila_header = r
                        col_nombres = c
                        break
                if fila_header != -1: break
            
            if fila_header == -1:
                st.error("No se encontró una columna válida de Nombres.")
            else:
                sheet.insert_cols(col_nombres + 1, 4)
                headers_nuevos = ["PRIMER APELLIDO", "SEGUNDO APELLIDO", "PRIMER NOMBRE", "SEGUNDO NOMBRE"]
                
                for i, h in enumerate(headers_nuevos):
                    col_actual = col_nombres + 1 + i
                    celda_origen = sheet.cell(row=fila_header, column=col_nombres)
                    celda_nueva = sheet.cell(row=fila_header, column=col_actual)
                    
                    celda_nueva.value = h
                    copiar_estilo(celda_origen, celda_nueva)
                    sheet.column_dimensions[get_column_letter(col_actual)].width = 19
                
                contador = 0
                for r in range(fila_header + 1, sheet.max_row + 1):
                    val = sheet.cell(row=r, column=col_nombres).value
                    if not val: continue
                    
                    pa, sa, pn, sn = analizar_nombre(val, opciones)
                    
                    for idx, txt in enumerate([pa, sa, pn, sn]):
                        celda = sheet.cell(row=r, column=col_nombres + 1 + idx)
                        celda.value = txt
                        copiar_estilo(sheet.cell(row=r, column=col_nombres), celda)
                    
                    contador += 1
                
                # Guardar el archivo en la memoria del navegador
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                st.success(f"¡Proceso completado! Se formatearon {contador} registros.")
                
                st.download_button(
                    label="Descargar Archivo Procesado",
                    data=output,
                    file_name=f"PROCESADO_{archivo_subido.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Error procesando el archivo: {e}")
