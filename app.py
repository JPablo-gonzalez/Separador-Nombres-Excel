import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
from copy import copy
import re
import io

# --- 1. DICCIONARIOS INTELIGENTES ---
CONECTORES = {"DE LA", "DEL", "DE", "SAN", "SANTA", "VON", "VAN", "LOS", "LAS"}

NOMBRES_COMUNES = {
    "JUAN", "CARLOS", "LUIS", "JOSE", "JOSÉ", "MARIA", "MARÍA", "ANDRES", "ANDRÉS", "DAVID", "ALEJANDRO", 
    "ARTURO", "FREDY", "EDWIN", "GERARDO", "WILSON", "EDISON", "JHON", "JORGE", "ALEXANDER", "JULIAN", 
    "DIEGO", "DANIEL", "MIGUEL", "ANGEL", "HERNANDO", "GUILLERMO", "GUSTAVO", "JAIME", "ALBERTO", "HECTOR", 
    "JAIRO", "CESAR", "JULIO", "TITO", "ERNESTO", "NELSON", "CONRADO", "ALIRIO", "LIBANIEL", "EDER", "DIANA", 
    "LUCIA", "WILLIAM", "ANTONIO", "ARMANDO", "MIRIAM", "GABRIELA", "ALFONSO", "JOSELIN", "JAIRO", "HERIBERTO",
    "REGINA", "AMPARO", "JUAN", "MANUEL", "HERNAN"
}

APELLIDOS_COMUNES = {
    "GOMEZ", "GÓMEZ", "ZAPATA", "PEREZ", "PÉREZ", "OSORIO", "VERA", "BETANCUR", "MORALES", "GALEANO", 
    "ESPINOSA", "GUARIN", "CELIS", "RAMIREZ", "RAMÍREZ", "HERNANDEZ", "HERNÁNDEZ", "TORO", "OCAMPO", 
    "ARROYAVE", "ARANGO", "MUÑOZ", "LONDOÑO", "AGUIRRE", "AMESQUITA", "MARIN", "MARÍN", "BETANCURT", 
    "TOBON", "TOBÓN", "GARCIA", "GARCÍA", "MEJIA", "MEJÍA", "ARANZAZU", "OSPINA", "SANCHEZ", "SÁNCHEZ", 
    "GAVIRIA", "CANO", "RUIZ", "BARRERA", "GALLO", "RAMOS", "GRAJALES", "GRISALES", "BOTERO", "CASTRO", 
    "BARRETO", "ZAMBRANO", "BUITRAGO", "OBANDO", "GALLEGO", "MESA", "ARIAS", "CASTAÑO", "HERRERA", 
    "MOLINA", "CARDONA", "PARRA", "TASCON", "TASCÓN", "DIAZ", "DÍAZ", "LOPEZ", "LÓPEZ", "MACIAS", "MACÍAS"
}

# --- 2. LÓGICA DE SEPARACIÓN CELDA POR CELDA ---
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

def analizar_nombre(texto, tipo_archivo):
    texto = str(texto).strip()
    orientacion = "A->N" 
    
    if tipo_archivo == "CON_LETRAS":
        if re.match(r"^\(B\)", texto, re.IGNORECASE):
            orientacion = "N->A" 
            texto = re.sub(r"^\(B\)\s*", "", texto, flags=re.IGNORECASE).strip()
        elif re.match(r"^\(A\)", texto, re.IGNORECASE):
            orientacion = "A->N" 
            texto = re.sub(r"^\(A\)\s*", "", texto, flags=re.IGNORECASE).strip()
    else:
        orientacion = "A->N"
        
    tokens = agrupar_conectores(texto.split())
    p_ape, s_ape, p_nom, s_nom = "", "", "", ""
    
    if len(tokens) == 1:
        if orientacion == "A->N": p_ape = tokens[0]
        else: p_nom = tokens[0]
        
    elif len(tokens) == 2:
        if orientacion == "A->N": p_ape, p_nom = tokens[0], tokens[1]
        else: p_nom, p_ape = tokens[0], tokens[1]
        
    elif len(tokens) == 3:
        t1, t2, t3 = tokens[0], tokens[1], tokens[2]
        
        if orientacion == "A->N":
            if t2.upper() in NOMBRES_COMUNES or t3.upper() in {"DE JESUS", "DEL CARMEN"}:
                p_ape, p_nom, s_nom = t1, t2, t3
            elif t2.upper() in APELLIDOS_COMUNES:
                p_ape, s_ape, p_nom = t1, t2, t3
            else:
                p_ape, s_ape, p_nom = t1, t2, t3 
        else: 
            if t2.upper() in APELLIDOS_COMUNES:
                p_nom, p_ape, s_ape = t1, t2, t3
            elif t2.upper() in NOMBRES_COMUNES:
                p_nom, s_nom, p_ape = t1, t2, t3
            else:
                p_nom, s_nom, p_ape = t1, t2, t3
                
    elif len(tokens) >= 4:
        if orientacion == "A->N":
            p_ape, s_ape, p_nom = tokens[0], tokens[1], tokens[2]
            s_nom = " ".join(tokens[3:])
        else:
            p_nom, s_nom, p_ape = tokens[0], tokens[1], tokens[2]
            s_ape = " ".join(tokens[3:])
            
    return p_ape, s_ape, p_nom, s_nom

def copiar_estilo_seguro(origen, destino):
    """Copia los estilos sin corromper el XML de Excel"""
    if origen.has_style:
        if origen.font: destino.font = copy(origen.font)
        if origen.border: destino.border = copy(origen.border)
        if origen.fill: destino.fill = copy(origen.fill)
        if origen.number_format: destino.number_format = copy(origen.number_format)
        if origen.protection: destino.protection = copy(origen.protection)
        if origen.alignment: destino.alignment = copy(origen.alignment)

# --- 3. INTERFAZ WEB STREAMLIT ---
st.set_page_config(page_title="Procesador de Nombres UTP", layout="centered", page_icon="📊")

st.title("📊 Separador Inteligente de Nombres")
st.markdown("""
Esta herramienta separa automáticamente los nombres en **4 columnas** solucionando el error de celdas combinadas para garantizar un documento 100% libre de errores.
""")

st.subheader("1. Selecciona el tipo de archivo:")
tipo_archivo = st.radio(
    "¿Cómo vienen los datos en la columna de nombres?", 
    options=["CON_LETRAS", "SIN_LETRAS"],
    format_func=lambda x: "🟢 El archivo contiene letras (A) o (B) al inicio del nombre." if x == "CON_LETRAS" else "🔵 El archivo NO contiene letras."
)

st.subheader("2. Sube tu archivo Excel:")
archivo_subido = st.file_uploader("Arrastra aquí el archivo (.xlsx)", type=["xlsx"])

if archivo_subido is not None:
    with st.spinner('Procesando datos y protegiendo la estética del documento...'):
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
                st.error("❌ No se encontró una columna válida de Nombres.")
            else:
                insert_idx = col_nombres + 1
                
                # PREVENCIÓN DE ERROR DE EXCEL: Descombinar celdas temporalmente
                merged_ranges = list(sheet.merged_cells.ranges)
                for m_range in merged_ranges:
                    sheet.unmerge_cells(str(m_range))
                
                # Insertar columnas
                sheet.insert_cols(insert_idx, 4)
                
                # RE-COMBINAR CELDAS: Ajustar las coordenadas para que no se dañen
                for m_range in merged_ranges:
                    min_col, min_row, max_col, max_row = m_range.bounds
                    
                    if max_col < insert_idx:
                        pass # No cruza la inserción
                    elif min_col >= insert_idx:
                        min_col += 4
                        max_col += 4
                    else:
                        max_col += 4 # Se estira para cubrir el espacio nuevo
                        
                    sheet.merge_cells(start_row=min_row, start_column=min_col, end_row=max_row, end_column=max_col)

                # Colocar encabezados
                headers_nuevos = ["PRIMER APELLIDO", "SEGUNDO APELLIDO", "PRIMER NOMBRE", "SEGUNDO NOMBRE"]
                for i, h in enumerate(headers_nuevos):
                    col_actual = insert_idx + i
                    celda_origen = sheet.cell(row=fila_header, column=col_nombres)
                    celda_nueva = sheet.cell(row=fila_header, column=col_actual)
                    
                    celda_nueva.value = h
                    copiar_estilo_seguro(celda_origen, celda_nueva)
                    sheet.column_dimensions[get_column_letter(col_actual)].width = 19
                
                # Procesar Nombres
                contador = 0
                for r in range(fila_header + 1, sheet.max_row + 1):
                    val = sheet.cell(row=r, column=col_nombres).value
                    if not val: continue
                    
                    pa, sa, pn, sn = analizar_nombre(val, tipo_archivo)
                    
                    for idx, txt in enumerate([pa, sa, pn, sn]):
                        celda = sheet.cell(row=r, column=insert_idx + idx)
                        celda.value = txt
                        copiar_estilo_seguro(sheet.cell(row=r, column=col_nombres), celda)
                    
                    contador += 1
                
                # Guardar el archivo limpio en memoria
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                st.success(f"✅ ¡Proceso impecable! Se evaluaron {contador} estudiantes y se protegió la estructura del archivo.")
                
                st.download_button(
                    label="📥 Descargar Archivo Procesado",
                    data=output,
                    file_name=f"OK_{archivo_subido.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Error procesando el archivo: {e}")
