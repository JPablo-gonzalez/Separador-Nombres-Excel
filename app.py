import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
import re
import copy
import io

# --- DICCIONARIOS MASIVOS DE DESAMBIGUACIÓN ---
CONECTORES = {"DE LA", "DEL", "DE", "SAN", "SANTA", "VON", "VAN", "LOS", "LAS"}
NOMBRES_COMUNES = {
    "JUAN", "CARLOS", "LUIS", "JOSE", "JOSÉ", "MARIA", "MARÍA", "ANDRES", "ANDRÉS", 
    "DAVID", "ALEJANDRO", "ARTURO", "FREDY", "EDWIN", "GERARDO", "WILSON", "EDISON", 
    "JHON", "JORGE", "ALEXANDER", "JULIAN", "DIEGO", "DANIEL", "MIGUEL", "ANGEL", 
    "HERNANDO", "GUILLERMO", "GUSTAVO", "JAIME", "ALBERTO", "HECTOR", "JAIRO", "CESAR", 
    "JULIO", "TITO", "ERNESTO", "NELSON", "CONRADO", "ALIRIO", "LIBANIEL", "EDER", 
    "DIANA", "LUCIA", "ANA", "ANTONIO", "PEDRO", "JESUS", "JESÚS", "MANUEL", "FRANCISCO", 
    "JAVIER", "FERNANDO", "ROSA", "ROBERTO", "MARTHA", "MARTA", "ELENA", "BLANCA", 
    "PATRICIA", "CARMEN", "LAURA", "VICTORIA", "EDUARDO", "RICARDO", "FELIPE", "RAUL", 
    "RAÚL", "PABLO", "GABRIEL", "RAFAEL", "OSCAR", "ÓSCAR", "TERESA", "ALICIA", "SANDRA", 
    "GLORIA", "SOFIA", "SOFÍA", "CAMILA", "VALENTINA", "ISABELLA", "MATEO", "SANTIAGO", 
    "SEBASTIAN", "SEBASTIÁN", "NICOLAS", "NICOLÁS", "SAMUEL", "LEONARDO", "MAURICIO", 
    "LINA", "PAULA", "DANIELA", "VALERIA", "MARCELA", "ANGELA", "ÁNGELA", "CAROLINA", 
    "CLAUDIA", "YURI", "YUDY", "LORENA", "MONICA", "MÓNICA", "TATIANA", "LIZETH", "PAOLA"
}
APELLIDOS_COMUNES = {
    "GOMEZ", "GÓMEZ", "ZAPATA", "PEREZ", "PÉREZ", "OSORIO", "VERA", "BETANCUR", 
    "MORALES", "GALEANO", "ESPINOSA", "GUARIN", "CELIS", "RAMIREZ", "RAMÍREZ", 
    "HERNANDEZ", "HERNÁNDEZ", "TORO", "OCAMPO", "ARROYAVE", "ARANGO", "MUÑOZ", 
    "LONDOÑO", "AGUIRRE", "AMESQUITA", "MARIN", "MARÍN", "BETANCURT", "TOBON", 
    "TOBÓN", "GARCIA", "GARCÍA", "MEJIA", "MEJÍA", "ARANZAZU", "OSPINA", "SANCHEZ", 
    "SÁNCHEZ", "GAVIRIA", "CANO", "RUIZ", "BARRERA", "GALLO", "RAMOS", "GRAJALES", 
    "GRISALES", "BOTERO", "CASTRO", "BARRETO", "ZAMBRANO", "BUITRAGO", "OBANDO", 
    "GALLEGO", "MESA", "ARIAS", "CASTAÑO", "HERRERA", "MOLINA", "CARDONA", "PARRA", 
    "TASCON", "TASCÓN", "DIAZ", "DÍAZ", "GONZALEZ", "GONZÁLEZ", "RODRIGUEZ", 
    "RODRÍGUEZ", "FERNANDEZ", "FERNÁNDEZ", "LOPEZ", "LÓPEZ", "MARTINEZ", "MARTÍNEZ", 
    "ROMERO", "SOSA", "ALVAREZ", "ÁLVAREZ", "TORRES", "FLORES", "ACOSTA", "BENITEZ", 
    "BENÍTEZ", "MEDINA", "SUAREZ", "SUÁREZ", "PEREYRA", "GIMENEZ", "GIMÉNEZ", "ROJAS", 
    "ORTIZ", "SILVA", "NUÑEZ", "NÚÑEZ", "LUNA", "JUAREZ", "JUÁREZ", "CABRERA", "RIOS", 
    "RÍOS", "GODOY", "MORENO", "FERREYRA", "DOMINGUEZ", "DOMÍNGUEZ", "CARRIZO", "PERALTA", 
    "CASTILLO", "LEDESMA", "QUIROGA", "VEGA", "OJEDA", "PONCE", "VILLALBA", "CARDOZO", 
    "NAVARRO", "CORONEL", "VAZQUEZ", "VÁZQUEZ", "VARGAS", "CACERES", "CÁCERES", 
    "FIGUEROA", "CORDOBA", "CÓRDOBA", "CORREA", "ZULUAGA", "RESTREPO", "ALZATE", "QUINTERO"
}

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
        token_medio = tokens[1].upper()
        if es_caso_b:
            p_nom = tokens[0]
            if token_medio in APELLIDOS_COMUNES and token_medio not in NOMBRES_COMUNES:
                p_ape, s_ape = tokens[1], tokens[2]
            elif token_medio in NOMBRES_COMUNES and token_medio not in APELLIDOS_COMUNES:
                s_nom, p_ape = tokens[1], tokens[2]
            else:
                if opciones["tres_palabras_default"] == "2A_1N": p_ape, s_ape = tokens[1], tokens[2]
                else: s_nom, p_ape = tokens[1], tokens[2]
        else:
            p_ape = tokens[0]
            if token_medio in NOMBRES_COMUNES and token_medio not in APELLIDOS_COMUNES:
                p_nom, s_nom = tokens[1], tokens[2]
            elif token_medio in APELLIDOS_COMUNES and token_medio not in NOMBRES_COMUNES:
                s_ape, p_nom = tokens[1], tokens[2]
            else:
                if opciones["tres_palabras_default"] == "2A_1N": s_ape, p_nom = tokens[1], tokens[2]
                else: p_nom, s_nom = tokens[1], tokens[2]
                    
    elif len(tokens) >= 4:
        if es_caso_b:
            p_nom, s_nom, p_ape = tokens[0], tokens[1], tokens[2]
            s_ape = " ".join(tokens[3:])
        else:
            p_ape, s_ape, p_nom = tokens[0], tokens[1], tokens[2]
            s_nom = " ".join(tokens[3:])
            
    return p_ape, s_ape, p_nom, s_nom

def copiar_estilo(origen, destino):
    """Clona absolutamente todos los atributos estéticos de una celda."""
    if origen.has_style:
        destino.font = copy.copy(origen.font)
        destino.border = copy.copy(origen.border)
        destino.fill = copy.copy(origen.fill)
        destino.alignment = copy.copy(origen.alignment)
        destino.number_format = copy.copy(origen.number_format)
        destino.protection = copy.copy(origen.protection)

# --- INTERFAZ WEB STREAMLIT ---
st.set_page_config(page_title="Procesador Impecable de Nombres", layout="wide")
st.title("Separador de Nombres (Sin Pérdida de Datos)")
st.write("Sube tu Excel. El sistema analizará la estructura de tu archivo y te permitirá elegir exactamente qué procesar para no alterar el resto de tu información.")

# 1. Configuración de reglas
st.sidebar.header("Reglas de Separación")
orden_val = st.sidebar.radio("El Excel viene por defecto con:", 
                             ["ApellidosPrimero", "NombresPrimero"],
                             format_func=lambda x: "Apellidos primero" if x == "ApellidosPrimero" else "Nombres primero")

tres_pal_val = st.sidebar.radio("Si un nombre tiene 3 palabras dudosas, asumir:", 
                                ["2A_1N", "1A_2N"],
                                format_func=lambda x: "2 Apellidos, 1 Nombre" if x == "2A_1N" else "1 Apellido, 2 Nombres")

ubicacion_columnas = st.sidebar.radio("¿Dónde colocar el resultado?", 
                                      ["Al lado", "Al final"],
                                      format_func=lambda x: "Insertar al lado de la original (Recomendado)" if x == "Al lado" else "Añadir al final de la tabla (Más seguro)")

opciones = {"orden_default": orden_val, "tres_palabras_default": tres_pal_val}

# 2. Carga interactiva
archivo_subido = st.file_uploader("1. Sube tu archivo Excel (.xlsx)", type=["xlsx"])

if archivo_subido is not None:
    try:
        wb = openpyxl.load_workbook(archivo_subido)
        nombres_hojas = wb.sheetnames
        
        st.subheader("2. Configura tu Archivo")
        col1, col2 = st.columns(2)
        
        with col1:
            hoja_seleccionada = st.selectbox("¿En qué hoja están los datos?", nombres_hojas)
            sheet = wb[hoja_seleccionada]
            
        with col2:
            fila_encabezados = st.number_input("¿En qué fila están los títulos (encabezados)?", min_value=1, max_value=100, value=1)
        
        # Leer encabezados para que el usuario elija
        encabezados = []
        for c in range(1, sheet.max_column + 1):
            val = sheet.cell(row=fila_encabezados, column=c).value
            encabezados.append(f"Columna {get_column_letter(c)}: {val if val else '[Vacía]'}")
            
        columna_seleccionada = st.selectbox("3. ¿Cuál es la columna exacta que contiene los nombres a separar?", encabezados)
        indice_columna = encabezados.index(columna_seleccionada) + 1
        
        if st.button("Procesar y Generar Archivo", type="primary"):
            with st.spinner('Procesando nombres y clonando formatos...'):
                
                # Determinar dónde insertar
                if ubicacion_columnas == "Al lado":
                    col_inicio_nuevas = indice_columna + 1
                    sheet.insert_cols(col_inicio_nuevas, 4)
                else:
                    col_inicio_nuevas = sheet.max_column + 1

                headers_nuevos = ["PRIMER APELLIDO", "SEGUNDO APELLIDO", "PRIMER NOMBRE", "SEGUNDO NOMBRE"]
                
                # Escribir títulos nuevos y clonar estética del encabezado
                for i, h in enumerate(headers_nuevos):
                    col_actual = col_inicio_nuevas + i
                    celda_origen = sheet.cell(row=fila_encabezados, column=indice_columna)
                    celda_nueva = sheet.cell(row=fila_encabezados, column=col_actual)
                    
                    celda_nueva.value = h
                    copiar_estilo(celda_origen, celda_nueva)
                    sheet.column_dimensions[get_column_letter(col_actual)].width = 20
                
                # Procesar filas
                contador = 0
                for r in range(fila_encabezados + 1, sheet.max_row + 1):
                    celda_origen = sheet.cell(row=r, column=indice_columna)
                    val = celda_origen.value
                    
                    if val is None or str(val).strip() == "":
                        continue
                        
                    pa, sa, pn, sn = analizar_nombre(str(val), opciones)
                    
                    for idx, txt in enumerate([pa, sa, pn, sn]):
                        celda_nueva = sheet.cell(row=r, column=col_inicio_nuevas + idx)
                        celda_nueva.value = txt
                        # Clona la estética celda por celda para que las filas de colores se mantengan intactas
                        copiar_estilo(celda_origen, celda_nueva)
                    
                    contador += 1
                
                # Guardar en memoria
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                st.success(f"¡Éxito! Se separaron {contador} registros perfectamente. Tu columna original y el resto de los datos están intactos.")
                
                st.download_button(
                    label="📥 Descargar Excel Impecable",
                    data=output,
                    file_name=f"IMPECABLE_{archivo_subido.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    except Exception as e:
        st.error(f"Error procesando el archivo: {e}. Asegúrate de que el archivo no esté protegido con contraseña.")
