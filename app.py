import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.worksheet.table import TableColumn
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
    "REGINA", "AMPARO", "MANUEL", "HERNAN"
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
    if origen.has_style:
        if origen.font: destino.font = copy(origen.font)
        if origen.border: destino.border = copy(origen.border)
        if origen.fill: destino.fill = copy(origen.fill)
        if origen.number_format: destino.number_format = copy(origen.number_format)
        if origen.protection: destino.protection = copy(origen.protection)
        if origen.alignment: destino.alignment = copy(origen.alignment)


# --- 2. ESTILOS (solo estética) ---
import os

NOMBRE_APP = "Onoma"
LEMA_APP = "Separador inteligente de nombres"
ICONO_APP = os.path.join(os.path.dirname(os.path.abspath(__file__)), "favicon.png")

LOGO_SVG = """
<svg class="logo" viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
  <defs>
    <linearGradient id="g" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0" stop-color="#1d4ed8"/><stop offset="1" stop-color="#10b981"/>
    </linearGradient>
  </defs>
  <rect width="64" height="64" rx="14" fill="url(#g)"/>
  <rect x="13" y="15" width="38" height="9" rx="4.5" fill="#fff"/>
  <rect x="30.5" y="27" width="3" height="4.5" fill="#fff" fill-opacity=".75"/>
  <path d="M27 31h10l-5 5.5z" fill="#fff" fill-opacity=".75"/>
  <rect x="13" y="39" width="7.4" height="10" rx="2.2" fill="#fff"/>
  <rect x="23" y="39" width="7.4" height="10" rx="2.2" fill="#fff"/>
  <rect x="33" y="39" width="7.4" height="10" rx="2.2" fill="#fff"/>
  <rect x="43" y="39" width="7.4" height="10" rx="2.2" fill="#fff"/>
</svg>
"""

# Paletas de color: el resto del CSS solo usa estas variables.
TEMA_CLARO = """
    color-scheme: light;
    --tinta: #0f172a;
    --gris: #64748b;
    --borde: #e2e8f0;
    --fondo-app: #f8fafc;
    --fondo-card: #ffffff;
    --fondo-suave: #f8fafc;
    --acento-texto: #1d4ed8;
    --punteado: #94a3b8;
    --glow-azul: rgba(59,130,246,.10);
    --glow-verde: rgba(16,185,129,.10);
    --sombra: rgba(15,23,42,.25);
"""

TEMA_OSCURO = """
    color-scheme: dark;
    --tinta: #e2e8f0;
    --gris: #94a3b8;
    --borde: #263449;
    --fondo-app: #0b1220;
    --fondo-card: #111a2e;
    --fondo-suave: #0e1627;
    --acento-texto: #7db1ff;
    --punteado: #475569;
    --glow-azul: rgba(59,130,246,.16);
    --glow-verde: rgba(16,185,129,.12);
    --sombra: rgba(0,0,0,.65);
"""

def css_tema(preferencia):
    """Claro / Oscuro fuerzan la paleta; Automático sigue la preferencia del navegador."""
    if preferencia == "claro":
        cuerpo = f":root {{ {TEMA_CLARO} }}"
    elif preferencia == "oscuro":
        cuerpo = f":root {{ {TEMA_OSCURO} }}"
    else:
        cuerpo = (
            f":root {{ {TEMA_CLARO} }}\n"
            f"@media (prefers-color-scheme: dark) {{ :root {{ {TEMA_OSCURO} }} }}"
        )
    return f"<style>\n{cuerpo}\n</style>"

ESTILOS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

:root {
    --verde: #10b981;
    --verde-osc: #059669;
    --azul: #3b82f6;
    --azul-osc: #1d4ed8;
}

html, body, [class*="css"], .stApp {
    font-family: 'Inter', sans-serif;
}

/* Fondo general */
.stApp {
    background:
        radial-gradient(1000px 500px at 10% -10%, var(--glow-azul), transparent 60%),
        radial-gradient(900px 500px at 100% 0%, var(--glow-verde), transparent 60%),
        var(--fondo-app) !important;
    transition: background-color .25s ease;
}
[data-testid="stAppViewContainer"], [data-testid="stMain"] { background: transparent !important; }

/* Ocultar elementos por defecto de Streamlit */
#MainMenu, footer, header[data-testid="stHeader"] { visibility: hidden; height: 0; }
.block-container { max-width: 780px; padding-top: 1.4rem; padding-bottom: 3rem; }

/* Selector de tema */
[data-testid="stButtonGroup"] { justify-content: flex-end; }
[data-testid="stButtonGroup"] button {
    background: var(--fondo-card);
    color: var(--gris);
    border: 1px solid var(--borde);
    font-weight: 600;
    font-size: .82rem;
}
[data-testid="stButtonGroup"] button p { color: inherit; font-size: .82rem; font-weight: 600; }
[data-testid="stButtonGroup"] button:hover { color: var(--tinta); border-color: var(--azul); }
[data-testid="stButtonGroup"] button[data-testid="stBaseButton-segmented_controlActive"] {
    background: linear-gradient(135deg, var(--azul), var(--verde));
    color: #fff;
    border-color: transparent;
}

/* HERO */
.hero {
    background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 55%, #059669 120%);
    border: 1px solid rgba(255,255,255,.06);
    border-radius: 24px;
    padding: 2.3rem 2rem 2.2rem 2rem;
    color: #fff;
    position: relative;
    overflow: hidden;
    box-shadow: 0 20px 40px -18px var(--sombra);
    margin-bottom: 1.4rem;
}
.hero::after {
    content: "";
    position: absolute; right: -60px; top: -60px;
    width: 220px; height: 220px; border-radius: 50%;
    background: rgba(255,255,255,.08);
}
.hero::before {
    content: "";
    position: absolute; right: 60px; bottom: -90px;
    width: 180px; height: 180px; border-radius: 50%;
    background: rgba(16,185,129,.25);
}
.hero .badge {
    display: inline-block;
    font-size: .72rem; font-weight: 600; letter-spacing: .08em;
    text-transform: uppercase;
    background: rgba(255,255,255,.14);
    border: 1px solid rgba(255,255,255,.25);
    padding: .3rem .7rem; border-radius: 999px; margin-bottom: 1rem;
    position: relative; z-index: 1;
}
.hero .marca { display: flex; align-items: center; gap: .85rem; position: relative; z-index: 1; }
.hero .logo { width: 52px; height: 52px; flex: none; filter: drop-shadow(0 8px 14px rgba(0,0,0,.35)); }
.hero h1 {
    font-size: 2.5rem; font-weight: 800; line-height: 1.1; letter-spacing: -.02em;
    margin: 0; padding: 0; color: #fff;
}
.hero .lema {
    font-size: 1.02rem; font-weight: 600; color: rgba(255,255,255,.92);
    margin: .55rem 0 .6rem 0; position: relative; z-index: 1;
}
.hero p {
    font-size: .98rem; color: rgba(255,255,255,.78);
    margin: 0; max-width: 560px; line-height: 1.55; position: relative; z-index: 1;
}

/* Chips de características */
.chips { display: flex; flex-wrap: wrap; gap: .5rem; margin-top: 1.1rem; position: relative; z-index: 1; }
.chip {
    font-size: .78rem; font-weight: 500; color: #fff;
    background: rgba(255,255,255,.12);
    border: 1px solid rgba(255,255,255,.2);
    padding: .3rem .75rem; border-radius: 999px;
}

/* Encabezados de paso */
.paso {
    display: flex; align-items: center; gap: .7rem;
    margin: 1.8rem 0 .7rem 0;
}
.paso .num {
    width: 30px; height: 30px; border-radius: 50%;
    background: linear-gradient(135deg, var(--azul), var(--verde));
    color: #fff; font-weight: 700; font-size: .9rem;
    display: flex; align-items: center; justify-content: center;
    box-shadow: 0 6px 14px -4px rgba(59,130,246,.55);
}
.paso .titulo { font-size: 1.05rem; font-weight: 700; color: var(--tinta); }
.paso .sub { font-size: .85rem; color: var(--gris); font-weight: 400; }

/* Radio como tarjetas */
div[role="radiogroup"] { gap: .6rem; }
div[role="radiogroup"] > label {
    background: var(--fondo-card);
    border: 1.5px solid var(--borde);
    border-radius: 14px;
    padding: .85rem 1rem !important;
    width: 100%;
    transition: all .18s ease;
    box-shadow: 0 1px 2px rgba(15,23,42,.04);
}
div[role="radiogroup"] > label:hover {
    border-color: var(--azul);
    transform: translateY(-1px);
    box-shadow: 0 8px 18px -10px rgba(59,130,246,.45);
}
div[role="radiogroup"] > label:has(input:checked) {
    border-color: var(--azul);
    background: linear-gradient(135deg, rgba(59,130,246,.10), rgba(16,185,129,.10));
    box-shadow: 0 8px 20px -12px rgba(59,130,246,.6);
}
div[role="radiogroup"] label p { font-size: .95rem; font-weight: 500; color: var(--tinta); }

/* Uploader */
[data-testid="stFileUploader"] section {
    background: var(--fondo-card);
    border: 2px dashed var(--punteado);
    border-radius: 18px;
    padding: 1.6rem 1rem;
    transition: all .2s ease;
}
[data-testid="stFileUploader"] section:hover {
    border-color: var(--verde);
    background: rgba(16,185,129,.06);
}
[data-testid="stFileUploader"] section span,
[data-testid="stFileUploader"] section small { color: var(--gris); }
[data-testid="stFileUploader"] button {
    border-radius: 10px;
    background: var(--fondo-card);
    border: 1.5px solid var(--azul);
    color: var(--acento-texto);
    font-weight: 600;
}
[data-testid="stFileUploader"] button:hover {
    background: var(--azul); color: #fff; border-color: var(--azul);
}
[data-testid="stFileUploaderFile"] { color: var(--tinta); }
[data-testid="stFileUploaderFile"] * { color: var(--tinta); }

/* Spinner */
[data-testid="stSpinner"] * { color: var(--tinta); }

/* Botón de descarga */
.stDownloadButton > button {
    width: 100%;
    background: linear-gradient(135deg, var(--azul) 0%, var(--verde) 100%);
    color: #fff; border: none;
    padding: .95rem 1.2rem;
    border-radius: 14px;
    font-weight: 700; font-size: 1.02rem;
    box-shadow: 0 14px 26px -12px rgba(16,185,129,.7);
    transition: all .2s ease;
}
.stDownloadButton > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 18px 30px -12px rgba(59,130,246,.7);
    color: #fff; border: none;
}
.stDownloadButton > button:active { transform: translateY(0); }

/* Tarjeta de resultado */
.resultado {
    background: var(--fondo-card);
    border: 1px solid var(--borde);
    border-left: 5px solid var(--verde);
    border-radius: 16px;
    padding: 1.2rem 1.3rem;
    margin: 1.2rem 0 1rem 0;
    box-shadow: 0 10px 24px -16px var(--sombra);
}
.resultado h3 { margin: 0 0 .25rem 0; font-size: 1.15rem; color: var(--tinta); }
.resultado p { margin: 0; color: var(--gris); font-size: .92rem; }
.stats { display: grid; grid-template-columns: repeat(3, 1fr); gap: .7rem; margin-top: 1rem; }
.stat {
    background: var(--fondo-suave); border: 1px solid var(--borde);
    border-radius: 12px; padding: .7rem .8rem; text-align: center;
}
.stat .valor { font-size: 1.35rem; font-weight: 800; color: var(--acento-texto); }
.stat .etq { font-size: .72rem; color: var(--gris); text-transform: uppercase; letter-spacing: .05em; }

/* Alertas nativas */
[data-testid="stAlert"] {
    border-radius: 14px;
    background: var(--fondo-card);
    border: 1px solid var(--borde);
}
[data-testid="stAlert"] * { color: var(--tinta); }

/* Pie de página */
.pie {
    text-align: center; color: var(--gris); font-size: .8rem;
    margin-top: 2.5rem; padding-top: 1.2rem; border-top: 1px solid var(--borde);
}

@media (max-width: 640px) {
    .hero { padding: 1.7rem 1.3rem; }
    .hero h1 { font-size: 2rem; }
    .hero .logo { width: 44px; height: 44px; }
    .stats { grid-template-columns: 1fr; }
}
</style>
"""

def paso(numero, titulo, subtitulo=""):
    sub = f'<div class="sub">{subtitulo}</div>' if subtitulo else ""
    st.markdown(
        f'<div class="paso"><div class="num">{numero}</div>'
        f'<div><div class="titulo">{titulo}</div>{sub}</div></div>',
        unsafe_allow_html=True,
    )

ETIQUETAS_TEMA = {
    "auto": ":material/brightness_auto: Auto",
    "claro": ":material/light_mode: Claro",
    "oscuro": ":material/dark_mode: Oscuro",
}

@st.fragment
def selector_tema():
    # Se ejecuta como fragmento: cambiar el tema no vuelve a procesar el archivo subido.
    # "Auto" (valor inicial) sigue la preferencia del navegador mediante prefers-color-scheme.
    eleccion = st.segmented_control(
        "Tema",
        options=list(ETIQUETAS_TEMA.keys()),
        format_func=lambda k: ETIQUETAS_TEMA[k],
        default="auto",
        key="tema_preferido",
        label_visibility="collapsed",
    )
    st.markdown(css_tema(eleccion or "auto"), unsafe_allow_html=True)

# --- INTERFAZ WEB STREAMLIT ---
st.set_page_config(page_title=f"{NOMBRE_APP} · Separador de nombres", layout="centered", page_icon=ICONO_APP)
st.markdown(ESTILOS, unsafe_allow_html=True)

_, col_tema = st.columns([2, 3])
with col_tema:
    selector_tema()

st.markdown(
    f"""
    <div class="hero">
        <div class="badge">Herramienta para Excel</div>
        <div class="marca">{LOGO_SVG}<h1>{NOMBRE_APP}</h1></div>
        <div class="lema">{LEMA_APP}</div>
        <p>Divide automáticamente los nombres completos en 4 columnas, manteniendo intactas
        las tablas de Excel, los formatos y los filtros.</p>
        <div class="chips">
            <span class="chip">✔ Conserva tablas y filtros</span>
            <span class="chip">✔ Respeta formatos</span>
            <span class="chip">✔ Detecta conectores (DE LA, DEL…)</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

paso(1, "Tipo de archivo", "Indica cómo vienen los datos en la columna de nombres")
tipo_archivo = st.radio(
    "Formato de la columna de nombres",
    options=["CON_LETRAS", "SIN_LETRAS"],
    format_func=lambda x: "🟢 El archivo contiene letras (A) o (B) al inicio." if x == "CON_LETRAS" else "🔵 El archivo NO contiene letras.",
    label_visibility="collapsed",
)

paso(2, "Sube tu archivo", "Formato Excel (.xlsx)")
archivo_subido = st.file_uploader(
    "Sube tu archivo Excel (.xlsx)", type=["xlsx"], label_visibility="collapsed"
)

if archivo_subido is not None:
    paso(3, "Resultado", "Revisa y descarga tu archivo procesado")
    with st.spinner('Procesando archivo sin errores de estructura...'):
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
                
                # 1. Descombinar celdas temporalmente para evitar corrupción
                merged_ranges = list(sheet.merged_cells.ranges)
                for m_range in merged_ranges:
                    sheet.unmerge_cells(str(m_range))
                
                # 2. Insertar las 4 columnas
                sheet.insert_cols(insert_idx, 4)
                
                # 3. Poner encabezados
                headers_nuevos = ["PRIMER APELLIDO", "SEGUNDO APELLIDO", "PRIMER NOMBRE", "SEGUNDO NOMBRE"]
                for i, h in enumerate(headers_nuevos):
                    col_actual = insert_idx + i
                    celda_origen = sheet.cell(row=fila_header, column=col_nombres)
                    celda_nueva = sheet.cell(row=fila_header, column=col_actual)
                    
                    celda_nueva.value = h
                    copiar_estilo_seguro(celda_origen, celda_nueva)
                    sheet.column_dimensions[get_column_letter(col_actual)].width = 19
                
                # 4. Actualizar dinámicamente las Tablas de Excel para que los filtros y totales no se rompan
                for table in list(sheet.tables.values()):
                    t_min_col, t_min_row, t_max_col, t_max_row = range_boundaries(table.ref)
                    if t_max_col >= insert_idx:
                        nuevo_max_col = t_max_col + 4
                        t_min_col_ajustado = t_min_col if t_min_col < insert_idx else t_min_col + 4
                        
                        table.ref = f"{get_column_letter(t_min_col_ajustado)}{t_min_row}:{get_column_letter(nuevo_max_col)}{t_max_row}"
                        
                        totals_rows = table.totalsRowCount if table.totalsRowCount else 0
                        data_max_row = t_max_row - totals_rows
                        ref_data_str = f"{get_column_letter(t_min_col_ajustado)}{t_min_row}:{get_column_letter(nuevo_max_col)}{data_max_row}"
                        
                        if table.autoFilter:
                            table.autoFilter.ref = ref_data_str
                        if table.sortState:
                            header_rows = table.headerRowCount if table.headerRowCount else 1
                            table.sortState.ref = f"{get_column_letter(t_min_col_ajustado)}{t_min_row + header_rows}:{get_column_letter(nuevo_max_col)}{data_max_row}"
                        
                        # Reconstruir columnas de la tabla para conservar propiedades de totales
                        nuevas_columnas = []
                        nombres_usados = set()
                        for c_idx in range(t_min_col_ajustado, nuevo_max_col + 1):
                            val_enc = str(sheet.cell(row=t_min_row, column=c_idx).value).strip()
                            if not val_enc or val_enc == "None": val_enc = f"Col_{c_idx}"
                            
                            nombre_final = val_enc
                            cnt = 1
                            while nombre_final in nombres_usados:
                                nombre_final = f"{val_enc}_{cnt}"
                                cnt += 1
                            nombres_usados.add(nombre_final)
                            
                            match_col = None
                            for old_c in table.tableColumns:
                                if old_c.name == val_enc:
                                    match_col = old_c
                                    break
                            
                            if match_col:
                                match_col.id = c_idx - t_min_col_ajustado + 1
                                nuevas_columnas.append(match_col)
                            else:
                                nuevas_columnas.append(TableColumn(id=c_idx - t_min_col_ajustado + 1, name=nombre_final))
                                
                        table.tableColumns = nuevas_columnas

                # 5. Recombinar celdas manteniendo la estructura original
                for m_range in merged_ranges:
                    min_col, min_row, max_col, max_row = m_range.bounds
                    if min_col >= insert_idx:
                        min_col += 4
                        max_col += 4
                    elif max_col >= insert_idx:
                        max_col += 4
                    sheet.merge_cells(start_row=min_row, start_column=min_col, end_row=max_row, end_column=max_row)
                
                # 6. Procesar los datos fila por fila
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
                
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                st.markdown(
                    f"""
                    <div class="resultado">
                        <h3>✅ ¡Proceso completado!</h3>
                        <p>Se separaron los nombres sin dañar filtros ni formatos.</p>
                        <div class="stats">
                            <div class="stat"><div class="valor">{contador}</div><div class="etq">Registros</div></div>
                            <div class="stat"><div class="valor">4</div><div class="etq">Columnas nuevas</div></div>
                            <div class="stat"><div class="valor">{"(A)/(B)" if tipo_archivo == "CON_LETRAS" else "Normal"}</div><div class="etq">Modo</div></div>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                st.download_button(
                    label="📥 Descargar Archivo Procesado",
                    data=output,
                    file_name=f"OK_{archivo_subido.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Error procesando el archivo: {e}")

st.markdown(f'<div class="pie">{NOMBRE_APP} · Tus archivos se procesan en memoria y no se almacenan</div>', unsafe_allow_html=True)
