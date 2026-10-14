"""Estética e interfaz visual: textos, colores, CSS y componentes visuales de Streamlit."""

import os

import streamlit as st

# =============================================================================
# 1. ESTILOS (solo estética)
# =============================================================================
# Sistema visual "AuthKit · Frosted glass cathedral at midnight": paleta monocromática fría, superficies
# de vidrio casi invisibles, bordes como trazos translúcidos de 1px (inset) y elevación con brillo interior
# + halo oscuro, nunca sombras convencionales. El único color vivo es Void Violet, solo en el botón
# principal (CTA). Radios: tarjetas 16px, insignias/entradas 6px, botones 999px, iconos 9999px.

NOMBRE_APP = "Desglosa"
LEMA_APP = "Separador inteligente de nombres"
_ICONO_ARCHIVO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "favicon.png")
ICONO_APP = _ICONO_ARCHIVO if os.path.exists(_ICONO_ARCHIVO) else "📊"

# Monocromo: los colores salen de las variables del tema (.fondo y .trazo en el CSS).
LOGO_SVG = """
<svg class="logo" viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
  <circle class="fondo" cx="32" cy="32" r="32"/>
  <rect class="trazo" x="13" y="15" width="38" height="9" rx="4.5"/>
  <rect class="trazo tenue" x="30.5" y="27" width="3" height="4.5"/>
  <path class="trazo tenue" d="M27 31h10l-5 5.5z"/>
  <rect class="trazo" x="13" y="39" width="7.4" height="10" rx="2.2"/>
  <rect class="trazo" x="23" y="39" width="7.4" height="10" rx="2.2"/>
  <rect class="trazo" x="33" y="39" width="7.4" height="10" rx="2.2"/>
  <rect class="trazo" x="43" y="39" width="7.4" height="10" rx="2.2"/>
</svg>
"""

# Paletas de color: el resto del CSS solo usa estas variables.
# Oscuro (AuthKit) es el tema por defecto; Claro se activa solo con el botón y usa los mismos tokens invertidos.
TEMA_OSCURO = """
    color-scheme: dark;
    --lienzo: #05060f;                 /* Midnight Canvas */
    --texto: #d1e4fa;                  /* Frost Glow: texto principal, iconos */
    --cuerpo: #c7d3ea;                 /* Moon Mist: texto de cuerpo, etiquetas */
    --tenue: #9da7ba;                  /* Fog Veil: texto secundario */
    --blanco: #ffffff;                 /* Pure White: texto de botones y entradas */
    --titulo: linear-gradient(0deg, #d8ecf8 0%, #98c0ef 100%);   /* Ice Highlight: solo la marca */
    --vidrio: rgba(186,214,247,.03);   /* tarjetas de vidrio */
    --vidrio-hover: rgba(186,214,247,.06);
    --fantasma: rgba(186,214,247,.06); /* botón pill fantasma */
    --relleno: #c7d3ea1f;              /* Luminous Fill: insignias, contenedores de icono */
    --borde: #bad7f71f;                /* Glass Edge: trazo inset de 1px */
    --borde-fuerte: rgba(186,215,247,.24);
    --brillo: rgba(216,236,248,.2);    /* brillo interior (inset) */
    --halo: rgba(0,0,0,.3);            /* halo oscuro exterior */
"""

TEMA_CLARO = """
    color-scheme: light;
    --lienzo: #ffffff;
    --texto: #05060f;
    --cuerpo: #2f343e;
    --tenue: #3f4959;
    --blanco: #05060f;
    --titulo: linear-gradient(0deg, #05060f 0%, #2f343e 100%);
    --vidrio: rgba(47,52,62,.03);
    --vidrio-hover: rgba(47,52,62,.06);
    --fantasma: rgba(47,52,62,.06);
    --relleno: rgba(47,52,62,.08);
    --borde: rgba(47,52,62,.14);
    --borde-fuerte: rgba(47,52,62,.28);
    --brillo: rgba(255,255,255,.6);
    --halo: rgba(5,6,15,.08);
"""

CSS_TEMA_OSCURO = f"<style>:root {{ {TEMA_OSCURO} }}</style>"
CSS_TEMA_CLARO = f"<style>:root {{ {TEMA_CLARO} }}</style>"

ESTILOS = """
<style>
/* Inter, Space Grotesk y JetBrains Mono se registran en .streamlit/config.toml ([[theme.fontFaces]]) y se
   sirven desde la propia app (static/fonts). No se usa Google Fonts: enviaba la IP del usuario a un tercero. */

:root {
    --violeta: #663af3;                /* Void Violet: el ÚNICO acento, solo para el CTA principal */
    --f-display: 'Space Grotesk', 'Inter', system-ui, sans-serif;
    --f-cuerpo: 'Inter', system-ui, -apple-system, 'Segoe UI', Roboto, sans-serif;
    --f-mono: 'JetBrains Mono', ui-monospace, 'SFMono-Regular', Menlo, monospace;
    /* Elevación: trazo inset 1px + brillo interior + halo oscuro (sin sombras convencionales) */
    --elevacion: inset 0 0 0 1px var(--borde), inset 0 1px 1px var(--brillo), 0 16px 32px var(--halo);
    --trazo: inset 0 0 0 1px var(--borde);
}

html, body, [class*="css"], .stApp, .stApp button, .stApp input, .stApp textarea {
    font-family: var(--f-cuerpo);
}
.stApp { color: var(--cuerpo); background: var(--lienzo) !important; transition: background-color .25s ease; }
[data-testid="stAppViewContainer"], [data-testid="stMain"] { background: transparent !important; }
.stApp p, .stApp li, .stApp label { color: var(--cuerpo); }

/* Ocultar elementos por defecto de Streamlit */
#MainMenu, footer, header[data-testid="stHeader"] { visibility: hidden; height: 0; }
.block-container { max-width: 780px; padding-top: 24px; padding-bottom: 48px; }

/* Botón discreto de tema (esquina inferior izquierda): contenedor de icono circular */
.st-key-selector_tema {
    position: fixed; left: 16px; bottom: 16px; z-index: 1000;
    width: auto !important;
}
.st-key-selector_tema button {
    width: 40px; height: 40px; min-height: 0; padding: 0;
    border: none; border-radius: 9999px;
    background: var(--fantasma);
    box-shadow: var(--trazo);
    color: var(--tenue);
    opacity: .6;
    transition: opacity .2s ease, box-shadow .2s ease, color .2s ease;
}
.st-key-selector_tema button p { color: inherit; font-size: 20px; line-height: 1; }
.st-key-selector_tema button:hover,
.st-key-selector_tema button:focus:not(:active) {
    opacity: 1; color: var(--texto); border: none;
    box-shadow: inset 0 0 0 1px var(--borde-fuerte), inset 0 1px 1px var(--brillo);
}

/* HERO: tarjeta de vidrio */
.hero {
    background: var(--vidrio);
    border-radius: 16px;
    padding: 40px 24px 32px 24px;
    box-shadow: var(--elevacion);
    margin-bottom: 24px;
}
.hero .badge {
    display: inline-block;
    font-family: var(--f-mono); font-size: 15px; font-weight: 400;
    letter-spacing: .10em; text-transform: uppercase;
    color: var(--tenue);
    margin-bottom: 16px;
}
.hero .marca { display: flex; align-items: center; gap: 16px; }
.hero .logo { width: 48px; height: 48px; flex: none; border-radius: 9999px; box-shadow: var(--trazo); }
.logo .fondo { fill: var(--relleno); }
.logo .trazo { fill: var(--texto); }
.logo .trazo.tenue { fill-opacity: .6; }
.hero h1 {
    font-family: var(--f-display); font-size: 48px; font-weight: 500;
    line-height: 1.1; letter-spacing: -.02em;
    margin: 0; padding: 0;
    background: var(--titulo);
    -webkit-background-clip: text; background-clip: text;
    -webkit-text-fill-color: transparent; color: transparent;
}
.hero .lema {
    font-size: 20px; font-weight: 500; color: var(--texto);
    margin: 16px 0 8px 0;
}
.hero p {
    font-size: 16px; color: var(--tenue);
    margin: 0; max-width: 560px; line-height: 1.6;
}

/* Insignias de características */
.chips { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 24px; }
.chip {
    font-size: 13px; font-weight: 500; color: var(--texto);
    background: var(--relleno);
    box-shadow: var(--trazo);
    padding: 4px 12px; border-radius: 6px;
}

/* Encabezados de paso */
.paso {
    display: flex; align-items: center; gap: 16px;
    margin: 48px 0 16px 0;
}
.paso .num {
    width: 40px; height: 40px; flex: none; border-radius: 9999px;
    background: var(--relleno);
    box-shadow: var(--trazo), inset 0 1px 1px var(--brillo);
    color: var(--texto);
    font-family: var(--f-mono); font-weight: 400; font-size: 15px;
    display: flex; align-items: center; justify-content: center;
}
.paso .titulo { font-family: var(--f-display); font-size: 28px; font-weight: 500; line-height: 1.2; color: var(--texto); }
.paso .sub { font-size: 14px; color: var(--tenue); font-weight: 400; margin-top: 4px; }

/* Zona de carga de archivo: tarjeta de vidrio con trazo translúcido (sin bordes punteados sólidos) */
[data-testid="stFileUploader"] section {
    display: flex; flex-direction: column; align-items: center; justify-content: center;
    gap: 16px;
    text-align: center;
    padding: 40px 24px;
    background: var(--vidrio);
    border: none;
    border-radius: 16px;
    box-shadow: var(--elevacion);
    transition: background-color .2s ease, box-shadow .2s ease;
}
[data-testid="stFileUploader"] section:hover {
    background: var(--vidrio-hover);
    box-shadow: inset 0 0 0 1px var(--borde-fuerte), inset 0 1px 1px var(--brillo), 0 16px 32px var(--halo);
}
/* icono */
[data-testid="stFileUploaderDropzoneInstructions"] {
    display: flex; flex-direction: column; align-items: center; gap: 16px;
}
[data-testid="stFileUploaderDropzoneInstructions"] svg,
[data-testid="stFileUploaderDropzoneInstructions"] [data-testid="stIconMaterial"] {
    box-sizing: content-box;
    width: 28px; height: 28px; font-size: 28px; line-height: 28px;
    padding: 16px;
    color: var(--texto); fill: var(--texto);
    overflow: hidden; white-space: nowrap;
    border-radius: 9999px;
    background: var(--relleno);
    box-shadow: var(--trazo), inset 0 1px 1px var(--brillo);
}
/* textos en español (reemplazan los de Streamlit) */
[data-testid="stFileUploaderDropzoneInstructions"] > div > * { display: none; }
[data-testid="stFileUploaderDropzoneInstructions"] > div::before {
    content: "Arrastra tu archivo Excel aquí";
    display: block;
    font-size: 18px; font-weight: 500; color: var(--texto);
}
[data-testid="stFileUploaderDropzoneInstructions"] > div::after {
    content: "o selecciónalo desde tu equipo · solo .xlsx";
    display: block; margin-top: 4px;
    font-size: 14px; font-weight: 400; color: var(--tenue);
}
/* botón: CTA violeta mientras no hay archivo. Solo el botón de subida (hijo directo de la zona), no los de
   las fichas de archivo de Streamlit ≥1.5x. En versiones donde el botón sigue visible con un archivo
   cargado, pasa a pill fantasma (el CTA pasa a ser "Descargar"). */
[data-testid="stFileUploader"] section > span > button,
[data-testid="stFileUploader"] section > button {
    font-size: 0 !important;
    padding: 12px 24px;
    border: none;
    border-radius: 999px;
    background: var(--violeta);
    box-shadow: inset 0 1px 1px rgba(216,236,248,.2), 0 16px 32px rgba(0,0,0,.3);
    transition: filter .2s ease, transform .2s ease;
}
[data-testid="stFileUploader"] section > span > button::after,
[data-testid="stFileUploader"] section > button::after {
    content: "Seleccionar archivo";
    font-size: 15px; font-weight: 500; color: #ffffff;
}
[data-testid="stFileUploader"] section > span > button *,
[data-testid="stFileUploader"] section > button * { display: none; }
[data-testid="stFileUploader"] section > span > button:hover,
[data-testid="stFileUploader"] section > button:hover { filter: brightness(1.1); transform: translateY(-1px); }
[data-testid="stFileUploader"] section > span > button:active,
[data-testid="stFileUploader"] section > button:active { transform: translateY(0); }
[data-testid="stFileUploader"]:has([data-testid="stFileUploaderFile"]) section > span > button,
[data-testid="stFileUploader"]:has([data-testid="stFileUploaderFile"]) section > button {
    background: var(--fantasma);
    box-shadow: var(--trazo);
}
[data-testid="stFileUploader"]:has([data-testid="stFileUploaderFile"]) section > span > button::after,
[data-testid="stFileUploader"]:has([data-testid="stFileUploaderFile"]) section > button::after { color: var(--blanco); }
/* ficha del archivo cargado (Streamlit ≥1.5x): insignia de vidrio + botones de icono circulares */
[data-testid="stFileUploader"]:has([data-testid="stFileChips"]) section { padding: 24px; }
[data-testid="stFileChips"] { gap: 16px; align-items: center; }
[data-testid="stFileChip"] {
    background: var(--vidrio) !important;
    border: none !important; border-radius: 6px;
    box-shadow: var(--trazo);
    padding: 8px 12px;
}
[data-testid="stFileChip"] * { color: var(--texto); }
[data-testid="stFileChipName"] { font-weight: 500; }
[data-testid="stFileChip"] > div:first-child { background: var(--relleno) !important; border-radius: 6px; }
[data-testid="stFileChip"] > div:first-child svg { fill: var(--texto); color: var(--texto); }
[data-testid="stFileChip"] small, [data-testid="stFileChip"] > div:nth-child(2) > div:last-child { color: var(--tenue); }
[data-testid="stFileChipDeleteBtn"] button,
[data-testid="stFileChips"] [data-testid="stBaseButton-borderlessIcon"] {
    border: none; border-radius: 9999px;
    background: var(--fantasma);
    box-shadow: var(--trazo);
    color: var(--texto);
}
[data-testid="stFileChips"] [data-testid="stBaseButton-borderlessIcon"] {
    width: 40px !important; height: 40px !important; min-width: 40px; min-height: 40px; padding: 0; flex: none;
}
[data-testid="stFileChipDeleteBtn"] button svg,
[data-testid="stFileChips"] [data-testid="stBaseButton-borderlessIcon"] * { color: var(--texto); fill: var(--texto); }
[data-testid="stFileChipDeleteBtn"] button:hover,
[data-testid="stFileChips"] [data-testid="stBaseButton-borderlessIcon"]:hover {
    background: var(--vidrio-hover);
    box-shadow: inset 0 0 0 1px var(--borde-fuerte), inset 0 1px 1px var(--brillo);
}
/* archivo ya cargado */
[data-testid="stFileUploaderFile"] {
    margin-top: 8px;
    padding: 12px 16px;
    background: var(--vidrio);
    border: none;
    box-shadow: var(--trazo);
    border-radius: 16px;
    color: var(--texto);
}
[data-testid="stFileUploaderFile"] * { color: var(--texto); }
[data-testid="stFileUploaderFile"] small { color: var(--tenue); }
[data-testid="stFileUploaderFile"] svg { fill: var(--tenue); color: var(--tenue); }
[data-testid="stFileUploaderDeleteBtn"] button { border-radius: 9999px; }
[data-testid="stFileUploaderDeleteBtn"] button:hover { background: var(--fantasma); }

/* Spinner */
[data-testid="stSpinner"] * { color: var(--cuerpo); }

/* Botón de descarga: CTA principal (el único violeta) */
.stElementContainer:has(> .stDownloadButton), .stDownloadButton { width: 100% !important; }
.stDownloadButton > button {
    width: 100%;
    background: var(--violeta);
    color: #ffffff; border: none;
    padding: 12px 24px;
    border-radius: 999px;
    font-weight: 500; font-size: 16px;
    box-shadow: inset 0 1px 1px rgba(216,236,248,.2), 0 16px 32px rgba(0,0,0,.3);
    transition: filter .2s ease, transform .2s ease;
}
.stDownloadButton > button p { color: #ffffff; font-size: 16px; font-weight: 500; }
.stDownloadButton > button:hover,
.stDownloadButton > button:focus:not(:active) {
    filter: brightness(1.1); transform: translateY(-1px);
    color: #ffffff; border: none; background: var(--violeta);
}
.stDownloadButton > button:active { transform: translateY(0); color: #ffffff; background: var(--violeta); }

/* Tarjeta de resultado */
.resultado {
    background: var(--vidrio);
    border-radius: 16px;
    padding: 24px;
    margin: 24px 0 16px 0;
    box-shadow: var(--elevacion);
}
.resultado h3 {
    margin: 0 0 8px 0; padding: 0;
    font-family: var(--f-cuerpo); font-size: 20px; font-weight: 600; color: var(--texto);
}
.resultado p { margin: 0; color: var(--tenue); font-size: 14px; }
.stats { display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; margin-top: 24px; }
.stat {
    background: var(--vidrio); box-shadow: var(--trazo), inset 0 1px 1px var(--brillo);
    border-radius: 16px; padding: 16px; text-align: center;
}
.stat .valor { font-family: var(--f-display); font-size: 28px; font-weight: 500; color: var(--texto); line-height: 1.2; }
.stat .etq {
    margin-top: 4px;
    font-family: var(--f-mono); font-size: 15px; font-weight: 400;
    color: var(--tenue); text-transform: uppercase; letter-spacing: .10em;
}

/* Consulta de edades */
.edades {
    background: var(--vidrio);
    border-radius: 16px;
    padding: 24px;
    margin: 8px 0 16px 0;
    box-shadow: var(--elevacion);
}
.edades .etq {
    font-family: var(--f-mono); font-size: 15px; font-weight: 400;
    letter-spacing: .10em; text-transform: uppercase; color: var(--tenue);
}
.edades .cifra { font-family: var(--f-display); font-size: 44px; font-weight: 400; color: var(--texto); line-height: 1.25; margin-top: 8px; }
.edades .cifra span { font-family: var(--f-cuerpo); font-size: 18px; font-weight: 500; color: var(--tenue); }
.edades .nota { font-size: 13px; color: var(--tenue); margin-top: 8px; }
.st-key-edad_limite { max-width: 340px; }
.st-key-edad_limite label p { font-size: 14px; font-weight: 500; color: var(--cuerpo); }
.st-key-edad_limite [data-testid="stNumberInputContainer"] {
    border: none !important; border-radius: 6px; overflow: hidden;
    box-shadow: var(--trazo);
}
.st-key-edad_limite [data-testid="stNumberInputContainer"]:focus-within { box-shadow: inset 0 0 0 1px var(--borde-fuerte); }
.st-key-edad_limite [data-testid="stNumberInputContainer"] { background: var(--vidrio) !important; }
.st-key-edad_limite input { background: transparent !important; color: var(--blanco) !important; font-size: 16px; }
.st-key-edad_limite button { background: transparent !important; color: var(--tenue) !important; border: none !important; }
.st-key-edad_limite button:hover { background: var(--fantasma) !important; color: var(--texto) !important; }

/* Desplegables (vista previa, filas por revisar) */
[data-testid="stExpander"] details {
    background: var(--vidrio);
    border: none; border-radius: 16px;
    box-shadow: var(--trazo);
}
[data-testid="stExpander"] summary { border-radius: 16px; padding: 16px 24px; background: transparent !important; }
[data-testid="stExpander"] summary p { font-size: 15px; font-weight: 500; color: var(--texto); }
[data-testid="stExpander"] summary svg { color: var(--tenue); fill: var(--tenue); }
[data-testid="stExpander"] summary:hover { background: var(--vidrio-hover) !important; }
[data-testid="stExpanderDetails"] { padding: 0 24px 24px 24px; border: none; box-shadow: inset 0 1px 0 var(--borde); padding-top: 16px; }
[data-testid="stDataFrame"] { border-radius: 6px; box-shadow: var(--trazo); overflow: hidden; }

/* Alertas nativas: vidrio monocromático, sin fondos de color */
[data-testid="stAlert"] > div,
[data-testid="stAlertContainer"] {
    border-radius: 16px;
    background: var(--vidrio) !important;
    border: none;
    box-shadow: var(--trazo), inset 0 1px 1px var(--brillo);
}
[data-testid="stAlert"] * { color: var(--cuerpo); }
[data-testid="stAlert"] svg { fill: var(--texto); color: var(--texto); }

/* Pie de página: separador como trazo translúcido */
.pie {
    text-align: center; color: var(--tenue); font-size: 13px;
    margin-top: 48px; padding-top: 24px; box-shadow: inset 0 1px 0 var(--borde);
}

@media (max-width: 640px) {
    .hero { padding: 32px 24px 24px 24px; }
    .hero h1 { font-size: 44px; }
    .hero .logo { width: 44px; height: 44px; }
    .paso .titulo { font-size: 24px; }
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


def _alternar_tema():
    st.session_state["tema_claro"] = not st.session_state.get("tema_claro", False)


@st.fragment
def selector_tema():
    # Fragmento: cambiar el tema no vuelve a procesar el archivo subido.
    claro = st.session_state.get("tema_claro", False)
    with st.container(key="selector_tema"):
        st.button(
            ":material/dark_mode:" if claro else ":material/light_mode:",
            key="btn_tema",
            on_click=_alternar_tema,
            help="Cambiar a tema oscuro" if claro else "Cambiar a tema claro",
        )
    if claro:
        st.markdown(CSS_TEMA_CLARO, unsafe_allow_html=True)



# =============================================================================
# 2. COMPONENTES DE LA INTERFAZ
# =============================================================================


def configurar_pagina():
    """Título, icono, ancho y hoja de estilos de la página. Debe ser lo primero que se dibuja."""
    st.set_page_config(page_title=f"{NOMBRE_APP} · Separador de nombres", layout="centered", page_icon=ICONO_APP)
    st.markdown(CSS_TEMA_OSCURO + ESTILOS, unsafe_allow_html=True)


def encabezado():
    """Portada: marca, lema, descripción y características."""
    st.markdown(
        f"""
        <div class="hero">
            <div class="badge">Herramienta para Excel</div>
            <div class="marca">{LOGO_SVG}<h1>{NOMBRE_APP}</h1></div>
            <div class="lema">{LEMA_APP}</div>
            <p>Divide automáticamente los nombres completos en 4 columnas. Deduce solo, fila por fila,
            si cada nombre viene como Nombres + Apellidos o Apellidos + Nombres, y deja el documento
            tal cual: imágenes, logos, colores, tablas y filtros.</p>
            <div class="chips">
                <span class="chip">✔ Entiende (A) y (B) por celda</span>
                <span class="chip">✔ Deduce el orden sin letras</span>
                <span class="chip">✔ Conserva imágenes y formato</span>
                <span class="chip">✔ Detecta conectores (DE LA, DEL…)</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def tarjeta_resultado(res):
    """Tarjeta de proceso completado con el conteo por tipo de celda y las cifras principales."""
    c = res["conteo"]
    n_revisar = len(res["revisar"])
    detalle = (f"Con (A): {c['A']} · con (B): {c['B']} · "
               f"sin letra, deducidas por el programa: {c['AUTO']}")

    st.markdown(
        f"""
                        <div class="resultado">
                            <h3>✅ ¡Proceso completado!</h3>
                            <p>{detalle}</p>
                            <div class="stats">
                                <div class="stat"><div class="valor">{res['registros']}</div><div class="etq">Registros</div></div>
                                <div class="stat"><div class="valor">4</div><div class="etq">Columnas nuevas</div></div>
                                <div class="stat"><div class="valor">{n_revisar}</div><div class="etq">Por revisar</div></div>
                            </div>
                        </div>
                        """,
        unsafe_allow_html=True,
    )


def tarjeta_edades(menores, total, limite, sin_fecha, no_validas, hoy):
    """Recuadro con cuántas personas son menores que la edad elegida."""
    excluidas = []
    if sin_fecha:
        excluidas.append(f"{sin_fecha} sin fecha de nacimiento")
    if no_validas:
        excluidas.append(f"{no_validas} con una fecha que no se pudo leer")
    nota = f"Edades calculadas al {hoy:%d/%m/%Y}."
    if excluidas:
        nota += " No se contaron: " + " y ".join(excluidas) + "."
    anios = "año" if limite == 1 else "años"
    st.markdown(
        f'<div class="edades"><div class="etq">Personas menores de {limite} {anios}</div>'
        f'<div class="cifra">{menores} <span>de {total}</span></div>'
        f'<div class="nota">{nota}</div></div>',
        unsafe_allow_html=True,
    )


def pie():
    """Pie de página."""
    st.markdown(f'<div class="pie">{NOMBRE_APP} · Tus archivos se procesan en memoria, no se guardan en disco y se descartan al quitarlos o cerrar la pestaña</div>', unsafe_allow_html=True)
