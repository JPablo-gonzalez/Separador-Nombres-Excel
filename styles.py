"""Estética e interfaz visual: textos, colores, CSS y componentes visuales de Streamlit."""

import os

import streamlit as st

# =============================================================================
# 1. ESTILOS (solo estética)
# =============================================================================

NOMBRE_APP = "Desglosa"
LEMA_APP = "Separador inteligente de nombres"
_ICONO_ARCHIVO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "favicon.png")
ICONO_APP = _ICONO_ARCHIVO if os.path.exists(_ICONO_ARCHIVO) else "📊"

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
# Oscuro es el tema por defecto; Claro se activa solo con el botón.
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

CSS_TEMA_OSCURO = f"<style>:root {{ {TEMA_OSCURO} }}</style>"
CSS_TEMA_CLARO = f"<style>:root {{ {TEMA_CLARO} }}</style>"

ESTILOS = """
<style>
/* Sin fuentes externas (Google Fonts): cargarlas enviaba la IP del usuario a un tercero en cada visita.
   Se usa Inter si está instalada y, si no, la fuente del sistema. */

:root {
    --verde: #10b981;
    --verde-osc: #059669;
    --azul: #3b82f6;
    --azul-osc: #1d4ed8;
}

html, body, [class*="css"], .stApp {
    font-family: 'Inter', system-ui, -apple-system, 'Segoe UI', Roboto, sans-serif;
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

/* Botón discreto de tema (esquina inferior izquierda) */
.st-key-selector_tema {
    position: fixed; left: 16px; bottom: 16px; z-index: 1000;
    width: auto !important;
}
.st-key-selector_tema button {
    width: 38px; height: 38px; min-height: 0; padding: 0;
    border-radius: 50%;
    background: var(--fondo-card);
    border: 1px solid var(--borde);
    color: var(--gris);
    opacity: .55;
    box-shadow: 0 6px 16px -8px var(--sombra);
    transition: opacity .2s ease, border-color .2s ease, transform .2s ease;
}
.st-key-selector_tema button p { color: inherit; font-size: 1.1rem; line-height: 1; }
.st-key-selector_tema button:hover {
    opacity: 1; color: var(--tinta); border-color: var(--azul); transform: scale(1.06);
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
.pregunta { margin: 1.3rem 0 .55rem 0; font-size: .9rem; font-weight: 600; color: var(--gris); }

/* Zona de carga de archivo */
[data-testid="stFileUploader"] section {
    display: flex; flex-direction: column; align-items: center; justify-content: center;
    gap: 1.1rem;
    text-align: center;
    padding: 2.2rem 1.2rem 2rem 1.2rem;
    background:
        radial-gradient(420px 160px at 50% 0%, rgba(59,130,246,.10), transparent 70%),
        var(--fondo-card);
    border: 2px dashed var(--punteado);
    border-radius: 22px;
    box-shadow: 0 14px 30px -22px var(--sombra);
    transition: all .2s ease;
}
[data-testid="stFileUploader"] section:hover {
    border-color: var(--verde);
    background:
        radial-gradient(420px 160px at 50% 0%, rgba(16,185,129,.14), transparent 70%),
        var(--fondo-card);
    transform: translateY(-2px);
    box-shadow: 0 20px 36px -22px rgba(16,185,129,.45);
}
/* icono */
[data-testid="stFileUploaderDropzoneInstructions"] {
    display: flex; flex-direction: column; align-items: center; gap: .8rem;
}
[data-testid="stFileUploaderDropzoneInstructions"] svg,
[data-testid="stFileUploaderDropzoneInstructions"] [data-testid="stIconMaterial"] {
    box-sizing: content-box;
    width: 1.9rem; height: 1.9rem; font-size: 1.9rem; line-height: 1.9rem;
    padding: .85rem;
    color: #fff; fill: #fff;
    overflow: hidden; white-space: nowrap;
    border-radius: 18px;
    background: linear-gradient(135deg, var(--azul), var(--verde));
    box-shadow: 0 12px 22px -10px rgba(59,130,246,.65);
}
/* textos en español (reemplazan los de Streamlit) */
[data-testid="stFileUploaderDropzoneInstructions"] > div > * { display: none; }
[data-testid="stFileUploaderDropzoneInstructions"] > div::before {
    content: "Arrastra tu archivo Excel aquí";
    display: block;
    font-size: 1.02rem; font-weight: 700; color: var(--tinta);
}
[data-testid="stFileUploaderDropzoneInstructions"] > div::after {
    content: "o selecciónalo desde tu equipo · solo .xlsx";
    display: block; margin-top: .3rem;
    font-size: .82rem; font-weight: 400; color: var(--gris);
}
/* botón */
[data-testid="stFileUploader"] section button {
    font-size: 0 !important;
    padding: .72rem 1.5rem;
    border: none;
    border-radius: 999px;
    background: linear-gradient(135deg, var(--azul) 0%, var(--verde) 100%);
    box-shadow: 0 14px 24px -12px rgba(16,185,129,.7);
    transition: transform .2s ease, box-shadow .2s ease;
}
[data-testid="stFileUploader"] section button::after {
    content: "Seleccionar archivo";
    font-size: .92rem; font-weight: 700; letter-spacing: .01em; color: #fff;
}
[data-testid="stFileUploader"] section button * { display: none; }
[data-testid="stFileUploader"] section button:hover {
    transform: translateY(-2px);
    box-shadow: 0 18px 28px -12px rgba(59,130,246,.75);
}
[data-testid="stFileUploader"] section button:active { transform: translateY(0); }
/* archivo ya cargado */
[data-testid="stFileUploaderFile"] {
    margin-top: .6rem;
    padding: .65rem .9rem;
    background: var(--fondo-card);
    border: 1px solid var(--borde);
    border-left: 4px solid var(--verde);
    border-radius: 14px;
    color: var(--tinta);
}
[data-testid="stFileUploaderFile"] * { color: var(--tinta); }
[data-testid="stFileUploaderFile"] small { color: var(--gris); }

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

/* Consulta de edades */
.edades {
    background: var(--fondo-card);
    border: 1px solid var(--borde);
    border-left: 5px solid var(--azul);
    border-radius: 16px;
    padding: 1.1rem 1.3rem;
    margin: .3rem 0 1rem 0;
    box-shadow: 0 10px 24px -16px var(--sombra);
}
.edades .etq { font-size: .92rem; color: var(--gris); }
.edades .cifra { font-size: 1.9rem; font-weight: 800; color: var(--acento-texto); line-height: 1.25; }
.edades .cifra span { font-size: 1.05rem; font-weight: 600; color: var(--gris); }
.edades .nota { font-size: .8rem; color: var(--gris); margin-top: .35rem; }
.st-key-edad_limite { max-width: 340px; }
.st-key-edad_limite label p { font-size: .9rem; font-weight: 600; color: var(--gris); }
.st-key-edad_limite [data-baseweb="input"] { border-color: var(--borde); }
.st-key-edad_limite [data-baseweb="input"],
.st-key-edad_limite [data-baseweb="input"] > div,
.st-key-edad_limite input,
.st-key-edad_limite button { background: var(--fondo-card) !important; color: var(--tinta) !important; }

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
