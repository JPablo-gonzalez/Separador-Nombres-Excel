import io
import math
import os
import re
import unicodedata
from collections import Counter
from copy import copy
from dataclasses import dataclass

import openpyxl
import streamlit as st
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token
from openpyxl.utils import column_index_from_string, get_column_letter, range_boundaries
from openpyxl.worksheet.cell_range import MultiCellRange
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.table import TableColumn

# =============================================================================
# 1. DICCIONARIOS SEMÁNTICOS Y CONECTORES
# =============================================================================


def _norm(texto):
    """Mayúsculas y sin tildes/diacríticos: MEJÍA -> MEJIA, MUÑOZ -> MUNOZ."""
    t = unicodedata.normalize("NFD", str(texto).upper())
    return "".join(c for c in t if unicodedata.category(c) != "Mn")


# Conectores que se pegan a la(s) palabra(s) siguiente(s): "DE LA CRUZ", "DEL RIO", "SAN JUAN"...
CONECTORES = {"DE", "DEL", "LA", "LOS", "LAS", "SAN", "SANTA", "VON", "VAN", "DA", "DOS", "DAS", "DI"}
# Los que, al abrir un bloque, delatan un apellido ("De La Cruz", "Del Rio", "Von Trapp").
# SAN y SANTA quedan fuera: también aparecen en nombres.
CONECTORES_APELLIDO = CONECTORES - {"SAN", "SANTA"}

# Bloques religiosos que cierran un nombre compuesto (cuentan como UNA sola palabra y son nombre).
RELIGIOSOS = {
    "DE JESUS", "DEL CARMEN", "DE MARIA", "DEL ROSARIO",
    "DEL PILAR", "DE LOS ANGELES", "DE LOS DOLORES",
}

_NOMBRES_BASE = {
    "JUAN", "CARLOS", "LUIS", "JOSE", "MARIA", "ANDRES", "DAVID", "ALEJANDRO", "ARTURO", "FREDY",
    "EDWIN", "GERARDO", "WILSON", "EDISON", "JHON", "JORGE", "ALEXANDER", "JULIAN", "DIEGO", "DANIEL",
    "MIGUEL", "ANGEL", "HERNANDO", "GUILLERMO", "GUSTAVO", "JAIME", "ALBERTO", "HECTOR", "JAIRO",
    "CESAR", "JULIO", "TITO", "ERNESTO", "NELSON", "CONRADO", "ALIRIO", "LIBANIEL", "EDER", "DIANA",
    "LUCIA", "WILLIAM", "ANTONIO", "ARMANDO", "MIRIAM", "GABRIELA", "ALFONSO", "JOSELIN", "HERIBERTO",
    "REGINA", "AMPARO", "MANUEL", "HERNAN",
    "PABLO", "PEDRO", "SANTIAGO", "SEBASTIAN", "NICOLAS", "FELIPE", "MATEO", "SAMUEL", "JESUS",
    "FERNANDO", "RICARDO", "ROBERTO", "OSCAR", "MAURICIO", "FABIAN", "CRISTIAN", "STEVEN", "BRAYAN",
    "KEVIN", "JHONATAN", "YEISON", "ESTEBAN", "EDUARDO", "ALEXIS", "ANDREA", "PAOLA", "CATALINA",
    "VALENTINA", "SOFIA", "ISABELLA", "CAMILA", "LAURA", "ANA", "CLAUDIA", "PATRICIA", "MARTHA",
    "SANDRA", "LUZ", "ELENA", "YESICA", "LEIDY", "NATALIA", "DANIELA", "MARIANA", "JULIANA",
    "PAULA", "ROSA", "LILIANA", "GLORIA", "BEATRIZ", "ADRIANA", "CAROLINA", "JOHAN", "SERGIO",
    "ISAAC", "EMMANUEL", "THOMAS", "SIMON", "JERONIMO", "MARTIN", "VICTOR", "RAFAEL", "ANDERSON",
    # ampliación
    "CAMILO", "GABRIEL", "LEONARDO", "LEON", "JAVIER", "IVAN", "FELIX", "MARCO", "MARCOS", "MARIO",
    "HUGO", "ORLANDO", "RAUL", "ROBINSON", "ROLANDO", "ELKIN", "EDGAR", "FABIO", "FRANCISCO", "HENRY",
    "JHONNY", "JOAQUIN", "LEONEL", "NESTOR", "OMAR", "OSWALDO", "RODRIGO", "SAUL", "TOMAS", "WALTER",
    "ALVARO", "ALAN", "DYLAN", "JUSTIN", "LUCAS", "BRYAN", "JEFFERSON", "SNEIDER", "YEFERSON", "JHOAN",
    "CRISTHIAN", "STIVEN", "DUVAN", "DEIBY", "YEISSON", "WILMAR", "WILDER", "FAVIO", "FABRICIO",
    "SARA", "NICOLE", "SHIRLEY", "KAREN", "LINA", "MONICA", "JENNIFER", "YULIANA", "MELISSA",
    "ALEJANDRA", "JOHANNA", "JOHANA", "TATIANA", "MARCELA", "ANGELA", "ANGIE", "JUANA", "LUISA",
    "FERNANDA", "ISABEL", "VERONICA", "DAYANA", "ESTEFANIA", "STEFANIA", "MARISOL", "YULIETH",
    "MAYERLY", "ALEXANDRA", "AURA", "BLANCA", "CARMEN", "CECILIA", "CLARA", "DORA", "ELSA", "EMILIA",
    "ESPERANZA", "EVA", "FLOR", "GISELA", "GRACIELA", "INES", "IRENE", "JUDITH", "KATHERINE",
    "LORENA", "LUCERO", "MABEL", "MARGARITA", "MARISELA", "MELANIE", "MERCEDES", "NANCY", "NORA",
    "OLGA", "PILAR", "RAQUEL", "RUTH", "SILVIA", "SOL", "SUSANA", "TERESA", "VANESSA", "VIVIANA",
    "XIMENA", "YOLANDA", "YURANI", "ZULEIDY", "MARIA", "KATHERIN", "KAROL", "LEIDY", "MAIRA",
    "ESTEFANY", "YENNY", "RAMIRO", "CIRO", "YESENIA", "JESSICA", "JESICA", "DAISY", "DEISY", "ERIKA", "ERICA", "SHARON",
}

_APELLIDOS_BASE = {
    "GOMEZ", "ZAPATA", "PEREZ", "OSORIO", "VERA", "BETANCUR", "MORALES", "GALEANO", "ESPINOSA",
    "GUARIN", "CELIS", "RAMIREZ", "HERNANDEZ", "TORO", "OCAMPO", "ARROYAVE", "ARANGO", "MUNOZ",
    "LONDONO", "AGUIRRE", "AMESQUITA", "MARIN", "BETANCURT", "TOBON", "GARCIA", "MEJIA", "ARANZAZU",
    "OSPINA", "SANCHEZ", "GAVIRIA", "CANO", "RUIZ", "BARRERA", "GALLO", "RAMOS", "GRAJALES",
    "GRISALES", "BOTERO", "CASTRO", "BARRETO", "ZAMBRANO", "BUITRAGO", "OBANDO", "GALLEGO", "MESA",
    "ARIAS", "CASTANO", "HERRERA", "MOLINA", "CARDONA", "PARRA", "TASCON", "DIAZ", "LOPEZ", "MACIAS",
    "RODRIGUEZ", "MARTINEZ", "GONZALEZ", "VARGAS", "JIMENEZ", "RESTREPO", "QUINTERO", "FIGUEROA",
    "SUAREZ", "CARDENAS", "ROJAS", "GUTIERREZ", "ORTIZ", "VELEZ", "DUQUE", "VALENCIA", "CORREA",
    "URIBE", "ALZATE", "MORENO", "SALAZAR", "CARDOSO", "MONTOYA", "JARAMILLO", "NARANJO", "PALACIO",
    "PALACIOS", "FRANCO", "HOYOS", "VILLA", "AGUDELO", "VASQUEZ", "TORRES", "FLOREZ", "ESCOBAR",
    "CASTILLO", "ROMERO", "ALVAREZ", "ACEVEDO", "SERNA", "USUGA", "ECHEVERRI", "POSADA", "LOAIZA",
    "HENAO", "MURILLO", "ZULUAGA", "OROZCO", "CARVAJAL", "BEDOYA", "RIVERA", "SIERRA", "PINEDA",
    # ampliación
    "ALVARADO", "ARBOLEDA", "ATEHORTUA", "AVENDANO", "BALLESTEROS", "BERRIO", "BLANDON", "CALLE",
    "CAMPUZANO", "CARO", "CASAS", "CEBALLOS", "CHAVARRIA", "COLORADO", "CORTES", "CUARTAS",
    "DELGADO", "DURANGO", "ESTRADA", "FERNANDEZ", "FUENTES", "GIL", "GIRALDO", "GUERRA", "GUZMAN",
    "HIGUITA", "HINCAPIE", "HOLGUIN", "LEON", "LOZANO", "MADRID", "MARQUEZ", "MAZO", "MEDINA",
    "MENDEZ", "MENDOZA", "MIRA", "MONSALVE", "MONTES", "MORA", "MUNERA", "NIETO", "NUNEZ", "OCHOA",
    "OLARTE", "OSSA", "PABON", "PAEZ", "PATINO", "PELAEZ", "PEREA", "PIEDRAHITA", "QUICENO", "QUIROZ",
    "RENDON", "RENTERIA", "REYES", "RIOS", "RIVAS", "ROLDAN", "SALDARRIAGA", "SOTO", "TABARES",
    "TAMAYO", "TREJOS", "TRUJILLO", "VALDERRAMA", "VALLEJO", "VARELA", "VELASQUEZ", "VERGARA",
    "VIDAL", "YEPES", "CRUZ", "MARTIN", "BEJARANO", "BERMUDEZ", "BUSTAMANTE", "CAICEDO", "CAMACHO",
    "CARRILLO", "CHACON", "CORDOBA", "COSSIO", "DOMINGUEZ", "ESPITIA", "GAITAN", "GALVIS", "GARZON",
    "GIRON", "GUEVARA", "IBARRA", "LARA", "LEAL", "LUNA", "MALDONADO", "MANRIQUE", "MARIN", "MEJIAS",
    "MONTAÑO", "MONTANO", "NAVARRO", "ORDONEZ", "PADILLA", "PAREDES", "PENA", "PRADA", "QUINTANA",
    "RANGEL", "SALGADO", "SANDOVAL", "SEPULVEDA", "SOLANO", "TELLO", "URREGO", "VACA",
    "VEGA", "VILLEGAS", "YEPEZ", "ZABALA", "ZUNIGA",
}

# Los sets exponen el contenido normalizado (sin tildes) para comparar sin importar el acento.
NOMBRES_COMUNES = {_norm(x) for x in _NOMBRES_BASE}
APELLIDOS_COMUNES = {_norm(x) for x in _APELLIDOS_BASE}

ORIENTACIONES = ("N->A", "A->N")
RE_PREFIJO = re.compile(r"^\(\s*([AaBb])\s*\)\s*")


def agrupar_conectores(tokens):
    """Une conectores hispanos con la palabra que les sigue ("DE LA" + "CRUZ" -> "DE LA CRUZ")
    y deja los bloques religiosos finales ("DE JESUS", "DEL CARMEN") como una sola palabra."""
    resultado = []
    n = len(tokens)
    i = 0
    while i < n:
        j = i
        while j < n - 1 and _norm(tokens[j]) in CONECTORES:
            j += 1
        resultado.append(" ".join(tokens[i:j + 1]))
        i = j + 1
    return resultado


# =============================================================================
# 2. MOTOR DE DECISIÓN (orientación y reparto deducidos celda por celda)
# =============================================================================
#
# Para cada celda se enumeran TODAS las lecturas posibles: cada orientación (N->A / A->N) y cada
# reparto de palabras entre nombres (k) y apellidos (m). Cada lectura recibe un puntaje y gana la mayor.
#
#   · Prefijo (A)/(B): fija la orientación de esa celda (y se borra). Solo se decide el reparto.
#   · Celda sin prefijo: las dos orientaciones compiten con la evidencia de las palabras:
#       - diccionarios NOMBRES_COMUNES / APELLIDOS_COMUNES (evidencia fuerte, ±2),
#       - palabras que el propio archivo enseña con sus filas más claras (evidencia media, ±1 a ±1.5),
#       - terminaciones típicas de apellido (-EZ, -IZ, -OZ, -AZ) y conectores iniciales (evidencia débil),
#       - estructura más frecuente (1-2 nombres, 2 apellidos) como último respaldo.
#   · Si una celda no ofrece ninguna evidencia, se usa la tendencia de las filas claras del mismo
#     archivo; si tampoco hay tendencia, el orden natural (Nombres -> Apellidos). Estas filas se
#     marcan para revisión.

MARGEN_ORIENTACION = 1.0      # diferencia mínima entre orientaciones para decidir por evidencia
MARGEN_APRENDER_PREFIJO = 1.5  # con prefijo, el reparto debe ser claro para enseñar palabras nuevas
MARGEN_APRENDER_AUTO = 3.0    # sin prefijo, la lectura debe ser muy clara para enseñar palabras nuevas
MARGEN_REVISAR = 1.0          # por debajo de esta confianza la fila se marca para revisión
ORIENTACION_NATURAL = "N->A"
MAX_ITERACIONES = 4
SUFIJOS_APELLIDO = ("EZ", "IZ", "OZ", "AZ")


@dataclass
class Resultado:
    texto: str                 # texto original de la celda
    p_ape: str = ""
    s_ape: str = ""
    p_nom: str = ""
    s_nom: str = ""
    orientacion: str = ORIENTACION_NATURAL
    origen: str = "AUTO"       # "A", "B" (prefijo) o "AUTO" (deducido)
    confianza: float = 0.0
    revisar: bool = False
    motivo: str = ""


class Lexico:
    """Palabras que el propio archivo enseña, además de los diccionarios base."""

    def __init__(self):
        self.nombres = Counter()
        self.apellidos = Counter()

    def agregar(self, contribucion):
        for (palabra, rol), cuantas in contribucion.items():
            (self.nombres if rol == "N" else self.apellidos)[palabra] += cuantas

    def peso(self, palabra, rol, excluir=None):
        base = NOMBRES_COMUNES if rol == "N" else APELLIDOS_COMUNES
        if palabra in base:
            return 2.0
        veces = (self.nombres if rol == "N" else self.apellidos)[palabra]
        if excluir:                                   # una fila no se enseña a sí misma
            veces -= excluir.get((palabra, rol), 0)
        return min(1.5, 0.75 + 0.25 * veces) if veces > 0 else 0.0


def _afinidades(token, lex, excluir=None):
    """(afinidad como nombre, afinidad como apellido) de una palabra o bloque."""
    clave = _norm(token)
    if clave in RELIGIOSOS:
        return 2.0, -2.0
    palabras = clave.split()
    nucleo = palabras[-1]
    wn = lex.peso(nucleo, "N", excluir)
    wa = lex.peso(nucleo, "A", excluir)
    if wn and wa:                                      # palabra ambigua (León, Martín...)
        an, aa = 0.5, 0.5
    elif wn:
        an, aa = wn, -wn
    elif wa:
        an, aa = -wa, wa
    elif len(nucleo) >= 5 and nucleo.endswith(SUFIJOS_APELLIDO):
        an, aa = -0.75, 0.75
    else:
        an, aa = 0.0, 0.0
    if len(palabras) > 1 and palabras[0] in CONECTORES_APELLIDO:
        an, aa = an - 1.0, aa + 1.0
    return an, aa


def _hipotesis(n):
    """Todas las lecturas posibles: (orientación, nº de nombres k, nº de apellidos m)."""
    if n == 1:
        return [("N->A", 1, 0), ("A->N", 0, 1)]
    return [(o, k, n - k) for o in ORIENTACIONES for k in range(1, n)]


def _prior_estructura(k, m):
    """Estructura más frecuente en Colombia: 2 apellidos y 1-2 nombres."""
    if k + m == 1:
        return 0.0
    return -1.0 * abs(m - 2) - 0.5 * abs(k - 2)


def _indices(n, orientacion, k, m):
    """Posiciones de las palabras que son nombres y de las que son apellidos."""
    if orientacion == "N->A":
        return list(range(k)), list(range(k, n))
    return list(range(m, n)), list(range(m))


def _puntuar(afin, orientacion, k, m):
    ni, ai = _indices(len(afin), orientacion, k, m)
    return (sum(afin[i][0] for i in ni) + sum(afin[i][1] for i in ai) + _prior_estructura(k, m))


@dataclass
class _Evaluacion:
    orientacion: str
    k: int
    m: int
    fallback: bool
    margen_est: float
    margen_or: float

    @property
    def confianza(self):
        return min(self.margen_est, self.margen_or)


def _evaluar(tokens, fija, lex, excluir, prior_archivo):
    n = len(tokens)
    afin = [_afinidades(t, lex, excluir) for t in tokens]
    puntos = {h: _puntuar(afin, *h) for h in _hipotesis(n)}
    mejor = {o: max(p for h, p in puntos.items() if h[0] == o) for o in ORIENTACIONES}

    if fija:
        orientacion, fallback, margen_or = fija, False, math.inf
    else:
        dif = mejor["N->A"] - mejor["A->N"]
        margen_or = abs(dif)
        if margen_or >= MARGEN_ORIENTACION:
            orientacion, fallback = ("N->A" if dif > 0 else "A->N"), False
        else:                                          # sin evidencia: tendencia del archivo u orden natural
            orientacion, fallback = (prior_archivo or ORIENTACION_NATURAL), True

    cand = sorted(((p, h) for h, p in puntos.items() if h[0] == orientacion),
                  key=lambda x: (-x[0], x[1][1]))
    _, (_, k, m) = cand[0]
    margen_est = cand[0][0] - cand[1][0] if len(cand) > 1 else math.inf
    return _Evaluacion(orientacion, k, m, fallback, margen_est, margen_or)


def _contribucion(tokens, ev):
    """Palabras desconocidas de una lectura clara: lo que esa fila le enseña al resto del archivo."""
    ni, ai = _indices(len(tokens), ev.orientacion, ev.k, ev.m)
    contribucion = Counter()
    for rol, indices in (("N", ni), ("A", ai)):
        for i in indices:
            clave = _norm(tokens[i])
            if clave in RELIGIOSOS:
                continue
            nucleo = clave.split()[-1]
            if len(nucleo) < 3 or nucleo in NOMBRES_COMUNES or nucleo in APELLIDOS_COMUNES:
                continue
            contribucion[(nucleo, rol)] += 1
    return contribucion


def _puede_ensenar(ev, fija):
    if fija:
        return ev.margen_est >= MARGEN_APRENDER_PREFIJO
    return (not ev.fallback) and ev.confianza >= MARGEN_APRENDER_AUTO


def separar_prefijo(texto):
    """(texto sin prefijo, orientación fijada por el prefijo o None, origen 'A' | 'B' | 'AUTO')."""
    texto = str(texto).strip()
    m = RE_PREFIJO.match(texto)
    if m:
        letra = m.group(1).upper()
        return texto[m.end():].strip(), ("A->N" if letra == "A" else "N->A"), letra
    return texto, None, "AUTO"


def _armar(tokens, ev):
    ni, ai = _indices(len(tokens), ev.orientacion, ev.k, ev.m)
    nombres = [tokens[i] for i in ni]
    apellidos = [tokens[i] for i in ai]
    return (apellidos[0] if apellidos else "", " ".join(apellidos[1:]),
            nombres[0] if nombres else "", " ".join(nombres[1:]))


def _motivo(origen, ev, prior_archivo):
    flecha = "Nombres → Apellidos" if ev.orientacion == "N->A" else "Apellidos → Nombres"
    if origen in ("A", "B"):
        texto = f"Prefijo ({origen}): {flecha}"
    elif not ev.fallback:
        texto = f"Deducido por los diccionarios: {flecha}"
    elif prior_archivo:
        texto = f"Sin palabras conocidas; se siguió la tendencia del archivo: {flecha}"
    else:
        texto = f"Sin palabras conocidas ni tendencia en el archivo; se asumió el orden natural: {flecha}"
    if ev.margen_est < MARGEN_REVISAR:
        texto += " (reparto de palabras poco claro)"
    return texto


def resolver_lista(textos):
    """Resuelve TODAS las celdas de una columna con un único flujo.

    Devuelve una lista de `Resultado` (uno por texto). Sirve igual para archivos mixtos (celdas con
    (A), con (B) y sin letra) que para archivos sin ninguna letra.
    """
    preparadas = []
    for texto in textos:
        limpio, fija, origen = separar_prefijo(texto)
        preparadas.append((texto, agrupar_conectores(limpio.split()), fija, origen))

    total = len(preparadas)
    lex, contribs, prior = Lexico(), [None] * total, None
    previas, evals = None, []
    for _ in range(MAX_ITERACIONES):
        evals = [(_evaluar(tk, fija, lex, contribs[i], prior) if tk else None)
                 for i, (_, tk, fija, _) in enumerate(preparadas)]
        if previas is not None and all(
                (a is None and b is None) or (a and b and (a.orientacion, a.k) == (b.orientacion, b.k))
                for a, b in zip(previas, evals)):
            break
        previas = evals
        # Se reconstruye lo aprendido a partir de las filas claras de esta vuelta.
        lex, contribs = Lexico(), [None] * total
        conteo = Counter()
        for i, ((_, tk, fija, _), ev) in enumerate(zip(preparadas, evals)):
            if ev is None:
                continue
            if _puede_ensenar(ev, fija):
                contribs[i] = _contribucion(tk, ev)
                lex.agregar(contribs[i])
            if fija is None and not ev.fallback and ev.confianza >= MARGEN_APRENDER_AUTO:
                conteo[ev.orientacion] += 1
        if conteo["N->A"] != conteo["A->N"]:
            prior = "N->A" if conteo["N->A"] > conteo["A->N"] else "A->N"
        else:
            prior = None

    resultados = []
    for (texto, tokens, fija, origen), ev in zip(preparadas, evals):
        if ev is None:
            resultados.append(Resultado(texto=str(texto), origen=origen, motivo="Celda sin nombre"))
            continue
        pa, sa, pn, sn = _armar(tokens, ev)
        resultados.append(Resultado(
            texto=str(texto), p_ape=pa, s_ape=sa, p_nom=pn, s_nom=sn,
            orientacion=ev.orientacion, origen=origen,
            confianza=ev.confianza if math.isfinite(ev.confianza) else 99.0,
            revisar=ev.fallback or ev.confianza < MARGEN_REVISAR,
            motivo=_motivo(origen, ev, prior),
        ))
    return resultados


def analizar_nombre(texto):
    """Una sola celda, sin el contexto del resto del archivo (útil para pruebas)."""
    return resolver_lista([texto])[0]


def copiar_estilo_seguro(origen, destino):
    if origen.has_style:
        if origen.font: destino.font = copy(origen.font)
        if origen.border: destino.border = copy(origen.border)
        if origen.fill: destino.fill = copy(origen.fill)
        if origen.number_format: destino.number_format = copy(origen.number_format)
        if origen.protection: destino.protection = copy(origen.protection)
        if origen.alignment: destino.alignment = copy(origen.alignment)


# =============================================================================
# 3. AJUSTE DE REFERENCIAS AL INSERTAR COLUMNAS
# =============================================================================
# openpyxl mueve las celdas al insertar columnas, pero NO actualiza las referencias que
# apuntan a ellas. Estas funciones replican lo que hace Excel: toda referencia a una
# columna >= insert_idx se desplaza n columnas hacia la derecha.

_RE_CELDA = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$")
_RE_COL = re.compile(r"^(\$?)([A-Za-z]{1,3})$")


def _desplazar_col(letras, ins, n):
    c = column_index_from_string(letras.upper())
    return get_column_letter(c + n) if c >= ins else letras


def _desplazar_operando(op, titulo, ins, n, hoja_propia):
    if "!" in op:
        pref, ref = op.rsplit("!", 1)
        if pref.strip("'").replace("''", "'") != titulo:
            return op                      # apunta a otra hoja
    else:
        pref, ref = None, op
        if not hoja_propia:
            return op                      # sin hoja explícita, pero la fórmula es de otra hoja
    if "[" in ref:
        return op                          # referencias estructuradas de tabla (por nombre)
    partes = ref.split(":")
    if len(partes) > 2:
        return op
    nuevas = []
    for p in partes:
        m = _RE_CELDA.match(p)
        if m:
            nuevas.append(f"{m.group(1)}{_desplazar_col(m.group(2), ins, n)}{m.group(3)}{m.group(4)}")
            continue
        m = _RE_COL.match(p)
        if m and len(partes) == 2:         # columnas completas (A:C); solas podrían ser un nombre
            nuevas.append(f"{m.group(1)}{_desplazar_col(m.group(2), ins, n)}")
            continue
        nuevas.append(p)
    ref_nueva = ":".join(nuevas)
    return f"{pref}!{ref_nueva}" if pref is not None else ref_nueva


def _desplazar_formula(texto, titulo, ins, n, hoja_propia):
    """texto incluye el '=' inicial. Si algo falla, devuelve el texto original."""
    try:
        tok = Tokenizer(texto)
        cambio = False
        for t in tok.items:
            if t.type == Token.OPERAND and t.subtype == Token.RANGE:
                nuevo = _desplazar_operando(t.value, titulo, ins, n, hoja_propia)
                if nuevo != t.value:
                    t.value = nuevo
                    cambio = True
        return tok.render() if cambio else texto
    except Exception:
        return texto


def _desplazar_rangos(texto, titulo, ins, n):
    """Lista de rangos separados por espacios (sqref) de la hoja propia."""
    return " ".join(_desplazar_operando(r, titulo, ins, n, True) for r in str(texto).split())


def desplazar_referencias(wb, sheet, ins, n):
    titulo = sheet.title

    # 1) Fórmulas de todas las hojas (una fórmula de otra hoja puede apuntar a esta)
    for ws in wb.worksheets:
        propia = ws is sheet
        for fila in ws.iter_rows():
            for c in fila:
                v = c.value
                if propia and c.hyperlink:
                    c.hyperlink.ref = c.coordinate      # el enlace viaja con su celda
                if isinstance(v, str) and v.startswith("="):
                    c.value = _desplazar_formula(v, titulo, ins, n, propia)
                elif isinstance(v, ArrayFormula):
                    v.text = _desplazar_formula(v.text, titulo, ins, n, propia)
                    if propia and v.ref:
                        v.ref = _desplazar_rangos(v.ref, titulo, ins, n)

    # 2) Formato condicional
    nuevo_cf = ConditionalFormattingList()
    for cf in sheet.conditional_formatting:
        rango = _desplazar_rangos(cf.sqref, titulo, ins, n)
        for regla in cf.rules:
            regla.formula = [_desplazar_formula("=" + f, titulo, ins, n, True)[1:] for f in regla.formula]
            nuevo_cf.add(rango, regla)
    sheet.conditional_formatting = nuevo_cf

    # 3) Validaciones de datos
    for dv in sheet.data_validations.dataValidation:
        dv.sqref = MultiCellRange(_desplazar_rangos(dv.sqref, titulo, ins, n))
        for attr in ("formula1", "formula2"):
            f = getattr(dv, attr)
            if f and not f.startswith('"'):
                setattr(dv, attr, _desplazar_formula("=" + f, titulo, ins, n, True)[1:])

    # 4) Nombres definidos (del libro y de la hoja)
    for dn in wb.defined_names.values():
        if dn.attr_text:
            dn.attr_text = _desplazar_formula("=" + dn.attr_text, titulo, ins, n, False)[1:]
    for ws in wb.worksheets:
        for dn in ws.defined_names.values():
            if dn.attr_text:
                dn.attr_text = _desplazar_formula("=" + dn.attr_text, titulo, ins, n, ws is sheet)[1:]

    # 5) Área de impresión, columnas repetidas y autofiltro de la hoja
    try:
        if sheet.print_area:
            partes = []
            for r in str(sheet.print_area).split(","):
                r = r.split("!")[-1].replace("$", "")
                partes.append(_desplazar_operando(r, titulo, ins, n, True))
            sheet.print_area = partes
        if sheet.print_title_cols:
            sheet.print_title_cols = _desplazar_operando(sheet.print_title_cols.split("!")[-1].replace("$", ""), titulo, ins, n, True)
    except Exception:
        pass
    if sheet.auto_filter and sheet.auto_filter.ref:
        sheet.auto_filter.ref = _desplazar_rangos(sheet.auto_filter.ref, titulo, ins, n)

    # 6) Anchos de columna (las columnas a la derecha conservan su ancho original)
    dims = []
    for k, d in list(sheet.column_dimensions.items()):
        mn = d.min or column_index_from_string(k)
        mx = d.max or mn
        dims.append((d, mn, mx))
    sheet.column_dimensions.clear()
    for d, mn, mx in dims:
        trozos = []
        if mx < ins:
            trozos.append((mn, mx))
        elif mn >= ins:
            trozos.append((mn + n, mx + n))
        else:                              # el rango atraviesa el punto de inserción: se parte en dos
            trozos.append((mn, ins - 1))
            trozos.append((ins + n, mx + n))
        for a, b in trozos:
            nd = copy(d)
            letra = get_column_letter(a)
            nd.index, nd.min, nd.max = letra, a, b
            sheet.column_dimensions[letra] = nd

    # 7) Paneles inmovilizados que quedan a la derecha del punto de inserción
    fp = sheet.freeze_panes
    if fp:
        m = _RE_CELDA.match(fp)
        if m and column_index_from_string(m.group(2)) > ins:
            sheet.freeze_panes = f"{get_column_letter(column_index_from_string(m.group(2)) + n)}{m.group(4)}"


# =============================================================================
# 4. PROCESAMIENTO DEL LIBRO (protección de tablas, filtros y celdas combinadas)
# =============================================================================

HEADERS_NUEVOS = ["PRIMER APELLIDO", "SEGUNDO APELLIDO", "PRIMER NOMBRE", "SEGUNDO NOMBRE"]
N_NUEVAS = len(HEADERS_NUEVOS)


class ColumnaNoEncontrada(Exception):
    pass


def localizar_columna_nombres(sheet):
    for r in range(1, min(30, sheet.max_row + 1)):
        for c in range(1, sheet.max_column + 1):
            val = str(sheet.cell(row=r, column=c).value).upper()
            if ("NOMBRE" in val or "APELLIDO" in val) and val not in HEADERS_NUEVOS:
                return r, c
    return None


def ajustar_tabla(sheet, table, col_nombres, fila_header):
    """Actualiza una Table de Excel (ref, autoFilter, sortState y tableColumns).
    Las referencias de la tabla siguen en coordenadas ORIGINALES (openpyxl no las toca)."""
    insert_idx = col_nombres + 1
    c1, r1, c2, r2 = range_boundaries(table.ref)

    if c2 < col_nombres:
        return                              # tabla totalmente a la izquierda: no cambia
    if c1 > col_nombres:                    # tabla totalmente a la derecha: solo se desplaza
        n1, n2 = c1 + N_NUEVAS, c2 + N_NUEVAS
        contiene_nombres = False
    else:                                   # la tabla contiene la columna de nombres: se ensancha
        n1, n2 = c1, c2 + N_NUEVAS
        contiene_nombres = True

    header_rows = 1 if table.headerRowCount is None else table.headerRowCount
    tot = table.totalsRowCount or 0
    data_max_row = r2 - tot

    if contiene_nombres and header_rows > 0 and r1 != fila_header:
        # encabezado de la tabla en una fila distinta a la detectada: también necesita los títulos nuevos
        for i, h in enumerate(HEADERS_NUEVOS):
            celda = sheet.cell(row=r1, column=insert_idx + i)
            celda.value = h
            copiar_estilo_seguro(sheet.cell(row=r1, column=col_nombres), celda)

    table.ref = f"{get_column_letter(n1)}{r1}:{get_column_letter(n2)}{r2}"
    if table.autoFilter:
        table.autoFilter.ref = f"{get_column_letter(n1)}{r1}:{get_column_letter(n2)}{data_max_row}"
    if table.sortState:
        table.sortState.ref = f"{get_column_letter(n1)}{r1 + header_rows}:{get_column_letter(n2)}{data_max_row}"

    # Reconstruir tableColumns: IDs consecutivos, nombres únicos y que coincidan con la celda de encabezado
    antiguas = {}
    for col in table.tableColumns:
        antiguas.setdefault(col.name, col)

    if header_rows > 0:
        bases = []
        for ci in range(n1, n2 + 1):
            v = sheet.cell(row=r1, column=ci).value
            v = "" if v is None else str(v).strip()
            bases.append(v if v else f"Col_{ci}")
    else:
        nombres_viejos = [c.name for c in table.tableColumns]
        k = col_nombres - c1 + 1 if contiene_nombres else len(nombres_viejos)
        bases = (nombres_viejos[:k] + HEADERS_NUEVOS + nombres_viejos[k:]) if contiene_nombres else nombres_viejos

    usados = set()
    nuevas_columnas = []
    for idx, base in enumerate(bases):
        nombre = base
        cnt = 1
        while nombre.lower() in usados:      # Excel compara nombres sin distinguir mayúsculas
            nombre = f"{base}_{cnt}"
            cnt += 1
        usados.add(nombre.lower())

        col = antiguas.pop(base, None)
        if col is not None:
            col.id = idx + 1
            col.name = nombre
        else:
            col = TableColumn(id=idx + 1, name=nombre)
        nuevas_columnas.append(col)

        if header_rows > 0:                  # la celda del encabezado debe decir exactamente lo mismo
            celda = sheet.cell(row=r1, column=n1 + idx)
            if celda.value != nombre:
                celda.value = nombre
    table.tableColumns = nuevas_columnas


def procesar_libro(wb):
    """Inserta las 4 columnas junto a la columna de nombres y rellena los datos.

    No recibe ninguna configuración: la orientación de cada celda se deduce sola.
    Devuelve estadísticas, una vista previa y la lista de filas que conviene revisar."""
    sheet = wb.active
    pos = localizar_columna_nombres(sheet)
    if pos is None:
        raise ColumnaNoEncontrada("No se encontró una columna válida de Nombres.")
    fila_header, col_nombres = pos
    insert_idx = col_nombres + 1

    # Filas de totales de tablas: no son estudiantes
    filas_omitir = set()
    for table in sheet.tables.values():
        _, _, _, tr2 = range_boundaries(table.ref)
        tot = table.totalsRowCount or 0
        if tot:
            filas_omitir.update(range(tr2 - tot + 1, tr2 + 1))

    # 1. Descombinar TODAS las celdas combinadas (se recombinan al final, ya desplazadas)
    rangos_combinados = [m.bounds for m in sheet.merged_cells.ranges]   # (min_col, min_row, max_col, max_row)
    for m_range in list(sheet.merged_cells.ranges):
        sheet.unmerge_cells(str(m_range))

    # 2. Insertar las 4 columnas y desplazar fórmulas, filtros, formatos condicionales, etc.
    sheet.insert_cols(insert_idx, N_NUEVAS)
    desplazar_referencias(wb, sheet, insert_idx, N_NUEVAS)

    # Autofiltro de hoja que terminaba justo en la columna de nombres: ahora debe abarcar las nuevas
    if sheet.auto_filter and sheet.auto_filter.ref:
        a1, ar1, a2, ar2 = range_boundaries(sheet.auto_filter.ref)
        if a2 == col_nombres:
            sheet.auto_filter.ref = f"{get_column_letter(a1)}{ar1}:{get_column_letter(a2 + N_NUEVAS)}{ar2}"

    # 3. Encabezados nuevos con el estilo del encabezado original
    celda_origen = sheet.cell(row=fila_header, column=col_nombres)
    for i, h in enumerate(HEADERS_NUEVOS):
        col_actual = insert_idx + i
        celda_nueva = sheet.cell(row=fila_header, column=col_actual)
        celda_nueva.value = h
        copiar_estilo_seguro(celda_origen, celda_nueva)
        sheet.column_dimensions[get_column_letter(col_actual)].width = 19

    # 4. Tablas de Excel
    for table in list(sheet.tables.values()):
        ajustar_tabla(sheet, table, col_nombres, fila_header)

    # 5. Recombinar celdas adaptando las coordenadas
    for min_col, min_row, max_col, max_row in rangos_combinados:
        if min_col >= insert_idx:
            min_col += N_NUEVAS
            max_col += N_NUEVAS
        elif max_col >= insert_idx:
            max_col += N_NUEVAS
        sheet.merge_cells(start_row=min_row, start_column=min_col, end_row=max_row, end_column=max_col)

    # 6. Datos: primero se leen todas las celdas (el motor aprende del archivo completo) y luego se escribe
    filas = []
    for r in range(fila_header + 1, sheet.max_row + 1):
        if r in filas_omitir:
            continue
        val = sheet.cell(row=r, column=col_nombres).value
        if val is None or not str(val).strip():
            continue
        if isinstance(val, str) and val.startswith("="):
            continue                          # fórmulas: no se tocan
        filas.append((r, val))

    resultados = resolver_lista([str(v) for _, v in filas])

    conteo = {"A": 0, "B": 0, "AUTO": 0}
    vista_previa, por_revisar = [], []
    for (r, val), res in zip(filas, resultados):
        conteo[res.origen] += 1
        celda_nombre = sheet.cell(row=r, column=col_nombres)
        for idx, txt in enumerate([res.p_ape, res.s_ape, res.p_nom, res.s_nom]):
            celda = sheet.cell(row=r, column=insert_idx + idx)
            celda.value = txt
            copiar_estilo_seguro(celda_nombre, celda)

        fila_resumen = {
            "Fila": r,
            "Original": str(val),
            "PRIMER APELLIDO": res.p_ape, "SEGUNDO APELLIDO": res.s_ape,
            "PRIMER NOMBRE": res.p_nom, "SEGUNDO NOMBRE": res.s_nom,
        }
        if len(vista_previa) < 10:
            vista_previa.append(fila_resumen)
        if res.revisar:
            por_revisar.append({**fila_resumen, "Criterio": res.motivo})

    return {"registros": len(filas), "conteo": conteo, "vista_previa": vista_previa, "revisar": por_revisar}


# =============================================================================
# 5. ESTILOS (solo estética)
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
# 6. INTERFAZ WEB STREAMLIT
# =============================================================================


def main():
    st.set_page_config(page_title=f"{NOMBRE_APP} · Separador de nombres", layout="centered", page_icon=ICONO_APP)
    st.markdown(CSS_TEMA_OSCURO + ESTILOS, unsafe_allow_html=True)

    st.markdown(
        f"""
        <div class="hero">
            <div class="badge">Herramienta para Excel</div>
            <div class="marca">{LOGO_SVG}<h1>{NOMBRE_APP}</h1></div>
            <div class="lema">{LEMA_APP}</div>
            <p>Divide automáticamente los nombres completos en 4 columnas. Deduce solo, fila por fila,
            si cada nombre viene como Nombres + Apellidos o Apellidos + Nombres, y mantiene intactas
            las tablas de Excel, los formatos y los filtros.</p>
            <div class="chips">
                <span class="chip">✔ Entiende (A) y (B) por celda</span>
                <span class="chip">✔ Deduce el orden sin letras</span>
                <span class="chip">✔ Conserva tablas y filtros</span>
                <span class="chip">✔ Detecta conectores (DE LA, DEL…)</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    paso(1, "Sube tu archivo", "Formato Excel (.xlsx) · no necesitas indicar nada más")
    archivo_subido = st.file_uploader(
        "Sube tu archivo Excel (.xlsx)", type=["xlsx"], label_visibility="collapsed"
    )

    if archivo_subido is not None:
        paso(2, "Resultado", "Revisa y descarga tu archivo procesado")
        with st.spinner("Procesando archivo sin errores de estructura..."):
            try:
                wb = openpyxl.load_workbook(archivo_subido)
                res = procesar_libro(wb)

                output = io.BytesIO()
                wb.save(output)
                output.seek(0)

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
                if n_revisar:
                    st.warning(
                        f"{n_revisar} fila(s) tenían poca evidencia (palabras que no están en los "
                        "diccionarios). El programa las separó con su mejor criterio; conviene mirarlas."
                    )
                    with st.expander("Filas para revisar"):
                        st.dataframe(res["revisar"], hide_index=True)
                if res["vista_previa"]:
                    with st.expander("Vista previa de las primeras filas"):
                        st.dataframe(res["vista_previa"], hide_index=True)
                st.download_button(
                    label="📥 Descargar Archivo Procesado",
                    data=output,
                    file_name=f"OK_{archivo_subido.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except ColumnaNoEncontrada:
                st.error("❌ No se encontró una columna válida de Nombres.")
            except Exception as e:
                st.error(f"Error procesando el archivo: {e}")

    st.markdown(f'<div class="pie">{NOMBRE_APP} · Tus archivos se procesan en memoria y no se almacenan</div>', unsafe_allow_html=True)

    selector_tema()


if __name__ == "__main__":
    main()
