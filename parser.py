"""Motor semántico de nombres.

Diccionarios de nombres y apellidos, agrupación de conectores ("DE LA", "DEL", "DE JESUS"...) y el
motor de decisión que deduce, celda por celda, la orientación (Nombres -> Apellidos o
Apellidos -> Nombres) y el reparto de palabras en las 4 columnas.
"""

import math
import re
import unicodedata
from collections import Counter
from dataclasses import dataclass

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
