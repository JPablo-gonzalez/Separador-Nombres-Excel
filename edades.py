"""Edades y menores de edad (solo para el resumen en pantalla: el Excel de salida no cambia).

Funciones puras (sin Excel ni Streamlit): reconocer el encabezado de la fecha de nacimiento,
interpretar la fecha tal como viene en la celda (fecha de Excel, texto "27/06/1953", "27 de junio de
1953"...), calcular la edad exacta y contar cuántas personas están por debajo de la edad elegida.
"""

import math
import re
import unicodedata
from datetime import date, datetime, timedelta, timezone

EDAD_LIMITE_DEFECTO = 18
EDAD_MAXIMA_VALIDA = 120          # una edad mayor delata una fecha mal escrita

# Colombia no tiene horario de verano: "hoy" es siempre la fecha en UTC-5, sin depender del reloj
# del servidor (Streamlit Cloud está en UTC) ni del navegador.
_ZONA_COLOMBIA = timezone(timedelta(hours=-5))


def fecha_hoy():
    return datetime.now(_ZONA_COLOMBIA).date()


# =============================================================================
# 1. ENCABEZADO
# =============================================================================

def _norm(texto):
    t = unicodedata.normalize("NFD", str(texto).upper())
    return "".join(c for c in t if unicodedata.category(c) != "Mn")


_NO_ES_FECHA = {"LUGAR", "CIUDAD", "MUNICIPIO", "PAIS", "DEPARTAMENTO", "SITIO", "PARTIDA", "REGISTRO"}


def puntaje_encabezado_fecha(texto):
    """0 si el texto no parece el encabezado de la fecha de nacimiento; más alto cuanto más claro.

    Acepta "FECHA DE NACIMIENTO", "FECHA NACIMIENTO", "F. NACIMIENTO", "FECHA NAC.", "FEC_NAC",
    "NACIMIENTO"... y descarta "LUGAR DE NACIMIENTO" o "CIUDAD DE NACIMIENTO"."""
    palabras = re.sub(r"[^A-Z0-9]+", " ", _norm(texto)).split()
    if not palabras or _NO_ES_FECHA.intersection(palabras):
        return 0
    nacimiento = any(p.startswith("NACIM") for p in palabras)
    fecha = any(p in ("FECHA", "FEC", "FCH", "F") for p in palabras)
    if nacimiento and fecha:
        return 3
    if fecha and "NAC" in palabras:
        return 2
    if nacimiento:
        return 1
    return 0


# =============================================================================
# 2. INTERPRETACIÓN DE LA FECHA
# =============================================================================

_MESES = {
    "ENE": 1, "JAN": 1, "FEB": 2, "MAR": 3, "ABR": 4, "APR": 4, "MAY": 5, "JUN": 6, "JUL": 7,
    "AGO": 8, "AUG": 8, "SEP": 9, "SET": 9, "OCT": 10, "NOV": 11, "DIC": 12, "DEC": 12,
}
_RE_HORA = re.compile(r"[\sT]\d{1,2}:\d{2}.*$")


def _crear(anio, mes, dia):
    try:
        return date(anio, mes, dia)
    except ValueError:
        return None


def _anio_completo(texto, hoy):
    anio = int(texto)
    if len(texto) <= 2:                       # "53" -> 1953, "12" -> 2012 (nunca en el futuro)
        anio += 2000 if 2000 + anio <= hoy.year else 1900
    return anio


def _desde_serial(serial, fecha1904):
    """Número de serie de Excel -> fecha (sistema 1900, o 1904 si el libro lo usa)."""
    dias = math.floor(serial)
    if fecha1904:
        base = date(1904, 1, 1)
    elif dias < 60:                           # Excel cuenta un 29/02/1900 que no existió
        base = date(1899, 12, 31)
    elif dias == 60:
        return None
    else:
        base = date(1899, 12, 30)
    try:
        return base + timedelta(days=dias)
    except OverflowError:
        return None


def _desde_numero(valor, fecha1904):
    if not math.isfinite(valor) or valor < 1:
        return None
    if valor == int(valor) and 10000000 <= valor <= 99999999:      # 19530627 o 27061953 escritos sin barras
        t = str(int(valor))
        return _crear(int(t[:4]), int(t[4:6]), int(t[6:])) or _crear(int(t[4:]), int(t[2:4]), int(t[:2]))
    if valor > 2958465:                                           # después del 31/12/9999
        return None
    return _desde_serial(valor, fecha1904)


def _desde_texto(texto, hoy, fecha1904):
    t = _norm(texto).strip()
    t = _RE_HORA.sub("", t).strip()           # "1953-06-27 00:00:00", "27/06/1953 12:00 A. M."
    if not t:
        return None
    if re.fullmatch(r"\d+(\.\d+)?", t):       # número guardado como texto
        return _desde_numero(float(t), fecha1904)

    m = re.fullmatch(r"(\d{4})[/\-. ](\d{1,2})[/\-. ](\d{1,2})", t)          # 1953-06-27
    if m:
        return _crear(int(m.group(1)), int(m.group(2)), int(m.group(3)))

    m = re.fullmatch(r"(\d{1,2})\s*[/\-. ]\s*(\d{1,2})\s*[/\-. ]\s*(\d{4}|\d{2})", t)   # 27/06/1953
    if m:
        a, b = int(m.group(1)), int(m.group(2))
        anio = _anio_completo(m.group(3), hoy)
        if a > 12 or b <= 12:                 # día primero, como se escribe en Colombia
            return _crear(anio, b, a)
        return _crear(anio, a, b)             # 06/27/1953: solo puede ser mes/día

    # Mes escrito con letras: "27 de junio de 1953", "27-JUN-1953", "junio 27 de 1953"
    palabras = re.findall(r"[A-Z]+|\d+", t)
    meses = [_MESES[p[:3]] for p in palabras if p.isalpha() and len(p) >= 3 and p[:3] in _MESES]
    numeros = [p for p in palabras if p.isdigit()]
    if len(meses) == 1 and len(numeros) == 2:
        dia_txt, anio_txt = (numeros if len(numeros[1]) >= len(numeros[0]) else numeros[::-1])
        if len(dia_txt) <= 2:
            return _crear(_anio_completo(anio_txt, hoy), meses[0], int(dia_txt))
    return None


def interpretar_fecha(valor, hoy, fecha1904=False):
    """Fecha de nacimiento a partir del valor de la celda: número (fecha de Excel) o texto.
    Devuelve None si no se puede interpretar."""
    if valor is None:
        return None
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    if isinstance(valor, (int, float)):
        return _desde_numero(float(valor), fecha1904)
    return _desde_texto(str(valor), hoy, fecha1904)


# =============================================================================
# 3. EDAD Y CONTEO
# =============================================================================

def calcular_edad(nacimiento, hoy):
    """Años cumplidos a la fecha `hoy` (quien nació un 29 de febrero los cumple el 1 de marzo)."""
    return hoy.year - nacimiento.year - ((hoy.month, hoy.day) < (nacimiento.month, nacimiento.day))


def edades_validas(valores, hoy, fecha1904=False):
    """Edad de cada persona a partir del valor de su celda de fecha.

    Devuelve (edades, sin_fecha, no_validas): la lista de edades que se pudieron calcular, cuántas
    personas no tienen fecha y cuántas tienen algo escrito que no es una fecha válida (no se
    entiende, está en el futuro o daría más de EDAD_MAXIMA_VALIDA años)."""
    lista, sin_fecha, no_validas = [], 0, 0
    for valor in valores:
        if valor is None or (isinstance(valor, str) and not valor.strip()):
            sin_fecha += 1
            continue
        nacimiento = interpretar_fecha(valor, hoy, fecha1904)
        edad = calcular_edad(nacimiento, hoy) if nacimiento and nacimiento <= hoy else None
        if edad is None or edad > EDAD_MAXIMA_VALIDA:
            no_validas += 1
        else:
            lista.append(edad)
    return lista, sin_fecha, no_validas


def contar_menores(edades, limite):
    """Cuántas de las edades son menores que la edad límite (17 años y 364 días cuenta como menor de 18)."""
    return sum(1 for e in edades if e < limite)
