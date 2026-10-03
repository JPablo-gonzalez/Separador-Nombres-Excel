import io
import os
import re
import unicodedata
from copy import copy

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


# Conectores de una sola palabra que se pegan a la(s) palabra(s) siguiente(s).
# Con esto "DE LA CRUZ", "DEL RIO", "SAN JUAN", "VON TRAPP" quedan como un solo bloque.
CONECTORES = {"DE", "DEL", "LA", "LOS", "LAS", "SAN", "SANTA", "VON", "VAN", "DA", "DOS", "DAS", "DI"}

# Bloques religiosos que cierran un nombre compuesto (cuentan como UNA sola palabra).
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
    # ampliación
    "PABLO", "PEDRO", "SANTIAGO", "SEBASTIAN", "NICOLAS", "FELIPE", "MATEO", "SAMUEL", "JESUS",
    "FERNANDO", "RICARDO", "ROBERTO", "OSCAR", "MAURICIO", "FABIAN", "CRISTIAN", "STEVEN", "BRAYAN",
    "KEVIN", "JHONATAN", "YEISON", "ESTEBAN", "EDUARDO", "ALEXIS", "ANDREA", "PAOLA", "CATALINA",
    "VALENTINA", "SOFIA", "ISABELLA", "CAMILA", "LAURA", "ANA", "CLAUDIA", "PATRICIA", "MARTHA",
    "SANDRA", "LUZ", "ELENA", "YESICA", "LEIDY", "NATALIA", "DANIELA", "MARIANA", "JULIANA",
    "PAULA", "ROSA", "LILIANA", "GLORIA", "BEATRIZ", "ADRIANA", "CAROLINA", "JOHAN", "SERGIO",
    "ISAAC", "EMMANUEL", "THOMAS", "SIMON", "JERONIMO", "MARTIN", "VICTOR", "RAFAEL", "ANDERSON",
}

_APELLIDOS_BASE = {
    "GOMEZ", "ZAPATA", "PEREZ", "OSORIO", "VERA", "BETANCUR", "MORALES", "GALEANO", "ESPINOSA",
    "GUARIN", "CELIS", "RAMIREZ", "HERNANDEZ", "TORO", "OCAMPO", "ARROYAVE", "ARANGO", "MUNOZ",
    "LONDONO", "AGUIRRE", "AMESQUITA", "MARIN", "BETANCURT", "TOBON", "GARCIA", "MEJIA", "ARANZAZU",
    "OSPINA", "SANCHEZ", "GAVIRIA", "CANO", "RUIZ", "BARRERA", "GALLO", "RAMOS", "GRAJALES",
    "GRISALES", "BOTERO", "CASTRO", "BARRETO", "ZAMBRANO", "BUITRAGO", "OBANDO", "GALLEGO", "MESA",
    "ARIAS", "CASTANO", "HERRERA", "MOLINA", "CARDONA", "PARRA", "TASCON", "DIAZ", "LOPEZ", "MACIAS",
    # ampliación
    "RODRIGUEZ", "MARTINEZ", "GONZALEZ", "VARGAS", "JIMENEZ", "RESTREPO", "QUINTERO", "FIGUEROA",
    "SUAREZ", "CARDENAS", "ROJAS", "GUTIERREZ", "ORTIZ", "VELEZ", "DUQUE", "VALENCIA", "CORREA",
    "URIBE", "ALZATE", "MORENO", "SALAZAR", "CARDOSO", "MONTOYA", "JARAMILLO", "NARANJO", "PALACIO",
    "PALACIOS", "FRANCO", "HOYOS", "VILLA", "AGUDELO", "VASQUEZ", "TORRES", "FLOREZ", "ESCOBAR",
    "CASTILLO", "ROMERO", "ALVAREZ", "ACEVEDO", "SERNA", "USUGA", "ECHEVERRI", "POSADA", "LOAIZA",
    "HENAO", "MURILLO", "ZULUAGA", "OROZCO", "CARVAJAL", "BEDOYA", "RIVERA", "SIERRA", "PINEDA",
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


def es_religioso(token):
    return _norm(token) in RELIGIOSOS


# =============================================================================
# 2. ORIENTACIÓN POR CELDA Y DIVISIÓN DEL NOMBRE
# =============================================================================


def detectar_orientacion(texto, tipo_archivo, orientacion_defecto):
    """Devuelve (texto_sin_prefijo, orientacion, origen).

    - CON_LETRAS: "(A) " -> A->N, "(B) " -> N->A (y se borra el prefijo, celda por celda).
    - Celda sin prefijo (o archivo SIN_LETRAS): se usa la orientación elegida en la interfaz.
    """
    if orientacion_defecto not in ORIENTACIONES:
        raise ValueError("La orientación por defecto debe ser 'N->A' o 'A->N'.")
    texto = str(texto).strip()
    if tipo_archivo == "CON_LETRAS":
        m = RE_PREFIJO.match(texto)
        if m:
            letra = m.group(1).upper()
            return texto[m.end():].strip(), ("A->N" if letra == "A" else "N->A"), letra
    return texto, orientacion_defecto, "DEFECTO"


def dividir_nombre(texto, orientacion):
    """Reparte las palabras en (primer_apellido, segundo_apellido, primer_nombre, segundo_nombre)."""
    tokens = agrupar_conectores(texto.split())
    n = len(tokens)
    p_ape = s_ape = p_nom = s_nom = ""
    nombres_primero = orientacion == "N->A"

    if n == 0:
        pass

    elif n == 1:
        if nombres_primero:
            p_nom = tokens[0]
        else:
            p_ape = tokens[0]

    elif n == 2:
        if nombres_primero:
            p_nom, p_ape = tokens
        else:
            p_ape, p_nom = tokens

    elif n == 3:
        t1, t2, t3 = tokens
        c2 = _norm(t2)
        if nombres_primero:
            if c2 in APELLIDOS_COMUNES:                  # Julián Mejía Ospina -> 1 nombre + 2 apellidos
                p_nom, p_ape, s_ape = t1, t2, t3
            elif c2 in NOMBRES_COMUNES or es_religioso(t2):  # Juan Pablo Pérez -> 2 nombres + 1 apellido
                p_nom, s_nom, p_ape = t1, t2, t3
            else:                                         # respaldo N->A: 1 nombre + 2 apellidos
                p_nom, p_ape, s_ape = t1, t2, t3
        else:
            if es_religioso(t3) or c2 in NOMBRES_COMUNES:   # Zapata Fredy Abelardo -> 1 apellido + 2 nombres
                p_ape, p_nom, s_nom = t1, t2, t3
            elif c2 in APELLIDOS_COMUNES:                 # Zapata Gómez Fredy -> 2 apellidos + 1 nombre
                p_ape, s_ape, p_nom = t1, t2, t3
            else:                                         # respaldo A->N: 2 apellidos + 1 nombre
                p_ape, s_ape, p_nom = t1, t2, t3

    else:  # 4 o más palabras
        resto = " ".join(tokens[3:])
        if nombres_primero:
            p_nom, s_nom, p_ape, s_ape = tokens[0], tokens[1], tokens[2], resto
        else:
            p_ape, s_ape, p_nom, s_nom = tokens[0], tokens[1], tokens[2], resto

    return p_ape, s_ape, p_nom, s_nom


def analizar_nombre(texto, tipo_archivo, orientacion_defecto):
    """Devuelve (p_ape, s_ape, p_nom, s_nom, origen) donde origen es 'A', 'B' o 'DEFECTO'."""
    limpio, orientacion, origen = detectar_orientacion(texto, tipo_archivo, orientacion_defecto)
    return (*dividir_nombre(limpio, orientacion), origen)


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


def procesar_libro(wb, tipo_archivo, orientacion_defecto):
    """Inserta las 4 columnas junto a la columna de nombres y rellena los datos.
    Devuelve un diccionario con estadísticas y una vista previa."""
    if tipo_archivo not in ("CON_LETRAS", "SIN_LETRAS"):
        raise ValueError("tipo_archivo debe ser 'CON_LETRAS' o 'SIN_LETRAS'.")
    if orientacion_defecto not in ORIENTACIONES:
        raise ValueError("La orientación por defecto debe ser 'N->A' o 'A->N'.")

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

    # 6. Datos fila por fila (la orientación se evalúa celda por celda)
    contador = 0
    conteo = {"A": 0, "B": 0, "DEFECTO": 0}
    vista_previa = []
    for r in range(fila_header + 1, sheet.max_row + 1):
        if r in filas_omitir:
            continue
        celda_nombre = sheet.cell(row=r, column=col_nombres)
        val = celda_nombre.value
        if val is None or not str(val).strip():
            continue
        if isinstance(val, str) and val.startswith("="):
            continue                          # fórmulas: no se tocan

        pa, sa, pn, sn, origen = analizar_nombre(val, tipo_archivo, orientacion_defecto)
        conteo[origen] += 1

        for idx, txt in enumerate([pa, sa, pn, sn]):
            celda = sheet.cell(row=r, column=insert_idx + idx)
            celda.value = txt
            copiar_estilo_seguro(celda_nombre, celda)

        if len(vista_previa) < 10:
            vista_previa.append({
                "Original": str(val),
                "PRIMER APELLIDO": pa, "SEGUNDO APELLIDO": sa,
                "PRIMER NOMBRE": pn, "SEGUNDO NOMBRE": sn,
            })
        contador += 1

    return {"registros": contador, "conteo": conteo, "vista_previa": vista_previa}


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

/*__RADIOS__*/

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
    div[role="radiogroup"] { grid-template-columns: 1fr; }
    div[role="radiogroup"] > label { min-height: 0; }
}
</style>
"""

_PLANTILLA_RADIO = """/* Tarjetas seleccionables (st.radio): __K__ */
.st-key-__K__ div[role="radiogroup"] {
    display: grid !important;
    grid-template-columns: 1fr 1fr;
    gap: .9rem;
}
.st-key-__K__ div[role="radiogroup"] > label {
    position: relative;
    display: flex; flex-direction: column; align-items: flex-start;
    margin: 0 !important;
    min-height: 120px;
    padding: 1.15rem 1.2rem 1.1rem 1.2rem !important;
    background: var(--fondo-card);
    border: 1.5px solid var(--borde);
    border-radius: 18px;
    cursor: pointer;
    transition: all .18s ease;
    box-shadow: 0 10px 24px -18px var(--sombra);
    overflow: hidden;
}
/* ocultar el circulito nativo del radio */
.st-key-__K__ div[role="radiogroup"] > label > *:first-child { display: none !important; }
.st-key-__K__ div[role="radiogroup"] > label input { position: absolute; opacity: 0; pointer-events: none; }

/* insignia de cada tarjeta */
.st-key-__K__ div[role="radiogroup"] > label::before {
    display: inline-block;
    margin-bottom: .75rem;
    padding: .28rem .65rem;
    border-radius: 999px;
    font-size: .74rem; font-weight: 700; letter-spacing: .04em;
    color: var(--acento-texto);
    background: rgba(59,130,246,.12);
    border: 1px solid rgba(59,130,246,.28);
}
.st-key-__K__ div[role="radiogroup"] > label:nth-of-type(1)::before { content: "__B1__"; }
.st-key-__K__ div[role="radiogroup"] > label:nth-of-type(2)::before { content: "__B2__"; }

/* título y descripción */
.st-key-__K__ div[role="radiogroup"] > label p { font-size: 1.02rem; font-weight: 700; color: var(--tinta); margin: 0; }
.st-key-__K__ div[role="radiogroup"] > label p::after {
    display: block; margin-top: .35rem;
    font-size: .82rem; font-weight: 400; line-height: 1.45; color: var(--gris);
}
.st-key-__K__ div[role="radiogroup"] > label:nth-of-type(1) p::after { content: "__D1__"; }
.st-key-__K__ div[role="radiogroup"] > label:nth-of-type(2) p::after { content: "__D2__"; }

/* marca de selección */
.st-key-__K__ div[role="radiogroup"] > label::after {
    content: "";
    position: absolute; top: 14px; right: 14px;
    width: 22px; height: 22px; border-radius: 50%;
    border: 1.5px solid var(--borde);
    background: transparent;
    transition: all .18s ease;
}
.st-key-__K__ div[role="radiogroup"] > label:hover {
    border-color: var(--azul);
    transform: translateY(-2px);
    box-shadow: 0 16px 28px -16px rgba(59,130,246,.55);
}
.st-key-__K__ div[role="radiogroup"] > label:has(input:checked) {
    border-color: var(--azul);
    background: linear-gradient(135deg, rgba(59,130,246,.14), rgba(16,185,129,.12)), var(--fondo-card);
    box-shadow: 0 0 0 3px rgba(59,130,246,.22), 0 16px 28px -16px rgba(59,130,246,.6);
}
.st-key-__K__ div[role="radiogroup"] > label:has(input:checked)::after {
    content: "✓";
    display: flex; align-items: center; justify-content: center;
    color: #fff; font-size: .8rem; font-weight: 800;
    background: linear-gradient(135deg, var(--azul), var(--verde));
    border-color: transparent;
}

"""


def _css_radio(clave, b1, b2, d1, d2):
    return (_PLANTILLA_RADIO.replace("__K__", clave).replace("__B1__", b1).replace("__B2__", b2)
            .replace("__D1__", d1).replace("__D2__", d2))


ESTILOS = ESTILOS.replace(
    "/*__RADIOS__*/",
    _css_radio("tipo_archivo", "(A) · (B)", "Sin prefijo",
               "(A) apellidos primero · (B) nombres primero",
               "Los nombres no llevan letra al inicio")
    + _css_radio("orientacion", "Nombres → Apellidos", "Apellidos → Nombres",
                 "Orden natural. Ej.: Julián Mejía Ospina",
                 "Ej.: Mejía Ospina Julián"),
)


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

    # ---- Paso 1: ¿hay prefijos? + orden estándar (el usuario controla el fallback) ----
    paso(1, "Tipo de archivo", "Indica cómo vienen los datos en la columna de nombres")
    tipo_archivo = st.radio(
        "Formato de la columna de nombres",
        options=["CON_LETRAS", "SIN_LETRAS"],
        format_func=lambda x: "Con letras (A) / (B)" if x == "CON_LETRAS" else "Sin letras",
        label_visibility="collapsed",
        key="tipo_archivo",
    )

    if tipo_archivo == "SIN_LETRAS":
        pregunta = "¿Cuál es la estructura estándar de esta lista?"
    else:
        pregunta = "Si alguna fila no trae (A) ni (B), ¿qué orden debe asumir por defecto?"
    st.markdown(f'<div class="pregunta">{pregunta}</div>', unsafe_allow_html=True)

    # Sin valor inicial: el usuario DEBE elegir. Esta selección alimenta la orientación por defecto.
    orientacion_defecto = st.radio(
        "Orden de los nombres",
        options=list(ORIENTACIONES),
        index=None,
        format_func=lambda x: "Nombres primero, luego apellidos" if x == "N->A" else "Apellidos primero, luego nombres",
        label_visibility="collapsed",
        key="orientacion",
    )

    # ---- Paso 2: archivo ----
    paso(2, "Sube tu archivo", "Formato Excel (.xlsx)")
    archivo_subido = st.file_uploader(
        "Sube tu archivo Excel (.xlsx)", type=["xlsx"], label_visibility="collapsed"
    )

    # ---- Paso 3: resultado ----
    if archivo_subido is not None:
        paso(3, "Resultado", "Revisa y descarga tu archivo procesado")
        if orientacion_defecto is None:
            st.warning("⚠️ Antes de procesar, elige en el paso 1 el orden en que vienen los nombres en tu lista.")
        else:
            with st.spinner("Procesando archivo sin errores de estructura..."):
                try:
                    wb = openpyxl.load_workbook(archivo_subido)
                    res = procesar_libro(wb, tipo_archivo, orientacion_defecto)

                    output = io.BytesIO()
                    wb.save(output)
                    output.seek(0)

                    if tipo_archivo == "CON_LETRAS":
                        modo = "(A)/(B)"
                        c = res["conteo"]
                        detalle = (f"Filas con (A): {c['A']} · con (B): {c['B']} · "
                                   f"sin prefijo ({'N → A' if orientacion_defecto == 'N->A' else 'A → N'}): {c['DEFECTO']}")
                    else:
                        modo = "N → A" if orientacion_defecto == "N->A" else "A → N"
                        detalle = "Se separaron los nombres sin dañar filtros ni formatos."

                    st.markdown(
                        f"""
                        <div class="resultado">
                            <h3>✅ ¡Proceso completado!</h3>
                            <p>{detalle}</p>
                            <div class="stats">
                                <div class="stat"><div class="valor">{res['registros']}</div><div class="etq">Registros</div></div>
                                <div class="stat"><div class="valor">4</div><div class="etq">Columnas nuevas</div></div>
                                <div class="stat"><div class="valor">{modo}</div><div class="etq">Modo</div></div>
                            </div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
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
