"""Compara dos .xlsx parte por parte (contenido descomprimido y metadatos de cada entrada del ZIP).

Uso: python pruebas/comparar_xlsx.py a.xlsx b.xlsx
Sale con código 0 si todas las partes son idénticas byte a byte.
"""

import sys
import zipfile


def partes(ruta):
    with zipfile.ZipFile(ruta) as z:
        return {i.filename: (z.read(i), i.date_time, i.compress_type, i.external_attr) for i in z.infolist()}, \
               [i.filename for i in z.infolist()]


def comparar(a, b):
    pa, oa = partes(a)
    pb, ob = partes(b)
    difs = []
    if oa != ob:
        difs.append(f"orden o lista de partes distinta: {sorted(set(oa) ^ set(ob)) or 'mismo conjunto, distinto orden'}")
    for nombre in oa:
        if nombre not in pb:
            continue
        for campo, x, y in zip(("contenido", "fecha", "compresión", "atributos"), pa[nombre], pb[nombre]):
            if x != y:
                difs.append(f"{nombre}: {campo} distinto")
    return difs


if __name__ == "__main__":
    d = comparar(sys.argv[1], sys.argv[2])
    print("IDÉNTICOS" if not d else "\n".join(d))
    sys.exit(1 if d else 0)
