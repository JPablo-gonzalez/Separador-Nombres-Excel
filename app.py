"""Desglosa · Separador inteligente de nombres (punto de entrada de Streamlit).

Contiene solo el flujo principal: recibe el archivo, invoca el motor de Excel y muestra el
resultado con los componentes visuales de styles.py.
"""

import gc

import streamlit as st

from edades import EDAD_LIMITE_DEFECTO, contar_menores, edades_validas, fecha_hoy
from excel_engine import LIMITE_SUBIDA_MB, ArchivoNoSoportado, ColumnaNoEncontrada, procesar_xlsx
from styles import configurar_pagina, encabezado, paso, pie, selector_tema, tarjeta_edades, tarjeta_resultado

# =============================================================================
# INTERFAZ WEB STREAMLIT (flujo principal)
# =============================================================================


@st.fragment
def consulta_edades(edades, total, sin_fecha, no_validas, hoy):
    # Fragmento: cambiar la edad límite solo recalcula el conteo; no vuelve a procesar el archivo.
    # Recibe solo números (edades), no nombres ni fechas.
    limite = st.number_input("¿Menores de cuántos años quieres contar?", min_value=1, max_value=120,
                             value=EDAD_LIMITE_DEFECTO, step=1, key="edad_limite")
    tarjeta_edades(contar_menores(edades, limite), total, limite, sin_fecha, no_validas, hoy)


def resumen_edades(res):
    """Conteo de menores de edad en pantalla. El Excel de salida no cambia."""
    fechas = res["fechas_nacimiento"]
    if fechas is None:
        st.info("No se encontró una columna de FECHA DE NACIMIENTO, así que no se pueden contar los menores de edad.")
        return
    hoy = fecha_hoy()
    edades, sin_fecha, no_validas = edades_validas(fechas["valores"], hoy, fechas["fecha1904"])
    consulta_edades(edades, res["registros"], sin_fecha, no_validas, hoy)


def main():
    configurar_pagina()

    encabezado()

    paso(1, "Sube tu archivo", "Formato Excel (.xlsx) · no necesitas indicar nada más")
    archivo_subido = st.file_uploader(
        "Sube tu archivo Excel (.xlsx)", type=["xlsx"], label_visibility="collapsed"
    )

    if archivo_subido is not None:
        paso(2, "Resultado", "Revisa y descarga tu archivo procesado")
        if archivo_subido.size > LIMITE_SUBIDA_MB * 1024 * 1024:
            st.error(f"❌ El archivo pesa más de {LIMITE_SUBIDA_MB} MB. Divídelo en partes más pequeñas e inténtalo de nuevo.")
        else:
            # Privacidad: el resultado no se guarda en ninguna caché (ni de Streamlit ni propia). Vive solo
            # durante esta ejecución y en el botón de descarga de esta sesión; desaparece al quitar el archivo
            # o al cerrar la pestaña.
            salida = res = None
            with st.spinner("Procesando el archivo sin alterar su formato..."):
                try:
                    salida, res = procesar_xlsx(archivo_subido.getvalue())

                    n_revisar = len(res["revisar"])
                    tarjeta_resultado(res)
                    resumen_edades(res)
                    for aviso in res["advertencias"]:
                        st.info(aviso)
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
                        data=salida,
                        file_name=f"OK_{archivo_subido.name}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        on_click="ignore",  # descargar no vuelve a ejecutar el script ni a procesar el archivo
                    )
                except (ColumnaNoEncontrada, ArchivoNoSoportado) as e:
                    st.error(f"❌ {e}")  # mensajes fijos, sin datos del archivo
                except Exception as e:
                    # El texto de una excepción puede incluir contenido del archivo (celdas, fórmulas, rutas
                    # internas): solo se muestra su tipo, y no se registra en ningún log.
                    st.error(f"No se pudo procesar el archivo ({type(e).__name__}). Revisa que sea un .xlsx válido.")
                finally:
                    # Suelta las referencias locales a los datos para que Python libere esa memoria ya.
                    del salida, res
                    gc.collect()

    pie()

    selector_tema()


if __name__ == "__main__":
    main()
