"""Desglosa · Separador inteligente de nombres (punto de entrada de Streamlit).

Contiene solo el flujo principal: recibe el archivo, invoca el motor de Excel y muestra el
resultado con los componentes visuales de styles.py.
"""

import streamlit as st

from excel_engine import LIMITE_SUBIDA_MB, ArchivoNoSoportado, ColumnaNoEncontrada, procesar_xlsx
from styles import configurar_pagina, encabezado, paso, pie, selector_tema, tarjeta_resultado

# =============================================================================
# INTERFAZ WEB STREAMLIT (flujo principal)
# =============================================================================


@st.cache_data(max_entries=2, ttl=900, show_spinner=False)
def _procesar_cacheado(datos):
    """Pulsar "Descargar" vuelve a ejecutar el script: el resultado se guarda para no reprocesar el archivo."""
    return procesar_xlsx(datos)


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
            with st.spinner("Procesando el archivo sin alterar su formato..."):
                try:
                    salida, res = _procesar_cacheado(archivo_subido.getvalue())

                    n_revisar = len(res["revisar"])
                    tarjeta_resultado(res)
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
                    )
                except (ColumnaNoEncontrada, ArchivoNoSoportado) as e:
                    st.error(f"❌ {e}")
                except Exception as e:
                    st.error(f"Error procesando el archivo: {e}")

    pie()

    selector_tema()


if __name__ == "__main__":
    main()
