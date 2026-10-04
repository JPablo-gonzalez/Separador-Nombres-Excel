# Desglosa · Separador inteligente de nombres

Aplicación web hecha con [Streamlit](https://streamlit.io) que recibe un archivo de Excel (`.xlsx`), encuentra la columna con los nombres completos e inserta a su derecha cuatro columnas nuevas: **PRIMER APELLIDO**, **SEGUNDO APELLIDO**, **PRIMER NOMBRE** y **SEGUNDO NOMBRE**. El orden de cada nombre se deduce solo, fila por fila, y el documento conserva su formato original.

Aplicación publicada: <https://separador-nombres-excel.streamlit.app/>

## Características principales

- **Caso (A):** si la celda empieza con `(A)`, el texto viene como Apellidos → Nombres. El prefijo se quita del resultado.
- **Caso (B):** si la celda empieza con `(B)`, el texto viene como Nombres → Apellidos. El prefijo se quita del resultado.
- **Sin letras:** las celdas sin prefijo se resuelven solas, sin preguntar nada al usuario. El programa compara todas las lecturas posibles usando diccionarios de nombres y apellidos comunes, terminaciones típicas de apellido (-EZ, -IZ, -OZ, -AZ) y las palabras que las filas más claras del mismo archivo le enseñan. Las filas con poca evidencia aparecen en una tabla "Filas para revisar".
- **Archivos mixtos:** un mismo archivo puede combinar celdas con (A), con (B) y sin letra.
- **Conectores:** "DE LA", "DEL", "DE", "SAN", "VON" y similares se unen a la palabra siguiente ("DE LA CRUZ"), y los bloques religiosos como "DE JESUS" o "DEL CARMEN" cuentan como una sola palabra.
- **Preservación del formato de Excel:** se conservan imágenes, logos, colores, estilos, celdas combinadas, tablas, filtros, fórmulas, formatos condicionales, validaciones y nombres definidos. Las tablas y filtros se amplían para incluir las columnas nuevas.
- **Archivos pesados:** el `.xlsx` se edita por dentro, sin cargar el libro completo en memoria, así que hojas con mucho formato o con celdas perdidas en filas lejanas no saturan el servidor.

## Arquitectura

| Archivo | Responsabilidad |
|---|---|
| `app.py` | Punto de entrada y controlador. Contiene solo el flujo principal de Streamlit: subir el archivo, procesarlo y mostrar el resultado. |
| `styles.py` | Estética e interfaz visual: nombre y logo de la app, paletas de color (tema oscuro y claro), CSS, encabezado, pasos, tarjeta de resultado, pie de página y botón de cambio de tema. |
| `parser.py` | Motor semántico de nombres: diccionarios de nombres y apellidos, conectores, agrupación de conectores y el motor de decisión (`resolver_lista` y `analizar_nombre`). |
| `excel_engine.py` | Motor de Excel: búsqueda del encabezado, inserción de las 4 columnas, desplazamiento de referencias y fórmulas, celdas combinadas, tablas, filtros, dibujos e imágenes (`procesar_xlsx`). |
| `requirements.txt` | Dependencias de Python. |
| `.streamlit/config.toml` | Tema base y límite de subida de Streamlit. |
| `favicon.png` | Icono de la pestaña del navegador. |

Las dependencias van en un solo sentido: `app.py` usa `styles.py` y `excel_engine.py`, y `excel_engine.py` usa `parser.py`.

## Instalación y ejecución local

Requiere Python 3.10 o superior.

```bash
git clone https://github.com/JPablo-gonzalez/Separador-Nombres-Excel.git
cd Separador-Nombres-Excel
pip install -r requirements.txt
streamlit run app.py
```

La aplicación se abre en <http://localhost:8501>.

## Despliegue en Streamlit Community Cloud

1. Sube el repositorio a GitHub.
2. Entra a <https://share.streamlit.io> y elige **Create app**.
3. Selecciona el repositorio, la rama `main` y el archivo principal `app.py`.
4. En **Advanced settings**, elige Python 3.10 o superior.
5. Pulsa **Deploy**. Streamlit instala `requirements.txt` y vuelve a desplegar solo cada vez que cambia `main`.
