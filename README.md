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
| `.streamlit/config.toml` | Tema base, límite de subida y telemetría de Streamlit desactivada. |
| `favicon.png` | Icono de la pestaña del navegador. |
| `index.html` | Solo para la versión stlite: página estática que ejecuta la misma `app.py` dentro del navegador (ver más abajo). |
| `.nojekyll` | Hace que GitHub Pages publique también `.streamlit/config.toml` (sin él, oculta las carpetas que empiezan por punto). |
| `pruebas/` | Scripts de las pruebas de paridad entre la versión de Streamlit Cloud y la versión stlite. La app no los usa. |

Las dependencias van en un solo sentido: `app.py` usa `styles.py` y `excel_engine.py`, y `excel_engine.py` usa `parser.py`.

## Privacidad de los archivos

Qué hace el programa con los datos (revisado en el código, no solo prometido):

- **Sin disco:** el `.xlsx` se recibe y se procesa en memoria (`io.BytesIO`). El código no crea archivos, ni siquiera temporales, y Streamlit guarda los archivos subidos en memoria, no en disco.
- **Sin cachés:** no se usa `st.cache_data` ni `st.cache_resource`. El resultado existe solo mientras se muestra en la sesión de quien lo subió: se descarta al quitar el archivo con la ✕, al subir otro o al cerrar la pestaña (Streamlit borra la sesión a partir de unos 2 minutos después de desconectarse).
- **Sin servicios externos:** no hay peticiones HTTP, APIs, bases de datos ni analítica en el código. La telemetría de Streamlit está desactivada (`gatherUsageStats = false`) y la página no carga fuentes de Google.
- **Sin logs con datos:** el programa no imprime nada. Los errores inesperados muestran solo el tipo de error, nunca el contenido de una celda.

Límites que no dependen del código:

- **En Streamlit Community Cloud el archivo sí sale del computador.** Viaja cifrado (HTTPS) hasta el servidor de Streamlit, se procesa en la memoria de ese servidor y el resultado vuelve al navegador. El código no lo guarda, pero la infraestructura (proxies, registros de acceso, volcados de memoria o swap del proveedor) no está bajo nuestro control.
- **Python no puede borrar la RAM de forma segura:** al soltar un dato, la memoria se marca libre y se reutiliza, pero sus bytes no se sobrescriben al instante. Lo que se garantiza es que el programa no conserva referencias a los datos.

Para que el archivo no salga del computador, usa la versión stlite (sección siguiente) o ejecuta la aplicación localmente.

## Versión stlite (el Excel se procesa dentro del navegador)

`index.html` carga [stlite](https://github.com/whitphx/stlite) 1.9.2, que trae Streamlit 1.62.0 y Python 3.13 compilado para el navegador (Pyodide 0.29.3), y ejecuta los mismos archivos `app.py`, `styles.py`, `parser.py`, `excel_engine.py`, `favicon.png` y `.streamlit/config.toml`, leídos del mismo sitio. No hay una segunda copia del programa: cualquier cambio en esos archivos llega también a esta versión.

Qué descarga el navegador al abrir la página (solo código, nunca datos del usuario):

- stlite y Pyodide desde `cdn.jsdelivr.net`;
- algunos paquetes de Python (openpyxl, protobuf y otros que Streamlit necesita) desde `pypi.org` y `files.pythonhosted.org`;
- los archivos de la app desde el propio sitio.

Son unos 45 MB sin comprimir (unos 27 MB por la red) la primera vez; después el navegador los guarda en caché. Esos servicios ven, como cualquier web, la IP y que se abrió la página, pero no el Excel.

Qué pasa con el Excel: se lee, se procesa y se descarga dentro de la pestaña. Comprobado en un navegador real: al subir, procesar y descargar no sale ninguna petición que envíe datos, y con toda la red bloqueada después de cargar la página la app procesa y descarga igual, con el mismo resultado byte a byte.

Diferencias conocidas con la versión de Streamlit Cloud:

- Al abrir la página aparece durante unos segundos la pantalla de carga de stlite mientras se prepara Python. Los avisos técnicos de stlite (en inglés) están ocultos y en su lugar se muestra "Preparando la aplicación…"; si la carga falla, el error de stlite sí se muestra.
- El procesamiento corre en el computador de quien usa la app; con archivos muy grandes es más lento que en el servidor (60.000 filas: unos 25 s frente a unos 13 s) y usa más memoria del navegador.
- `index.html` añade, además del aviso de carga, una única regla de CSS de compatibilidad: stlite cambia la prioridad de una regla global de Streamlit y la caja de subida quedaba 5 px más baja con un archivo cargado; la regla devuelve el valor original.

Publicación en GitHub Pages: en **Settings → Pages**, elige **Deploy from a branch**, la rama `stlite` y la carpeta `/ (root)`. La página queda en `https://<usuario>.github.io/Separador-Nombres-Excel/`.

Para probarla en tu equipo sin publicarla, sirve la carpeta con cualquier servidor estático, por ejemplo `python -m http.server 8000`, y abre <http://localhost:8000/>.

## Instalación y ejecución local

Requiere Python 3.10 o superior.

```bash
git clone https://github.com/JPablo-gonzalez/Separador-Nombres-Excel.git
cd Separador-Nombres-Excel
pip install -r requirements.txt
streamlit run app.py --server.address localhost
```

La aplicación se abre en <http://localhost:8501>. `--server.address localhost` hace que solo tu computador pueda abrirla; sin esa opción, Streamlit también acepta conexiones de otros equipos de tu red. En modo local los archivos no salen de tu equipo: la única conexión es la del navegador con el propio programa.

## Despliegue en Streamlit Community Cloud

1. Sube el repositorio a GitHub.
2. Entra a <https://share.streamlit.io> y elige **Create app**.
3. Selecciona el repositorio, la rama `main` y el archivo principal `app.py`.
4. En **Advanced settings**, elige Python 3.10 o superior.
5. Pulsa **Deploy**. Streamlit instala `requirements.txt` y vuelve a desplegar solo cada vez que cambia `main`.
