"""Prueba de paridad en un navegador real (Chromium con Playwright).

Abre una versión de la app (Streamlit normal o la página stlite), sube cada Excel, espera el resultado, lo descarga
y guarda: tiempos, textos de la pantalla, mensajes de error, capturas de pantalla y todas las peticiones de red.

Uso:
    python pruebas/paridad_navegador.py <nombre> <url> <carpeta_excels> <archivo1.xlsx,archivo2.xlsx,...> [--cortar-red]

    <nombre>       carpeta de resultados (pruebas/resultados/<nombre>)
    --cortar-red   tras cargar la página, bloquea toda petición salvo el código estático de stlite: si la app aún
                   procesa y descarga, el Excel no necesitó salir del navegador.

Variables opcionales para probar sin acceso a cdn.jsdelivr.net (sirve copias locales con las mismas URLs):
    STLITE_LOCAL   carpeta "build" del paquete npm @stlite/browser@1.9.2
    PYODIDE_LOCAL  carpeta de la distribución completa de Pyodide 0.29.3
    CHROMIUM       ruta de un ejecutable de Chromium
    PRUEBA_PROXY   proxy HTTPS por el que salir (entornos sin conexión directa); PyPI se pide a través de él
"""

import asyncio
import json
import os
import sys
import time
from pathlib import Path

from playwright.async_api import async_playwright

CDN_STLITE = "https://cdn.jsdelivr.net/npm/@stlite/browser@1.9.2/build/"
CDN_PYODIDE = "https://cdn.jsdelivr.net/pyodide/v0.29.3/full/"
TIPOS = {".js": "text/javascript", ".mjs": "text/javascript", ".css": "text/css", ".wasm": "application/wasm",
         ".json": "application/json"}


async def main(nombre, url, carpeta, archivos, cortar_red):
    out = Path(__file__).parent / "resultados" / nombre
    out.mkdir(parents=True, exist_ok=True)
    espejos = [(CDN_STLITE, os.environ.get("STLITE_LOCAL")), (CDN_PYODIDE, os.environ.get("PYODIDE_LOCAL"))]
    info, peticiones, bloqueadas, fase = {"url": url, "archivos": {}}, [], [], {"v": "carga"}

    async with async_playwright() as p:
        proxy = os.environ.get("PRUEBA_PROXY")
        navegador = await p.chromium.launch(executable_path=os.environ.get("CHROMIUM"),
                                            proxy={"server": proxy, "bypass": "localhost,127.0.0.1"} if proxy else None)
        ctx = await navegador.new_context(viewport={"width": 1280, "height": 900}, accept_downloads=True,
                                          ignore_https_errors=bool(proxy))
        if proxy:
            async def pypi(route):  # el proxy quita la cabecera CORS que PyPI sí envía
                r = await route.fetch()
                await route.fulfill(response=r, headers={**r.headers, "access-control-allow-origin": "*"})
            await ctx.route("https://pypi.org/**", pypi)
            await ctx.route("https://files.pythonhosted.org/**", pypi)

        async def espejo(route):
            u = route.request.url
            for prefijo, local in espejos:
                if local and u.startswith(prefijo):
                    f = Path(local) / u[len(prefijo):].split("?")[0]
                    return await route.fulfill(status=200, body=f.read_bytes(), headers={
                        "content-type": TIPOS.get(f.suffix, "application/octet-stream"), "access-control-allow-origin": "*"})
            await route.fallback()

        await ctx.route("https://cdn.jsdelivr.net/**", espejo)
        ctx.on("request", lambda r: peticiones.append({
            "fase": fase["v"], "metodo": r.method, "url": r.url, "bytes_enviados": len(r.post_data_buffer or b"")}))
        page = await ctx.new_page()

        t0 = time.time()
        await page.goto(url)
        await page.wait_for_selector('[data-testid="stFileUploader"] section button', timeout=300_000)
        info["carga_s"] = round(time.time() - t0, 2)
        await page.wait_for_timeout(2000)
        await page.screenshot(path=str(out / "01_inicio.png"))

        if cortar_red:
            async def cortar(route):
                if route.request.method == "GET" and route.request.url.startswith(CDN_STLITE):
                    return await espejo(route)
                bloqueadas.append(f"{route.request.method} {route.request.url[:120]}")
                await route.abort()
            await ctx.unroute("https://cdn.jsdelivr.net/**")
            await ctx.route("**/*", cortar)

        for i, archivo in enumerate(archivos):
            fase["v"] = f"proceso:{archivo}"
            if i:
                quitar = await page.query_selector('[data-testid="stFileChipDeleteBtn"] button')
                if quitar:
                    await quitar.click()
                    await page.wait_for_selector(".stDownloadButton", state="detached")
            t1 = time.time()
            await page.set_input_files('input[type="file"]', str(Path(carpeta) / archivo))
            await page.wait_for_timeout(300)
            await page.wait_for_selector('[data-testid="stSpinner"]', state="detached", timeout=600_000)
            await page.wait_for_function(  # resultado, aviso de error o error del propio cuadro de subida
                """() => document.querySelector('.stDownloadButton button, [data-testid="stAlert"]')
                      || /Error/.test(document.querySelector('[data-testid="stFileUploader"]')?.innerText || '')""",
                timeout=600_000)
            res = {"proceso_s": round(time.time() - t1, 2)}
            await page.wait_for_timeout(1000)
            res["mensajes"] = await page.eval_on_selector_all(
                '[data-testid="stAlert"], [data-testid="stFileUploader"] [role="alert"]', "els => els.map(e => e.innerText)")
            res["texto"] = await page.inner_text('[data-testid="stMainBlockContainer"]')
            if await page.query_selector(".stDownloadButton button"):
                fase["v"] = f"descarga:{archivo}"
                t2 = time.time()
                async with page.expect_download() as d:
                    await page.click(".stDownloadButton button")
                descarga = await d.value
                await descarga.save_as(str(out / descarga.suggested_filename))
                res.update(descarga_s=round(time.time() - t2, 2), nombre_descarga=descarga.suggested_filename)
            if i == 0:
                await page.screenshot(path=str(out / "02_resultado.png"))
                for exp in await page.query_selector_all('[data-testid="stExpander"] summary'):
                    await exp.click()
                    await page.wait_for_timeout(800)
                await page.screenshot(path=str(out / "03_expanders.png"))
                await page.click(".st-key-selector_tema button")
                await page.wait_for_timeout(1500)
                await page.screenshot(path=str(out / "04_claro.png"))
                await page.click(".st-key-selector_tema button")
                await page.wait_for_timeout(1000)
            info["archivos"][archivo] = res
        info["titulo"] = await page.title()
        await navegador.close()

    info["peticiones"], info["bloqueadas"] = peticiones, bloqueadas
    info["peticiones_que_envian_datos"] = [r for r in peticiones if r["metodo"] not in ("GET", "HEAD") or r["bytes_enviados"]]
    (out / "info.json").write_text(json.dumps(info, indent=1, ensure_ascii=False))
    for archivo, res in info["archivos"].items():
        print(archivo, res["proceso_s"], "s", res.get("nombre_descarga") or res["mensajes"])
    print("Peticiones que envían datos:", len(info["peticiones_que_envian_datos"]), "· bloqueadas:", bloqueadas)


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    asyncio.run(main(args[0], args[1], args[2], args[3].split(","), "--cortar-red" in sys.argv))
