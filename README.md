# Bitácora de Issues

## Ejecutar la app localmente

1. Abre la carpeta del proyecto.
2. Instala las dependencias con `pip install -r requirements.txt`.
3. Doble clic en start.bat.
4. Abre en el navegador:
   - http://127.0.0.1:8000/index.html

Pillow permite insertar las evidencias como imágenes en el archivo Excel generado por el servidor local.
En GitHub Pages, el Excel se genera directamente en el navegador; ambos modos respetan los filtros activos y crean las hojas `Bitácora` y `Resumen`.
Las imágenes nuevas se guardan con un máximo de 2200 píxeles por lado y calidad JPEG del 92 %, y se insertan en Excel con su proporción original. Las evidencias ya guardadas no recuperan los detalles perdidos antes de este cambio.
Cada issue puede tener un documento Word `.doc` o `.docx` de hasta 3 MB. El archivo queda disponible para descarga desde la tabla, se puede reemplazar o quitar al editar el issue y se incluye en los respaldos JSON; el Excel conserva sus columnas y evidencias visuales actuales.
Los campos de título, descripción, novedades, comentarios y nombre de módulo incluyen un corrector ortográfico local de español, con sugerencias e ignorado de palabras. El texto se revisa en el navegador y no se envía a servicios externos; el diccionario va incluido en el archivo local `spellcheck.bundle.js`.
Para reconstruir el corrector después de cambiar su código fuente, instala las dependencias con `npm install` y ejecuta `npm run build:spellcheck`. El diccionario español se distribuye bajo MPL 1.1 o posterior; revisa `THIRD_PARTY_NOTICES.md`.
La tabla muestra la antigüedad de cada issue abierto y alerta cuando supera el límite configurable (7 días de forma predeterminada).

## Publicar en GitHub Pages

La app ya está preparada para funcionar como sitio estático.

1. Sube esta carpeta a un repositorio de GitHub.
2. En GitHub, entra a Settings > Pages.
3. En Source, selecciona Deploy from a branch.
4. Elige la rama principal y la carpeta /root.
5. Guarda y espera a que se publique.

Tu enlace será algo como:
- https://<tu-usuario>.github.io/<nombre-del-repositorio>/

## Datos

- En GitHub Pages, los datos se guardan en el navegador del usuario mediante almacenamiento local.
- En modo local, también puedes usar el servidor incluido en server.py, que guarda los datos en data/storage.json.
