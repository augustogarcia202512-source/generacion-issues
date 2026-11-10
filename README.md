# Bitácora de Issues

## Ejecutar con Docker (recomendado)

Requisito: Docker Desktop instalado.

1. Abre una terminal en la carpeta del proyecto.
2. Ejecuta `docker compose up -d --build`.
3. Abre en el navegador:
   - http://localhost:8000/index.html
4. Para detenerla: `docker compose down`.

Los datos se guardan en una base de datos SQLite (`storage.db`) en la carpeta `C:\generacion-issues-data` de tu equipo, montada como volumen: la información sobrevive aunque el contenedor se detenga o se reconstruya. Esa ruta está fuera de OneDrive a propósito, para evitar bloqueos de sincronización con el archivo de base de datos; puedes cambiarla editando el volumen en `docker-compose.yml`. Si ya tenías datos en `data/storage.json`, se importan automáticamente la primera vez.

En Docker (Linux) el Excel se genera con las evidencias incrustadas como imágenes visibles en las celdas; los objetos OLE con icono requieren Windows con Excel de escritorio (ver ejecución local).

## Ejecutar la app localmente

1. Abre la carpeta del proyecto.
2. Instala las dependencias con `pip install -r requirements.txt`.
3. Doble clic en start.bat.
4. Abre en el navegador:
   - http://127.0.0.1:8000/index.html

El servidor local usa Microsoft Excel para insertar las evidencias como objetos OLE con icono en la hoja `Bitácora`; este modo requiere Windows, Excel de escritorio y las dependencias de `requirements.txt` actualizadas. GitHub Pages genera esos objetos directamente en el navegador, sin un servidor de exportación aparte; requiere el bundle `excel-ole.bundle.js` actualizado. En ambos casos, al hacer doble clic en un icono de Excel se abre la imagen incrustada.
Ambos modos respetan los filtros activos y crean las hojas `Bitácora` y `Resumen`. El detalle incluye la prioridad de cada issue, y el resumen agrega gráficas por estado, severidad y prioridad. También se pueden importar archivos Excel `.xlsx` o `.xls`: se agregan los Test ID nuevos, se omiten los duplicados y se generan automáticamente los códigos de las filas que dejan `Test ID` en blanco, según el módulo. La importación lee los campos de la hoja `Bitácora` (o la primera hoja) y no recupera imágenes incrustadas.
Desde la tabla o el tablero Kanban se puede descargar un informe Word independiente por issue. El documento conserva el encabezado y la geometría de página del formato de evidencia de pruebas. El caso y su título aparecen en tablas con bordes; los datos, el resumen, la descripción, las evidencias y las novedades quedan dentro de la tabla sombreada `EVIDENCIA / PRUEBAS`. No modifica ni reemplaza el documento Word adjunto. El encabezado se obtiene del archivo `issue-report-header.zip`.
En la tabla y en Kanban, los títulos extensos se muestran en una vista previa de hasta tres líneas para facilitar la lectura. Si el título se recorta, la opción `Ver título completo` lo abre completo en una ventana.
Cada issue admite hasta 10 imágenes, que se conservan en el respaldo y se incluyen en el Excel. Las capturas de pantalla y las imágenes PNG/JPEG pegadas o cargadas conservan su resolución original; otros formatos se convierten a JPEG al 92 % y se limitan a 2200 píxeles por lado. El visor muestra las dimensiones reales y permite ver la imagen a tamaño original. Las evidencias ya guardadas no recuperan los detalles perdidos antes de este cambio.
Cada issue puede tener un documento Word `.doc` o `.docx` de hasta 3 MB. El archivo queda disponible para vista previa desde la tabla o al editar, se puede reemplazar o quitar y se incluye en los respaldos JSON; la vista previa local muestra `.docx` y permite descargarlo desde el visor. El formato `.doc` antiguo muestra un aviso y se puede descargar para abrirlo en Word. El Excel conserva sus columnas de evidencia.
Los campos de título, descripción, novedades, comentarios y nombre de módulo incluyen un corrector ortográfico local de español, con sugerencias e ignorado de palabras. El texto se revisa en el navegador y no se envía a servicios externos; el diccionario va incluido en el archivo local `spellcheck.bundle.js`.
Para reconstruir el corrector, la vista previa local de Word, el lector de Excel, el generador OLE o los informes Word después de cambiar el código fuente, instala las dependencias con `npm install` y ejecuta `npm run build:spellcheck`, `npm run build:docx-preview`, `npm run build:excel-import`, `npm run build:excel-ole` y `npm run build:issue-report`. El diccionario español se distribuye bajo MPL 1.1 o posterior; revisa `THIRD_PARTY_NOTICES.md`.
La tabla muestra la antigüedad de cada issue abierto y alerta cuando supera el límite configurable (7 días de forma predeterminada). También muestra la fecha de cierre en el detalle, la tabla y el Excel. Se registra al agregar una solución o al cambiar el estado a Closed, incluso desde Kanban.
La lista de personas asignables se administra desde el botón `Personas`, se guarda en el almacenamiento de la aplicación y se incluye en los respaldos JSON.

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
- Con Docker, los datos se guardan en la base de datos SQLite `C:\generacion-issues-data\storage.db`, que queda en tu equipo mediante el volumen configurado.
- En modo local (start.bat), el servidor incluido en server.py usa la misma base de datos SQLite en `data/storage.db`; si existe un `data/storage.json` anterior, lo migra automáticamente y lo renombra a `storage.json.migrated`.
- Para extraer las imágenes de los issues como archivos, ejecuta `python export-images.py`; se crea la carpeta `imagenes-exportadas` con una subcarpeta por issue.
- Para extraer los documentos Word adjuntos, ejecuta `python export-documents.py`; se crea la carpeta `documentos-exportados` con una subcarpeta por issue y el nombre original de cada archivo.
