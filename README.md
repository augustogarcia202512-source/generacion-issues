# Bitácora de Issues

## Ejecutar la app localmente

1. Abre la carpeta del proyecto.
2. Instala las dependencias con `pip install -r requirements.txt`.
3. Doble clic en start.bat.
4. Abre en el navegador:
   - http://127.0.0.1:8000/index.html

El servidor local usa Microsoft Excel para insertar las evidencias como objetos OLE con icono en la hoja `Bitácora`; este modo requiere Windows, Excel de escritorio y las dependencias de `requirements.txt` actualizadas. GitHub Pages genera esos objetos directamente en el navegador, sin un servidor de exportación aparte; requiere el bundle `excel-ole.bundle.js` actualizado. En ambos casos, al hacer doble clic en un icono de Excel se abre la imagen incrustada.
Ambos modos respetan los filtros activos y crean las hojas `Bitácora` y `Resumen`. El detalle incluye la prioridad de cada issue, y el resumen agrega gráficas por estado, severidad y prioridad. También se pueden importar archivos Excel `.xlsx` o `.xls`: se agregan los Test ID nuevos, se omiten los duplicados y se generan automáticamente los códigos de las filas que dejan `Test ID` en blanco, según el módulo. La importación lee los campos de la hoja `Bitácora` (o la primera hoja) y no recupera imágenes incrustadas.
Cada issue admite hasta 10 imágenes, que se conservan en el respaldo y se incluyen en el Excel. Las capturas de pantalla y las imágenes PNG/JPEG pegadas o cargadas conservan su resolución original; otros formatos se convierten a JPEG al 92 % y se limitan a 2200 píxeles por lado. El visor muestra las dimensiones reales y permite ver la imagen a tamaño original. Las evidencias ya guardadas no recuperan los detalles perdidos antes de este cambio.
Cada issue puede tener un documento Word `.doc` o `.docx` de hasta 3 MB. El archivo queda disponible para vista previa desde la tabla o al editar, se puede reemplazar o quitar y se incluye en los respaldos JSON; la vista previa local muestra `.docx` y permite descargarlo desde el visor. El formato `.doc` antiguo muestra un aviso y se puede descargar para abrirlo en Word. El Excel conserva sus columnas de evidencia.
Los campos de título, descripción, novedades, comentarios y nombre de módulo incluyen un corrector ortográfico local de español, con sugerencias e ignorado de palabras. El texto se revisa en el navegador y no se envía a servicios externos; el diccionario va incluido en el archivo local `spellcheck.bundle.js`.
Para reconstruir el corrector, la vista previa local de Word, el lector de Excel o el generador OLE después de cambiar el código fuente, instala las dependencias con `npm install` y ejecuta `npm run build:spellcheck`, `npm run build:docx-preview`, `npm run build:excel-import` y `npm run build:excel-ole`. El diccionario español se distribuye bajo MPL 1.1 o posterior; revisa `THIRD_PARTY_NOTICES.md`.
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
- En modo local, también puedes usar el servidor incluido en server.py, que guarda los datos en data/storage.json.
