# Bitácora de Issues

## Ejecutar la app localmente

1. Abre la carpeta del proyecto.
2. Doble clic en start.bat.
3. Abre en el navegador:
   - http://127.0.0.1:8000/index.html

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
