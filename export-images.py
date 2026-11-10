"""Extrae las imágenes de los issues guardadas en la base de datos SQLite
y las guarda como archivos en una carpeta, organizadas por issue.

Uso:
    python export-images.py [carpeta_destino]

Si no se indica carpeta destino, se crea "imagenes-exportadas" junto a la base de
datos (C:/generacion-issues-data/imagenes-exportadas), fuera de OneDrive.
La base de datos se lee de C:/generacion-issues-data/storage.db (la misma que usa
Docker) o de data/storage.db si existe (modo local con start.bat).
"""

import base64
import json
import sqlite3
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DOCKER_DB = Path(r'C:\generacion-issues-data\storage.db')
LOCAL_DB = ROOT / 'data' / 'storage.db'

MIME_EXTENSIONS = {
    'image/jpeg': '.jpg',
    'image/png': '.png',
    'image/gif': '.gif',
    'image/bmp': '.bmp',
    'image/webp': '.webp',
    'image/tiff': '.tif',
}


def find_database():
    for candidate in (DOCKER_DB, LOCAL_DB):
        if candidate.exists():
            return candidate
    sys.exit('No se encontró la base de datos (ni en Docker ni en data\\storage.db).')


def main():
    db_path = find_database()
    output_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else db_path.parent / 'imagenes-exportadas'
    print(f'Base de datos: {db_path}')

    conn = sqlite3.connect(f'file:{db_path}?mode=ro', uri=True)
    keys = [row[0] for row in conn.execute(
        "SELECT key FROM storage WHERE key LIKE 'issue-images:%' ORDER BY key"
    )]
    if not keys:
        sys.exit('No hay imágenes guardadas en la base de datos.')

    # Mapa id de issue -> Test ID, para nombrar las carpetas de forma legible.
    test_ids = {}
    row = conn.execute("SELECT value FROM storage WHERE key = 'issues-list'").fetchone()
    if row:
        try:
            issues = json.loads(json.loads(row[0]) if isinstance(json.loads(row[0]), str) else row[0])
            for issue in issues:
                if isinstance(issue, dict) and issue.get('id'):
                    test_ids[str(issue['id'])] = str(issue.get('testid') or issue['id'])
        except Exception:
            pass

    total = 0
    for key in keys:
        issue_id = key.split(':', 1)[1]
        raw = conn.execute('SELECT value FROM storage WHERE key = ?', (key,)).fetchone()[0]
        value = json.loads(raw)
        if isinstance(value, str):
            value = json.loads(value)
        if not isinstance(value, list):
            continue

        folder_name = f'{test_ids.get(issue_id, issue_id)}_{issue_id}'
        issue_dir = output_dir / folder_name
        saved = 0
        for index, data_url in enumerate(value, 1):
            if not isinstance(data_url, str):
                continue
            prefix, separator, encoded = data_url.partition(',')
            if not separator or not prefix.startswith('data:image/'):
                continue
            mime_type = prefix[5:].split(';', 1)[0].lower()
            extension = MIME_EXTENSIONS.get(mime_type, '.png')
            try:
                image_data = base64.b64decode(encoded, validate=True)
            except Exception:
                print(f'  ! No se pudo decodificar la imagen {index} de {key}')
                continue
            issue_dir.mkdir(parents=True, exist_ok=True)
            (issue_dir / f'evidencia-{index}{extension}').write_bytes(image_data)
            saved += 1
        if saved:
            print(f'  {folder_name}: {saved} imagen(es)')
            total += saved

    print(f'\nListo: {total} imagen(es) exportadas a {output_dir}')


if __name__ == '__main__':
    main()
