"""Extrae los documentos Word adjuntos a los issues guardados en la base de datos
SQLite y los guarda como archivos en una carpeta, organizados por issue.

Uso:
    python export-documents.py [carpeta_destino]

Si no se indica carpeta destino, se crea "documentos-exportados" junto a la base de
datos (C:/generacion-issues-data/documentos-exportados), fuera de OneDrive.
La base de datos se lee de C:/generacion-issues-data/storage.db (la misma que usa
Docker) o de data/storage.db si existe (modo local con start.bat).
"""

import base64
import json
import re
import sqlite3
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DOCKER_DB = Path(r'C:\generacion-issues-data\storage.db')
LOCAL_DB = ROOT / 'data' / 'storage.db'


def find_database():
    for candidate in (DOCKER_DB, LOCAL_DB):
        if candidate.exists():
            return candidate
    sys.exit('No se encontró la base de datos (ni en Docker ni en data\\storage.db).')


def safe_filename(name):
    cleaned = re.sub(r'[<>:"/\\|?*]', '_', str(name)).strip()
    return cleaned or 'documento.docx'


def main():
    db_path = find_database()
    output_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else db_path.parent / 'documentos-exportados'
    print(f'Base de datos: {db_path}')

    conn = sqlite3.connect(f'file:{db_path}?mode=ro', uri=True)
    keys = [row[0] for row in conn.execute(
        "SELECT key FROM storage WHERE key LIKE 'issue-document:%' ORDER BY key"
    )]
    if not keys:
        sys.exit('No hay documentos guardados en la base de datos.')

    # Mapa id de issue -> Test ID, para nombrar las carpetas de forma legible.
    test_ids = {}
    row = conn.execute("SELECT value FROM storage WHERE key = 'issues-list'").fetchone()
    if row:
        try:
            issues = json.loads(row[0])
            if isinstance(issues, str):
                issues = json.loads(issues)
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
        if not isinstance(value, dict):
            continue

        data_url = value.get('dataUrl', '')
        prefix, separator, encoded = data_url.partition(',')
        if not separator or not prefix.startswith('data:'):
            print(f'  ! El documento de {key} no tiene un data URL válido')
            continue
        try:
            document_data = base64.b64decode(encoded, validate=True)
        except Exception:
            print(f'  ! No se pudo decodificar el documento de {key}')
            continue

        folder_name = f'{test_ids.get(issue_id, issue_id)}_{issue_id}'
        issue_dir = output_dir / folder_name
        issue_dir.mkdir(parents=True, exist_ok=True)
        filename = safe_filename(value.get('name') or f'documento-{issue_id}.docx')
        (issue_dir / filename).write_bytes(document_data)
        print(f'  {folder_name}: {filename} ({len(document_data):,} bytes)')
        total += 1

    print(f'\nListo: {total} documento(s) exportados a {output_dir}')


if __name__ == '__main__':
    main()
