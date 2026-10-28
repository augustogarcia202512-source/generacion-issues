from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
import json
import os
from pathlib import Path
import base64
import io

try:
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    Workbook = None

ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / 'data'
DATA_DIR.mkdir(exist_ok=True)
STORAGE_FILE = DATA_DIR / 'storage.json'

if not STORAGE_FILE.exists():
    STORAGE_FILE.write_text('{}', encoding='utf-8')

class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == '/api/storage':
            self.handle_storage_get(parsed)
            return
        super().do_GET()

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path == '/api/storage':
            self.handle_storage_post(parsed)
            return
        if parsed.path == '/api/export-excel':
            self.handle_export_excel(parsed)
            return
        super().do_POST()

    def do_DELETE(self):
        parsed = urlparse(self.path)
        if parsed.path == '/api/storage':
            self.handle_storage_delete(parsed)
            return
        super().do_DELETE()

    def handle_storage_get(self, parsed):
        params = parse_qs(parsed.query)
        key = params.get('key', [''])[0]
        data = self.load_storage()
        value = data.get(key)
        payload = {'value': value} if value is not None else {}
        self.send_json(payload)

    def handle_storage_post(self, parsed):
        length = int(self.headers.get('Content-Length', '0'))
        body = self.rfile.read(length).decode('utf-8') if length else '{}'
        payload = {}
        try:
            payload = json.loads(body)
        except json.JSONDecodeError:
            try:
                payload = {k: v[0] if len(v) == 1 else v for k, v in parse_qs(body).items()}
            except Exception:
                payload = {}

        key = payload.get('key')
        value = payload.get('value')
        if not key:
            self.send_json({'ok': False}, status=400)
            return

        data = self.load_storage()
        data[key] = value
        self.save_storage(data)
        self.send_json({'ok': True})

    def handle_export_excel(self, parsed):
        if Workbook is None:
            self.send_json({'error': 'OpenPyXL no está disponible en el servidor.'}, status=500)
            return

        length = int(self.headers.get('Content-Length', '0'))
        try:
            payload = json.loads(self.rfile.read(length).decode('utf-8'))
            issues = payload.get('issues', [])
            images = payload.get('images', {})
            output = self.build_excel(issues, images)
        except Exception as exc:
            self.send_json({'error': f'No se pudo generar el Excel: {exc}'}, status=500)
            return

        body = output.getvalue()
        filename = f"issues-{__import__('datetime').date.today().isoformat()}.xlsx"
        self.send_response(200)
        self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def build_excel(self, issues, images):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = 'Issues'
        sheet.freeze_panes = 'A2'
        sheet.sheet_view.showGridLines = False

        headers = ['Código del caso', 'Módulo afectado', 'Título del hallazgo',
                   'Descripción del hallazgo', 'Estado', 'Severity', 'Priority',
                   'Asignado a', 'Fecha de creación', 'Evidencia 1', 'Evidencia 2',
                   'Evidencia 3', 'Evidencia 4', 'Evidencia 5', 'Trazabilidad']
        sheet.append(headers)

        header_fill = PatternFill('solid', fgColor='E8A33D')
        thin_border = Border(*(Side(style='thin', color='D4DEE7') for _ in range(4)))
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = Font(bold=True, color='FFFFFF')
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = thin_border
        sheet.row_dimensions[1].height = 28

        width_map = [18, 24, 24, 32, 16, 16, 16, 18, 20, 18, 18, 18, 18, 18, 42]
        for index, width in enumerate(width_map, 1):
            sheet.column_dimensions[get_column_letter(index)].width = width

        for issue_index, issue in enumerate(issues, 2):
            trace = '\n\n'.join(
                f"{item.get('type', 'Novedad')} — {item.get('date', '')}\n{item.get('text', '')}"
                for item in issue.get('novedades', [])
            )
            values = [issue.get('testid', ''), issue.get('modulo', ''), issue.get('formulario', ''),
                      issue.get('obs', ''), issue.get('status', ''), issue.get('severity', ''),
                      issue.get('priority', ''), issue.get('assigned', ''), issue.get('createdAt', ''),
                      '', '', '', '', '', trace]
            sheet.append(values)
            sheet.row_dimensions[issue_index].height = 95
            for cell in sheet[issue_index]:
                cell.alignment = Alignment(vertical='top', wrap_text=True)
                cell.border = thin_border

            issue_images = images.get(str(issue.get('id')), [])[:5]
            for image_index, data_url in enumerate(issue_images):
                try:
                    image = self.decode_excel_image(data_url)
                    if image:
                        sheet.add_image(image, f'{get_column_letter(10 + image_index)}{issue_index}')
                except Exception:
                    continue

        metrics = workbook.create_sheet('Metricas')
        metrics.sheet_view.showGridLines = False
        metrics.freeze_panes = 'A3'
        self.populate_metrics(metrics, issues, images)

        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        return output

    def decode_excel_image(self, data_url):
        prefix, separator, payload = data_url.partition(',')
        if not separator or not prefix.startswith('data:image/'):
            return None
        raw = base64.b64decode(payload)
        image = Image(io.BytesIO(raw))
        image.width = 110
        image.height = 82
        return image

    def populate_metrics(self, metrics, issues, images):
        def counts(field):
            result = {}
            for issue in issues:
                key = str(issue.get(field) or 'Sin definir')
                result[key] = result.get(key, 0) + 1
            return result

        status_counts = counts('status')
        severity_counts = counts('severity')
        module_counts = counts('modulo')
        closed = sum(issue.get('status') == 'Closed' for issue in issues)
        evidence_count = sum(len(image_list) for image_list in images.values())

        metrics.append(['Resumen general', 'Valor'])
        general = [
            ('Total de issues', len(issues)),
            ('Issues cerrados', closed),
            ('Issues abiertos', len(issues) - closed),
            ('Porcentaje cerrado', f'{round(closed * 100 / len(issues)) if issues else 0}%'),
            ('Evidencias adjuntas', evidence_count)
        ]
        for row in general:
            metrics.append(row)

        metrics.append([])
        metrics.append(['Por estado', 'Cantidad', 'Por severidad', 'Cantidad', 'Por módulo', 'Cantidad'])
        metric_groups = [status_counts, severity_counts, module_counts]
        max_rows = max([len(group) for group in metric_groups] + [0])
        for index in range(max_rows):
            values = []
            for group in metric_groups:
                entry = list(group.items())[index] if index < len(group) else ('', '')
                values.extend(entry)
            metrics.append(values)

        metrics.column_dimensions['A'].width = 28
        metrics.column_dimensions['B'].width = 18
        metrics.column_dimensions['C'].width = 28
        metrics.column_dimensions['D'].width = 18
        metrics.column_dimensions['E'].width = 28
        metrics.column_dimensions['F'].width = 18
        dark_fill = PatternFill('solid', fgColor='142534')
        section_fill = PatternFill('solid', fgColor='D8EAF3')
        for cell in metrics[1]:
            cell.fill = dark_fill
            cell.font = Font(bold=True, color='FFFFFF')
            cell.alignment = Alignment(horizontal='center')
        for row in (7,):
            for cell in metrics[row]:
                cell.fill = section_fill
                cell.font = Font(bold=True)
        for row in metrics.iter_rows():
            for cell in row:
                cell.border = Border(*(Side(style='thin', color='D4DEE7') for _ in range(4)))
                cell.alignment = Alignment(vertical='top', wrap_text=True)

    def handle_storage_delete(self, parsed):
        params = parse_qs(parsed.query)
        key = params.get('key', [''])[0]
        if not key:
            self.send_json({'ok': False}, status=400)
            return

        data = self.load_storage()
        data.pop(key, None)
        self.save_storage(data)
        self.send_json({'ok': True})

    def load_storage(self):
        try:
            with STORAGE_FILE.open('r', encoding='utf-8') as fh:
                return json.load(fh)
        except Exception:
            return {}

    def save_storage(self, data):
        with STORAGE_FILE.open('w', encoding='utf-8') as fh:
            json.dump(data, fh, ensure_ascii=False, indent=2)

    def send_json(self, payload, status=200):
        body = json.dumps(payload).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', '8000'))
    server = ThreadingHTTPServer(('0.0.0.0', port), Handler)
    print(f'Serving at http://127.0.0.1:{port}')
    server.serve_forever()
