import * as XLSX from '@e965/xlsx';

const HEADER_ALIASES = {
  testid: ['test id', 'testid', 'id', 'codigo del caso', 'codigo', 'issue key', 'issue id'],
  formulario: ['titulo del hallazgo', 'titulo', 'title', 'summary', 'issue title'],
  modulo: ['modulo', 'module', 'componente', 'component'],
  obs: ['descripcion del hallazgo', 'descripcion', 'description', 'observaciones', 'detalle', 'detail'],
  status: ['estado', 'status'],
  severity: ['severidad', 'severity', 'impacto'],
  priority: ['prioridad', 'priority'],
  assigned: ['asignado a', 'asignado', 'responsable', 'assigned to', 'assignee'],
  fecha: ['fecha registro', 'fecha', 'created', 'created at', 'date'],
  closedAt: ['fecha cierre', 'fecha de cierre', 'closed at'],
  trace: ['comentarios trazabilidad', 'comentarios', 'novedades', 'comments', 'history']
};

const STATUS_VALUES = new Map([
  ['open', 'Open'],
  ['blocked', 'Blocked'],
  ['ready for retest', 'ReadyRetest'],
  ['readyretest', 'ReadyRetest'],
  ['retest', 'ReadyRetest'],
  ['in progress', 'InProgress'],
  ['inprogress', 'InProgress'],
  ['closed', 'Closed'],
  ['re open', 'ReOpen'],
  ['reopen', 'ReOpen']
]);

function normalizeHeader(value) {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLocaleLowerCase('es')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function findHeaderRow(rows) {
  for (let rowIndex = 0; rowIndex < Math.min(rows.length, 20); rowIndex++) {
    const values = new Set(rows[rowIndex].map(normalizeHeader));
    if (HEADER_ALIASES.testid.some(alias => values.has(alias))) return rowIndex;
  }
  throw new Error('No se encontró una columna Test ID o Código del caso en la hoja.');
}

function parseStatus(value, rowNumber) {
  const normalized = normalizeHeader(value);
  if (!normalized) return 'Open';
  const status = STATUS_VALUES.get(normalized);
  if (!status) {
    throw new Error(`El estado "${value}" de la fila ${rowNumber} no es compatible.`);
  }
  return status;
}

export function parseExcelIssues(data) {
  let workbook;
  try {
    workbook = XLSX.read(data, { type: 'array', cellDates: true, dateNF: 'dd/mm/yyyy' });
  } catch (error) {
    throw new Error(`No se pudo leer el archivo Excel: ${error.message}`);
  }

  if (!workbook.SheetNames.length) throw new Error('El archivo Excel no contiene hojas.');
  const sheetName = workbook.SheetNames.find(name => normalizeHeader(name) === 'bitacora') || workbook.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
    header: 1,
    raw: false,
    defval: '',
    blankrows: false
  });
  const headerRowIndex = findHeaderRow(rows);
  const headerValues = rows[headerRowIndex].map(normalizeHeader);
  const columns = Object.fromEntries(Object.entries(HEADER_ALIASES).map(([field, aliases]) => [
    field,
    headerValues.findIndex(value => aliases.includes(value))
  ]));
  const testIdColumn = columns.testid;
  const issues = [];

  for (let rowIndex = headerRowIndex + 1; rowIndex < rows.length; rowIndex++) {
    const row = rows[rowIndex];
    const testid = String(row[testIdColumn] ?? '').trim();
    if (!testid) {
      if (!row.some(value => String(value ?? '').trim())) continue;
    }
    const cell = field => columns[field] < 0 ? '' : String(row[columns[field]] ?? '').trim();
    const rowNumber = rowIndex + 1;
    const trace = cell('trace');
    issues.push({
      sourceRow: rowNumber,
      testid,
      formulario: cell('formulario'),
      modulo: cell('modulo') || 'Sin módulo',
      obs: cell('obs'),
      status: parseStatus(cell('status'), rowNumber),
      severity: cell('severity'),
      priority: cell('priority'),
      assigned: cell('assigned'),
      fecha: cell('fecha'),
      closedAt: cell('closedAt'),
      novedades: trace ? [trace] : [],
      solucion: trace
    });
  }

  if (!issues.length) throw new Error('No se encontraron issues en la hoja.');
  return issues;
}
