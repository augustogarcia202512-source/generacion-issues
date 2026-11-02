import assert from 'node:assert/strict';
import test from 'node:test';
import * as XLSX from '@e965/xlsx';
import {parseExcelIssues} from '../src/excel-import.js';

function workbookBytes(sheets) {
  const workbook = XLSX.utils.book_new();
  for (const [name, rows] of sheets) {
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), name);
  }
  return XLSX.write(workbook, {type: 'array', bookType: 'xlsx'});
}

test('reads the Bitácora sheet and maps exported issue fields', () => {
  const bytes = workbookBytes([
    ['Resumen', [['Indicador', 'Valor'], ['Total', 99]]],
    ['Bitácora', [
      ['Test ID', 'Título del hallazgo', 'Módulo', 'Descripción del hallazgo', 'Estado', 'Severidad', 'Prioridad', 'Asignado a', 'Fecha registro', 'Fecha cierre', 'Comentarios / trazabilidad'],
      ['QA-001', 'Error al guardar', 'Pagos', 'El botón no responde', 'In Progress', 'Alta', 'Media', 'Ana', '02/11/2026', '', 'Revisado por QA'],
      ['QA-002', 'Issue reabierto', 'Pagos', 'Validación', 'Re-Open', '', '', '', '', '', '']
    ]]
  ]);
  const issues = parseExcelIssues(bytes);
  assert.equal(issues.length, 2);
  assert.deepEqual(issues[0], {
    sourceRow: 2,
    testid: 'QA-001',
    formulario: 'Error al guardar',
    modulo: 'Pagos',
    obs: 'El botón no responde',
    status: 'InProgress',
    severity: 'Alta',
    priority: 'Media',
    assigned: 'Ana',
    fecha: '02/11/2026',
    closedAt: '',
    novedades: ['Revisado por QA'],
    solucion: 'Revisado por QA'
  });
  assert.equal(issues[1].status, 'ReOpen');
});

test('accepts blank Test ID for automatic assignment, localized aliases, and skips blank rows', () => {
  const bytes = workbookBytes([
    ['Issues', [
      ['Código del caso', 'Título', 'Descripción'],
      ['QA-002', 'Issue sin módulo', 'Detalle'],
      ['', 'Issue con código automático', 'Detalle'],
      ['', '', '']
    ]]
  ]);
  const issues = parseExcelIssues(bytes);
  assert.equal(issues.length, 2);
  assert.equal(issues[0].status, 'Open');
  assert.equal(issues[0].modulo, 'Sin módulo');
  assert.equal(issues[1].testid, '');
  assert.equal(issues[1].formulario, 'Issue con código automático');
});

test('rejects unsupported statuses and workbooks without a Test ID column', () => {
  const unsupportedStatus = workbookBytes([
    ['Issues', [['Test ID', 'Estado'], ['QA-003', 'En revisión']]]
  ]);
  assert.throws(() => parseExcelIssues(unsupportedStatus), /fila 2 no es compatible/);

  const missingId = workbookBytes([
    ['Issues', [['Título', 'Estado'], ['Issue', 'Open']]]
  ]);
  assert.throws(() => parseExcelIssues(missingId), /No se encontró una columna Test ID/);

});
