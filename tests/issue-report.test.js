import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import JSZip from 'jszip';
import {Packer} from 'docx';
import mammoth from 'mammoth';
import {createIssueReportBlob, createIssueReportDocument, issueReportFileName} from '../src/issue-report.js';

const tinyPng = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/l2cAAAAASUVORK5CYII=';
const issue = {
  id: 'report-test',
  testid: 'AF-001',
  formulario: 'Registrar requerimiento',
  modulo: 'Activos fijos',
  obs: `${'Descripción extensa sin perder saltos de línea. '.repeat(30)}\nDetalle adicional.`,
  status: 'Open',
  severity: 'Alta',
  priority: 'Alta',
  assigned: 'Equipo QA',
  fecha: '03/11/2026',
  closedAt: '',
  novedades: [{type: 'Seguimiento', createdAt: '2026-11-03T10:00:00Z', text: 'Validación pendiente.'}]
};

const documentWithEvidence = await createIssueReportDocument(issue, [tinyPng, tinyPng]);
const bytesWithEvidence = await Packer.toBuffer(documentWithEvidence);
const extractedWithEvidence = await mammoth.extractRawText({buffer: bytesWithEvidence});

for (const expected of [
  'CASO DE PRUEBAS',
  'EVIDENCIA / PRUEBAS',
  'AF-001',
  'Activos fijos',
  'Registrar requerimiento',
  'Descripción extensa sin perder saltos de línea.',
  'Evidencia 1',
  'Evidencia 2',
  'Seguimiento',
  'Validación pendiente.'
]) {
  assert.ok(extractedWithEvidence.value.includes(expected), `Generated report contains "${expected}"`);
}

const documentWithoutEvidence = await createIssueReportDocument({...issue, novedades: []}, []);
const bytesWithoutEvidence = await Packer.toBuffer(documentWithoutEvidence);
const extractedWithoutEvidence = await mammoth.extractRawText({buffer: bytesWithoutEvidence});
assert.ok(extractedWithoutEvidence.value.includes('Este issue no tiene imágenes adjuntas.'));

const headerTemplate = await readFile(new URL('../issue-report-header.zip', import.meta.url));
const sourceHeaderZip = await JSZip.loadAsync(headerTemplate);
assert.deepEqual(
  Object.keys(sourceHeaderZip.files).filter(path => !sourceHeaderZip.files[path].dir).sort(),
  ['word/_rels/header1.xml.rels', 'word/header1.xml', 'word/media/image1.png']
);
const reportBlob = await createIssueReportBlob(issue, [tinyPng], headerTemplate);
const reportZip = await JSZip.loadAsync(await reportBlob.arrayBuffer());
const finalHeaderXml = await reportZip.file('word/header1.xml').async('string');
const normalizedHeaderXml = finalHeaderXml
  .replaceAll(' w:dirty="true"', '')
  .replace(
    /(<w:instrText\b[^>]*>PAGE<\/w:instrText>[\s\S]*?<w:fldChar\b[^>]*w:fldCharType="separate"[^>]*\/>[\s\S]*?<w:t(?:\s[^>]*)?>)1(<\/w:t>)/,
    (_match, before, after) => `${before}2${after}`
  );
assert.equal(normalizedHeaderXml, await sourceHeaderZip.file('word/header1.xml').async('string'));
assert.equal((finalHeaderXml.match(/w:dirty="true"/g) || []).length, 2);
assert.deepEqual(
  await reportZip.file('word/media/report-template-image1.png').async('uint8array'),
  await sourceHeaderZip.file('word/media/image1.png').async('uint8array')
);
const reportDocumentXml = await reportZip.file('word/document.xml').async('string');
assert.match(reportDocumentXml, /<w:headerReference w:type="default" r:id="rId\d+"\/>/);
assert.match(reportDocumentXml, /<w:pgSz w:w="12242" w:h="15842"\/>/);
assert.match(reportDocumentXml, /<w:pgMar w:top="567" w:right="902" w:bottom="851" w:left="851" w:header="709" w:footer="369" w:gutter="0"\/>/);
const reportRelationships = await reportZip.file('word/_rels/document.xml.rels').async('string');
assert.match(reportRelationships, /Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/header" Target="header1\.xml"/);
const reportContentTypes = await reportZip.file('[Content_Types].xml').async('string');
assert.match(reportContentTypes, /PartName="\/word\/header1\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.header\+xml"/);
const reportSettings = await reportZip.file('word/settings.xml').async('string');
assert.match(reportSettings, /<w:updateFields w:val="true"\/>/);
const mergedText = await mammoth.extractRawText({buffer: Buffer.from(await reportBlob.arrayBuffer())});
assert.ok(mergedText.value.includes('AF-001'));

const nativeFetch = globalThis.fetch;
globalThis.fetch = async path => {
  assert.equal(path, 'issue-report-header.zip');
  return new Response(headerTemplate, {status: 200});
};
try {
  const defaultTemplateReport = await createIssueReportBlob(issue, []);
  const defaultTemplateZip = await JSZip.loadAsync(await defaultTemplateReport.arrayBuffer());
  assert.ok(defaultTemplateZip.file('word/header1.xml'));
} finally {
  globalThis.fetch = nativeFetch;
}

assert.equal(issueReportFileName(issue), `informe-AF-001-${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, '0')}-${String(new Date().getDate()).padStart(2, '0')}.docx`);
assert.match(issueReportFileName({...issue, testid: 'área / AF:001'}), /^informe-area-AF-001-\d{4}-\d{2}-\d{2}\.docx$/);
await assert.rejects(() => createIssueReportDocument(issue, ['invalid-image']), /formato que no se puede incluir/);

console.log('Issue Word report tests passed.');
