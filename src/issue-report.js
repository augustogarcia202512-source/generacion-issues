import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  ImageRun,
  PageNumber,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType
} from 'docx';
import JSZip from 'jszip';

const COLORS = {
  navy: '1F3A5F',
  amber: 'E8A33D',
  text: '263746',
  muted: '64788A',
  line: 'D4DEE7',
  soft: 'EEF3F7',
  white: 'FFFFFF'
};
const CONTENT_WIDTH = 10489;
const IMAGE_MAX_WIDTH = 490;
const IMAGE_MAX_HEIGHT = 680;
const REPORT_PAGE_SIZE = {width: 12242, height: 15842};
const REPORT_PAGE_MARGINS = {top: 567, right: 902, bottom: 851, left: 851, header: 709, footer: 369, gutter: 0};
const HEADER_RELATIONSHIP_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header';
const HEADER_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml';
const HEADER_IMAGE_PATH = 'word/media/report-template-image1.png';

const STATUS_LABELS = {
  Open: 'Open',
  Blocked: 'Blocked',
  ReadyRetest: 'Ready for Retest',
  InProgress: 'In Progress',
  Closed: 'Closed',
  ReOpen: 'Re-Open'
};

function safeText(value, fallback = 'Sin definir') {
  const text = String(value ?? '').trim();
  return text || fallback;
}

function textParagraph(text, options = {}) {
  return new Paragraph({
    spacing: {after: options.after ?? 120, line: 280},
    alignment: options.alignment ?? AlignmentType.LEFT,
    children: [new TextRun({
      text: String(text ?? ''),
      font: 'Arial',
      size: options.size ?? 21,
      bold: options.bold ?? false,
      color: options.color ?? COLORS.text,
      italics: options.italics ?? false
    })]
  });
}

function sectionHeading(text) {
  return new Paragraph({
    spacing: {before: 180, after: 80},
    keepNext: true,
    children: [new TextRun({
      text,
      font: 'Arial',
      size: 20,
      bold: true,
      color: '000000'
    })]
  });
}

function detailHeading(text) {
  return new Paragraph({
    spacing: {before: 150, after: 70},
    keepNext: true,
    children: [new TextRun({
      text,
      font: 'Arial',
      size: 22,
      bold: true,
      color: '000000'
    })]
  });
}

function reportCell(children, {shading, margins = {top: 100, bottom: 100, left: 120, right: 120}} = {}) {
  return new TableCell({
    width: {size: CONTENT_WIDTH, type: WidthType.DXA},
    margins,
    shading: shading ? {type: ShadingType.CLEAR, fill: shading} : undefined,
    borders: {
      top: {style: BorderStyle.SINGLE, size: 6, color: '000000'},
      bottom: {style: BorderStyle.SINGLE, size: 6, color: '000000'},
      left: {style: BorderStyle.SINGLE, size: 6, color: '000000'},
      right: {style: BorderStyle.SINGLE, size: 6, color: '000000'}
    },
    children
  });
}

function borderedTable(rows) {
  const border = {style: BorderStyle.SINGLE, size: 6, color: '000000'};
  return new Table({
    width: {size: CONTENT_WIDTH, type: WidthType.DXA},
    columnWidths: [CONTENT_WIDTH],
    borders: {top: border, bottom: border, left: border, right: border, insideHorizontal: border, insideVertical: border},
    rows
  });
}

function centeredTitle(text, size, after = 0) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: {after},
    children: [new TextRun({text, font: 'Arial', size, bold: true, color: '000000'})]
  });
}

function metadataParagraph(label, value) {
  return new Paragraph({
    spacing: {after: 0, line: 240},
    children: [
      new TextRun({text: `${label}: `, font: 'Arial', size: 20, bold: true, color: '000000'}),
      new TextRun({text: String(value ?? ''), font: 'Arial', size: 20, color: '000000'})
    ]
  });
}

function decodeImageDataUrl(dataUrl) {
  const match = /^data:image\/(png|jpe?g|gif|bmp|webp);base64,([A-Za-z0-9+/]+=*)$/i.exec(dataUrl || '');
  if (!match) throw new Error('Una evidencia tiene un formato que no se puede incluir en el informe Word.');
  const mimeExtension = match[1].toLowerCase();
  const extension = mimeExtension === 'jpeg' ? 'jpg' : mimeExtension;
  const binary = atob(match[2]);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index++) bytes[index] = binary.charCodeAt(index);
  return {bytes, extension};
}

function imageDimensions(bytes, extension) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (extension === 'png' && bytes.length >= 24) {
    return {width: view.getUint32(16), height: view.getUint32(20)};
  }
  if (extension === 'gif' && bytes.length >= 10) {
    return {width: view.getUint16(6, true), height: view.getUint16(8, true)};
  }
  if (extension === 'bmp' && bytes.length >= 26) {
    return {width: Math.abs(view.getInt32(18, true)), height: Math.abs(view.getInt32(22, true))};
  }
  if (extension === 'jpg' && bytes.length >= 4) {
    let offset = 2;
    while (offset + 9 < bytes.length) {
      if (bytes[offset] !== 0xff) {
        offset++;
        continue;
      }
      const marker = bytes[offset + 1];
      const segmentLength = view.getUint16(offset + 2);
      if ([0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf].includes(marker)) {
        return {width: view.getUint16(offset + 7), height: view.getUint16(offset + 5)};
      }
      if (segmentLength < 2) break;
      offset += 2 + segmentLength;
    }
  }
  return null;
}

async function prepareImage(dataUrl, index) {
  let {bytes, extension} = decodeImageDataUrl(dataUrl);
  if (extension === 'webp') {
    if (typeof createImageBitmap !== 'function') {
      throw new Error(`No se pudo convertir la evidencia ${index + 1} al formato compatible con Word.`);
    }
    const bitmap = await createImageBitmap(new Blob([bytes], {type: 'image/webp'}));
    const canvas = typeof OffscreenCanvas === 'function'
      ? new OffscreenCanvas(bitmap.width, bitmap.height)
      : Object.assign(document.createElement('canvas'), {width: bitmap.width, height: bitmap.height});
    const context = canvas.getContext('2d');
    if (!context) throw new Error(`No se pudo convertir la evidencia ${index + 1} al formato compatible con Word.`);
    context.drawImage(bitmap, 0, 0);
    bitmap.close();
    const png = typeof canvas.convertToBlob === 'function'
      ? await canvas.convertToBlob({type: 'image/png'})
      : await new Promise(resolve => canvas.toBlob(resolve, 'image/png'));
    if (!png) throw new Error(`No se pudo convertir la evidencia ${index + 1} al formato compatible con Word.`);
    bytes = new Uint8Array(await png.arrayBuffer());
    extension = 'png';
  }

  const dimensions = imageDimensions(bytes, extension);
  if (!dimensions?.width || !dimensions?.height) {
    throw new Error(`No se pudieron leer las dimensiones de la evidencia ${index + 1}.`);
  }
  const scale = Math.min(IMAGE_MAX_WIDTH / dimensions.width, IMAGE_MAX_HEIGHT / dimensions.height, 1);
  return {
    bytes,
    extension,
    width: Math.max(1, Math.round(dimensions.width * scale)),
    height: Math.max(1, Math.round(dimensions.height * scale))
  };
}

function trackingContent(issue) {
  const entries = Array.isArray(issue.novedades) ? issue.novedades : [];
  return entries.flatMap(entry => {
    const text = typeof entry === 'string' ? entry : entry?.text;
    if (!String(text || '').trim()) return [];
    const type = typeof entry === 'string' ? 'Novedad' : safeText(entry.type, 'Novedad');
    const date = typeof entry === 'string' ? '' : entry.createdAt || entry.date;
    const metadata = [type, date].filter(Boolean).join(' · ');
    return [
      ...(metadata ? [textParagraph(metadata, {size: 17, color: COLORS.muted, bold: true, after: 25})] : []),
      textParagraph(text, {after: 125})
    ];
  });
}

async function loadHeaderTemplate(templateBytes) {
  if (templateBytes) return templateBytes;
  const response = await fetch('issue-report-header.zip');
  if (!response.ok) {
    throw new Error(`No se pudo cargar el encabezado del formato Word (${response.status}).`);
  }
  return response.arrayBuffer();
}

function nextRelationshipId(relationshipsXml) {
  const ids = [...relationshipsXml.matchAll(/\bId="rId(\d+)"/g)].map(match => Number(match[1]));
  return `rId${Math.max(0, ...ids) + 1}`;
}

async function applyOriginalHeader(reportBlob, templateBytes) {
  const [reportZip, templateZip] = await Promise.all([
    JSZip.loadAsync(reportBlob),
    JSZip.loadAsync(await loadHeaderTemplate(templateBytes))
  ]);
  const requiredParts = ['word/header1.xml', 'word/_rels/header1.xml.rels', 'word/media/image1.png'];
  const parts = await Promise.all(requiredParts.map(async path => {
    const file = templateZip.file(path);
    if (!file) throw new Error(`Falta el componente "${path}" del formato Word original.`);
    return file.async(path.endsWith('.xml') || path.endsWith('.rels') ? 'string' : 'uint8array');
  }));
  const [sourceHeaderXml, sourceHeaderRelationships, headerImage] = parts;
  let headerXml = sourceHeaderXml.replace(
    /(<w:fldChar\b[^>]*w:fldCharType="begin")\/>/g,
    '$1 w:dirty="true"/>'
  );
  const firstPageHeader = headerXml.replace(
    /(<w:instrText\b[^>]*>PAGE<\/w:instrText>[\s\S]*?<w:fldChar\b[^>]*w:fldCharType="separate"[^>]*\/>[\s\S]*?<w:t(?:\s[^>]*)?>)\d+(<\/w:t>)/,
    (_match, before, after) => `${before}1${after}`
  );
  if (firstPageHeader === headerXml) {
    throw new Error('No se pudo inicializar la numeración de páginas del encabezado original.');
  }
  headerXml = firstPageHeader;
  const headerRelationships = sourceHeaderRelationships.replace(
    'Target="media/image1.png"',
    'Target="media/report-template-image1.png"'
  );
  if (headerRelationships === sourceHeaderRelationships) {
    throw new Error('No se encontró la imagen del encabezado original.');
  }

  const documentRelationshipsFile = reportZip.file('word/_rels/document.xml.rels');
  const documentFile = reportZip.file('word/document.xml');
  const contentTypesFile = reportZip.file('[Content_Types].xml');
  const settingsFile = reportZip.file('word/settings.xml');
  if (!documentRelationshipsFile || !documentFile || !contentTypesFile || !settingsFile) {
    throw new Error('El informe Word generado no contiene las partes necesarias para aplicar el formato original.');
  }

  let documentRelationships = await documentRelationshipsFile.async('string');
  const headerRelationshipId = nextRelationshipId(documentRelationships);
  documentRelationships = documentRelationships.replace(
    '</Relationships>',
    `<Relationship Id="${headerRelationshipId}" Type="${HEADER_RELATIONSHIP_TYPE}" Target="header1.xml"/></Relationships>`
  );

  let documentXml = await documentFile.async('string');
  const sectionMatch = documentXml.match(/<w:sectPr(?:\s[^>]*)?>[\s\S]*?<\/w:sectPr>/);
  if (!sectionMatch) throw new Error('No se encontró la configuración de página del informe Word.');
  let sectionXml = sectionMatch[0];
  sectionXml = sectionXml.replace(
    /<w:pgSz\b[^>]*\/>/,
    `<w:pgSz w:w="${REPORT_PAGE_SIZE.width}" w:h="${REPORT_PAGE_SIZE.height}"/>`
  );
  sectionXml = sectionXml.replace(
    /<w:pgMar\b[^>]*\/>/,
    `<w:pgMar w:top="${REPORT_PAGE_MARGINS.top}" w:right="${REPORT_PAGE_MARGINS.right}" w:bottom="${REPORT_PAGE_MARGINS.bottom}" w:left="${REPORT_PAGE_MARGINS.left}" w:header="${REPORT_PAGE_MARGINS.header}" w:footer="${REPORT_PAGE_MARGINS.footer}" w:gutter="${REPORT_PAGE_MARGINS.gutter}"/>`
  );
  sectionXml = sectionXml.replace(
    '<w:sectPr>',
    `<w:sectPr><w:headerReference w:type="default" r:id="${headerRelationshipId}"/>`
  );
  if (!sectionXml.includes('<w:headerReference ')) {
    throw new Error('No se pudo enlazar el encabezado original al informe Word.');
  }
  documentXml = documentXml.replace(sectionMatch[0], sectionXml);

  let contentTypes = await contentTypesFile.async('string');
  if (!contentTypes.includes('PartName="/word/header1.xml"')) {
    contentTypes = contentTypes.replace(
      '</Types>',
      `<Override PartName="/word/header1.xml" ContentType="${HEADER_CONTENT_TYPE}"/></Types>`
    );
  }
  let settingsXml = await settingsFile.async('string');
  if (!settingsXml.includes('<w:updateFields')) {
    settingsXml = settingsXml.replace('</w:settings>', '<w:updateFields w:val="true"/></w:settings>');
  }

  reportZip.file('word/header1.xml', headerXml);
  reportZip.file('word/_rels/header1.xml.rels', headerRelationships);
  reportZip.file(HEADER_IMAGE_PATH, headerImage);
  reportZip.file('word/_rels/document.xml.rels', documentRelationships);
  reportZip.file('word/document.xml', documentXml);
  reportZip.file('[Content_Types].xml', contentTypes);
  reportZip.file('word/settings.xml', settingsXml);
  return reportZip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  });
}

export async function createIssueReportDocument(issue, evidence = []) {
  if (!issue || typeof issue !== 'object') throw new TypeError('No se encontró el issue para generar el informe.');
  const generatedAt = new Intl.DateTimeFormat('es-CO', {dateStyle: 'long'}).format(new Date());
  const images = await Promise.all(evidence.map(prepareImage));
  const content = [
    metadataParagraph('Código', safeText(issue.testid)),
    metadataParagraph('Módulo', safeText(issue.modulo)),
    metadataParagraph('Fecha de ejecución', safeText(issue.fecha, generatedAt)),
    metadataParagraph('Estado', STATUS_LABELS[issue.status] || safeText(issue.status)),
    metadataParagraph('Severidad', safeText(issue.severity)),
    metadataParagraph('Prioridad', safeText(issue.priority || issue.Priority || issue.prioridad)),
    metadataParagraph('Asignado a', safeText(issue.assigned, 'Sin asignar')),
    metadataParagraph('Fecha de cierre', safeText(issue.closedAtText || issue.closedAt, '—')),
    sectionHeading('RESUMEN DEL CASO'),
    detailHeading('1. Título'),
    textParagraph(safeText(issue.formulario, 'Sin título')),
    detailHeading('2. Descripción del problema'),
    textParagraph(safeText(issue.obs, 'Sin descripción registrada.')),
    detailHeading('3. Evidencia y comportamiento:')
  ];

  if (images.length) {
    for (let index = 0; index < images.length; index++) {
      const image = images[index];
      content.push(
        textParagraph(`${index + 1}. Evidencia ${index + 1}:`, {size: 20, bold: true, after: 60}),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: {after: 120},
          children: [new ImageRun({
            data: image.bytes,
            type: image.extension,
            transformation: {width: image.width, height: image.height},
            altText: {
              name: `Evidencia ${index + 1}`,
              title: `Evidencia ${index + 1} - ${safeText(issue.testid)}`,
              description: `Imagen adjunta al issue ${safeText(issue.testid)}`
            }
          })]
        })
      );
    }
  } else {
    content.push(textParagraph('Este issue no tiene imágenes adjuntas.', {color: '000000', italics: true}));
  }

  const updates = trackingContent(issue);
  if (updates.length) {
    content.push(sectionHeading('Seguimiento y novedades'), ...updates);
  }
  content.push(
    new Paragraph({
      spacing: {before: 300},
      alignment: AlignmentType.RIGHT,
      children: [new TextRun({text: `Generado el ${generatedAt}`, font: 'Arial', size: 16, color: '404040', italics: true})]
    })
  );

  const reportTable = borderedTable([
    new TableRow({
      cantSplit: true,
      children: [reportCell([centeredTitle('CASO DE PRUEBAS', 24)], {margins: {top: 40, bottom: 40, left: 80, right: 80}})]
    }),
    new TableRow({
      cantSplit: true,
      children: [reportCell([centeredTitle(safeText(issue.formulario, 'Sin título'), 20, 0)], {margins: {top: 40, bottom: 40, left: 80, right: 80}})]
    }),
    new TableRow({
      cantSplit: true,
      children: [reportCell([centeredTitle('EVIDENCIA / PRUEBAS', 22)], {margins: {top: 20, bottom: 20, left: 80, right: 80}})]
    }),
    new TableRow({
      children: [reportCell(content, {shading: 'FFFFFF', margins: {top: 90, bottom: 90, left: 45, right: 45}})]
    })
  ]);

  return new Document({
    creator: 'Bitácora de Issues / QA',
    title: `Informe de pruebas - ${safeText(issue.testid)}`,
    description: 'Informe de evidencia generado desde la Bitácora de Issues.',
    sections: [{
      properties: {
        page: {
          size: {width: 12240, height: 15840},
          margin: {top: 720, right: 800, bottom: 720, left: 800}
        }
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            border: {top: {style: BorderStyle.SINGLE, size: 4, color: COLORS.line, space: 5}},
            children: [
              new TextRun({text: 'Bitácora de Issues / QA   ·   Página ', font: 'Arial', size: 16, color: COLORS.muted}),
              PageNumber.CURRENT
            ]
          })]
        })
      },
      children: [reportTable]
    }]
  });
}

export function issueReportFileName(issue) {
  const code = safeText(issue?.testid, 'issue')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9_-]+/g, '-')
    .replace(/^-+|-+$/g, '') || 'issue';
  const today = new Date();
  const date = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
  return `informe-${code}-${date}.docx`;
}

export async function createIssueReportBlob(issue, evidence = [], headerTemplateBytes) {
  const report = await Packer.toBlob(await createIssueReportDocument(issue, evidence));
  return applyOriginalHeader(report, headerTemplateBytes);
}
