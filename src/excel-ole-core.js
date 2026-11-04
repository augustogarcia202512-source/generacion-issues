import * as CFB from 'cfb';

const PNG_SIGNATURE = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);

function findBytes(data, signature, start = 0) {
  for (let index = start; index <= data.length - signature.length; index++) {
    let matches = true;
    for (let offset = 0; offset < signature.length; offset++) {
      if (data[index + offset] !== signature[offset]) {
        matches = false;
        break;
      }
    }
    if (matches) return index;
  }
  return -1;
}

function replaceBytes(data, search, replacement) {
  if (search.length !== replacement.length) {
    throw new Error('Los nombres de archivo del objeto OLE deben conservar su longitud.');
  }
  const result = new Uint8Array(data);
  let index = 0;
  while ((index = findBytes(result, search, index)) >= 0) {
    result.set(replacement, index);
    index += replacement.length;
  }
  return result;
}

function fileExtensionBytes(extension, encoding = 'ascii') {
  const text = `.${extension}`;
  if (encoding === 'utf16le') {
    const bytes = new Uint8Array(text.length * 2);
    for (let index = 0; index < text.length; index++) {
      bytes[index * 2] = text.charCodeAt(index);
    }
    return bytes;
  }
  return new TextEncoder().encode(text);
}

export function createOlePackage(imageBytes, extension, templateBytes) {
  if (!(imageBytes instanceof Uint8Array) || imageBytes.length === 0) {
    throw new TypeError('La evidencia para el objeto OLE está vacía o no es válida.');
  }
  const normalizedExtension = extension === 'jpg' ? 'jpg' : extension === 'png' ? 'png' : null;
  if (!normalizedExtension) {
    throw new TypeError('El objeto OLE solo acepta evidencias PNG o JPG.');
  }

  const signature = normalizedExtension === 'png'
    ? PNG_SIGNATURE
    : new Uint8Array([0xff, 0xd8, 0xff]);
  if (findBytes(imageBytes, signature) !== 0) {
    throw new TypeError(`El contenido de la evidencia no coincide con el formato ${normalizedExtension.toUpperCase()}.`);
  }

  const compoundFile = CFB.read(templateBytes, {type: 'array'});
  const nativeStream = compoundFile.FileIndex.find(entry => entry.name === '\u0001Ole10Native');
  if (!nativeStream || !nativeStream.content) {
    throw new Error('La plantilla OLE no contiene el archivo de evidencia esperado.');
  }

  const oldStream = Uint8Array.from(nativeStream.content);
  const imageOffset = findBytes(oldStream, PNG_SIGNATURE);
  if (imageOffset < 4) {
    throw new Error('No se encontró la evidencia de plantilla dentro del objeto OLE.');
  }
  const oldImageLength = new DataView(oldStream.buffer, oldStream.byteOffset, oldStream.byteLength)
    .getUint32(imageOffset - 4, true);
  const tailOffset = imageOffset + oldImageLength;
  if (tailOffset > oldStream.length) {
    throw new Error('La evidencia de plantilla del objeto OLE está incompleta.');
  }

  let prefix = oldStream.slice(0, imageOffset - 4);
  let tail = oldStream.slice(tailOffset);
  if (normalizedExtension === 'jpg') {
    prefix = replaceBytes(prefix, fileExtensionBytes('png'), fileExtensionBytes('jpg'));
    tail = replaceBytes(tail, fileExtensionBytes('png', 'utf16le'), fileExtensionBytes('jpg', 'utf16le'));
  }

  const updatedStream = new Uint8Array(prefix.length + 4 + imageBytes.length + tail.length);
  updatedStream.set(prefix);
  const view = new DataView(updatedStream.buffer);
  view.setUint32(prefix.length, imageBytes.length, true);
  updatedStream.set(imageBytes, prefix.length + 4);
  updatedStream.set(tail, prefix.length + 4 + imageBytes.length);
  view.setUint32(0, updatedStream.length - 4, true);

  CFB.utils.cfb_add(compoundFile, nativeStream.name, updatedStream);
  const packageBytes = CFB.write(compoundFile, {type: 'array', fileType: 'cfb'});
  return packageBytes instanceof Uint8Array ? packageBytes : new Uint8Array(packageBytes);
}
