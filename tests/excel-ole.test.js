import assert from 'node:assert/strict';
import fs from 'node:fs';
import * as CFB from 'cfb';
import {createOlePackage} from '../src/excel-ole-core.js';

const template = fs.readFileSync(new URL('../src/ole-package-template.bin', import.meta.url));
const templateCfb = CFB.read(template, {type: 'buffer'});
const templateNative = templateCfb.FileIndex.find(entry => entry.name === '\u0001Ole10Native');
const templateStream = Buffer.from(templateNative.content);
const templateImageOffset = templateStream.indexOf(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]));
const templateImageLength = templateStream.readUInt32LE(templateImageOffset - 4);
const samplePng = templateStream.subarray(templateImageOffset, templateImageOffset + templateImageLength);
const sampleJpg = Buffer.from([0xff, 0xd8, 0xff, 0x11, 0x22, 0xff, 0xd9]);

function readNativePayload(packageBytes, expectedExtension) {
  const cfb = CFB.read(packageBytes, {type: 'buffer'});
  const native = cfb.FileIndex.find(entry => entry.name === '\u0001Ole10Native');
  assert.ok(native, 'Package OLE contains Ole10Native stream');
  const stream = Buffer.from(native.content);
  const signature = expectedExtension === 'jpg'
    ? Buffer.from([0xff, 0xd8, 0xff])
    : Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
  const offset = stream.indexOf(signature);
  assert.ok(offset >= 4, 'Package OLE contains the expected image signature');
  const length = stream.readUInt32LE(offset - 4);
  assert.equal(stream.readUInt32LE(0), stream.length - 4);
  return {stream, payload: stream.subarray(offset, offset + length)};
}

const pngPackage = createOlePackage(samplePng, 'png', template);
const pngResult = readNativePayload(pngPackage, 'png');
assert.deepEqual(pngResult.payload, samplePng);

const jpgPackage = createOlePackage(sampleJpg, 'jpg', template);
const jpgResult = readNativePayload(jpgPackage, 'jpg');
assert.deepEqual(jpgResult.payload, sampleJpg);
assert.ok(Buffer.from(jpgResult.stream).includes(Buffer.from('.jpg')));
assert.ok(Buffer.from(jpgResult.stream).includes(Buffer.from('.jpg', 'utf16le')));

assert.throws(() => createOlePackage(sampleJpg, 'gif', template), /PNG o JPG/);
assert.throws(() => createOlePackage(new Uint8Array(), 'png', template), /vacía/);

console.log('Excel OLE package tests passed.');
