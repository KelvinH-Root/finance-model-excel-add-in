import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, existsSync } from 'node:fs';
import { deflateRawSync, crc32 } from 'node:zlib';
import vm from 'node:vm';

// Load src/zip.js the way the task pane does (a classic script that sets globalThis.HfgZip).
function loadZip() {
  const context = vm.createContext({ TextDecoder, DecompressionStream, Blob, Response, Uint8Array, DataView, Buffer, btoa });
  vm.runInContext(readFileSync(new URL('../src/zip.js', import.meta.url), 'utf8'), context);
  return context.HfgZip;
}

// Build a small zip in memory: one deflated entry and one stored entry.
function makeZip(files) {
  const locals = [];
  const centrals = [];
  let offset = 0;
  for (const { name, data, deflate } of files) {
    const nameBytes = Buffer.from(name, 'utf8');
    const raw = Buffer.from(data, 'utf8');
    const body = deflate ? deflateRawSync(raw) : raw;
    const crc = crc32(raw);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0); local.writeUInt16LE(20, 4); local.writeUInt16LE(0, 6);
    local.writeUInt16LE(deflate ? 8 : 0, 8); local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(body.length, 18); local.writeUInt32LE(raw.length, 22);
    local.writeUInt16LE(nameBytes.length, 26); local.writeUInt16LE(0, 28);
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0); central.writeUInt16LE(20, 4); central.writeUInt16LE(20, 6);
    central.writeUInt16LE(deflate ? 8 : 0, 10); central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(body.length, 20); central.writeUInt32LE(raw.length, 24);
    central.writeUInt16LE(nameBytes.length, 28); central.writeUInt32LE(offset, 42);
    locals.push(local, nameBytes, body);
    centrals.push(central, nameBytes);
    offset += 30 + nameBytes.length + body.length;
  }
  const centralBytes = Buffer.concat(centrals);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(files.length, 8); end.writeUInt16LE(files.length, 10);
  end.writeUInt32LE(centralBytes.length, 12); end.writeUInt32LE(offset, 16);
  const all = Buffer.concat([...locals, centralBytes, end]);
  return all.buffer.slice(all.byteOffset, all.byteOffset + all.byteLength);
}

const WORKBOOK = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Cover" sheetId="1" r:id="rId1"/><sheet name="Time" sheetId="2" state="hidden" r:id="rId2"/><sheet name="Sites &amp; Stages" sheetId="3" r:id="rId3"/></sheets></workbook>`;

test('lists sheets in order with their visibility', async () => {
  const zip = loadZip();
  const buffer = makeZip([
    { name: '[Content_Types].xml', data: '<Types/>', deflate: false },
    { name: 'xl/workbook.xml', data: WORKBOOK, deflate: true }
  ]);
  const sheets = await zip.listSheets(buffer);
  assert.deepEqual(JSON.parse(JSON.stringify(sheets.map(s => [s.name, s.state]))), [['Cover', 'visible'], ['Time', 'hidden'], ['Sites & Stages', 'visible']]);
  assert.deepEqual(JSON.parse(JSON.stringify(zip.listParts(buffer))), ['[Content_Types].xml', 'xl/workbook.xml']);
});

test('reports a file that is not a workbook', async () => {
  const zip = loadZip();
  const buffer = makeZip([{ name: 'readme.txt', data: 'hello', deflate: false }]);
  await assert.rejects(zip.listSheets(buffer), /xl\/workbook\.xml/);
  await assert.rejects(zip.listSheets(new ArrayBuffer(10)), /Not a zip file/);
});

test('base64 matches Node for binary data', () => {
  const zip = loadZip();
  const bytes = new Uint8Array(70000).map((_, i) => (i * 31) % 256);
  assert.equal(zip.base64FromBuffer(bytes.buffer), Buffer.from(bytes).toString('base64'));
});

test('reads a real template when one is available locally', async t => {
  const path = process.env.HFG_TEMPLATE;
  if (!path || !existsSync(path)) { t.skip('Set HFG_TEMPLATE to a local copy of the template'); return; }
  const zip = loadZip();
  const file = readFileSync(path);
  const sheets = await zip.listSheets(file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength));
  assert.ok(sheets.length > 5);
  assert.ok(sheets.some(s => s.name === 'Time'));
});
