/*
 * Minimal read-only zip reader for the probe: lists the sheets in an .xlsx
 * so the user can choose which ones to insert. Works in the task pane
 * (WebView2, WKWebView, browsers) and in Node 18+ for tests.
 */
(function (root) {
  'use strict';

  function u16(view, at) { return view.getUint16(at, true); }
  function u32(view, at) { return view.getUint32(at, true); }

  function findEndOfCentralDirectory(view) {
    const min = Math.max(0, view.byteLength - 65557);
    for (let i = view.byteLength - 22; i >= min; i--) {
      if (u32(view, i) === 0x06054b50) return i;
    }
    throw new Error('Not a zip file: end of central directory not found');
  }

  function entries(buffer) {
    const view = new DataView(buffer);
    const eocd = findEndOfCentralDirectory(view);
    const count = u16(view, eocd + 10);
    let at = u32(view, eocd + 16);
    const decoder = new TextDecoder('utf-8');
    const list = [];
    for (let n = 0; n < count; n++) {
      if (u32(view, at) !== 0x02014b50) throw new Error('Bad central directory entry');
      const method = u16(view, at + 10);
      const compressedSize = u32(view, at + 20);
      const size = u32(view, at + 24);
      const nameLength = u16(view, at + 28);
      const extraLength = u16(view, at + 30);
      const commentLength = u16(view, at + 32);
      const localOffset = u32(view, at + 42);
      const name = decoder.decode(new Uint8Array(buffer, at + 46, nameLength));
      list.push({ name, method, compressedSize, size, localOffset });
      at += 46 + nameLength + extraLength + commentLength;
    }
    return list;
  }

  async function inflateRaw(bytes) {
    if (typeof DecompressionStream === 'undefined') {
      throw new Error('This Excel build has no DecompressionStream, so the sheet list cannot be read');
    }
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  async function readEntry(buffer, entry) {
    const view = new DataView(buffer);
    const at = entry.localOffset;
    if (u32(view, at) !== 0x04034b50) throw new Error('Bad local header for ' + entry.name);
    const start = at + 30 + u16(view, at + 26) + u16(view, at + 28);
    const raw = new Uint8Array(buffer, start, entry.compressedSize);
    if (entry.method === 0) return raw;
    if (entry.method === 8) return inflateRaw(raw);
    throw new Error('Unsupported compression method ' + entry.method);
  }

  function decodeXml(text) {
    return text
      .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
      .replace(/&amp;/g, '&');
  }

  function attribute(tag, name) {
    const match = new RegExp('\\s' + name + '="([^"]*)"').exec(tag);
    return match ? decodeXml(match[1]) : null;
  }

  /** Returns [{name, state}] for every sheet in workbook order. */
  async function listSheets(buffer) {
    const entry = entries(buffer).find(e => e.name === 'xl/workbook.xml');
    if (!entry) throw new Error('No xl/workbook.xml: is this an .xlsx or .xlsm file?');
    const xml = new TextDecoder('utf-8').decode(await readEntry(buffer, entry));
    const tags = xml.match(/<(?:\w+:)?sheet\s[^>]*>/g) || [];
    return tags.map(tag => ({ name: attribute(tag, 'name'), state: attribute(tag, 'state') || 'visible' }));
  }

  /** Lists every part name, for reporting what a file carries. */
  function listParts(buffer) {
    return entries(buffer).map(e => e.name);
  }

  function base64FromBuffer(buffer) {
    const bytes = new Uint8Array(buffer);
    let binary = '';
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
      binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    }
    return (typeof btoa === 'function') ? btoa(binary) : Buffer.from(binary, 'binary').toString('base64');
  }

  const api = { listSheets, listParts, base64FromBuffer, entries };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.HfgZip = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
