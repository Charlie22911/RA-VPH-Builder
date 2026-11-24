const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const { join } = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const html = readFileSync(process.env.RA_VPH_TEST_HTML || join(__dirname, '..', 'RA-VPH-Summary-Builder-Standalone.html'), 'utf8');
const scripts = [...html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g)].map(match => match[1]);
const headers = 'Asset ID,System Acronym,Script Name,Severity,IP Address,Last Detected Date,Remedy Available,Remediation Date,Plugin ID';
const row = (system, id) => `1,${system},Same title,High,192.0.2.1,2026-11-20,TRUE,2026-12-04,${id}`;
const csv = (name, ...lines) => ({ name, async arrayBuffer() { return new TextEncoder().encode(lines.join('\n')).buffer; } });

function openBuilder() {
  const elements = new Map();
  const createElement = () => ({ value: '', textContent: '', style: {}, classList: { toggle() {} }, addEventListener() {}, append() {}, replaceChildren() {} });
  const getElementById = id => {
    if (!elements.has(id)) elements.set(id, { ...createElement(), checked: id === 'filterRA', querySelector: tag => getElementById(tag) });
    return elements.get(id);
  };
  const context = vm.createContext({
    document: { createElement, getElementById, querySelectorAll: () => [] }, window: { addEventListener() {} }, console,
    ArrayBuffer, Uint8Array, Blob, Response, TextDecoder, TextEncoder, CompressionStream, DecompressionStream
  });
  scripts.forEach(script => vm.runInContext(script, context));
  const downloads = [];
  context.triggerBlobDownload = (blob, filename) => downloads.push({ blob, filename });
  return { context, downloads, getElementById };
}

async function readExport(builder) {
  await builder.context.exportWorkbook();
  return builder.downloads.at(-1).blob.arrayBuffer();
}

test('a failed upload keeps the previous batch and names the missing columns', async () => {
  const builder = openBuilder();
  await builder.context.handleFiles([csv('valid.csv', headers, row('First', 101))]);
  await builder.context.handleFiles([csv('invalid.csv', 'Asset ID,Severity', '1,High')]);
  assert.match(builder.getElementById('statusBar').textContent, /invalid\.csv.*Plugin ID.*Remedy Available/);
  const workbook = builder.context.XLSX.read(new Uint8Array(await readExport(builder)), { type: 'array' });
  assert.equal(workbook.Sheets.ALL.P2.v, 101);
});

test('the most recent upload wins and its date is independent of an earlier batch', async () => {
  const builder = openBuilder();
  await builder.context.handleFiles([csv('dated.csv', 'CUI 2024-01-02', headers, row('Dated', 99))]);
  let finishFirst;
  const slowUpload = builder.context.handleFiles([{ name: 'slow.csv', arrayBuffer: () => new Promise(resolve => { finishFirst = resolve; }) }]);
  await builder.context.handleFiles([csv('new.csv', headers, row('New', 202))]);
  finishFirst(await csv('slow.csv', headers, row('Slow', 303)).arrayBuffer());
  await slowUpload;
  assert.equal(builder.getElementById('bannerInfo').textContent, '');
  assert.ok(!builder.getElementById('outFilename').value.includes('02JAN24'));
  const workbook = builder.context.XLSX.read(new Uint8Array(await readExport(builder)), { type: 'array' });
  assert.equal(workbook.Sheets.ALL.P2.v, 202);
  assert.equal(workbook.Sheets.ALL.P3, undefined);
});

test('a multi-file drop keeps every file after the drop data store is released', async () => {
  const builder = openBuilder();
  const second = csv('second.csv', headers, row('Test', 202));
  let files;
  const first = {
    name: 'first.csv',
    async arrayBuffer() {
      await Promise.resolve();
      files.pop();
      return csv('first.csv', headers, row('Test', 101)).arrayBuffer();
    }
  };
  files = [first, second];
  await builder.context.handleFiles(files);
  const workbook = builder.context.XLSX.read(new Uint8Array(await readExport(builder)), { type: 'array' });
  assert.equal(workbook.Sheets.ALL.P3?.v, 202);
});

test('bulk reports unassigned rows and retains every named system with colliding filenames', async () => {
  const builder = openBuilder();
  await builder.context.handleFiles([csv('systems.csv', headers, row('A/B', 101), row('A:B', 202), row('a_b', 303), row('', 404), row('Café', 505))]);
  assert.match(builder.getElementById('bulkWarning').textContent, /1 row.*without a System Acronym/i);
  const archive = await readExport(builder);
  const entries = await builder.context.readZipEntries(archive);
  assert.equal(entries.size, 4);
  assert.equal(new Set([...entries.keys()].map(name => name.toLowerCase())).size, entries.size);
  assert.ok([...entries.keys()].some(name => name.includes('Café')));
  const view = new DataView(archive);
  let offset = view.getUint32(archive.byteLength - 6, true);
  let sawUtf8 = false;
  for (const name of entries.keys()) {
    if (name.includes('Café')) {
      assert.equal(view.getUint16(offset + 8, true) & 2048, 2048);
      assert.equal(view.getUint16(view.getUint32(offset + 42, true) + 6, true) & 2048, 2048);
      sawUtf8 = true;
    }
    offset += 46 + view.getUint16(offset + 28, true) + view.getUint16(offset + 30, true) + view.getUint16(offset + 32, true);
  }
  assert.equal(sawUtf8, true);
  const ids = [...entries.values()].map(bytes => builder.context.XLSX.read(bytes, { type: 'array' }).Sheets.ALL.P2.v);
  assert.deepEqual(ids.sort((a, b) => a - b), [101, 202, 303, 505]);
});

test('POAM counts are keyed to the row Plugin ID and the workbook keeps its eight-sheet layout', async () => {
  const builder = openBuilder();
  await builder.context.handleFiles([csv('counts.csv', headers, row('Test', 101), row('Test', 101), row('Test', 202))]);
  const buffer = await readExport(builder);
  const workbook = builder.context.XLSX.read(new Uint8Array(buffer), { type: 'array' });
  assert.deepEqual(Array.from(workbook.SheetNames), ['Summary', 'Per Asset', 'ALL', 'Crit', 'High', 'Med', 'Low', 'POAM']);
  const poam = workbook.Sheets.POAM;
  assert.equal(poam['!ref'], 'A1:I3');
  assert.deepEqual([poam.C2.v, poam.C3.v], [2, 1]);
  assert.deepEqual([poam.C2.f, poam.C3.f], ['COUNTIF(ALL!P:P, 101)', 'COUNTIF(ALL!P:P, 202)']);
  assert.equal(workbook.Sheets.Summary.C2.f, 'COUNTIF(ALL!P:P, F2)');
  const entries = await builder.context.readZipEntries(buffer);
  const decoder = new TextDecoder();
  assert.match(decoder.decode(entries.get('xl/tables/table8.xml')), /ref="A1:I3".*TableStyleMedium15/);
  assert.match(decoder.decode(entries.get('xl/styles.xml')), /vertical="top" wrapText="1"/);
});

test('CSV dates and formatted XLSX dates stay readable in the exported sheets', async () => {
  const builder = openBuilder();
  const xlsx = builder.context.XLSX;
  const input = xlsx.utils.book_new();
  input.Workbook = { WBProps: { date1904: true } };
  const dateCell = { t: 'n', v: 44884, z: 'yyyy-mm-dd' };
  const data = [headers.split(','), [1, 'Test', 'Typed dates', 'High', '192.0.2.1', dateCell, true, dateCell, 202]];
  xlsx.utils.book_append_sheet(input, xlsx.utils.aoa_to_sheet(data), 'VRAM');
  xlsx.utils.book_append_sheet(input, xlsx.utils.aoa_to_sheet([['Unused']]), 'Other');
  await builder.context.handleFiles([
    csv('dates.csv', headers, row('Test', 101)),
    { name: 'dates.xlsx', async arrayBuffer() { return xlsx.write(input, { type: 'array', bookType: 'xlsx' }); } }
  ]);
  const workbook = xlsx.read(new Uint8Array(await readExport(builder)), { type: 'array' });
  const sheet = workbook.Sheets.ALL;
  assert.deepEqual([sheet.K2.v, sheet.O2.v, sheet.K3.v, sheet.O3.v], ['2026-11-20', '2026-12-04', '2026-11-20', '2026-11-20']);
});
