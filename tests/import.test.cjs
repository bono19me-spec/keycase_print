const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const XLSX = require('../vendor/xlsx.full.min.js');
const source = fs.readFileSync(path.join(__dirname, '../app.js'), 'utf8');

function setup(rows) {
  const sheet = XLSX.utils.aoa_to_sheet([]);
  XLSX.utils.sheet_add_aoa(sheet, rows.map(([room, name]) => [room, null, null, name]), { origin: 'U13' });
  const elements = new Map();
  const context = vm.createContext({
    XLSX, sheet,
    localStorage: { getItem: () => null },
    document: {
      getElementById(id) {
        if (!elements.has(id)) elements.set(id, { value: '', classList: { toggle() {} } });
        return elements.get(id);
      }
    }
  });
  vm.runInContext(source.replace('\ninit();', '\n'), context);
  const run = (code) => vm.runInContext(code, context);
  run('const parsed = extractRecords(sheet);');
  return { run, elements };
}

test('room-only Excel rows remain printable without a name or honorific in every copy mode', () => {
  const { run } = setup([['701', ''], ['702', '山田 太郎']]);
  assert.equal(run('parsed.valid.length'), 2);
  assert.equal(run('parsed.valid[0].outputNames.length'), 0);
  assert.equal(run('parsed.warnings.length'), 1);
  assert.match(run('parsed.warnings[0].reason'), /部屋番号のみ印刷/);
  assert.equal(run('parsed.warnings[0].room'), '701');
  for (const mode of ['room', 'guest', 'individual']) {
    run(`settings.printCopyMode = '${mode}';`);
    assert.equal(run('getPrintableRecords(parsed.valid, "all").length'), 2);
    assert.equal(run('getPrintPageCount(parsed.valid)'), 2);
    const html = run('buildPrintHtml(getPrintableRecords([parsed.valid[0]], "all"))');
    assert.match(html, /<p class="room">701<\/p>/);
    assert.doesNotMatch(html, /<p class="name"/);
    assert.doesNotMatch(html, /様/);
  }
});

test('print paper mode keeps B6 unchanged and centers it horizontally on A4 landscape', () => {
  const { run } = setup([['701', '山田 太郎']]);
  assert.equal(run('getPrintPageConfig().widthMm'), 182);
  assert.equal(run('getPrintPageConfig().heightMm'), 128);
  assert.equal(run('getPrintPageConfig().canvasOffsetXMm'), 0);
  assert.equal(run('getPrintPageConfig().canvasOffsetYMm'), 0);

  run("settings.printPaperSize = 'a4';");
  assert.equal(run('getPrintPageConfig().widthMm'), 297);
  assert.equal(run('getPrintPageConfig().heightMm'), 210);
  assert.equal(run('getPrintPageConfig().canvasOffsetXMm'), 57.5);
  assert.equal(run('getPrintPageConfig().canvasOffsetYMm'), 0);
  const html = run('buildPrintHtml(getPrintableRecords(parsed.valid, "all"))');
  assert.match(html, /size: 297mm 210mm/);
  assert.match(html, /left: 57.5mm/);
  assert.match(html, /用紙 A4（横）/);
});

test('continuation names attach to the latest room even when its first name cell is empty', () => {
  const { run } = setup([['701', '山田 太郎'], ['702', ''], ['', '山田 花子']]);
  assert.equal(run('parsed.valid.length'), 2);
  assert.equal(run('parsed.valid[0].rawName'), '山田 太郎');
  assert.equal(run('parsed.valid[1].room'), '702');
  assert.equal(run('parsed.valid[1].rawName'), '山田 花子');
  assert.equal(run('parsed.warnings.length'), 0);
});

test('a name before the first room is still excluded with a warning', () => {
  const { run } = setup([['', '山田 太郎'], ['701', '']]);
  assert.equal(run('parsed.valid.length'), 1);
  assert.equal(run('parsed.valid[0].room'), '701');
  assert.equal(run('parsed.warnings.length'), 2);
  assert.equal(run('parsed.warnings[0].reason'), '部屋番号が空です');
});

test('a sheet containing only room numbers can proceed and displays the print warning', () => {
  const { run, elements } = setup([['701', ''], ['702', '']]);
  run(`
    currentWorkbook = { SheetNames: ['名簿'], Sheets: { '名簿': sheet } };
    resetPrintSelection = () => {};
    renderTable = () => {};
    parseSelectedSheet();
  `);
  assert.equal(run('records.length'), 2);
  assert.equal(elements.get('simpleNextUpload').disabled, false);
  assert.match(elements.get('simpleStatus').textContent, /警告 2件/);
  assert.match(elements.get('simpleStatus').textContent, /部屋番号のみ印刷/);
});
