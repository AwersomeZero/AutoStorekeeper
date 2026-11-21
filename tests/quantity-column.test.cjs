const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { test } = require('node:test');

// Run the actual page handlers with a minimal DOM and an inspectable Excel adapter.
function page() {
  function element() {
    return {
      textContent: '', innerHTML: '', style: {}, children: [], events: {},
      classList: { add() {}, remove() {} },
      appendChild(child) { this.children.push(child); },
      addEventListener(type, handler) { this.events[type] = handler; }
    };
  }
  const nodes = new Map();
  const get = id => {
    if (!nodes.has(id)) nodes.set(id, element());
    return nodes.get(id);
  };
  const inputs = ['Total', 'Missing'].map(value => Object.assign(element(), { value }));
  const readers = [];
  let saved;
  const XLSX = {
    utils: {
      json_to_sheet: (rows, options) => ({ rows, headers: options.header }),
      book_new: () => ({}),
      book_append_sheet: (book, sheet) => { book.sheet = sheet; }
    },
    writeFile: book => { saved = book.sheet; }
  };
  const window = {
    XLSX, SORTING_LIST_DATA: [{ en: 'Stone', ru: 'Камень', size: 64 }],
    addEventListener: (_, handler) => handler()
  };
  const context = {
    window, XLSX, console: { error() {} },
    document: {
      getElementById: get, querySelectorAll: () => inputs,
      createElement: element, createDocumentFragment: element,
      createTextNode: text => ({ textContent: text })
    },
    FileReader: class {
      readAsText(file) { readers.push(() => this.onload({ target: { result: file.text } })); }
    }
  };
  vm.runInNewContext(fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8'), context);
  return {
    get,
    upload(text, name = 'materials.txt') {
      get('fileInput').events.change({ target: { files: [{ name, text, size: text.length }] } });
    },
    finish(index = 0) { readers.splice(index, 1)[0](); },
    select(value) {
      const input = inputs.find(input => input.value === value);
      input.checked = true;
      input.events.change();
    },
    download() {
      saved = undefined;
      get('downloadBtn').events.click();
      return saved;
    }
  };
}

test('switching TXT recalculates counts, storage, preview header and Excel', () => {
  const p = page();
  p.upload('| Item | Available | Missing | Total |\n| Stone | 5 | 65 | 3,520 |');
  p.finish();
  assert.equal(p.download().rows[0].Total, 3520);
  assert.equal(p.get('statChests').textContent, 2);
  assert.equal(p.get('statShulkers').textContent, 3);
  p.select('Missing');
  assert.equal(p.get('downloadBtn').disabled, true);
  assert.equal(p.download(), undefined);
  p.finish();
  const sheet = p.download();
  assert.equal(sheet.rows[0].Missing, 65);
  assert.equal(sheet.rows[0].Total, undefined);
  assert.equal(sheet.headers[2], 'Missing');
  assert.equal(p.get('statChests').textContent, 1);
  assert.equal(p.get('statShulkers').textContent, 1);
  assert.match(p.get('quantityHeader').textContent, /Missing/);
  p.select('Total');
  p.finish();
  assert.equal(p.download().rows[0].Total, 3520);
});

test('switching TXT to JSON hides the selector and uses Total, then restores TXT selection', () => {
  const p = page();
  const text = '| Item | Total | Missing |\n| Stone | 500 | 10 |';
  p.upload(text, 'materials.TXT');
  p.finish();
  assert.equal(p.get('quantitySelector').hidden, false);
  p.select('Missing');
  p.finish();
  assert.equal(p.download().rows[0].Missing, 10);
  for (const json of [
    '{"Materials":[{"Item":"Stone","Total":500}]}',
    '[{"Item":"Stone","Total":500}]',
    '{"Stone":500}'
  ]) {
    p.upload(json, 'materials.JSON');
    assert.equal(p.get('quantitySelector').hidden, true);
    p.finish();
    assert.equal(p.download().rows[0].Total, 500);
    assert.equal(p.download().headers[2], 'Total');
    assert.equal(p.get('quantityHeader').textContent, 'Всего (Total)');
  }
  p.upload(text);
  p.finish();
  assert.equal(p.get('quantitySelector').hidden, false);
  assert.equal(p.download().rows[0].Missing, 10);
});

test('TXT rejects alternative names for Missing', () => {
  for (const field of ['Missed', 'missed', 'missing', 'Пропущено', 'недостающие', 'недостаёт', 'недостает']) {
    for (const [text, name] of [
      [`| Item | Total | ${field} |\n| Stone | 500 | 10 |`, 'materials.txt']
    ]) {
      const p = page();
      p.select('Missing');
      p.upload(text, name);
      p.finish();
      assert.equal(p.get('downloadBtn').disabled, true);
      assert.equal(p.download(), undefined);
    }
  }
});

test('missing source column clears previous results without falling back to Total', () => {
  for (const [text, name] of [
    ['| Item | Total |\n| Stone | 100 |', 'materials.txt']
  ]) {
    const p = page();
    p.upload(text, name);
    p.finish();
    assert.equal(p.download().rows[0].Total, 100);
    p.select('Missing');
    p.finish();
    assert.equal(p.get('downloadBtn').disabled, true);
    assert.equal(p.download(), undefined);
    assert.equal(p.get('statRows').textContent, '0');
    assert.equal(p.get('dataTable').style.display, 'none');
  }
});

test('a late file read cannot overwrite the latest selected column', () => {
  const p = page();
  p.upload('| Item | Total | Missing |\n| Stone | 500 | 10 |');
  p.select('Missing');
  p.finish(1);
  p.finish(0);
  assert.equal(p.download().rows[0].Missing, 10);
});

test('repository sample parses its Missing column', () => {
  const text = fs.readFileSync(path.join(__dirname, '..', '_internal', 'lists', 'material_list_2025-11-13_18.36.46.txt'), 'utf8');
  const p = page();
  p.select('Missing');
  p.upload(text);
  p.finish();
  assert.ok(p.download().rows.some(row => row.Missing === 37676));
});
