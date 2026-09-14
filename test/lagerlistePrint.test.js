const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

const source = fs.readFileSync(path.join(__dirname, '../assets/js/lagerliste.js'), 'utf8');
const printSource = source.slice(source.indexOf('function exportLagerlistePdf()'), source.indexOf('async function refreshLagerlisteSnapshotList()'));

function setup(failPrint = false) {
    const calls = [];
    let html = '';
    const section = { style: {} };
    const detail = { style: {} };
    const clone = {
        innerHTML: '<table><tr><td>August snapshot 123,45</td></tr></table>',
        querySelectorAll(selector) {
            if (selector.includes('section')) return [section];
            if (selector.includes('detail')) return [detail];
            return [{ remove: () => calls.push('tools removed') }];
        }
    };
    const handlers = {};
    const frame = {
        style: {}, setAttribute() {}, remove: () => calls.push('removed'),
        contentWindow: {
            document: { open() {}, write(value) { html = value; }, close() {}, fonts: { ready: Promise.resolve() } },
            addEventListener(event, handler) { handlers[event] = handler; },
            focus() {}, print() { calls.push('print'); if (failPrint) throw new Error('printer unavailable'); }
        }
    };
    const context = vm.createContext({
        document: {
            getElementById(id) { return id === 'lagerlisteResults' ? { cloneNode: () => clone } : null; },
            createElement(tag) { assert.equal(tag, 'iframe'); return frame; },
            body: { appendChild: () => calls.push('attached') }
        },
        window: { open() { throw new Error('Popup must not be used'); } },
        lagerlisteDisplayedLabel: '2026-08', lagerlisteEscape: value => value,
        alert: message => calls.push(message)
    });
    vm.runInContext(printSource + '\nexportLagerlistePdf();', context);
    return { frame, calls, handlers, section, detail, html };
}

test('Lagerliste prints selected snapshot without popups, then cleans up after print/cancel', async () => {
    const result = setup();
    assert.match(result.html, /Lagerliste - 2026-08/);
    assert.match(result.html, /August snapshot 123,45/);
    assert.equal(result.section.style.display, 'block');
    assert.equal(result.detail.style.display, 'table-row');
    assert.ok(!result.calls.includes('print'));
    await result.frame.onload();
    assert.ok(result.calls.includes('print'));
    result.handlers.afterprint();
    assert.equal(result.calls.at(-1), 'removed');
});

test('Lagerliste cleans up and reports a print failure', async () => {
    const result = setup(true);
    await result.frame.onload();
    assert.ok(result.calls.includes('removed'));
    assert.match(result.calls.at(-1), /printer unavailable/);
});
