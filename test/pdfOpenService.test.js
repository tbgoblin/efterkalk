const test = require('node:test');
const assert = require('node:assert/strict');
const { normalizePdfTarget, openPdfTarget } = require('../services/pdfOpenService');

test('accepts absolute drive, UNC and HTTPS PDF targets', () => {
    assert.deepEqual(normalizePdfTarget('C:/drawings/part.PDF'), {
        kind: 'path',
        value: 'C:\\drawings\\part.PDF'
    });
    assert.deepEqual(normalizePdfTarget('\\\\fileserver\\drawings\\part.pdf'), {
        kind: 'path',
        value: '\\\\fileserver\\drawings\\part.pdf'
    });
    assert.deepEqual(normalizePdfTarget('file://fileserver/drawings/part.pdf'), {
        kind: 'path',
        value: '\\\\fileserver\\drawings\\part.pdf'
    });
    assert.equal(
        normalizePdfTarget('https://example.test/drawings/part.pdf?revision=2').kind,
        'url'
    );
});

test('rejects relative, non-PDF and command-shaped targets', () => {
    assert.throws(() => normalizePdfTarget('drawings\\part.pdf'), /absolut PDF-sti/);
    assert.throws(() => normalizePdfTarget('C:\\drawings\\part.pdf.exe'), /absolut PDF-sti/);
    assert.throws(() => normalizePdfTarget('C:\\drawings\\part.pdf\r\ncalc.exe'), /Ugyldig PDF-sti/);
    assert.throws(() => normalizePdfTarget('ftp://example.test/part.pdf'), /absolut PDF-sti/);
});

test('Electron opens UNC paths with the default viewer', async () => {
    const calls = [];
    const target = '\\\\fileserver\\drawings\\part.pdf';

    await openPdfTarget(target, {
        openPath: async value => {
            calls.push(value);
            return '';
        }
    });

    assert.deepEqual(calls, [target]);
});

test('plain Node fallback uses explorer without a shell and passes one exact argument', async () => {
    const calls = [];
    let unrefCalled = false;
    const target = 'C:\\drawings\\report & final.pdf';

    await openPdfTarget(target, {
        spawn(command, args, options) {
            calls.push({ command, args, options });
            return { unref() { unrefCalled = true; } };
        }
    });

    assert.equal(calls[0].command, 'explorer.exe');
    assert.deepEqual(calls[0].args, [target]);
    assert.equal(calls[0].options.shell, false);
    assert.equal(unrefCalled, true);
});

test('viewer errors are propagated to the API layer', async () => {
    await assert.rejects(
        openPdfTarget('C:\\drawings\\part.pdf', { openPath: async () => 'No application is associated' }),
        /No application is associated/
    );
});
