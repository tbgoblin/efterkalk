const test = require('node:test');
const assert = require('node:assert/strict');
const zlib = require('zlib');
const { createBomService } = require('../services/bomService');

function createCompressedPdf(content) {
    const compressed = zlib.deflateSync(Buffer.from(content, 'latin1'));
    return Buffer.concat([
        Buffer.from('%PDF-1.4\n1 0 obj\n<< /Type /Page /MediaBox [0 0 595 842] /Contents 2 0 R >>\nendobj\n2 0 obj\n<< /Length ' + compressed.length + ' /Filter /FlateDecode >>\nstream\n', 'latin1'),
        compressed,
        Buffer.from('\nendstream\nendobj\n%%EOF', 'latin1')
    ]);
}

test('compressed PDF yields dimensions, page size and preview vectors', () => {
    const service = createBomService({});
    const pdf = createCompressedPdf([
        'BT /F1 12 Tf 10 10 Td (112,37 x 225,80) Tj ET',
        '10 10 m 210 10 l 210 110 l 10 110 l h S'
    ].join('\n'));

    const result = service.analyzeDrawingFile('part.pdf', pdf);

    assert.equal(result.format, 'pdf');
    assert.equal(result.widthMm, 112.37);
    assert.equal(result.lengthMm, 225.8);
    assert.ok(result.pageWidthMm > 209 && result.pageWidthMm < 211);
    assert.ok(result.pageHeightMm > 296 && result.pageHeightMm < 298);
    assert.ok(result.previewSegments.length >= 4);
});
