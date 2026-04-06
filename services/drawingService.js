const fs = require('fs');
const path = require('path');
const sql = require('mssql/msnodesqlv8');

function isHttpUrl(value) {
    return /^https?:\/\//i.test(String(value || '').trim());
}

function isAbsoluteWindowsPath(value) {
    const normalized = String(value || '').trim();
    return /^[a-zA-Z]:[\\/]/.test(normalized) || /^\\\\/.test(normalized);
}

function normalizeWindowsPath(value) {
    return String(value || '').trim().replace(/\//g, '\\');
}

function buildImageItems(webPg, pictFNm) {
    const items = [];
    const seen = new Set();

    function pushItem(type, value, label) {
        const cleanedValue = String(value || '').trim();
        if (!cleanedValue) return;
        const key = type + '|' + cleanedValue.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        items.push({ type, value: cleanedValue, label });
    }

    const webPgValue = String(webPg || '').trim();
    const pictFNmValue = String(pictFNm || '').trim();

    if (webPgValue && pictFNmValue) {
        if (isHttpUrl(webPgValue)) {
            const baseUrl = webPgValue.replace(/\\/g, '/').replace(/\/+$/, '');
            const fileName = pictFNmValue.replace(/^[/\\]+/, '');
            pushItem('url', baseUrl + '/' + fileName, 'WebPg + PictFNm');
        } else {
            const basePath = normalizeWindowsPath(webPgValue).replace(/[\\/]+$/, '');
            const fileName = pictFNmValue.replace(/^[\\/]+/, '');
            pushItem('file', basePath + '\\' + fileName, 'WebPg + PictFNm');
        }
    }

    if (pictFNmValue) {
        pushItem(isHttpUrl(pictFNmValue) ? 'url' : 'file', isHttpUrl(pictFNmValue) ? pictFNmValue : normalizeWindowsPath(pictFNmValue), 'PictFNm');
    }

    if (webPgValue) {
        pushItem(isHttpUrl(webPgValue) ? 'url' : 'file', isHttpUrl(webPgValue) ? webPgValue : normalizeWindowsPath(webPgValue), 'WebPg');
    }

    return items;
}

function isSupportedImagePath(filePath) {
    return /\.(png|jpe?g|gif|webp|bmp|tiff?)$/i.test(String(filePath || '').trim());
}

function normalizeToken(value) {
    return String(value || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function isLikelyDrawingFileForProdNo(fileName, prodNo) {
    const fileBase = String(fileName || '').replace(/\.pdf$/i, '');
    const fileToken = normalizeToken(fileBase);
    const prodToken = normalizeToken(prodNo);
    if (!fileToken || !prodToken) return false;

    const prodNoLStripped = prodToken.endsWith('L') ? prodToken.slice(0, -1) : prodToken;
    return fileToken === prodToken
        || fileToken.startsWith(prodToken)
        || fileToken === prodNoLStripped
        || (prodNoLStripped && fileToken.startsWith(prodNoLStripped));
}

function findNewestPdfInFolder(baseDir, prodNo, maxDepth = 2) {
    const root = String(baseDir || '').trim();
    if (!root) return null;
    if (!fs.existsSync(root)) return null;

    const queue = [{ dir: root, depth: 0 }];
    let best = null;

    while (queue.length > 0) {
        const current = queue.shift();
        let entries = [];
        try {
            entries = fs.readdirSync(current.dir, { withFileTypes: true });
        } catch (_) {
            continue;
        }

        for (const entry of entries) {
            const fullPath = path.join(current.dir, entry.name);
            if (entry.isDirectory()) {
                if (current.depth < maxDepth) {
                    queue.push({ dir: fullPath, depth: current.depth + 1 });
                }
                continue;
            }

            if (!entry.isFile() || !/\.pdf$/i.test(entry.name)) continue;
            if (!isLikelyDrawingFileForProdNo(entry.name, prodNo)) continue;

            let mtimeMs = 0;
            try {
                mtimeMs = Number(fs.statSync(fullPath).mtimeMs || 0);
            } catch (_) {}

            if (!best || mtimeMs > best.mtimeMs) {
                best = { fullPath, mtimeMs };
            }
        }
    }

    return best ? best.fullPath : null;
}

function resolveDrawingPathFromWebPg(prodNo, webPg) {
    const raw = String(webPg || '').trim();
    if (!raw) return null;

    const normalized = raw.replace(/\//g, '\\');

    if (/\.pdf(\?|#|$)/i.test(raw)) {
        if (/^https?:\/\//i.test(raw) || /^file:\/\//i.test(raw)) return raw;
        if (fs.existsSync(normalized)) return normalized;
        return null;
    }

    if (fs.existsSync(normalized)) {
        try {
            const stat = fs.statSync(normalized);
            if (stat.isDirectory()) {
                return findNewestPdfInFolder(normalized, prodNo, 3);
            }
            if (stat.isFile() && /\.pdf$/i.test(normalized)) {
                return normalized;
            }
        } catch (_) {}
    }

    const withPdf = normalized + '.pdf';
    if (fs.existsSync(withPdf)) return withPdf;

    return null;
}

function isUsablePdfPath(webPg) {
    const value = String(webPg || '').trim();
    if (!value || !/\.pdf(\?|#|$)/i.test(value)) return false;

    if (/^https?:\/\//i.test(value) || /^file:\/\//i.test(value)) {
        return true;
    }

    const normalized = value.replace(/\//g, '\\');
    if (normalized.startsWith('\\\\') || /^[A-Za-z]:\\/.test(normalized)) {
        try {
            return fs.existsSync(normalized);
        } catch (_) {
            return false;
        }
    }

    return true;
}

async function getLatestDrawingByProdNo(pool, prodNos, logEvent = null) {
    const uniqueProdNos = Array.from(new Set(
        (prodNos || [])
            .map(p => String(p || '').trim())
            .filter(Boolean)
    ));

    if (uniqueProdNos.length === 0) return new Map();

    try {
        const colsResult = await pool.request().query(`
            SELECT COLUMN_NAME
            FROM INFORMATION_SCHEMA.COLUMNS
            WHERE TABLE_NAME = 'FreeInf2'
        `);

        const colNameMap = new Map((colsResult.recordset || []).map(r => {
            const name = String(r.COLUMN_NAME || '').trim();
            return [name.toLowerCase(), name];
        }));
        const hasDateTimePair = colNameMap.has('chdt') && colNameMap.has('chtm');

        let orderByExpr = 'CONVERT(VARCHAR(1000), WebPg) DESC';
        if (hasDateTimePair) {
            orderByExpr = '[' + colNameMap.get('chdt') + '] DESC, [' + colNameMap.get('chtm') + '] DESC';
        } else {
            const singleCandidates = ['moddt', 'upddt', 'regdt', 'insdt', 'crtdt', 'findt', 'id', 'recno', 'lnno'];
            const chosen = singleCandidates.find(c => colNameMap.has(c));
            if (chosen) {
                const quoted = '[' + colNameMap.get(chosen) + ']';
                orderByExpr = quoted + ' DESC';
            }
        }

        const req = pool.request();
        const placeholders = uniqueProdNos.map((prodNo, i) => {
            const param = 'prod' + i;
            req.input(param, sql.VarChar(100), prodNo);
            return '@' + param;
        });

        const result = await req.query(`
            WITH x AS (
                SELECT
                    LTRIM(RTRIM(CONVERT(VARCHAR(100), ProdNo))) AS ProdNo,
                    LTRIM(RTRIM(CONVERT(VARCHAR(1000), WebPg))) AS WebPg,
                    ROW_NUMBER() OVER (
                        PARTITION BY LTRIM(RTRIM(CONVERT(VARCHAR(100), ProdNo)))
                        ORDER BY ${orderByExpr}
                    ) AS rn
                FROM FreeInf2
                WHERE LTRIM(RTRIM(CONVERT(VARCHAR(100), ProdNo))) IN (${placeholders.join(',')})
                  AND WebPg IS NOT NULL
                  AND LTRIM(RTRIM(CONVERT(VARCHAR(1000), WebPg))) <> ''
            )
            SELECT ProdNo, WebPg
            FROM x
            WHERE rn <= 12
            ORDER BY ProdNo, rn
        `);

        const map = new Map();
        for (const row of (result.recordset || [])) {
            const key = String(row.ProdNo || '').trim().toUpperCase();
            const webPg = String(row.WebPg || '').trim();
            if (!key || !webPg || map.has(key)) continue;

            const resolved = resolveDrawingPathFromWebPg(key, webPg);
            if (resolved && isUsablePdfPath(resolved)) {
                map.set(key, resolved);
            }
        }
        return map;
    } catch (err) {
        if (typeof logEvent === 'function') {
            logEvent('WARNING drawing lookup skipped: ' + err.message);
        }
        return new Map();
    }
}

module.exports = {
    isHttpUrl,
    isAbsoluteWindowsPath,
    normalizeWindowsPath,
    buildImageItems,
    isSupportedImagePath,
    normalizeToken,
    isLikelyDrawingFileForProdNo,
    findNewestPdfInFolder,
    resolveDrawingPathFromWebPg,
    isUsablePdfPath,
    getLatestDrawingByProdNo
};
