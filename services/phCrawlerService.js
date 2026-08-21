// ── Personalehåndbog crawler ────────────────────────────────────────────────
// Estratto verbatim da routes/apiRoutes.js: crawl, indicizzazione e ricerca
// full-text della personalehåndbog intranet (http://apv/GHB/).
const http = require('http');

let phIndex   = [];          // [{url, title, text}]
let phStatus  = 'idle';      // 'idle' | 'indexing' | 'ready' | 'error'
let phIndexedAt = null;
let phError   = null;
const PH_BASE = 'http://apv/GHB/';

function phFetch(url) {
    return new Promise((resolve, reject) => {
        const req = http.get(url, { headers: { 'User-Agent': 'Gantech-Crawler/1.0' } }, (res) => {
            if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
                return resolve({ redirect: res.headers.location });
            }
            let body = '';
            res.setEncoding('utf8');
            res.on('data', c => body += c);
            res.on('end', () => resolve({ body, status: res.statusCode }));
        });
        req.setTimeout(10000, () => { req.destroy(); reject(new Error('timeout')); });
        req.on('error', reject);
    });
}

function phLinks(html, base) {
    const out = new Set();
    const re = /href=["']([^"']+)["']/gi;
    let m;
    while ((m = re.exec(html)) !== null) {
        try {
            const abs = new URL(m[1], base).href.split('?')[0].split('#')[0];
            if (abs.startsWith(PH_BASE) && !abs.match(/\.(jpg|jpeg|png|gif|svg|pdf|zip|docx?|xlsx?|css|js|ico|woff2?)$/i)) {
                out.add(abs);
            }
        } catch {}
    }
    return [...out];
}

function phTitle(html) {
    const m = html.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
    return m ? m[1].replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim() : '';
}

function phText(html) {
    return html
        .replace(/<script[\s\S]*?<\/script>/gi, ' ')
        .replace(/<style[\s\S]*?<\/style>/gi, ' ')
        .replace(/<nav[\s\S]*?<\/nav>/gi, ' ')
        .replace(/<header[\s\S]*?<\/header>/gi, ' ')
        .replace(/<footer[\s\S]*?<\/footer>/gi, ' ')
        .replace(/<[^>]+>/g, ' ')
        .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#\d+;/g, ' ')
        .replace(/\s+/g, ' ').trim();
}

async function crawlPH() {
    if (phStatus === 'indexing') return;
    phStatus = 'indexing';
    phError = null;
    phIndex = [];
    const visited = new Set();
    const queue = [PH_BASE];
    console.log('[PH-CRAWL] Starting crawl of', PH_BASE);
    while (queue.length > 0) {
        const url = queue.shift();
        if (visited.has(url)) continue;
        visited.add(url);
        try {
            const r = await phFetch(url);
            if (r.redirect) {
                try {
                    const abs = new URL(r.redirect, url).href.split('?')[0].split('#')[0];
                    if (abs.startsWith(PH_BASE) && !visited.has(abs)) queue.push(abs);
                } catch {}
                continue;
            }
            if (r.status !== 200) continue;
            const title = phTitle(r.body);
            const text  = phText(r.body);
            phIndex.push({ url, title, text });
            for (const link of phLinks(r.body, url)) {
                if (!visited.has(link)) queue.push(link);
            }
        } catch { /* skip unreachable page */ }
    }
    phStatus = 'ready';
    phIndexedAt = new Date().toISOString();
    console.log('[PH-CRAWL] Done:', phIndex.length, 'pages indexed');
}

function markPhError(message) {
    phStatus = 'error';
    phError = String(message || 'unknown');
}

function getPhStatus() {
    return { status: phStatus, count: phIndex.length, indexedAt: phIndexedAt, error: phError };
}

function searchPh(rawQuery) {
    const q = String(rawQuery || '').toLowerCase().trim();
    if (!q) return { results: [], status: phStatus };
    if (phStatus !== 'ready') return { results: [], status: phStatus };
    const terms = q.split(/\s+/).filter(Boolean);
    const results = [];
    for (const page of phIndex) {
        const haystack = (page.title + ' ' + page.text).toLowerCase();
        const score = terms.reduce((s, t) => s + (haystack.split(t).length - 1), 0);
        if (score === 0) continue;
        const firstTerm = terms[0];
        const idx = page.text.toLowerCase().indexOf(firstTerm);
        let snippet = '';
        if (idx >= 0) {
            const s = Math.max(0, idx - 80);
            const e = Math.min(page.text.length, idx + 200);
            snippet = (s > 0 ? '…' : '') + page.text.slice(s, e) + (e < page.text.length ? '…' : '');
        } else {
            snippet = page.text.slice(0, 240) + '…';
        }
        results.push({ url: page.url, title: page.title || page.url, snippet, score });
    }
    results.sort((a, b) => b.score - a.score);
    return { results: results.slice(0, 40), status: phStatus };
}

function isPhIndexing() {
    return phStatus === 'indexing';
}

// Start crawl in background after module load
setTimeout(() => crawlPH().catch(e => { markPhError(e.message); console.error('[PH-CRAWL] Error:', e.message); }), 3000);

module.exports = {
    crawlPH,
    getPhStatus,
    searchPh,
    isPhIndexing,
    markPhError
};
