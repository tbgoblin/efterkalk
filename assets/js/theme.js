(function (root) {
    'use strict';

    const palette = {
        surface: '#202322', raised: '#292d2b', hover: '#343a37', border: '#49534d',
        text: '#edf1ee', muted: '#b7c2bc', blue: '#95caff', green: '#8fd5ac',
        red: '#ffaaa4', amber: '#f1ce83', violet: '#d7b8ef',
        blueSurface: '#253743', greenSurface: '#243d31', redSurface: '#442c2b',
        amberSurface: '#403722', violetSurface: '#382f42'
    };
    const normalize = value => ['light', 'dark', 'system'].includes(value) ? value : 'light';

    function colorRole(red, green, blue, kind) {
        const high = Math.max(red, green, blue), low = Math.min(red, green, blue);
        const spread = high - low;
        const brightness = (high + low) / 2;
        let family = 'blue';
        if (spread > 24) {
            if (red >= high && green > blue + spread * 0.35) family = 'amber';
            else if (red >= high && blue > green + spread * 0.4) family = 'violet';
            else if (red >= high) family = 'red';
            else if (green >= high && green > blue + 8) family = 'green';
        }
        if (kind === 'border') return spread > 70 && brightness < 175 ? null : 'border';
        if (kind === 'background') {
            if (brightness < 170) return null;
            if (spread > 32) return family + 'Surface';
            return brightness > 248 ? 'surface' : brightness > 232 ? 'raised' : 'hover';
        }
        if (brightness > 210) return null;
        if (family === 'blue' && brightness < 85) return 'text';
        if (spread > 55 && high > 70) return family;
        return brightness < 85 ? 'text' : 'muted';
    }

    const model = { normalize, colorRole, palette };
    if (typeof module === 'object' && module.exports) module.exports = model;
    if (!root.document) return;
    const document = root.document;
    const system = root.matchMedia('(prefers-color-scheme: dark)');
    let preference = 'light';
    let changePreference = null;
    let started = false;
    let pending = false;
    const styledSheets = new WeakSet();
    const pendingElements = new Set();
    const probe = document.createElement('span').style;
    const colorPattern = /url\([^)]*\)|var\([^)]*\)|#[\da-f]{3,8}\b|rgba?\([^)]*\)|\b(?:white|black|gray|grey|red|green|blue|navy|silver)\b/gi;

    function translate(value, kind) {
        if (value.includes('url(')) return value;
        return value.replace(colorPattern, literal => {
            if (/^(url|var)\(/i.test(literal)) return literal;
            probe.color = '';
            probe.color = literal;
            const parsed = probe.color.match(/^rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)$/);
            if (!parsed) return literal;
            const alpha = parsed[4] === undefined ? 1 : Number(parsed[4]);
            if (alpha < 0.5) return literal;
            const role = colorRole(Number(parsed[1]), Number(parsed[2]), Number(parsed[3]), kind);
            return role ? 'var(--goh-' + role + ', ' + literal + ')' : literal;
        });
    }

    function adaptStyle(style, textFill = false) {
        for (const property of Array.from(style)) {
            const kind = property === 'color' || property === '-webkit-text-fill-color' || property === 'fill' && textFill ? 'text'
                : property === 'background-color' || property === 'background-image' ? 'background'
                    : /^(border.*color|outline-color)$/.test(property) ? 'border' : null;
            if (!kind) continue;
            const before = style.getPropertyValue(property);
            const after = translate(before, kind);
            if (after !== before) style.setProperty(property, after, style.getPropertyPriority(property));
        }
    }

    function adaptSheets() {
        function walk(rules) {
            for (const rule of rules) {
                if (rule.media?.mediaText === 'print') continue;
                if (rule.style) adaptStyle(rule.style, /label|axis|text|tick/.test(rule.selectorText || ''));
                if (rule.cssRules) walk(rule.cssRules);
            }
        }
        for (const sheet of document.styleSheets) {
            if (styledSheets.has(sheet) || sheet.ownerNode?.dataset.gohTheme !== undefined) continue;
            try { walk(sheet.cssRules); styledSheets.add(sheet); } catch (_) {}
        }
    }

    function adaptTree(element) {
        if (element.nodeType !== 1 || element.closest('[data-goh-theme-preserve]')) return;
        if (element.style) adaptStyle(element.style, element.localName === 'text');
        for (const child of element.querySelectorAll('[style]')) {
            if (!child.closest('[data-goh-theme-preserve]')) adaptStyle(child.style, child.localName === 'text');
        }
    }

    function startAdapter() {
        if (started) return;
        started = true;
        adaptSheets();
        adaptTree(document.documentElement);
        const observer = new MutationObserver(records => {
            let sheetsChanged = false;
            for (const record of records) {
                if (record.type === 'attributes') pendingElements.add(record.target);
                for (const node of record.addedNodes) {
                    if (node.nodeType !== 1) continue;
                    pendingElements.add(node);
                    if (node.matches('style,link[rel=stylesheet]') || node.querySelector('style,link[rel=stylesheet]')) sheetsChanged = true;
                }
            }
            if (sheetsChanged) adaptSheets();
            if (pending) return;
            pending = true;
            queueMicrotask(() => {
                const elements = [...pendingElements];
                pendingElements.clear(); pending = false;
                for (const element of elements) if (element.isConnected) adaptTree(element);
            });
        });
        observer.observe(document.documentElement, { subtree: true, childList: true, attributes: true, attributeFilter: ['style'] });
        document.addEventListener('load', event => {
            if (event.target.matches?.('link[rel=stylesheet]')) adaptSheets();
        }, true);
    }

    function apply(value) {
        preference = normalize(value);
        const dark = preference === 'dark' || preference === 'system' && system.matches;
        if (dark) startAdapter();
        document.documentElement.dataset.gohTheme = dark ? 'dark' : 'light';
        for (const input of document.querySelectorAll('[name=gohTheme]')) input.checked = input.value === preference;
    }

    function bind(value, save, status = 'saved') {
        changePreference = save;
        apply(value);
        for (const input of document.querySelectorAll('[name=gohTheme]')) input.disabled = !save;
        const state = document.getElementById('gohThemeStatus');
        if (state) {
            state.textContent = !save ? 'Log ind for at v\u00e6lge tema' : ({ loading: 'Henter GOH-profil...', saving: 'Gemmer i GOH...', saved: 'Gemt i GOH', error: 'Ikke gemt i GOH. Se dashboard.', conflict: 'Profilkonflikt. Se dashboard.' }[status] || '');
            state.dataset.error = String(status === 'error' || status === 'conflict');
        }
    }

    document.addEventListener('change', event => {
        if (!event.target.matches('[name=gohTheme]') || !changePreference) return;
        const value = normalize(event.target.value);
        apply(value);
        changePreference(value);
    });
    system.addEventListener('change', () => { if (preference === 'system') apply(preference); });
    root.GohTheme = { model, apply, bind };

    if (document.currentScript?.dataset.standalone !== undefined) {
        fetch('/ui/theme', { credentials: 'same-origin', cache: 'no-store', signal: AbortSignal.timeout(8000) })
            .then(response => response.ok ? response.json() : null)
            .then(result => { if (result?.ok) apply(result.theme); }).catch(() => {});
    }
})(typeof window === 'undefined' ? globalThis : window);