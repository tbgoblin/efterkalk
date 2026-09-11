(function () {
    'use strict';

    const SORTABLE_CLASS = 'universal-sortable';
    const STICKY_CLASS = 'universal-sticky-header';
    const ASC_CLASS = 'universal-sort-asc';
    const DESC_CLASS = 'universal-sort-desc';
    const TABLE_CLASS = 'universal-table-behavior';
    const RESIZABLE_CLASS = 'universal-resizable-window';
    const MODAL_SELECTOR = '.kf-modal,.settings-modal,.modal-box,.oversigt-modal-shell,.order-detail-modal-shell,.modal-backdrop > .modal';
    const OVERLAY_SELECTOR = '.kf-modal-overlay,.settings-modal-overlay,.modal-overlay,.oversigt-modal-overlay,.order-detail-modal-overlay,.modal-backdrop';
    let modalPointerGesture = null;
    let suppressOverlayClickUntil = 0;

    function injectStyles() {
        if (document.getElementById('universalTableSortStyles')) return;
        const style = document.createElement('style');
        style.id = 'universalTableSortStyles';
        style.textContent =
            'th.' + SORTABLE_CLASS + '{cursor:pointer;user-select:none;}' +
            'table thead th.' + STICKY_CLASS + '{position:sticky!important;top:0!important;z-index:6;}' +
            'table.' + TABLE_CLASS + '>thead{position:sticky!important;top:0!important;z-index:6;}' +
            '.' + RESIZABLE_CLASS + '{resize:both!important;min-width:min(360px,90vw);min-height:min(240px,70vh);max-width:calc(100vw - 24px)!important;max-height:calc(100vh - 24px)!important;}' +
            'th.' + SORTABLE_CLASS + '::after{content:" ↕";opacity:.55;font-size:.85em;}' +
            'th.' + ASC_CLASS + '::after{content:" ▲";opacity:1;}' +
            'th.' + DESC_CLASS + '::after{content:" ▼";opacity:1;}' +
            'th.' + SORTABLE_CLASS + ':focus-visible{outline:2px solid #ffca55;outline-offset:-2px;}';
        document.head.appendChild(style);
    }

    function isSpecialHeader(th) {
        return th.hasAttribute('onclick') ||
            th.hasAttribute('data-no-sort') ||
            th.hasAttribute('data-sort-field') ||
            Number(th.getAttribute('colspan') || 1) !== 1 ||
            Boolean(th.querySelector('button,a,input,select,[onmousedown],.via-col-resizer'));
    }

    function decorateTable(table) {
        if (!table || table.nodeName !== 'TABLE' || table.dataset.universalSort === 'off') return;
        const tbody = table.tBodies && table.tBodies[0];
        if (!tbody) return;
        const hasFixedSections = table.classList.contains('kf-customer-summary-table');
        if (!hasFixedSections) table.classList.add(TABLE_CLASS);
        const headers = table.tHead ? Array.from(table.tHead.querySelectorAll('th')) : [];
        if (!headers.length) return;
        headers.forEach(th => {
            if (!hasFixedSections) th.classList.add(STICKY_CLASS);
            if (isSpecialHeader(th)) return;
            th.classList.add(SORTABLE_CLASS);
            th.tabIndex = th.tabIndex >= 0 ? th.tabIndex : 0;
            if (!th.title) th.title = 'Klik for at sortere kolonnen';
        });
    }

    function decorateWithin(root) {
        if (!root || root.nodeType !== 1) return;
        if (root.nodeName === 'TABLE') decorateTable(root);
        else if (root.closest) decorateTable(root.closest('table'));
        if (root.querySelectorAll) root.querySelectorAll('table').forEach(decorateTable);
        decorateResizableWindows(root);
    }

    function decorateResizableWindow(element) {
        if (!element || element.classList.contains(RESIZABLE_CLASS)) return;
        element.classList.add(RESIZABLE_CLASS);
        element.title = element.title || 'Træk i nederste højre hjørne for at ændre vinduets størrelse';
        const computed = window.getComputedStyle(element);
        if (computed.overflow === 'visible') element.style.overflow = 'auto';
    }

    function decorateResizableWindows(root) {
        if (!root || root.nodeType !== 1) return;
        if (root.matches && root.matches(MODAL_SELECTOR)) decorateResizableWindow(root);
        if (root.querySelectorAll) root.querySelectorAll(MODAL_SELECTOR).forEach(decorateResizableWindow);
    }

    function parseDate(text) {
        let match = text.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})$/);
        if (match) return Number(match[3] + String(match[2]).padStart(2, '0') + String(match[1]).padStart(2, '0'));
        match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
        if (match) return Number(match[1] + String(match[2]).padStart(2, '0') + String(match[3]).padStart(2, '0'));
        return null;
    }

    function parseNumber(text) {
        let value = text.replace(/\u00a0/g, ' ').trim();
        if (!value) return null;
        const negativeParentheses = /^\(.*\)$/.test(value);
        value = value.replace(/[()]/g, '').replace(/\s+/g, '');
        value = value.replace(/DKK|MIO|KR|%/gi, '');
        if (!/^[+\-]?\d[\d.,]*$/.test(value)) return null;
        if (value.includes(',') && value.includes('.')) value = value.replace(/\./g, '').replace(',', '.');
        else if (value.includes(',')) value = value.replace(',', '.');
        const parsed = Number(value);
        return Number.isFinite(parsed) ? (negativeParentheses ? -parsed : parsed) : null;
    }

    function cellValue(cell) {
        const explicit = cell.getAttribute('data-sort-value');
        const text = String(explicit !== null ? explicit : cell.textContent || '').trim();
        if (!text || text === '—' || text === '-') return { type: 'empty', value: '' };
        const date = parseDate(text);
        if (date !== null) return { type: 'date', value: date };
        const number = parseNumber(text);
        if (number !== null) return { type: 'number', value: number };
        return { type: 'text', value: text.toLocaleLowerCase('da-DK') };
    }

    function compareValues(left, right, direction) {
        if (left.type === 'empty' && right.type !== 'empty') return 1;
        if (right.type === 'empty' && left.type !== 'empty') return -1;
        let result = 0;
        if (left.type === right.type && (left.type === 'number' || left.type === 'date')) result = left.value - right.value;
        else result = String(left.value).localeCompare(String(right.value), 'da-DK', { numeric: true, sensitivity: 'base' });
        return direction === 'asc' ? result : -result;
    }

    function sortByHeader(th) {
        const table = th.closest('table');
        const tbody = table && table.tBodies ? table.tBodies[0] : null;
        if (!table || !tbody) return;
        const columnIndex = th.cellIndex;
        const allRows = Array.from(tbody.rows);
        const fixedRows = allRows.filter(row => row.classList.contains('total-row') || row.hasAttribute('data-sort-fixed'));
        const sortableRows = allRows.filter(row => !fixedRows.includes(row));
        if (sortableRows.length < 2 || sortableRows.some(row => row.cells.length <= columnIndex || Array.from(row.cells).some(cell => cell.colSpan > 1 || cell.rowSpan > 1))) return;

        const direction = th.getAttribute('aria-sort') === 'ascending' ? 'desc' : 'asc';
        const ranked = sortableRows.map((row, index) => ({ row, index, value: cellValue(row.cells[columnIndex]) }));
        ranked.sort((a, b) => compareValues(a.value, b.value, direction) || a.index - b.index);
        ranked.forEach(item => tbody.appendChild(item.row));
        fixedRows.forEach(row => tbody.appendChild(row));

        table.querySelectorAll('thead th').forEach(header => {
            header.classList.remove(ASC_CLASS, DESC_CLASS);
            header.removeAttribute('aria-sort');
        });
        th.classList.add(direction === 'asc' ? ASC_CLASS : DESC_CLASS);
        th.setAttribute('aria-sort', direction === 'asc' ? 'ascending' : 'descending');
    }

    document.addEventListener('click', event => {
        const th = event.target.closest && event.target.closest('th');
        const table = th && th.closest('table');
        if (!th || !table || table.dataset.universalSort === 'off' || isSpecialHeader(th) ||
            event.target.closest('button,a,input,select,[onmousedown],.via-col-resizer')) return;
        th.classList.add(SORTABLE_CLASS);
        if (!table.classList.contains('kf-customer-summary-table')) th.classList.add(STICKY_CLASS);
        sortByHeader(th);
    }, true);

    document.addEventListener('keydown', event => {
        if (event.key !== 'Enter' && event.key !== ' ') return;
        const th = event.target.closest && event.target.closest('th');
        const table = th && th.closest('table');
        if (!th || !table || table.dataset.universalSort === 'off' || isSpecialHeader(th)) return;
        th.classList.add(SORTABLE_CLASS);
        if (!table.classList.contains('kf-customer-summary-table')) th.classList.add(STICKY_CLASS);
        event.preventDefault();
        sortByHeader(th);
    });

    // Native CSS resize can finish with the pointer outside the dialog. Without
    // this guard, the resulting click lands on the overlay and closes the modal.
    document.addEventListener('pointerdown', event => {
        const modal = event.target.closest && event.target.closest('.' + RESIZABLE_CLASS);
        modalPointerGesture = modal ? {
            modal,
            x: event.clientX,
            y: event.clientY,
            moved: false
        } : null;
    }, true);

    document.addEventListener('pointermove', event => {
        if (!modalPointerGesture || modalPointerGesture.moved) return;
        if (Math.abs(event.clientX - modalPointerGesture.x) > 3 || Math.abs(event.clientY - modalPointerGesture.y) > 3) {
            modalPointerGesture.moved = true;
        }
    }, true);

    document.addEventListener('pointerup', () => {
        if (modalPointerGesture && modalPointerGesture.moved) suppressOverlayClickUntil = Date.now() + 700;
        modalPointerGesture = null;
    }, true);

    document.addEventListener('click', event => {
        if (modalPointerGesture && modalPointerGesture.moved) {
            suppressOverlayClickUntil = Date.now() + 700;
            modalPointerGesture = null;
        }
        if (Date.now() > suppressOverlayClickUntil) return;
        const overlay = event.target.closest && event.target.closest(OVERLAY_SELECTOR);
        if (!overlay || event.target !== overlay) return;
        event.preventDefault();
        event.stopImmediatePropagation();
        suppressOverlayClickUntil = 0;
    }, true);

    function initialize() {
        injectStyles();
        decorateWithin(document.body);
        new MutationObserver(mutations => {
            mutations.forEach(mutation => mutation.addedNodes.forEach(decorateWithin));
        }).observe(document.body, { childList: true, subtree: true });
    }

    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', initialize);
    else initialize();
})();
