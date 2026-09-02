(function(root, factory) {
    const api = factory();
    if (typeof module === 'object' && module.exports) {
        module.exports = api;
    }
    if (root) {
        root.AftercalcCostExclusions = api;
    }
})(typeof globalThis !== 'undefined' ? globalThis : this, function() {
    'use strict';

    function finiteNumber(value) {
        const parsed = Number(value);
        return Number.isFinite(parsed) ? parsed : 0;
    }

    function getLineKey(line, index) {
        const lineNumber = Number(line && line.LnNo);
        if (Number.isInteger(lineNumber) && lineNumber > 0) {
            return 'line:' + String(lineNumber);
        }
        return 'index:' + String(Number.isFinite(Number(index)) ? Number(index) : 0);
    }

    function getLineSalesPrice(line) {
        return finiteNumber(line && line.DPrice) * finiteNumber(line && line.NoFin);
    }

    function isLinkedProductionLine(line) {
        return finiteNumber(line && line.PurcNo) > 0
            && finiteNumber(line && line.LinkedOrderType) !== 6
            && !(line && line.IsDiscountLine);
    }

    function getLineCostContribution(line) {
        if (!line) return 0;
        if (isLinkedProductionLine(line)) {
            return finiteNumber(line.ProductionOrderTotalCost ?? line.EffectiveLineCost ?? line.LineCost);
        }
        return finiteNumber(line.EffectiveLineCost ?? line.LineCost);
    }

    function calculateAdjustedCost(totalCost, salesOrderLines, excludedLineKeys) {
        const lines = Array.isArray(salesOrderLines) ? salesOrderLines : [];
        const excludedKeys = excludedLineKeys instanceof Set
            ? excludedLineKeys
            : new Set(Array.isArray(excludedLineKeys) ? excludedLineKeys.map(String) : []);
        const lineEntries = lines.map(function(line, index) {
            return { line, key: getLineKey(line, index) };
        });
        const matchedKeys = lineEntries
            .filter(function(entry) { return excludedKeys.has(entry.key); })
            .map(function(entry) { return entry.key; });
        const linkedGroups = new Map();
        let excludedCost = 0;
        let deferredSharedLineCount = 0;
        const deferredKeys = [];

        for (const entry of lineEntries) {
            if (!isLinkedProductionLine(entry.line)) continue;
            const purcNo = String(Number(entry.line.PurcNo));
            if (!linkedGroups.has(purcNo)) linkedGroups.set(purcNo, []);
            linkedGroups.get(purcNo).push(entry);
        }

        for (const entry of lineEntries) {
            if (!excludedKeys.has(entry.key) || isLinkedProductionLine(entry.line)) continue;
            excludedCost += getLineCostContribution(entry.line);
        }

        for (const groupEntries of linkedGroups.values()) {
            const selectedEntries = groupEntries.filter(function(entry) { return excludedKeys.has(entry.key); });
            if (selectedEntries.length === 0) continue;
            if (selectedEntries.length !== groupEntries.length) {
                deferredSharedLineCount += selectedEntries.length;
                selectedEntries.forEach(function(entry) { deferredKeys.push(entry.key); });
                continue;
            }
            excludedCost += getLineCostContribution(groupEntries[0].line);
        }

        const originalCost = finiteNumber(totalCost);
        return {
            originalCost,
            excludedCost,
            adjustedCost: originalCost - excludedCost,
            excludedLineCount: matchedKeys.length,
            deferredSharedLineCount,
            deferredKeys,
            matchedKeys
        };
    }

    return Object.freeze({
        getLineKey,
        getLineSalesPrice,
        getLineCostContribution,
        calculateAdjustedCost
    });
});
