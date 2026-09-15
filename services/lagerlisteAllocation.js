// Allocate shared orders once, using the same packing evidence as Lagerliste2.
function allocateSharedOrders(finishedRows, viaRows, states) {
    const stateByOrder = new Map(states.map(row => [Number(row.orderNo), row]));
    const finishedByOrder = new Map(finishedRows.map(row => [Number(row.OrdNo), row]));
    const viaByOrder = new Map(viaRows.map(row => [Number(row.OrdNo), row]));
    const round = value => Math.round(value * 100) / 100;
    const audit = [];
    const finished = finishedRows.map(row => {
        if (row.AllocationApplied) return { ...row };
        const state = stateByOrder.get(Number(row.OrdNo));
        if (!viaByOrder.has(Number(row.OrdNo))) return { ...row };
        if (!state) throw new Error('Missing packing evidence for order ' + row.OrdNo);
        const ratio = Math.max(0, Math.min(1, Number(state.packedRatio)));
        if (!Number.isFinite(ratio)) throw new Error('Invalid packing evidence for order ' + row.OrdNo);
        audit.push({ ordNo: row.OrdNo, packedRatio: ratio, originalFinished: row.Value,
            originalVia: viaByOrder.get(Number(row.OrdNo)).Value });
        return { ...row, OriginalValue: row.Value, Value: round(row.Value * ratio), AllocationApplied: true };
    });
    const via = viaRows.map(row => {
        if (row.AllocationApplied) return { ...row };
        if (!finishedByOrder.has(Number(row.OrdNo))) return { ...row };
        const ratio = Math.max(0, Math.min(1, Number(stateByOrder.get(Number(row.OrdNo)).packedRatio)));
        const result = { ...row, OriginalValue: row.Value, AllocationApplied: true };
        for (const key of ['MaterialCost', 'StangCost', 'PurchasedPartCost', 'TimeCost']) {
            result[key] = round(Number(row[key] || 0) * (1 - ratio));
        }
        if (Array.isArray(row.PurchasedPartDetails)) {
            result.PurchasedPartDetails = row.PurchasedPartDetails.map(detail => ({
                ...detail,
                originalCountedValue: Number(detail.countedValue || 0),
                countedValue: round(Number(detail.countedValue || 0) * (1 - ratio))
            }));
        }
        result.Value = round(result.MaterialCost + result.StangCost + result.PurchasedPartCost + result.TimeCost);
        return result;
    });
    return { finished, via, audit };
}

// Remove only quantities already valued as follow-up stock. A reservation is
// not proof of consumption: any other physical quantity must remain valued.
function allocateComponentStock(componentRows, followUpRows) {
    const availableByProduct = new Map();
    for (const row of followUpRows) {
        const key = String(row.ProdNo).trim();
        const valuedQuantity = row.PoPhStB === undefined ? row.Beholdning : row.PoPhStB;
        availableByProduct.set(key, (availableByProduct.get(key) || 0) + Math.max(0, Number(valuedQuantity || 0)));
    }
    const audit = [];
    const round = value => Math.round(value * 100) / 100;
    const rows = componentRows.map(row => {
        if (row.StockAllocationApplied) return { ...row };
        const key = String(row.ProdNo).trim();
        const overlap = Math.min(Math.max(0, Number(row.Quantity || 0)), availableByProduct.get(key) || 0);
        if (!overlap) return { ...row };
        availableByProduct.set(key, availableByProduct.get(key) - overlap);
        const quantity = Number(row.Quantity) - overlap;
        const fifoValue = round(quantity * Number(row.UnitCost || 0));
        audit.push({ prodNo: row.ProdNo, overlapQty: overlap, removedValue: round(Number(row.FifoValue || 0) - fifoValue) });
        return { ...row, StockAllocationApplied: true, OverlapQuantity: overlap, OriginalQuantity: row.Quantity, OriginalFifoValue: row.FifoValue,
            Quantity: quantity, Value: round(quantity * Number(row.StandardPrice || 0)), FifoValue: fifoValue };
    }).filter(row => Number(row.Quantity) !== 0);
    return { rows, audit };
}

function allocatePurchasedPartsFromFollowUpStock(followUpRows, viaRows) {
    const requestedByProduct = new Map();
    for (const viaRow of Array.isArray(viaRows) ? viaRows : []) {
        for (const detail of Array.isArray(viaRow.PurchasedPartDetails) ? viaRow.PurchasedPartDetails : []) {
            const key = String(detail.prodNo || '').trim();
            const quantity = Math.max(0, Number(detail.stockTransferQty || 0));
            if (key && quantity > 0) requestedByProduct.set(key, (requestedByProduct.get(key) || 0) + quantity);
        }
    }

    const round = value => Math.round(Number(value || 0) * 100) / 100;
    const audit = [];
    const rows = (Array.isArray(followUpRows) ? followUpRows : []).map(row => {
        if (row.PurchasedPartAllocationApplied) return { ...row };
        const key = String(row.ProdNo || '').trim();
        const requestedQty = requestedByProduct.get(key) || 0;
        if (requestedQty <= 0) return { ...row };
        const physicalQty = Math.max(0, Number(row.PoPhStB || 0));
        const availableQty = Math.max(0, Number(row.Beholdning || 0));
        const reservedPhysicalQty = Math.min(
            Math.max(0, physicalQty - availableQty),
            Math.max(0, Number(row.ShpRsv || 0))
        );
        const transferredQty = Math.min(requestedQty, reservedPhysicalQty);
        if (transferredQty <= 0) return { ...row };
        const remainingQty = Math.max(0, physicalQty - transferredQty);
        const fifoPrice = Number(row.PhCstPr || 0);
        const transferredValue = round(transferredQty * fifoPrice);
        const remainingValue = round(remainingQty * fifoPrice);
        const availableValue = round(availableQty * fifoPrice);
        requestedByProduct.set(key, Math.max(0, requestedQty - transferredQty));
        audit.push({ prodNo: row.ProdNo, transferredQty: round(transferredQty), transferredValue });
        return {
            ...row,
            PurchasedPartAllocationApplied: true,
            ValuedQuantity: round(remainingQty),
            TransferredToViaQty: round(transferredQty),
            TransferredToViaValue: transferredValue,
            RemainingReservedValue: round(Math.max(0, Number(row.ReservedValue || 0) - transferredValue)),
            Value: remainingValue,
            _preciseValue: remainingQty * fifoPrice,
            Diff: round(remainingValue - availableValue)
        };
    });
    return { rows, audit };
}

function validateValuation(payload) {
    if (payload?.valuationVersion !== 35) throw new Error('Opdater lagerlisten før lukning: beregningen er fra en ældre version.');
    const c = payload.categories;
    if (!c || !payload.totals) throw new Error('Ufuldstændig lagerberegning.');
    const follow = new Map((c.opfolgningvare || []).map(r => [
        String(r.ProdNo).trim(),
        Math.max(0, Number(r.PoPhStB === undefined ? r.Beholdning : r.PoPhStB) || 0)
    ]));
    for (const row of c.gr5Items || []) {
        const overlap = follow.get(String(row.ProdNo).trim()) || 0;
        if (overlap > 0 && Number(row.Quantity) > 0) {
            const expected = Number(row.OriginalQuantity) - Math.min(Math.max(0, Number(row.OriginalQuantity)), overlap);
            if (!row.StockAllocationApplied || !Number.isFinite(expected) || Math.abs(expected - Number(row.Quantity)) > 0.000001) {
                throw new Error('Dobbelt lagerbeholdning: ' + row.ProdNo);
            }
        }
    }
    const via = new Map((c.salgordreVia || []).map(r => [Number(r.OrdNo), r]));
    for (const row of c.finishedNotInvoiced || []) {
        const other = via.get(Number(row.OrdNo));
        if (other && (!row.AllocationApplied || !other.AllocationApplied)) throw new Error('Dobbelt ordreværdi: ' + row.OrdNo);
    }
    for (const [category, field] of [['finishedNotInvoiced', 'finishedNotInvoiced'], ['salgordreVia', 'salgordreVia']]) {
        const sum = (c[category] || []).reduce((n, r) => n + Number(r.Value), 0);
        if (!Number.isFinite(sum) || !Number.isFinite(Number(payload.totals[field])) || Math.abs(sum - Number(payload.totals[field])) > 0.011) {
            throw new Error('Total stemmer ikke med rækkerne: ' + category);
        }
    }
    return true;
}

function validateClosure(payload, month, now = new Date()) {
    validateValuation(payload);
    const generated = new Date(payload.generatedAt);
    const monthInDenmark = date => new Intl.DateTimeFormat('sv-SE', { timeZone: 'Europe/Copenhagen', year: 'numeric', month: '2-digit' }).format(date);
    if (!Number.isFinite(generated.getTime()) || !/^\d{4}-(0[1-9]|1[0-2])$/.test(month) || monthInDenmark(generated) !== month || monthInDenmark(now) !== month) {
        throw new Error('Måneden passer ikke til beregningsdatoen. Historiske lukninger kræver en særskilt revision.');
    }
    if (now - generated > 5 * 60 * 1000 || generated - now > 60000) throw new Error('Lagerberegningen er for gammel. Tryk Opdater lagerliste før lukning.');
}

module.exports = { allocateSharedOrders, allocateComponentStock, allocatePurchasedPartsFromFollowUpStock, validateValuation, validateClosure };
