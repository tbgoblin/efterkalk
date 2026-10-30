// Manual no-charge classification. Visma orders and invoice amounts are read-only.
function createUbOrdersService({ getConnection, sql, orderNotes }) {
    const fail = (message, statusCode = 400) => Object.assign(new Error(message), { statusCode });
    function monthKey(value) {
        const month = String(value || '').trim();
        if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(month)) throw fail('Vælg en gyldig måned.');
        return month;
    }
    function orderNumber(value) {
        const number = Number(value);
        if (!Number.isSafeInteger(number) || number <= 0 || number > 2147483647) throw fail('Angiv et gyldigt ordrenummer.');
        return number;
    }
    async function readOrders(ordNos) {
        if (!ordNos.length) return [];
        const pool = await getConnection();
        const result = await pool.request()
            .input('ordNos', sql.NVarChar(sql.MAX), ordNos.join(','))
            .query(`SELECT o.OrdNo, o.OrdDt, o.CustNo, o.LstInvDt, o.InvoNo, o.InvoAm,
                           c.Nm AS CustomerName
                    FROM Ord o LEFT JOIN Actor c ON o.CustNo > 0 AND c.CustNo=o.CustNo
                    WHERE o.TrTp=1 AND o.OrdTp=1
                      AND o.OrdNo IN (SELECT TRY_CAST(value AS int) FROM STRING_SPLIT(@ordNos, ','))
                    ORDER BY o.OrdNo DESC`);
        return result.recordset || [];
    }
    function registrationMonth(note, order) {
        if (note.ubMonth) return note.ubMonth;
        const date = String(order.LstInvDt || order.OrdDt || '');
        return date.slice(0, 4) + '-' + date.slice(4, 6);
    }
    async function list(month) {
        month = monthKey(month);
        const notes = orderNotes.getAllNotes();
        const ordNos = Object.keys(notes).filter(key => notes[key].isUB === true
            && (!notes[key].ubMonth || notes[key].ubMonth === month));
        const orders = await readOrders(ordNos);
        return orders.filter(order => registrationMonth(notes[String(order.OrdNo)], order) === month)
            .map(order => {
                const note = notes[String(order.OrdNo)];
                return { ...order, month, text:note.ubText || note.text || '' };
            });
    }
    async function add({ ordNo, month, text = '' }) {
        ordNo = orderNumber(ordNo);
        month = monthKey(month);
        const [order] = await readOrders([ordNo]);
        if (!order) throw fail('Salgsordren findes ikke.', 404);
        const existing = orderNotes.getNote(ordNo) || {};
        if (existing.isUB && existing.ubMonth && existing.ubMonth !== month) {
            throw fail('Ordren er allerede registreret som U/B i ' + existing.ubMonth + '.', 409);
        }
        const note = orderNotes.setNote(ordNo, { ...existing, isUB:true, ubMonth:month, ubText:String(text).slice(0, 2000) });
        return { order, note };
    }
    async function remove({ ordNo, month }) {
        ordNo = orderNumber(ordNo);
        month = monthKey(month);
        const existing = orderNotes.getNote(ordNo);
        if (!existing || !existing.isUB) return null;
        const [order] = await readOrders([ordNo]);
        if (!order || registrationMonth(existing, order) !== month) throw fail('Ordren er ikke registreret som U/B i den valgte måned.', 409);
        return orderNotes.setNote(ordNo, { ...existing, isUB:false, ubMonth:'', ubText:'' });
    }
    return { list, add, remove };
}
module.exports = { createUbOrdersService };
