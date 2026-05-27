const getConnection = require('./db.js');

async function chain(pool, ordNo, depth, visited) {
    if (depth > 6 || visited.has(ordNo)) return;
    visited.add(ordNo);

    const hdr = await pool.request().query(
        'SELECT OrdNo, TrTp FROM Ord WHERE OrdNo = ' + ordNo
    );
    const tp = hdr.recordset[0] ? hdr.recordset[0].TrTp : '?';
    const indent = '  '.repeat(depth);
    console.log(indent + '=== OrdNo ' + ordNo + ' (TrTp=' + tp + ') ===');

    const lns = await pool.request().query(
        'SELECT LnNo, ProdNo, ProdTp4, PurcNo, DPrice, NoFin, NoOrg, CCstPr, ' +
        'CAST(NoFin * CCstPr AS DECIMAL(10,2)) AS LineCost ' +
        'FROM OrdLn WHERE OrdNo = ' + ordNo + ' ORDER BY LnNo'
    );
    for (const r of lns.recordset) {
        const pn = r.PurcNo ? Number(r.PurcNo) : null;
        const flag = pn ? ' --> PurcNo=' + pn : '';
        console.log(
            indent + '  Ln' + r.LnNo +
            ' [' + r.ProdNo + ']' +
            ' tp4=' + r.ProdTp4 +
            ' NoFin=' + r.NoFin +
            ' CCstPr=' + r.CCstPr +
            ' LineCost=' + r.LineCost +
            flag
        );
        if (pn && !visited.has(pn)) {
            await chain(pool, pn, depth + 1, visited);
        }
    }
}

async function run() {
    const pool = await getConnection();
    await chain(pool, 400893, 0, new Set());
    process.exit(0);
}

run().catch(e => { console.error(e); process.exit(1); });
