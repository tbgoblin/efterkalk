const getConnection = require('./db.js');

async function inspect() {
    try {
        const pool = await getConnection();
        const request = pool.request();
        
        const query = "SELECT OrdNo, LnNo, ProdNo, ProdTp4, PurcNo, TrInf2, TrInf4, NoOrg, NoFin, CCstPr, DPrice " +
                      "FROM OrdLn " +
                      "WHERE OrdNo IN (400526, 400529, 400572) " +
                      "AND (ProdNo LIKE '1005401473%' OR ProdNo LIKE '%L')";
        
        const result = await request.query(query);

        console.log('Results:');
        result.recordset.forEach(ln => {
            console.log(JSON.stringify(ln));
        });
        
        const ord400572 = result.recordset.filter(ln => ln.OrdNo === 400572);
        if (ord400572.length > 0) {
            console.log('Order 400572 found, ' + ord400572.length + ' lines.');
        } else {
            console.log('Order 400572 NOT found or no matches.');
        }

    } catch (err) {
        console.error(err);
    } finally {
        // Not strictly necessary for a quick script, but cleaner
        process.exit();
    }
}

inspect();
