const getConnection = require("./db.js");
const sql = require("mssql/msnodesqlv8");

async function run() {
    try {
        const pool = await getConnection();
        
        const ordRes = await pool.request().query("SELECT OrdNo, InvoAm, DInvoIF, Gr4 FROM Ord WHERE OrdNo = 401985");
        console.log("--- Ord Header ---");
        console.table(ordRes.recordset);

        const lnRes = await pool.request().query("SELECT LnNo, ProdNo, ProdTp4, PurcNo, DPrice, NoOrg, NoFin, CCstPr FROM OrdLn WHERE OrdNo = 401985");
        const lines = lnRes.recordset.map(l => ({
            LnNo: l.LnNo,
            ProdNo: l.ProdNo,
            ProdTp4: l.ProdTp4,
            PurcNo: l.PurcNo,
            DPrice: l.DPrice,
            NoOrg: l.NoOrg,
            NoFin: l.NoFin,
            CCstPr: l.CCstPr,
            lineCost: (l.NoFin || 0) * (l.CCstPr || 0)
        }));
        console.log("--- Sales Order Lines ---");
        console.table(lines);

        const purcNos = [...new Set(lines.map(l => l.PurcNo).filter(p => p != null && p !== 0))];
        console.log("--- Distinct PurcNo ---");
        console.log(purcNos);

        console.log("--- Cost Inquiry for PurcNos ---");
        for (const pNo of purcNos) {
            const pRes = await pool.request().query("SELECT LnNo, ProdNo, ProdTp4, PurcNo, DPrice, NoOrg, NoFin, CCstPr FROM OrdLn WHERE OrdNo = " + pNo);
            const pLines = pRes.recordset;
            
            let totalCost = 0;
            let zeroNoFinCount = 0;

            pLines.forEach(l => {
                if (l.LnNo === 1) return;
                const tp4 = l.ProdTp4 ? l.ProdTp4.toString() : "";
                if (["0", "3", "5"].includes(tp4)) return;

                totalCost += (l.NoFin || 0) * (l.CCstPr || 0);

                if (tp4 === "1" && (l.NoFin === 0 || l.NoFin === null) && l.NoOrg > 0) {
                    zeroNoFinCount++;
                }
            });

            console.log("PurcNo " + pNo + ": TotalCost=" + totalCost + ", SuspiciousCount (NoFin=0, NoOrg>0, ProdTp4=1)=" + zeroNoFinCount);
        }

        process.exit(0);
    } catch (err) {
        console.error(err);
        process.exit(1);
    }
}
run();
