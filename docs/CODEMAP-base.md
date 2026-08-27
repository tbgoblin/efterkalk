# CODEMAP-base — Gantech Operations Hub

Opdateret: 2026-08-25 (v1.1.41). Basiskort over kodebasen: filer, moduler, endpoints og Visma-tabeller.

## Arkitektur

```
Electron (electron-main.js)
  └─ Node/Express server (server.js, port 3000)
       ├─ routes/apiRoutes.js  → createApiRouter (alle endpoints, monteret UDEN /api-prefix)
       ├─ services/            → domænelogik + SQL (Visma DB F0001, 10.2.0.3\VISMA, msnodesqlv8)
       ├─ diskCache.js         → JSON-cache på disk (cache/), versionerede nøgler
       ├─ db.js                → getConnection() singleton pool
       └─ assets/js/*.js       → klientmoduler (serveret statisk, script-tag med ?v=version)
```

- `server.js` er også HTML-shell: al side-CSS + inline JS + dashboard-markup ligger her.
- Auth: bearer tokens (`users.json`). `requireModulePermission('x')` er ægte middleware; `requireSuperadmin(req,res)` returnerer user/null og SKAL wrappes: `(req,res,next)=>{ if(!requireSuperadmin(req,res)) return; next(); }`.

## Moduler (dashboard-kategorier)

| Kategori | Modul | Klient | Backend |
|---|---|---|---|
| Salg | Efterkalkulation | inline i server.js | services/aftercalcService.js |
| Salg | SalgOrdre VIA | assets/js/via.js | services/viaService.js |
| Salg | Omsætning | inline | services/omsaetningService.js |
| Salg | Ordreindgang | inline | services/ordreindgangService.js |
| Bogholderi | Faktura | (planlagt — SharePoint-flows i Fakturasystem/) | — |
| Bogholderi | Lagerliste | assets/js/lagerliste.js | services/lagerlisteService.js |
| Produktion | Ordreoversigt | inline | inline i apiRoutes |
| Produktion | Belastning | inline | services/belastningService.js |
| Produktion | BOM/Beregner | assets/bom/bom-{core,views,beregner,main}.js | services/bomService.js |
| HR | Personalehåndbog/QMS | assets/js/qms-ph.js | services/phCrawlerService.js, qmsService.js |

## Lagerliste (services/lagerlisteService.js)

- Cache-nøgle: `lagerliste_v27`; `currentMemoryCache` holder seneste payload (snapshots kræver den).
- Kategorier: plateGroups (Gr6=1, '3%'), restPlateGroups (FreeInf1 FrInfTp=120 OG Gr7=1), stang (Gr6=2, ProdTr; total værdisat med FIFO/PhCstPr), gr5Items (Gr5=11, FIFO), opfolgningvare (Gr9=1, Bal+StcInc−ShpRsv), nestingCutting/"Plader VIA" (Ord.Gr3=2, sidste 3 mdr., plade TrTp=5 NoFin>0 + alle produkter TrTp=7 NoFin=0). Lagerliste 2 løser SalesOrdNo via `OrdLn.R4` → seneste `ProdTr.R4` for samme produkt/linje → `Ord.R4` → `TrInf2/OrdBasNo`; `TrInf1='Søg…'` markeres som uregistreret restkilde. Negative pladelinjer er estimeret rest: vises, men CountedValue=0 indtil rest er registreret. finishedNotInvoiced, salgordreVia.
- Permanente Lagerliste-eksklusioner: OrdNo `61423`, `75330`, `131790`, `140134`, `331368` må aldrig medtages i Rest plader, Plader VIA, Færdige SO eller Salgsordre VIA.
- Snapshots: måned (`data/lagerliste/YYYY-MM.json`) + dags-snapshot; gemmes i baggrund via setImmediate, kompakt JSON.
- `lookupProduct(prodNo)` → Vareopslag: Prod+StcBal nøgletal, ShpBal-partier (RestBal≠0), aktive reservationer (Rsv NoRsv>0 + Ord/Actor), åbne ordrelinjer (OrdLn NoOrg−NoFin≠0).
- Endpoint: `GET /lagerliste/vareopslag/:prodno` (module-perm lagerliste).
- Periodesammenligning: `GET /lagerliste/snapshot-months` (liste af YYYY-MM) + klient `lagerlisteComparePeriods()` (Periode A/B: aktuel/måned/snapshot → kategori-tabel + bevægelser pr. produkt/ordre med Ny/Udgået/Ændret + Ind/Ud/Netto). `Fra → Til` sporer observerede Plader VIA↔materiale-VIA og VIA→Færdige SO flyt pr. Salgsordre. Materialebalance (FIFO) tæller kun lagerfald mod positivt materialetilløb i VIA; ekskluderer VIA Tid og lagerindgange. Manuelle afstemninger gemmes i bruger-skrivbar `lagerliste_reconciliations.json` via superadmin API; kræver note og ændrer aldrig Visma eller Lagerliste-totaller.
- TEMP: `GET /lagerliste/reservations-debug` (superadmin, read-only, `?table=X&prodno=Y` itererbar) — fjernes efter reservationsfeature.

## Salgsordre VIA (services/viaService.js)

- Cache: `salgordre_via_v27`. Komponenter pr. ordre: TimeCost (ResourceMinutes), MaterialCost/"VIA Laser" (ProdTp4=2, NOT LIKE '%L'), StangCost (Gr6=2), PurchasedPartCost/"Indkøbt dele".
- Indkøbt dele-regel (= aftercalcService.isPurchasedPartLine): ProdTp4='2' AND PurcNo>0 AND ikke '%L'-produkt; pris = COALESCE(NULLIF(DPrice,0), CCstPr) fra linket indkøbsordre (TrTp=6); qty = NoFin || NoOrg. Verificeret live: 410749=53.000, 410759=0.

## Visma-tabeller (verificeret mod live DB)

- `Ord`: OrdNo, TrTp (1=salg, 5/7=produktion, 6=indkøb), CustNo, DelDt, Gr3 (2=nesting), Gr4, InvoAm, OrdDt, R4 (link produktionsordre→salgsordre).
- `OrdLn`: LnNo, ProdNo, TrTp, ProdTp4 (linjetype), NoOrg (bestilt), NoFin (forbrugt/færdigmeldt — starter altid på 0, tæller op når dele tages fra lager og indgår i produktet), NoRsv (reserveret!), NoPic, DPrice, CstPr, IncCst, PurcNo (link til indkøbsordre), TrInf2 (kildeordre), TrInf4 (rute).
- `StcBal` (lagersaldo, StcNo=1): Bal, StcInc, ShpRsv (= SUM af aktive reservationer), ShpRsvIn, PicNotR, PoPhStB (fysisk), PhCstPr (FIFO).
- `ShpBal` (vareparti/lot): ShpNo, Loc, RestBal, NoRsv, CstPr/CCstPr, RecDt, SupNo, OrdNo (modtagelsesordre), OriOrdNo/OriOLnNo.
- `Rsv` (reservation → ordre-link, opdaget 2026-08-25): OrdNo, OrdLnNo, ProdNo, ShpNo, NoRsv (aktiv rest), NoPic, NoFin, NoRlz, CstPr/CCstPr, PurcNo. Aktiv reservation = NoRsv>0. SUM(Rsv.NoRsv) pr. produkt ≈ StcBal.ShpRsv.
- `Prod`: Descr, Gr5 (11=komponentlager), Gr6 (1=plade, 2=stang), Gr9 (1=opfølgning), ProdGr, Inf (standardpris, komma-decimal), NWgtU, HgtU/WdtU/LgtU.
- `Actor`: CustNo→Nm (kunde), EmpNo→Nm.
- `FreeInf1` FrInfTp=120: restplader. `ProdTr`: lagerbevægelser (StcMov, FrStc, FinDt).

## Konventioner

- Cache-invalidering = bump versionsnøgle (lagerliste_vN, salgordre_via_vN, order_margin_vN, aftercalc_vN).
- Alle SQL: `WITH(NOLOCK)`, `TRY_CONVERT(decimal(18,6), ...)`, parametriserede inputs.
- UI-sprog: dansk (åøæ). Beløb: `Intl.NumberFormat('da-DK')` + ' DKK'. Visma-datoer: int YYYYMMDD.
- Server genstarter IKKE automatisk — kræver kill node/electron + `npm start`.
- Version bump: package.json (electron-builder artefakt + /health).
