# CODEMAP-base — Gantech Operations Hub

Opdateret: 2026-09-04 (v1.1.49). Basiskort over kodebasen: filer, moduler, endpoints og Visma-tabeller.

## Indeks

1. [Arkitektur](#arkitektur)
2. [Moduler](#moduler-dashboard-kategorier)
3. [Omsætning](#omsætning)
4. [BOM laserparametre](#bom-laserparametre)
5. [BOM bukberegning](#bom-bukberegning)
6. [Lagerliste](#lagerliste-serviceslagerlisteservicejs)
7. [Salgsordre VIA](#salgsordre-via-servicesviaservicejs)
8. [Visma-tabeller](#visma-tabeller-verificeret-mod-live-db)
9. [Konventioner](#konventioner)

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

## BOM laserparametre

- Parametrene vedligeholdes manuelt i GOH-tabellen `dbo.BomLaserParameters`; de kommer ikke fra Visma `FreeInf2`.
- Tabellen oprettes idempotent af `services/gohDataService.js` ved første brug. Unik nøgle: `(ProdNo, Machine)`.
- BOM → Parametre læser og skriver tabellen uden cache. Skrivning via programmet kræver superadmin.
- Endpoint: `GET/POST /bom/calculators/laser-params`. Beregneren kræver en aktiv, eksakt række for varenr. og maskine, medmindre lasertiden er overstyret manuelt.
- Superadmin kan køre en eksplicit engangsimport fra `BOM.xlsm`/`skæreparametre` via `POST /bom/calculators/laser-params/import-excel` eller knappen i Parametre. Importen indsætter som standard kun manglende `(ProdNo, Machine)` og bevarer eksisterende GOH-værdier.
- Eventuelle rækker fra den tidligere fejlagtige kilde med `Source = 'visma-seed'` ignoreres.

## BOM bukberegning

- GOH-tabeller: `dbo.BomBendingMachines`, `dbo.BomBendingHandlingBands` og `dbo.BomBendingActualSamples`; de oprettes idempotent af `services/gohDataService.js`.
- Hvis appens Windows-login ikke har DDL-rettigheder, køres `docs/sql/2026-09-04-bom-calculator-parameters.sql` én gang af en databaseadministrator. Appen skal ikke tildeles permanent `CREATE TABLE`.
- Maskine og håndteringsklasser redigeres i BOM → Parametre eller direkte i GOH. Programskrivning kræver superadmin; læsning sker uden cache.
- Automatisk tid pr. emne = maskincyklus pr. buk + bagstop/vinkelkorrektion + pålæg/aflæg + rotationer/vendinger. Opstart holdes separat pr. ordre.
- Maskinkontrol bruger gennemsnitlig bukkelængde og estimeret luftbukkekraft. Beregningen afvises ved overskredet længde eller sikker kraftkapacitet.
- Buk-input: antal, samlet bukkelængde, gennemsnitsvinkel, V-åbning, trækstyrke, rotationer og vendinger. Minutter og opstart kan overstyres manuelt.
- Endpoints: `GET /bom/calculators/bending-params`; superadmin `POST /bom/calculators/bending-machines`, `/bending-handling-bands`, `/bending-actual-samples`.
- Test: `test/bomBendingCalculation.test.js` dækker klassevalg, tidsberegning og kapacitetsafvisning.

## Omsætning

- UI og beregning ligger inline i `server.js`; kalenderlogik ligger i `assets/js/omsaetning-daily-thresholds.js`; omsætningsdata hentes via `services/omsaetningService.js`.
- Administration (kun superadmin) har en årsvælger og 12 redigerbare arbejdsdage. Kalenderforslaget udelader weekender, danske helligdage og registrerede virksomhedsferieuger. Manuelle værdier er heltal `0-31` pr. `YYYY-MM`.
- Dagsmål: `0-punkt / dag (DKK)` og `Budget / dag (DKK)`. Månedens grænser beregnes som `dagsværdi × arbejdsdage / 1.000.000`.
- Periodeoversigten summerer kun måneder med bogført omsætning. Den viser samlet omsætning, 0-punkt, budget samt forskel i Mio/% til både 0-punkt og budget.
- GOH `dbo.AppState` er primær lagring. `omsaetning_working_days` gemmer `{ months, updatedAt }`; `omsaetning_daily_budget_settings` gemmer `{ useDailyBudget, dailyBreakEvenDkk, dailyBudgetDkk, updatedAt, updatedBy }`; `omsaetning_thresholds` gemmer kundeoverride for manuelle/daglige tærskler.
- Kundegrænser er database-first i `services/omsaetningThresholdsService.js`: API svarer først succes efter GOH `MERGE`; lokal `omsaetning_thresholds.json` er kun mirror/fallback efter en vellykket DB-skrivning.
- Endpoints: `GET /omsaetning/working-days`, `GET|POST /omsaetning/daily-budget-settings`, `GET|POST /omsaetning/customer-threshold/:custno`, `GET|POST /admin/working-days` (admin-endpoints kræver superadmin).
- Test: `test/omsaetningDailyThresholds.test.js` dækker danske helligdage, ferieuger, månedlige override, normalisering og beregnede månedsmål.

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
