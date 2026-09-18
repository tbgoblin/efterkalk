# CODEMAP-base — Gantech Operations Hub

Opdateret: 2026-09-15 (v1.1.68). Basiskort over kodebasen: filer, moduler, endpoints og Visma-tabeller.

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

- `server.js` er også HTML-shell: hovedparten af side-CSS og inline JS ligger her; dashboardens model/controller er i `assets/js/dashboard.js`, med separat `assets/dashboard.css`.
- Auth: bearer tokens (`users.json`). `requireModulePermission('x')` er ægte middleware; `requireSuperadmin(req,res)` returnerer user/null og SKAL wrappes: `(req,res,next)=>{ if(!requireSuperadmin(req,res)) return; next(); }`.

## Moduler (dashboard-kategorier)

Widgetindstillinger (2026-09-16): `boards[].options[widgetId]` normaliseres med `widgetOptions`; kun afvigelser fra standard gemmes. Alle widgets får egne titel-/søgeindstillinger og relevante felter i `gohWidgetSettings`. Gem/klik i kundegraf tilpasser den valgte visning pr. bruger, også standardvisninger med deres oprindelige id. `data(widget)` beregner fra autoriserede fælles kilder med lokale widgetfiltre; CSV bruger samme udsnit. Belastning: `loadHorizons` deduplikerer synlige horisonter (1-90, standard20); serverbridgen kræver og cacher dem separat og sender `loads[days]`, aldrig en afkortet kopi af en anden horisont. Indstillingsændringer henter ikke data under dialogredigering, kun ved Gem.

Kundegraf: `customerShares` grupperer bogført omsætning pr. kundenummer, alternativt forskellige fakturaordrenumre fra Efterkalk. Top3-6 + Andre, SVG-ringsegmenter med keyboard/tooltip og samme detail-/CSV-udsnit. Negative kundesaldi vises særskilt med nettototal og indgår ikke som positive arealer. Aktivitet kræver Efterkalk+Omsætning og henter ikke ordrekost. Kvartals-KPI viser de tre måneder fra samme bogførte datasæt, inklusive korrektioner; fremtidige måneder er null.

Seneste ordreindgang: `ordreindgangService.getRecentOrders` og GET `/ordreindgang/recent?from=&to=` kræver Ordreindgang+Omsætning. En parameteriseret query på aktiv database, TrTp1/OrdTp1, OrdDt (ikke teknisk oprettelsestid), højst90dage, ordredato/nummer faldende, entydig kundetilknytning. Aktuel værdi `(InvoSF+InvoIF)*(ExRt/100)` bevares; kreditfilter i dashboard. Bridgen cacher seneste90dage i15min; widget filtrerer1-90dage og5/10/20/50rækker lokalt. `recent-orders` har egen source-id.

GOH-præferencesvar og PUT bruger `schemaVersion:4`; klienten afviser ældre backend før PUT, og serveren afviser gamle klienter uden schema4, så standardtilpasninger ikke bliver tabt. `normalizeConfig` bevarer 12 `custom-*` dashboards plus overrides for `economy`, `production`, `sales` og `management`. `availableBoards` samler tilladte standardvisninger (med personlige overrides) og personlige faner i samme værktøjslinje; dropdown er fjernet. `saveBoard` bevarer fanernes rækkefølge og gemmer via GOH-store. Standardreset sletter kun brugerens override; grundskabelonerne og andre profiler er uændrede. Side-menuens modulknapper er alfabetiske efter dansk navn. Genstart og genindlæsning kræves ved backendopdatering. Browserfixtures verificerer oprettelse, standardredigering, widgetvalg/layout, brugerskift, reset/annuller, GOH-fejl/retry og 16 faner ved 1440/390 px; Node-suite og nyt GOH-schema live er ikke verificeret i editorsessionen.

Omsætningens stacked-månedsdiagram bruger `GohReportCharts.revenueStacked` fra `assets/js/report-charts.js`, også i Ledelsesrapport. Positive/negative stakke, tomme måneder, netto-totaler og konto-farver bevares. Netto-totalen placeres over hele den positive stak med hvid tekstkontur, så kreditposter ikke flytter labelen ind i søjlen. Kunde-sammenligning og total-trend ligger fortsat i `renderOmsaetningCharts`; `appendMonthTotalLabel` bruges i kunde-sammenligningen. `printOmsaetningReport` kopierer SVG-`outerHTML`, så labelen er en del af den trykte graf.

Seneste rapportændring (2026-09-18): Omsætning vises nu altid i ét diagram med én legenda for hele den valgte periode (1-36 måneder); den tidligere 12-månedersopdeling nedenfor er erstattet. Browserkontrol: 1/12/13/15/24/36 måneder med kreditter og alle nettoværdier; ingen labeloverlap ved 15/36 måneder. A4-landskab med 15 måneder: ét panel på 272 px inden for 726 px sidehøjde. Klientversion `ledelsesrapport.js?v=17`.

Uafhængige rapportperioder: `customerFrom/customerTo` er inklusive måneder (1-36), uafhængigt af `from/to`. Samme periode genbruger revenue-tasken; forskellige perioder henter separat kundesummary med samme kontofilter. `loadFrom/loadTo` er inklusive plandatoer (1-180 dage), sendt som valgt start og spænd til den eksisterende Belastning-query; efterfølgende rækker fjernes, rest før start bevares i grafen. `viaFrom/viaTo` er et valgfrit ordredatofilter på aktuelt åbne ordrer, ikke historisk kost. Manglende ordredato tælles særskilt ved datofilter. Klienten kontrollerer returnerede perioder og afviser gammel backend. Ændring af felter alene udfører ingen dataquery; først Dan rapport henter data. Browserfixture med årsomsætning, en måneds kunder, uafhængige uger og planår er kontrolleret uden Visma-kald.

Login/session (2026-09-16): `GET /auth/session` genopretter token og safeUser fra en gyldig HttpOnly-cookie ved hovedsidens opstart, med no-store også på 401. Sessionscookie-navn scopes efter lokal port; sessioner udløber efter 8 timer, logout eller servergenstart. Electron åbner interne sider i samme sessionspartition; Tilbage fra et barn til `/` lukker barnet og fokuserer hovedvinduet. `electron-preload.js` eksponerer kun info/restore/remember/forget, og IPC afviser andre webContents, frames og origins end hovedsidens `/`. `services/desktopCredentialService.js` bruger Electron safeStorage/Windows DPAPI til krypterede bytes i appens Windows-userData, aldrig GOH eller fælles cache. Scope inkluderer Windows-bruger/domæne, servermaskine, clientnavn og RDS/lokal; Chromium-partition inkluderer også sessionnavn. Ét senest valgt app-login pr. scope; username er med i den krypterede payload. Husk kode er opt-in efter godkendt login, skjult uden Windows-kryptering/identificerbar RDS-client. Native restore returnerer kun auth-resultat, aldrig kode. Native handlinger serialiseres; logout glemmer login og afviser sene sessionssvar. Browserudgaven bruger adgangskodeadministratorens autocomplete-felter. Testtilføjelser i `test/authService.test.js` dækker cookies, session-endpoint, klientrestore/cancel/opt-in og mock-krypteret lagring. Node-tests og ægte DPAPI/Electron/RDS-retur er endnu ikke kørt; kræver genstart og driftskontrol.

Ledelsesrapport (2026-09-18): `assets/ledelsesrapport.html` og `assets/js/ledelsesrapport.js`; GET `/ledelsesrapport/config` og `/ledelsesrapport/data` kræver `ledelsesrapport`-rettighed. Administrationens fane Ledelsesrapport · Standard gemmer et globalt preset i GOH `dbo.AppState` under `ledelsesrapport_defaults`; GET/POST `/admin/ledelsesrapport-defaults` kræver superadmin. Presettet omfatter omsætnings- og kundeperioder, kundeantal, Ordreindgang-uger og synlige linjer, Belastning-datoer og en ordnet liste af valgte `ResGr` samt VIA-periode. Ressourcer kan trækkes med drag & drop eller flyttes med op/ned-kontroller i både administration og rapportdialog; hver valgt række viser side og placering. Op til 32 dage bruges fire ressourcer pr. printside, over 32 dage to pr. side. Arrayrækkefølgen gemmes og styrer rækkefølgen af Belastning-graferne. `/ledelsesrapport/config` leverer presettet til alle brugere med moduladgang, men brugeren kan ændre felterne, linjerne, ressourcerne og deres rækkefølge for den enkelte rapport. Gamle preset uden valgene viser alle ressourcer og linjer. Konti er ikke del af presettet, fordi alle tilgængelige omsætningskonti altid anvendes. Ordreindgangs søjler er fast grundlag; linjerne Ordre, MA3, Budget og periodegennemsnit kan vælges enkeltvis. Nuluger udelades fra MA3. Belastning filtreres lokalt på stabil `ResGr`, efter data er hentet. Omsætning vises samlet, og største kunder viser omsætning samt DB i DKK og procent. Den tidligere VIA Kostfordeling er erstattet af en statisk Ordrebeholdning-graf, som genbruger `omsaetningService.getOrderFlow()` og viser primo, tilgang, faktureret og ultimo fordelt på tidligere og nye ordrer. Grafen bruger rapportens `to`-måned, dog højst den aktuelle måned; API'et sender kun de summerede Order Flow-tal. VIA-data hentes fortsat til rapportens eksisterende KPI'er. Test: `test/ledelsesrapport.test.js`. Backendændringer kræver genstart.

Dashboard-layout: `normalizeLayout` begrænser x/y/w/h på 12 kolonner; `arrangeWidgets` pakker uden overlap og udfylder plads under korte widgets. Direkte drag fra header/håndtag og resize fra hjørnet gemmer ved pointerup; pointercancel gemmer ikke. Piletaster har samme funktion. `Indret` bruger en eksplicit `layoutDraft` med Gem/Annuller. Standardvisninger beholder id/navn ved layoutændringer og gemmes som brugerens override. `boards[].layout` er bagudkompatibel med `wide`. Fast grid-række 24 px + gap 16 px, widget-body ruller, header/footer forbliver synlige. Layoutændringer skriver kun præferencer, ikke produktionsdata; ingen Visma-læsning under drag. Ordreflow-DOM, fokus og scrolldata genbruges ved normale renders.

Dashboard-profiler: GET/PUT `/dashboard/preferences?profile=<id>` kræver login. Serveren afleder en SHA-256-nøgle fra kanonisk brugernavn samt aktiv SQL-server/database, aldrig klientens brugerfelt. GOH `dbo.AppState` gemmer `dashboard_v1_<hash>` med `{version,config}`; `setAppState` bruger HOLDLOCK og opt-in `expectedVersion`, så en forældet version giver 409. `getAppState({strict:true})` skelner mellem manglende profil og DB-fejl (503); eksisterende fail-soft kald er uændrede. `createPreferenceStore` indlæser GOH før render, importerer kun legacy ved tom GOH-profil, serialiserer writes og ignorerer sene svar efter scope-skift. Den lokale nøgle med suffiks `:goh` indeholder kun recovery `{version,config,pending}`. Filtre inkl. Ordreflow-state gemmes; søgning debounces 450 ms og afsluttes ved pagehide/reset. Fejl og konflikter kræver eksplicit retry/hentning, aldrig tavs overskrivning.

Dashboard Belastning-søgning: `loadSearchFilters(query, unfilteredLoad)` vælger lokal ressourcematch eller backend `kunde`/`ord`. `dashboardLoadQuery` går via `ensureSources(..., query)` til kildeparametre/cache-nøgler, og kontekstens `loadQuery` forhindrer visning af forrige søgning. `buildData` anvender ikke et ekstra ressourcenavnfilter på et allerede kunde-/ordrefiltreret payload (`kunde`/`ord`). Eksisterende 20-dages Belastning-query og beregningsformler ændres ikke. UI-debounce 450 ms + scoped kildecache. Verificeret live `logi`: 56 aggregerede rækker, 561,53051405 dagtimer og 1420 kapacitetstimer på de returnerede ressourcer (15.09.2026 + 20 dage).

Interaktivt Ordreflow: `assets/js/order-flow.js` eksporterer den fælles `build`/`summarize`-model og browserens `mount`; `assets/order-flow.css` styrer desktopgrafen. Widget-id `order-flow` har modulrettighed `omsaetning`, men egen datakilde `source: 'order-flow'`, egen måned og CSV. Standard i Økonomi/Salg/Ledelse; personlige layouts ændres ikke automatisk. Samme komponent monteres i `renderOmsaetningMonthDetail`. Barer, kohorter, færdigfaktureret-filter, søgning og pagination er lokale. Måneds-/scope-token afviser sene svar. Præferencer er kun i hukommelsen pr. bruger/database/view og nulstilles ved logout.

`GET /omsaetning/order-flow?month=YYYY-MM&customers=...` kræver Omsætning; `getOrderFlow` i `services/omsaetningService.js` udfører én SQL-batch. Salgsordrer (`TrTp=1, OrdTp=1`) før cutoff, med rest eller faktura/ordredato fra månedens start; positive ordreværdier, kundeafgrænsning og beskyttet Actor-join. Fakturaer kobles entydigt via Ord/CustTr, på faktura+kunde, også mod ordrer uden for kandidatlisten. AcTr bruger alle `10_Omsætning`-konti, SrcTp 1/9, AcYrPr før/i valgt periode; indeværende måned stopper ved i dag. Samlet fakturahistorik kontrolleres mod Ord.InvoSF i DKK med 1 DKK tolerance. Manglende/tvetydig historik giver udeladte ordrer og tydelig delsum, ikke opdigtet primo. Historik er rekonstruktion på aktuelle ordreværdier, ikke snapshots. Regulering fastholder regnskabsligningen ved fx overfakturering. `fetchOrderFlowCached` deler 15-minutters kildecache/in-flight/error-backoff mellem Home og månedspanel, afgrænset til bruger/database/måned/kunder. Manuel opdatering omgår TTL; filtre foretager ingen forespørgsler. Test i `test/omsaetningMonthDetail.test.js` og `test/dashboard.test.js`.

SalgOrdre VIA med Omsætning-adgang sammensættes fra `getOrderFlow` for indeværende måned/alle kunder: `getOpenViaOrders` vælger modellens `closing > 0.01`, ikke SQL-kandidatlisten. `buildViaBacklog` bevarer ordrerne, beriger med kost fra eksplicit parameteriserede ordre-id'er og returnerer `RemainingSalesValue`, `CostDataAvailable`, `unknownCount` og `excludedResidualDkk`. KPI, tabel og CSV summerer de samme filtrerede restbeløb. Ukendt kost bliver ikke nul; dublerede rækker afvises. Uden Omsætning-adgang bruges det historiske produktionsudvalg. Lagerliste kalder fortsat getterens historiske standardfilter (`Gr12 <> 10` og aktive produktionsstatusser); kost-/reservations-/overlapregler er ikke ændret. Cache er adskilt pr. server/database/dato/grundlag, med 5 min TTL og in-flight-deduplikering. `cached=1` udfører ingen SQL. Enkeltordre-svar og klienten kontrollerer ordre-id; opdateringen ugyldiggør fuldlistens cache. Klienten afviser gamle backend-kontrakter, ignorerer sene sessionssvar og bevarer tidligere rækker med synlig fejl. Regressioner: `test/viaService.test.js`.

Margin-konsistens: `/order-margin/:ordno` går altid gennem `getOrComputeOrderMargin`, som kontrollerer både kosteksklusioner og den aktuelle cached Efterkalk-summary, før memory/disk-margin genbruges. Route-laget må ikke læse eller overskrive margin-cache direkte. Ordredetaljen kalder `synchronizeDashboardOrderMargin` med effektiv kost inkl. stykliste/eksklusioner; både kalender- og regnskabsårsdata opdateres inden for bruger/database-scope. Sene kostsvar må ikke overskrive allerede opdateret kost. Regression: ordre 407940, faktura 5718,80 og kost 4812,49 giver DB +906,31 DKK (15,85 %), ikke gammel kost 7848,49.

Ansvarlig-widgetten (`sellers`) viser fakturabeløb og DB % som to bjælker pr. ansvarlig. `buildData` beregner `dbPercent = SUM(DB)/SUM(revenue)*100` kun for kendt kost og eksponerer `costCount/count` for delvise resultater. Ingen positiv omsætning i beregningsgrundlaget giver `null`, negative DB bevares. Widgetten udløser samme begrænsede kostkø som best/risk/coverage og eksporterer DB % samt kostdækning til CSV.

Manuel dashboard-opdatering: `refreshVisibleSources` deduplikerer synlige, tilladte widgetkilder og samler fejl pr. kilde. `refreshDashboardSources` bruger `createSourceCache.load(..., true)` og Omsætnings `forceRefresh`, genhenter ordrenoter og bruger VIA `loadSalgordreVia(true, { loadReservations: false })`. Efterkalk-payload markerer `refreshCosts` for at springe lokal margin-cache over uden at rydde server-cache. Klientknappen fastholder layout/filtre, låses under opdatering og ignorerer sene statusbeskeder efter scope-skift.

`isCreditOrder(row, notes)` udelukker `InvoAm < 0` og `orderNotesCache[OrdNo].isCreditNote === true` fra dashboardens Efterkalk-widgets, optællinger, CSV og kostkø. Bridgen sender aktuelle noter og opdaterer efter indlæsning/ændring/sletning. Andre nulbeløb og negative DB bevares; bogført Omsætning ændres ikke. SQL-kilden kræver `O.TrTp = 1`.

`openDashboardOrder()` sætter `orderDetailReturnModule = 'dashboard'`, når Efterkalk åbner. `closeOrderDetailModal()` bevarer oprindelsen ved rapport-til-ordre, men forbruger og nulstiller den ved endelig lukning: dashboard, SalgOrdre VIA eller normal Efterkalk-liste. `goToDashboard()` nulstiller også oprindelsen ved direkte Home-navigation.

Dashboard har fire rettighedsfiltrerede skabeloner (Økonomi/Produktion/Salg/Ledelse) og op til 12 personlige layouts. `GohDashboard.update()` modtager brugernavn, databaseprofil, `canAccessModule`, `economicPeriod` og kildedata fra bridgen i `server.js`. Økonomiske widgets kræver Omsætning-adgang. `ensureDashboardSources()` modtager den valgte periode og bruger bruger/database-afgrænset 15-minutters cache med samtidighedsdeling og 60 sekunders fejl-backoff. `economicRange(today, period)` returnerer som standard regnskabsåret fra juli; kalenderåret starter 1. januar med Visma-perioderne `[(YYYY-1)07, YYYY07)`. Efterkalk henter alle salgsfakturaordrer fra kildens startdato til i dag, aldrig den begrænsede `orderListData`; kreditfilteret bevares. Omsætning genbruger `fetchOmsaetningSummaryCached` med standardkonti. Datointervaller indgår i cache-nøgler; klienten skjuler tidligere økonomiperioders data under skift. `buildData()` bruger `periodFrom` til rækker, måneder og viste datoer. Kost indlæses via `hydrateOrderCosts` og `/order-margin/:ordno`, højst tre ad gangen med annullering ved scope-skift. Belastning genbruger modulernes loader ved samme filtre; dashboardens standard er i dag + 20 dage, uafhængigt af økonomiperioden. Dagarbejde, kapacitet og aften konverteres fra minutter til timer; rest før i dag vises separat. Præferencer gemmes under `goh:dashboard:v1:<database>:<username>`, uden forretningsdata. `test/dashboard.test.js` dækker kalender-/regnskabsår og kildeparametre, datagrundlag over 150 ordrer, DB, kapacitet, rettigheder, cache/præferenceisolering og genereret HTML/JS.

| Kategori | Modul | Klient | Backend |
|---|---|---|---|
| Salg | Efterkalkulation | inline i server.js | services/aftercalcService.js |
| Salg | SalgOrdre VIA | assets/js/via.js | services/viaService.js |
| Salg | Omsætning | inline | services/omsaetningService.js |
| Salg | Ordreindgang | inline | services/ordreindgangService.js |
| Ledelse | Ledelsesrapport | assets/js/ledelsesrapport.js, assets/js/report-charts.js | routes/apiRoutes.js (genbruger modulservices) |
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

## MånedsDB / Efterkalk-snapshots

- Månedsdialogen ligger inline i `server.js`; fakturaordrer kommer fra `GET /efterkalk/customer-invoices`, mens marginen pr. ordre kommer fra `GET /order-margin/:ordno`.
- Der findes to eksplicitte kostmodeller: `registreret` = `totalCost`; `korrigeret` = `totalCost + styklisteFallbackCost`. Sidstnævnte medregner `NoOrg * CCstPr` for relevante stykliste-operationer, når den effektive registrerede tid er 0.
- GOHCache gemmer månedsrækker i `dbo.EfterkalkSnapshotRun` og `dbo.EfterkalkOrderSnapshot`. `StyklisteFallbackCost` oprettes idempotent af applikationen og findes også i `docs/GOH-EFTERKALK-SNAPSHOT-SCHEMA.sql`.
- Kunde- og ordre-CSV indeholder begge kostmodeller samt fallbackbeløbet. Procenter eksporteres med et `%`-tegn, så Excel ikke fortolker fx 7,80 som 780%.
- Endpoints: `GET|POST /efterkalk/month-snapshot`, `GET /efterkalk/customer-trend-snapshots`, `POST /efterkalk/snapshot-backfill`.

## Lagerliste (services/lagerlisteService.js)

- Cache-nøgle: `lagerliste_v36`; tidligere cache med forkert VIA-perimeter genbruges ikke. Værdiansættelsesschema og historiske snapshots er uændrede. `currentMemoryCache` holder seneste payload (snapshots kræver den).
- Kategorier: plateGroups (Gr6=1, '3%'), restPlateGroups (FreeInf1 FrInfTp=120 OG Gr7=1), stang (Gr6=2, ProdTr; total værdisat med FIFO/PhCstPr), gr5Items (Gr5=11, FIFO), opfolgningvare (Gr9=1, fysisk PoPhStB × FIFO minus verificeret overførsel af modtagne indkøbsdele til VIA; fysisk værdi og overført værdi bevares som auditfelter), nestingCutting/"Plader VIA" (Ord.Gr3=2, sidste 3 mdr., plade TrTp=5 NoFin>0 + alle produkter TrTp=7 NoFin=0). Lagerliste 2 løser SalesOrdNo via `OrdLn.R4` → seneste `ProdTr.R4` for samme produkt/linje → `Ord.R4` → `TrInf2/OrdBasNo`; `TrInf1='Søg…'` markeres som uregistreret restkilde. Negative pladelinjer er estimeret rest: vises, men CountedValue=0 indtil rest er registreret. finishedNotInvoiced, salgordreVia.
- Permanente Lagerliste-eksklusioner: OrdNo `61423`, `75330`, `131790`, `140134`, `331368` må aldrig medtages i Rest plader, Plader VIA, Færdige SO eller Salgsordre VIA.
- Snapshots: måned (`data/lagerliste/YYYY-MM.json`) + dags-snapshot; gemmes i baggrund via setImmediate, kompakt JSON. Superadmin-knappen `Kopiér lokale måneder til GOH` kalder `POST /lagerliste/migrate-local-to-goh`, verificerer hver GOH-skrivning og bevarer altid de lokale originaler; modstridende GOH-data overskrives kun efter ekstra bekræftelse og sikkerhedskopieres først.
- `lookupProduct(prodNo)` → Vareopslag: Prod+StcBal nøgletal, ShpBal-partier (RestBal≠0), aktive reservationer (Rsv NoRsv>0 + Ord/Actor), åbne ordrelinjer (OrdLn NoOrg−NoFin≠0).
- Den aktuelle Lagerliste viser `Reserveret til ordre · info` efter Opfølgningsvarer: kun aktive Rsv med verificeret ShpBal-parti og åben salgsordre. Tabellen ændrer ikke selv totalen. Modtagne indkøbsdele med aktiv reservation flyttes fra Opfølgningsvarer til VIA ved FIFO; øvrige reservationer forbliver lager. Historiske snapshots viser ikke aktuelle reservationer.
- Vareopslag: `GET /lagerliste/vareopslag/:prodno` viser Prod/StcBal, ShpBal-partier, aktive Rsv og åbne ordrelinjer (module-perm `lagerliste`).
- Snapshot-endpoints: `GET|POST /lagerliste/snapshot/:month`, `GET|POST /lagerliste/snapshots`, `GET|DELETE /lagerliste/snapshots/:id`. Skrivning/sletning kræver Superadmin. `POST /lagerliste/migrate-local-to-goh` verificerer kopier, bevarer lokale filer og sikkerhedskopierer før godkendt overskrivning.
- Lagerliste 2-endpoints: `GET /lagerliste2/routes/current`, `GET /lagerliste2/reservations/current`, `POST /lagerliste2/movement-evidence` (module-perm `lagerliste`).
- Periodesammenligning: `GET /lagerliste/snapshot-months` (liste af YYYY-MM) + klient `lagerlisteComparePeriods()` (Periode A/B: aktuel/måned/snapshot → kategori-tabel + bevægelser pr. produkt/ordre med Ny/Udgået/Ændret + Ind/Ud/Netto). `Fra → Til` sporer observerede Plader VIA↔materiale-VIA og VIA→Færdige SO flyt pr. Salgsordre. Materialebalance (FIFO) tæller kun lagerfald mod positivt materialetilløb i VIA; ekskluderer VIA Tid og lagerindgange. Manuelle afstemninger gemmes i bruger-skrivbar `lagerliste_reconciliations.json` via superadmin API; kræver note og ændrer aldrig Visma eller Lagerliste-totaller.
- TEMP: `GET /lagerliste/reservations-debug` (superadmin, read-only, `?table=X&prodno=Y` itererbar) — fjernes efter reservationsfeature.

## Salgsordre VIA (services/viaService.js)

- Cache: `salgordre_via_v35_<databasehash>_<dato>_backlog|production`. Komponenter pr. ordre: TimeCost (ResourceMinutes), MaterialCost/"VIA Laser" (ProdTp4=2, NOT LIKE '%L'), StangCost (Gr6=2), PurchasedPartCost/"Indkøbte dele til ordre".
- Indkøbte dele-regel: ProdTp4='2' AND PurcNo>0 AND ikke '%L'-produkt. `NoOrg` alene tælles aldrig. Forbrugt `NoFin` værdisættes med indkøbsprisen. Modtaget, endnu ikke forbrugt mængde tælles kun inden for bestilt antal; for Gr9 kræves også fysisk beholdning og aktiv Rsv på samme produktionslinje, og overførslen værdisættes med FIFO, så samme værdi trækkes fra Opfølgningsvarer.
- Linjer med samme salgsordre + indkøbsordre + produkt aggregeres, så modtaget mængde ikke gentages. `PurchasedPartDetails` følger samme `NoPac`-fordeling som `PurchasedPartCost` ved overlap med Færdige SO.
- UI-total, KPI, sortering og CSV medregner alle `PurchasedPartCost`. Detaljen viser indkøbsordre, produkt, bestilt, modtaget, forbrugt, enhedspris og medregnet VIA.
- Endpoints: `GET /salgordre-via` og `GET /salgordre-via/reservations` (module-perm `salgordreVia`). Hovedlisten har 5 min cache; det historiske produktionsudvalg har 10 min baggrundswarmup. Kommerciel restsaldo kræver desuden Omsætning-adgang og indlæses ved behov.

## Visma-tabeller (verificeret mod live DB)

- `Ord`: OrdNo, TrTp (1=salg, 5/7=produktion, 6=indkøb), CustNo, DelDt, Gr3 (2=nesting), Gr4, InvoAm, OrdDt, R4 (link produktionsordre→salgsordre).
- `OrdLn`: LnNo, ProdNo, TrTp, ProdTp4 (linjetype), NoOrg (bestilt), NoFin (forbrugt/færdigmeldt — starter altid på 0, tæller op når dele tages fra lager og indgår i produktet), NoRsv (reserveret!), NoPic, DPrice, CstPr, IncCst, PurcNo (link til indkøbsordre), TrInf2 (kildeordre), TrInf4 (rute).
- `StcBal` (lagersaldo, StcNo=1): Bal, StcInc, ShpRsv (= SUM af aktive reservationer), ShpRsvIn, PicNotR, PoPhStB (fysisk), PhCstPr (FIFO).
- `ShpBal` (vareparti/lot): ShpNo, Loc, RestBal, NoRsv, CstPr/CCstPr, RecDt, SupNo, OrdNo (modtagelsesordre), OriOrdNo/OriOLnNo.
- `Rsv` (reservation → ordre-link, opdaget 2026-08-25): OrdNo, OrdLnNo, ProdNo, ShpNo, NoRsv (aktiv rest), NoPic, NoFin, NoRlz, CstPr/CCstPr, PurcNo. Aktiv reservation = NoRsv>0. SUM(Rsv.NoRsv) pr. produkt ≈ StcBal.ShpRsv. Lagerliste 2 og Salgsordre VIA viser aktiv Rsv-værdi som `Reserveret lager` kun når `Rsv.ShpNo` findes i `ShpBal`, og ordrelinket via `OrdLn.R4` → `Ord.R4` → direkte salgsordre findes og er åbent efter samme statusregel som Salgsordre VIA. Parti-CstPr bruges først; Rsv-CstPr og FIFO er fallback. Beløbet er en informativ delmængde af fysisk lager og summeres ikke igen. Ugyldige/gamle links vises som ikke medregnet.
- `Prod`: Descr, Gr5 (11=komponentlager), Gr6 (1=plade, 2=stang), Gr9 (1=opfølgning), ProdGr, Inf (standardpris, komma-decimal), NWgtU, HgtU/WdtU/LgtU.
- `Actor`: CustNo→Nm (kunde), EmpNo→Nm.
- `FreeInf1` FrInfTp=120: restplader. `ProdTr`: lagerbevægelser (StcMov, FrStc, FinDt).

## Konventioner

- Cache-invalidering = bump versionsnøgle (lagerliste_vN, salgordre_via_vN, order_margin_vN, aftercalc_vN).
- Alle SQL: `WITH(NOLOCK)`, `TRY_CONVERT(decimal(18,6), ...)`, parametriserede inputs.
- UI-sprog: dansk (åøæ). Beløb: `Intl.NumberFormat('da-DK')` + ' DKK'. Visma-datoer: int YYYYMMDD.
- Server genstarter IKKE automatisk — kræver kill node/electron + `npm start`.
- Version bump: package.json (electron-builder artefakt + /health).
