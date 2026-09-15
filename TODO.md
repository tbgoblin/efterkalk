# GOH: grafisk og operationel videreudvikling

Opdateret 2026-09-15. Desktop er målplatformen; mobilarbejde er ikke bestilt.
Dette er en prioriteret revision, ikke en erklæring om at alle moduler er redesignet.

## Færdigt i denne iteration

- [x] Dashboard-søgning i Belastning understøtter kunde/ordre samt lokal ressourcefiltrering. Live-kontrol: `logi` giver planlagte timer; visningen bliver ikke tom på grund af et ekstra ressourcenavnfilter.
- [x] Personlige dashboard-positioner og størrelser: træk, resize, tastatur, pak tæt, Gem/Annuller. Standardskabeloner kopieres; layouts isoleres pr. bruger/database.
- [x] Direkte header-drag og hjørne-resize med automatisk lagring i GOH pr. bruger/tilsluttet database. Widgetvalg, navne, sidste visning, filtre og Ordreflow-state følger profilen. Legacy-import, lokal recovery, serialiserede writes, versionskonflikter og sene sessionssvar håndteres.
- [x] Indstillinger pr. widget: titel/filter, rækkeantal, Belastning-horisont/rest/aften, DB-grænse, leveringshorisont og KPI-detaljer. GOH schema2 beskytter mod gamle backends, der ikke bevarer options.
- [x] Kvartalsomsætning med tre afstemte månedstal; kunderinggraf efter omsætning/antal fakturaordrer, Top+Andre, signeret saldo og klik til kundetabel/CSV.
- [x] Seneste ordreindgang med separat kilde efter Visma-ordredato, egen dagperiode og rækkeantal. Begge nye widgets i Salg/bibliotek; eksisterende personlige visninger ændres ikke.
- [ ] Efter servergenstart: kontroller nye widget-options med GOH schema2, samtidige Belastning-horisonter og `/ordreindgang/recent` mod live-data. Browserfixtures kontrollerer data-/layoutkontrakter; fuld Node-suite er endnu ikke kørt.
- [ ] Afsluttende driftskontrol efter servergenstart: bekræft GET/PUT-rundtur i GOH og genåbning fra anden session. Fuld Node-testsuite afventer et miljø med kommandokørsel; isolerede browserkontroller er gennemført.
- [x] Fast widgethøjde med intern rulning og tæt placering under kortere widgets.
- [x] Ordreflow: måned/kohorte, søjle til ordreudsnit, kun færdigfakturerede, søgning, pagination og CSV. Samme komponent i Home og månedspanelet; fokus og scrolldata bevares ved baggrundsopdateringer.

## Fælles principper

1. Overblik, interaktiv graf, præcis tabel og ordredetalje er én sammenhængende arbejdsgang. Et grafklik ændrer et synligt filter; tilbage bevarer søgning, periode og rulleposition.
2. Graf, KPI, tabel og CSV bruger samme fulde datasæt og samme filtrering. Top-N begrænser kun visningen. Delvis datadækning mærkes; ukendt er ikke nul.
3. Hver værdi har enhed, periode, kilde og beregningsgrundlag. Salgsværdi, kost, planlagte timer, kapacitet, lager og VIA må ikke blandes på én akse eller summeres på tværs.
4. Formler og Gantech-regler bevares. Designarbejde ændrer ikke Visma, DB-procenter, regnskabsår, kosteksklusioner, reservationer eller værdiansættelse. Eventuelle dataproblemer behandles og testes særskilt.
5. Tabeller forbliver tilgængelige: højrejusterede tabulære tal, enhed i overskrift, rolig header, én sorteringsmarkør, fastholdt nøglekolonne hvor relevant og kolonnetilvalg. Lange tabeller får lokal søgning og pagination/virtualisering efter datamængde.
6. Få farver med fast betydning: blå for valg/tidligere data, grønblå for aktuelle værdier, rav for forbehold, rød for tab/overskridelse. Fortegn og tekst bærer betydningen også uden farve. Ingen dekorative kort omkring hele sider.
7. Grafvalg følger opgaven: tidsserie for udvikling, stablede søjler for kostdele, divergerende søjler for ændringer, ressource/dag-matrix for belastning. Ingen cirkeldiagrammer til mange kunder og ingen graf til almindelige adgangs- eller vareformularer.
8. Samme tastaturbetjening, datatilstande, diskrete bevægelser og reduceret-motion-støtte. Værktøjsknapper bruger ikoner med navn/tooltip. Filtre er lokale, når grunddata allerede er indlæst; andre hentninger er debounced og cachede.

## Næste moduler

### 1. Belastning

- [ ] Samlet ressource/dag-matrix med timer og kapacitet; klik på en celle bruger eksisterende dag-/ordredetalje. Genbrug `renderBelastningBars`, `buildBelastningClusterSvg` og `renderBelastningDetailTable` i `server.js`.
- [ ] Tydelig forskel på samlet periodebelastning og enkelte overbelastede dage. Bevar aften separat; vis nul kapacitet eksplicit.
- [ ] Erstat kortenes store rammer med sammenhængende produktionsoversigt; fasthold kunde/ordre-filter ved skift fra Home.
- [ ] Separat datakontrol før nye tidsafhængige KPI'er: SQL bruger inklusiv `@Dage + 1`; tidligere live-data havde kapacitet før valgt startdato. Afklar datogrænse/tidszone uden at ændre periodedefinition i et designpatch.

### 2. SalgOrdre VIA

- [ ] Stablede kostsøjler pr. ressource eller kunde over ordretabellen; klik filtrerer de allerede indlæste rækker. Brug `getSalgordreViaVisibleRows` og `renderSalgordreVia` i `assets/js/via.js`.
- [ ] Vis Materiale, Stang, Indkøbte dele og Tid én gang. Samlet kost er ikke en femte stablet komponent; salgsværdi har eget sammenligningsfelt.
- [ ] Flyt købte dele til en eksplicit detaljevisning frem for mange indlejrede tabeller. Bevar Bestilt/Modtaget/Forbrugt/Medregnet og de verificerede reservationer.
- [ ] Synlige kolonner, leveringsstatus og progression kan gemmes som personlige visninger; eksisterende kolonnebredder bevares.

### 3. Omsætning og Ordreindgang

- [ ] Knyt eksisterende måneds-/kontosøjler og budgettærskler til samme månedspanel som Ordreflow. Undgå flere gentagelser af samme månedstotal. Brug `renderOmsaetningCharts` og `renderOmsaetningMonthDetail`.
- [ ] Kundesammenligning som sorterede bjælker med beløb og ændring, derefter kundens fakturaordrer. Kreditposteringer skal fortsat indgå i bogført omsætning.
- [ ] Ordreindgang: klik på uge/søjle filtrerer uge- og kundetabel; behold særskilte serier for ordreværdi, tilbud, budget og MA3. Brug `renderOrdreindgangTrendChart`, `renderOrdreindgangWeeklyTable`, `renderOrdreindgangCustomersTable`.
- [ ] Kontrollér visning af negative ugebeløb: eksisterende grafens `toY` klamper værdier til nul. En akseændring kræver en negativ fixture, ikke ændring af ordrebeløbet.
- [ ] Hold månedens nye ordrer, tidligere ordrebeholdning og bogførte fakturaer adskilt. Ordreflow bruger alle omsætningskonti, øvrige omsætningsgrafer følger kontofilteret.

### 4. Efterkalkulation

- [ ] Kompakt ordreheader med salg, effektiv kost, DB DKK, DB % og datadækning; kostfordeling og tabel kan vælges i samme arbejdsflade.
- [ ] Klik på en kostdel åbner relevante salgs-/produktionslinjer. Laser-, operations- og ordretotaler må ikke stables sammen, hvis de overlapper. Genbrug ejerens effektive kost inkl. stykliste og kosteksklusioner.
- [ ] Kundens fakturaforløb og ordresammenligning genbruger fulde autoriserede periodedata, ikke listens 150 rækker.
- [ ] Rapporter/print beholder fulde tabeller og samme totaler; skærmens widgethøjde må ikke klippe udskriften.

### 5. Lagerliste og Lagerliste 2

- [ ] Oversigt med kategoriandele og ændring fra valgt snapshot; klik åbner eksisterende kategoritabel i `assets/js/lagerliste.js`.
- [ ] Stablede værdier bruger kun gensidigt udelukkende poster. `Varelager`, `Vare i arbejde` og `TOTAL` er summer, ikke ekstra komponenter.
- [ ] FIFO/standardpris og aktuel/snapshot er eksplicitte valg. Modtagne reserverede købte dele flyttes Lager til VIA én gang; reservationer er separat information.
- [ ] Lagerliste 2: vis kun dokumenterede bevægelser som flow. Uforklarede differencer forbliver kontrolposter; ingen konstrueret balance for grafens skyld.
- [ ] Kontrolvisningen forbliver Beta, og officiel Lagerliste 1/udskrift bevares.

### 6. Ordreoversigt, BOM og øvrige moduler

- [ ] Ordreoversigt: produktionsrækkefølge og afvigelser i en kompakt statuslinje; vare-/rutelister forbliver primære, især i print.
- [ ] BOM: brug eksisterende `renderCostBreakdown` i `assets/bom/bom-beregner.js` som indgang til materiale/operationer. Prismatrix får sammenlignelig pris/antal-visning med samme motor, opstartfordeling og enheder.
- [ ] Administration: bevar faner, labels og bekræftet lagring; komprimer formularer og gør status tydelig. Ingen dekorative analysegrafer for brugerrettigheder.
- [ ] QMS/Personalehåndbog: forbedr søgning, dokumenthierarki og læseflade, ikke økonomiske widgets. Versions- og kildeoplysninger forbliver synlige.
- [ ] Planlagte fakturafunktioner er ikke en del af den nuværende aktive UI-omlægning.

## Verifikation pr. modul

- [ ] Fast fixture og stikprøve fra virkelig cache: graf = filtreret KPI = tabel = CSV; korrektioner, tomme data og manglende kost medtages i testen.
- [ ] Ingen rettigheds-/databaseoverførsel i cache eller personlige visninger; ingen dataadgang ved rent visuelle ændringer.
- [ ] Desktopkontrol ved mindst 1280 og 1920 px: ingen overlap, korrekt intern rulning og navigation tilbage. Fokus og valg bevares under opdateringer.
- [ ] Print/PDF kontrolleres separat, før gamle visninger erstattes.
- [ ] Kør hele Node-testsuiten i et miljø med terminaladgang. Denne iteration er valideret med browserkontroller og live Belastning-data; `npm test` er endnu ikke kørt her.