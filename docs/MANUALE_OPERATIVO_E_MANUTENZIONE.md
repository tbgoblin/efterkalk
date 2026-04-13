# Manuale operativo e di manutenzione — Gantech Efterkalk

## 1. Scopo dell’app

`Gantech Efterkalk` è una app desktop Windows basata su **Electron + Express + SQL Server** usata per:

- vedere gli ultimi ordini fatturati
- analizzare costi, ricavi e margini
- esplodere gli ordini di produzione collegati
- controllare operazioni, nesting e materiale laser
- aprire rapidamente i disegni PDF (`Vis tegning`)

L’interfaccia utente usa testi prevalentemente **danesi**, adatti al contesto operativo di fabbrica.

---

## 2. Requisiti

### Sistema
- Windows
- Node.js installato per sviluppo/build
- accesso al database SQL Server aziendale
- permessi per leggere eventuali cartelle rete/UNC dei disegni

### Connessione database
Configurata in `db.js`:

- **server:** `10.2.0.3\\VISMA`
- **database:** `F0001`
- driver: `msnodesqlv8`
- autenticazione: `trustedConnection: true`

> Se la macchina non vede il server SQL o non ha i driver nativi corretti, l’app non caricherà i dati.

---

## 3. Avvio dell’app

### Modalità sviluppo desktop

```bash
npm install
npm run desktop
```

Questo:
1. avvia il server locale
2. apre la finestra Electron
3. carica la UI da `http://localhost:<porta>`

### Modalità server

```bash
npm start
```

### Porta usata
- il server usa `process.env.PORT` se impostata
- altrimenti `electron-main.js` calcola una porta per sessione utente/RDS nel range `3000-3999`
- fallback standard server: `3000`

---

## 4. Uso quotidiano

### 4.1 Accesso iniziale
All’apertura compare una finestra `Adgangskode`.

**Codice attuale:** `12345`

⚠️ Questo controllo è **solo lato interfaccia** e non sostituisce una sicurezza reale server-side.

### 4.2 Barra principale
Nella parte alta sono disponibili:

- `Søg` → apre il dettaglio ordine per numero
- `Opdater...` → azioni rapide su cache/lista/program
- `Skift marginberegning` → cambia formula di visualizzazione del margine
- `Skjul kundeliste` / `Vis kundeliste`
- `Ryd cache` → cancella la cache persistente
- filtro `Alle brugere`
- campo `Søg kunde i listen...`

### 4.3 Lista ordini
La lista mostra gli ultimi ordini fatturati:

- finestra temporale: **30 giorni**
- massimo righe caricate: **150**
- colonne tipiche: `Bruger`, `Ordrenr.`, `Kunde`, data fattura, importo, margine, refresh

La lista:
- può essere ordinata
- si aggiorna automaticamente
- usa cache locale per partire più velocemente
- per i `MultiOrdre` (`Ord.Gr4 = 3`) mostra un badge tondo `M` con tooltip `MultiOrdre`

### 4.4 Dettaglio ordine
Aprendo un ordine si vedono tipicamente:

1. **testata ordine**
2. **righe ordine vendita**
3. **ordini di produzione collegati**
4. **operazioni** con subtotali `Delsum`
5. **summary** finale con ricavo, costo totale e margine

Note operative:
- le righe `Ydelse` / righe con `PurcNo` collegato possono essere aperte per vedere l’ordine di produzione figlio
- nei `MultiOrdre` compare anche la logica dedicata `NestMultiPris` nelle viste laser

### 4.5 Disegni e immagini
Se per il prodotto esiste un disegno, appare il pulsante `Vis tegning`.

Comportamento:
- l’app prova prima ad aprire il file tramite backend `POST /open-drawing`
- se non riesce, tenta apertura tramite URL/path lato client
- supporta percorsi locali, UNC e URL HTTP/HTTPS

### 4.6 Laser / nesting
Sono presenti viste e metriche dedicate al laser:

- endpoint `GET /laser-route-metrics`
- riepilogo materiale/lastre/sfrido
- dettaglio nesting per prodotto tramite `GET /nesting-detail/:ordno/:prodno`
- per i `MultiOrdre` (`Ord.Gr4 = 3`) la colonna speciale si chiama `NestMultiPris`
- nei `MultiOrdre`, il costo laser viene calcolato come **`kg forbrugt × media CstPr delle righe TrTp=5 della route`**
- il calcolo avviene **prima per singola `rute`**, poi i risultati vengono sommati su tutti i `nestingordre` collegati
- se lo stesso prodotto è distribuito su più `nestingordre`, il riepilogo li aggrega tutti
- anche negli ordini standard il popup laser può mostrare più righe (`nestingordre` / `rute`) per lo stesso prodotto

### 4.7 Interpretazione dei costi laser
Per evitare ambiguità durante i controlli:

- `NestKost pr. stk` nel popup laser è il costo unitario della **riga/route mostrata**
- `Samlet kost` nel popup è il costo totale della singola riga aggregata (`qta × costo unitario` oppure `QuotaCosto`)
- la riga principale `Materiale Laser` può mostrare un prezzo unitario diverso dal popup anche quando il **totale** è lo stesso, perché la divisione può avvenire su una quantità diversa (solo la riga madre vs somma di più `nestingordre`)
- `Ryd cache` forza il refresh dei dati memorizzati, ma non cambia le differenze dovute alla formula o alla quantità usata nel riparto

---

## 5. Regole di calcolo complete (fonti, formule, manipolazioni)

Questa sezione descrive **come viene calcolato ogni importo visualizzato**, da quali campi DB arriva e in quali casi il valore viene **modificato / ricalcolato / sostituito**.

> Tutte le formule sotto sono allineate alla logica attuale in `services/aftercalcService.js`, `server.js` e `utils/productRules.js`.

### 5.1 Campi sorgente usati dal sistema

| Campo | Provenienza | Significato operativo | Uso nel calcolo |
|---|---|---|---|
| `Ord.InvoAm` | testata ordine | totale fatturato ordine vendita | base del ricavo totale |
| `Ord.Gr4` | testata ordine | tipo ordine (`MultiOrdre` ecc.) | cambia la logica laser |
| `OrdLn.NoFin` | riga ordine | quantità / minuti dichiarati come finiti | base standard per quantità e costi |
| `OrdLn.NoOrg` | riga ordine | quantità originale / `Stykliste Minutter` | fallback quando `NoFin = 0` |
| `OrdLn.NoInvo` | riga ordine | quantità fatturata / fatturabile | base prioritaria per `Ydelse` |
| `OrdLn.NoInvoAb` | riga ordine | quantità acquistata/fatturata lato acquisto | usata per warning di fattura mancante |
| `OrdLn.DPrice` | riga ordine | prezzo unitario della riga | usato come prezzo vendita oppure riferimento esterno |
| `OrdLn.CCstPr` | riga ordine | costo unitario standard | base per `LineCost` e molti fallback |
| `OrdLn.PurcNo` | riga ordine | ordine figlio collegato | collega ordini vendita/produzione |
| `OrdLn.ProdTp4` | riga ordine | gruppo logico (`1`, `2`, `4`, `6`...) | decide la formula da usare |
| `OrdLn.TrInf2` / `TrInf4` | riga ordine | riferimenti ordine/ruta | usati soprattutto nel laser |
| `OrdLn.CstPr` | righe nesting/laser | costo materia sulla route | usato per media `CstPr` nel laser |
| `OrdLn.Free3` | righe nesting/laser | peso storico/unitario | usato per stimare kg attesi nel laser |
| `Struct.NoPerStr` | distinta base | peso atteso per struttura | supporto al calcolo kg laser |

### 5.2 Grandezze derivate interne

L’app costruisce e visualizza alcune grandezze derivate, non sempre presenti direttamente nel DB:

- `LineCost` = **`NoFin × CCstPr`**
  - è il costo “grezzo” di partenza della riga
- `EffectiveLineCost`
  - è il **vero costo usato in UI e totali** dopo tutte le regole speciali
- `DisplayQuantity`
  - è la quantità/minuti mostrata in tabella dopo eventuali fallback (`NoOrg`, `NoInvo`, child order ecc.)
- `DisplayUnitCost`
  - quando possibile = **`EffectiveLineCost / DisplayQuantity`**
  - viene usato per mostrare un costo unitario coerente col totale

### 5.3 Esclusioni e filtri globali

Queste regole vengono applicate **prima** dei totali:

- i prodotti globalmente esclusi via `isGloballyExcludedProdNo(...)` non entrano nei calcoli
- nelle **operazioni** (`ProdTp4 = 1`) i codici **`R1090`** e **`R8200`** sono esclusi da costo e visualizzazione
- in **`Produkt dele`** (`ProdTp4 = 4`) tutti i prodotti che iniziano per **`R`** sono nascosti e non conteggiati
- la stessa esclusione `R*` vale anche ricorsivamente nei sottoordini
- nelle somme ordine di produzione, la riga **`LnNo = 1`** è trattata come riga principale del prodotto e **non entra nei subtotali di gruppo**

### 5.4 `Salgsordrelinjer` (righe ordine vendita)

Per ogni riga ordine vendita la UI mostra:

| Colonna UI | Formula / sorgente | Note |
|---|---|---|
| `Færdigmeldt` | `DisplayQuantity` | normalmente `NoFin`; se `NoFin = 0` e `NoOrg > 0`, mostra `NoOrg` |
| `Kostpris` | se c’è `PurcNo`: `ProductionOrderTotalCost / NoFin`; altrimenti `CCstPr` | nelle righe collegate a produzione mostra il costo unitario del figlio |
| `Samlet kost` | `EffectiveLineCost` | è il costo effettivo finale |
| `Salgspris/enhed` | `DPrice` | prezzo vendita unitario |
| `Salgspris` | `DPrice × NoFin` | totale vendita della riga |
| `Margin (%)` | dipende dalla modalità margine scelta in UI | vedi § 5.11 |
| `Prod.ordre` | `PurcNo` | se presente, apre l’ordine di produzione collegato |

Regole aggiuntive sulle righe vendita:

1. **Riga sconto / riga a zero**
   - se `DPrice × NoFin = 0` **e** non esiste fallback materiale/tubo, la riga viene trattata come `IsDiscountLine = true`
   - in questo caso `EffectiveLineCost = 0`

2. **Sostituzione costo con ordine di produzione**
   - se la riga ha `PurcNo` valorizzato e non è una riga sconto, il costo mostrato **non è più il costo grezzo della riga vendita**
   - viene sostituito da:

```text
ProductionOrderTotalCost = totalCost dell’ordine di produzione figlio
EffectiveLineCost = ProductionOrderTotalCost
```

3. **Fallback tubo/materiale incoerente**
   - se il prodotto inizia per `3`, `NoFin = 0` e `NoOrg > 0`, il costo può essere ricalcolato con:

```text
NoOrg × CCstPr
```

### 5.5 Ordini di produzione e `Delsum`

Le righe dell’ordine di produzione vengono raggruppate per `ProdTp4`:

| Chiave | Significato UI |
|---|---|
| `1` | Operation |
| `2` | Materiale Laser |
| `4` | Produkt dele |
| `5` | Rute |
| `6` | Ydelse |
| `7` | Underleverandør |
| `8` | Materiale fast antal |
| `NA` | non classificato |

Regole di aggregazione:

- le righe `ProdTp4 = 3` vengono **accorpate al gruppo `1 - Operation`**
- le righe `LnNo = 1`, `ProdTp4 = 0`, `3`, `5` **non entrano** nei subtotali gruppo
- `Delsum` è la somma dei `EffectiveLineCost` delle righe visibili del gruppo
- `Total ordre` è la somma di tutti i `Delsum` visibili del blocco produzione

### 5.6 `1 - Operation` (operazioni)

Formula base:

```text
EffectiveLineCost = EffectiveOperationMinutes × CCstPr
```

Dove:

- `Stykliste Minutter` mostrato in UI = `NoOrg`
- `Færdigmeldt minutter` mostrato in UI = `EffectiveOperationMinutes`

Regole speciali:

1. **Fallback minuti con icona `🕒`**
   - se una operazione `R*` ha `NoFin = 0` ma `NoOrg > 0`, il sistema usa `NoOrg` come minuti effettivi
   - la UI mostra l’icona `🕒`

2. **Esclusioni**
   - `R1090` e `R8200` non vengono conteggiati

3. **`R6200`**
   - nei subtotali operazioni viene trattato come:

```text
NoOrg × CCstPr
```

4. **`R1100` + `LASER EAGLE`**
   - se `ProdNo = R1100`, `ProdTp4 = 1` e l’operatore contiene `LASER EAGLE`, il costo operativo viene raddoppiato da `adjustOperationLinePricing(...)`

### 5.7 `2 - Materiale Laser`

Per i prodotti laser (`ProdNo` che termina con `L`) il costo non viene letto solo da `CCstPr`, ma può essere ricalcolato per route.

#### Formula laser specializzata
Per ogni `route`:

```text
Costo unitario laser = kg forbrugt per pezzo × media CstPr delle righe TrTp = 5 della stessa route
```

Poi:

```text
EffectiveLineCost = costo unitario laser × quantità finita
```

Dettagli importanti:

- `kg forbrugt` è ricostruito dai dati `TrTp = 5/7`, `Free3` e, quando serve, `Struct.NoPerStr`
- se il prodotto compare su più `nestingordre` / `route`, il sistema aggrega i costi
- per `MultiOrdre` (`Ord.Gr4 = 3`) la colonna viene etichettata `NestMultiPris`
- per ordini standard la colonna resta `Kostpris nesting`

Se non esiste un costo laser specializzato valido, il fallback è:

- `NestingCost × NoFin`, se `NestingCost > 0`
- altrimenti `LineCost`

#### Incoerenza materiale/tubo
Se `ProdTp4 = 2`, il prodotto inizia per `3`, `NoFin = 0` e `NoOrg > 0`, il costo viene ricalcolato come:

```text
NoOrg × CCstPr
```

con warning di incoerenza.

### 5.8 `4 - Produkt dele`

Questo gruppo rappresenta componenti / sottoordini.

Regole:

- tutti i `R*` vengono esclusi
- se la riga ha `PurcNo`, l’app apre ricorsivamente l’ordine figlio e usa il suo totale:

```text
EffectiveLineCost = childSummary.totalCost
```

- gli eventuali warning del figlio vengono propagati al padre

### 5.9 `6 - Ydelse` (lavorazioni esterne / ordine di acquisto)

Questa è la regola più importante da preservare.

**Interpretazione funzionale:** `Ydelse` non è una vendita; rappresenta una **lavorazione esterna / acquisto esterno sul prodotto**.

Per questo motivo in UI:

- la colonna dedicata è `Ydelse pris/enhed`
- non vanno mostrate colonne `Kostpris/enhed` o `Nesting/enhed` nel popup filtrato `Ydelse`
- il popup deve mostrare **solo il prodotto cliccato**

#### Quantità usata per il costo `Ydelse`
La quantità effettiva è:

```text
NoInvo, se NoInvo > 0
altrimenti NoFin
```

Se `NoInvo = 0` e si usa `NoFin`, la UI mostra warning `🧾`.

#### Sorgente autoritativa del costo unitario `Ydelse`
Se la riga `Ydelse` ha un `PurcNo` verso un child order, la sorgente corretta del costo è la **riga figlia corrispondente nel child order**, non il valore grezzo del parent.

Formula attuale:

```text
matchedChildLine = riga del child order con stesso ProdNo
matchedChildUnitCost = matchedChildLine.EffectiveLineCost / matchedChildLine.DisplayQuantity
EffectiveLineCost = effectiveQuantity × matchedChildUnitCost
DisplayUnitCost = EffectiveLineCost / DisplayQuantity
```

Fallback se il child non fornisce tutto:

- prima `matchedChildLine.CCstPr`
- poi `matchedChildLine.DPrice`
- poi `matchedChildLine.DisplayUnitCost`
- infine `line.CCstPr`

Questo è il motivo per cui il valore corretto di `Ydelse pris/enhed` può essere diverso dal `DPrice` grezzo della riga padre.

### 5.10 `7 - Underleverandør`, `8 - Materiale fast antal` e altri gruppi

Se non entra una regola speciale (`Operation`, `Laser`, `Produkt dele`, `Ydelse`), il calcolo standard è:

```text
EffectiveLineCost = LineCost
DisplayUnitCost = EffectiveLineCost / DisplayQuantity   (se la quantità > 0)
altrimenti fallback a CCstPr
```

### 5.11 Warning, icone e testo mostrato in UI

Le icone visualizzate nella UI hanno il seguente significato:

| Icona | Significato |
|---|---|
| `🕒` | `Færdigmeldt` era 0 e il sistema ha usato `Stykliste Minutter / NoOrg` |
| `🧾` | fattura mancante / `NoInvo = 0`, quindi è stato usato `NoFin` |
| `⚠️` | incoerenza generica (es. materiale/rør con `NoFin = 0` ma `NoOrg > 0`) |
| `🏭` | warning proveniente da ordine di produzione collegato (solo se esposto in UI) |

Il testo tooltip non è generico: viene costruito da `WarningText` / `warningText` e descrive il motivo reale.

### 5.12 Totali finali ordine e margine

#### Costo totale ordine
Il totale finale mostrato nella testata ordine viene calcolato così:

```text
salesNoPOTotalCost = somma EffectiveLineCost delle righe vendita senza PurcNo
productionTotalCost = somma totalCost degli ordini di produzione collegati a righe vendita non-sconto
totalCost = salesNoPOTotalCost + productionTotalCost
```

#### Ricavo totale ordine

```text
totalRevenue = Ord.InvoAm
```

#### Margine in DKK

```text
margin = totalRevenue - totalCost
```

#### Percentuale margine
La UI supporta due modalità:

1. **Klassisk**
```text
((Salg - Kost) / Salg) × 100
```

2. **Ny**
```text
(Salg / Kost) × 100
```

La stessa logica viene usata sia per il margine ordine sia per il badge margine sulle singole righe.

### 5.13 Regola di manutenzione documentale

Ogni volta che si modifica una formula in:

- `services/aftercalcService.js`
- `server.js`
- `utils/productRules.js`

deve essere aggiornato anche questo capitolo, specificando:

- **campo sorgente**
- **formula finale**
- **eventuale fallback / manipolazione**
- **icona warning associata**

---

## 6. Architettura tecnica

### 6.1 Componenti principali

| File/modulo | Responsabilità |
|---|---|
| `electron-main.js` | finestra desktop, auto-start, updater, compatibilità RDS |
| `server.js` | bootstrap Express, HTML UI, warmup/cache, orchestrazione |
| `routes/apiRoutes.js` | API backend principali |
| `services/aftercalcService.js` | calcoli ordine, produzione, margini |
| `services/drawingService.js` | ricerca PDF/disegni/immagini |
| `utils/productRules.js` | regole dedicate ai prodotti |
| `utils/logger.js` | log su file + console |
| `diskCache.js` | cache persistente JSON su disco |
| `db.js` | connessione MSSQL |

### 6.2 Flusso semplificato
1. Electron avvia il server locale.
2. `server.js` carica la lista ordini dalla cache o dal DB.
3. Parte il warmup in background di margini e aftercalc.
4. La UI interroga le API Express per lista, dettaglio, summary e laser.
5. I risultati vengono memorizzati su disco per gli avvii successivi.

### 6.3 Endpoint principali

| Endpoint | Scopo |
|---|---|
| `GET /health` | check rapido stato server |
| `GET /order-list` | elenco ordini recenti |
| `GET /order-list-check-time` | verifica se la lista va aggiornata |
| `GET /aftercalc/:ordno` | dettaglio completo aftercalc ordine |
| `GET /order-margin/:ordno` | costo/ricavo per badge margine |
| `GET /production-summary/:ordno` | riepilogo ordine di produzione |
| `GET /laser-route-metrics` | metriche laser/nesting |
| `GET /nesting-detail/:ordno/:prodno` | dettaglio nesting per prodotto |
| `POST /cache-refresh-order/:ordno` | refresh cache singolo ordine |
| `GET /cache-refresh-order-status/:ordno` | stato refresh ordine |
| `POST /cache-clear` | svuota cache persistente |
| `GET /cache-status` | elenco elementi in cache |
| `GET /warmup-status` | stato warmup iniziale |
| `POST /open-drawing` | apertura disegno PDF |
| `POST /desktop-update-check` | avvia controllo aggiornamenti desktop |

---

## 7. Cache e performance

### Cache utilizzate
- **order list cache** in memoria
- **margin cache** in memoria
- **aftercalc cache** persistente su file JSON
- **production summary cache** persistente
- **laser metrics cache** persistente

### TTL attuali
Da `server.js`:

- `aftercalc`: **30 min**
- `production summary`: **30 min**
- `order margin`: **30 min**
- `laser metrics`: **60 min**
- `order list cache`: **10 min**

### Posizione cache
`diskCache.js` cerca in ordine:

1. `GANTECH_CACHE_DIR`
2. `C:\GantechCache`
3. `C:\cache\Gantech`
4. `%LOCALAPPDATA%\Gantech Efterkalk\cache`
5. `%APPDATA%\Gantech Efterkalk\cache`
6. `./cache`
7. cartella temporanea di sistema

### Quando usare `Ryd cache`
Usarlo solo se:
- i dati sembrano incoerenti o bloccati
- la lista non si aggiorna
- dopo modifiche importanti o test di manutenzione

⚠️ Dopo la pulizia, il ricaricamento può essere lento finché il warmup non ricostruisce la cache.

---

## 8. Logging

I log vengono scritti in `gantech.log`.

Percorsi tipici:
- `GANTECH_LOG_DIR` se definita
- `%LOCALAPPDATA%\Gantech Efterkalk\gantech.log`
- `%APPDATA%\Gantech Efterkalk\gantech.log`
- cartella progetto

In modalità desktop, `electron-main.js` prova anche percorsi condivisi come:
- `C:\GantechCache`
- `C:\cache\Gantech`

Controllare il log per:
- errori SQL
- problemi di warmup/cache
- errori apertura PDF
- problemi updater
- startup e porta usata

---

## 9. Ambiente RDS / desktop condiviso

L’app contiene accorgimenti specifici:

- GPU disabilitata per maggiore compatibilità
- sandbox disattivata in alcuni casi RDS
- auto-start saltato in ambiente RDS condiviso
- porta locale calcolata usando utente/sessione/client per ridurre collisioni
- `electron-main.js` usa `app.requestSingleInstanceLock()` per impedire aperture duplicate della stessa app

Questo è importante se più utenti aprono l’app sullo stesso host.

---

## 10. Build, deploy e release

### 10.1 Build installer

```bash
npm run build:win
```

Genera:
- `dist/Gantech-Efterkalk-Setup-<version>.exe`
- `dist/latest.yml`
- blockmap per auto-update

### 10.2 Avvio desktop

```bash
npm run desktop
```

### 10.3 Dipendenze chiave
Da `package.json`:
- `express`
- `mssql`
- `msnodesqlv8`
- `electron-updater`
- `electron`
- `electron-builder`

### 10.4 Publish completo

```powershell
.\publish.ps1
```

Lo script esegue:
1. controllo stato git
2. commit delle modifiche
3. `git push`
4. `npm version patch`
5. `git push --follow-tags`
6. `npm run build:win`
7. publish release GitHub via `release-github.ps1`

### Prerequisiti per publish
- `gh` (GitHub CLI) installato
- `gh auth login` già eseguito
- spazio libero sufficiente su disco `C:`

> Se NSIS fallisce con errori di scrittura, verificare subito lo spazio disco e pulire `dist/`.

### Auto-update
Configurato con provider GitHub:
- owner: `tbgoblin`
- repo: `efterkalk`

All’avvio desktop, l’app controlla se esiste una release più nuova e notifica l’utente quando è pronta.

---

## 11. Manutenzione ordinaria

### Checklist consigliata

#### Giornaliera / al bisogno
- verificare che la lista ordini si apra correttamente
- controllare che `Vis tegning` funzioni sui prodotti principali
- usare `Ryd cache` solo se necessario

#### Dopo modifiche codice
- controllare `GET /health`
- provare `GET /order-list`
- testare almeno un `GET /aftercalc/<ordNo reale>`
- testare almeno un `GET /production-summary/<prodOrdNo reale>`

#### Prima di una release
- confermare build `npm run build:win`
- verificare presenza file in `dist/`
- verificare release GitHub pubblicata
- verificare download/update su una macchina di test

---

## 12. Troubleshooting rapido

### Problema: la lista ordini non compare
Controllare:
1. server locale attivo
2. `GET /health` risponde `200`
3. connessione SQL disponibile
4. file `gantech.log`
5. eventuale pulizia cache

### Problema: il disegno PDF non si apre
Controllare:
1. valore `WebPg` / `PictFNm`
2. accesso a cartelle rete/UNC
3. esistenza del PDF sul path risolto
4. permessi dell’utente Windows/RDS

### Problema: build Windows fallisce
Cause comuni:
- poco spazio su disco `C:`
- artefatti vecchi in `dist/`
- dipendenze native SQL non allineate

Azioni:
- pulire `dist/`
- rilanciare `npm run postinstall`
- rilanciare la build

### Problema: dati lenti o startup lento
- attendere fine warmup iniziale
- controllare cartella cache
- verificare accesso al DB
- controllare se il log segnala errori di query o timeout

---

## 13. Note per sviluppatori

- La refactor attuale ha separato logger, product rules, drawing service, aftercalc service e API routes.
- `server.js` resta il composition root e contiene ancora la UI HTML inline.
- `views/indexPage.js` risulta al momento **non integrato** e va ignorato finché non viene completato correttamente.
- In questo progetto è importante fare **refactor strutturali senza cambiare la logica**.

### Regola pratica
Prima di dichiarare conclusa una modifica, verificare sempre con chiamate reali o build reali, non solo con supposizioni.

---

## 14. File utili già presenti

- `DESKTOP_DEPLOY.md` → note rapide di deployment Windows
- `AUTO_UPDATE_SETUP.md` → configurazione aggiornamenti automatici
- `publish.ps1` → rilascio automatizzato
- `release-github.ps1` → pubblicazione asset su GitHub Releases

---

## 15. Contatti / ownership

Autore indicato nel progetto: **Gantech**.

Per modifiche business-critical, validare sempre con chi conosce i flussi reali di produzione e costing.