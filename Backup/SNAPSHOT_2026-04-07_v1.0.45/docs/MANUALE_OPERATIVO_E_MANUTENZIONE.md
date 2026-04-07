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

### 4.4 Dettaglio ordine
Aprendo un ordine si vedono tipicamente:

1. **testata ordine**
2. **righe ordine vendita**
3. **ordini di produzione collegati**
4. **operazioni** con subtotali `Delsum`
5. **summary** finale con ricavo, costo totale e margine

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

---

## 5. Regole business importanti

Queste regole sono già implementate e **non vanno cambiate senza validazione funzionale**.

### 5.1 `R1090`
- escluso globalmente dai costi e dalle operazioni rilevanti
- motivo: il suo costo è già distribuito altrove
- non deve falsare subtotali e totale costi

### 5.2 `R6200`
- per alcune operazioni il costo effettivo usa `NoOrg * CCstPr`
- utile dove i minuti pianificati sono il riferimento corretto

### 5.3 `R1100` + `LASER EAGLE`
In `utils/productRules.js`:
- se `ProdNo = R1100`
- e `ProdTp4 = 1`
- e operatore contiene `LASER EAGLE`

allora costo/prezzo operativo viene raddoppiato.

### 5.4 Logica ricorsiva ordini figli
`services/aftercalcService.js` contiene la funzione ricorsiva `loadProductionOrderDetails(prodOrdNo, visited = new Set())`.

Questa logica serve per:
- seguire ordini di produzione collegati
- evitare loop con `visited`
- riportare il costo del figlio sul padre

> Non rimuovere o “semplificare” questa parte senza testare i casi reali di produzione.

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