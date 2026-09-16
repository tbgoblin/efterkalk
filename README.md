# Gantech Efterkalk

App desktop per **efterkalkulation** e analisi margini ordini, pensata per uso interno in ambiente produzione/fabbrica.

Il prodotto è cresciuto fino a diventare un **hub operativo interno**: oltre al costing degli ordini riunisce dati commerciali e produttivi provenienti da Visma/SQL Server e comprende aree dedicate a BOM/preventivazione, magazzino, QMS, carico produttivo, fatturato e VIA.

**Versione attuale:** `1.1.68` (fonte: `package.json`)

---

## ✨ Funzioni principali

- elenco rapido degli ultimi ordini fatturati
- ricerca per `OrdNo`
- calcolo costi/ricavi/margine per ordine
- dettaglio ordini di produzione collegati
- apertura **tegning/PDF** con pulsante `Vis tegning`
- **Lagerliste 1**: valore fisico, chiusure mensili, snapshot giornalieri, confronto periodi, PDF, ricerca articolo e prenotazioni verificate
- **SalgOrdre VIA**: materiale, stang, tempo e componenti acquistati; distingue quantità ordinate, ricevute e consumate
- **Lagerliste 2 (Beta/Shadow)** separata: riconcilia i movimenti tra periodi, distingue REST previsto/registrato/svalutazione e usa `NoPac` per evitare doppioni VIA/Færdige senza cambiare Lagerliste 1
- cache locale + warmup automatico per velocizzare l’avvio
- pacchetto desktop Windows con aggiornamento automatico via GitHub Releases

### Aggiornamenti verificati (`2026-04-07`)

- protezione **single-instance** in `electron-main.js` per evitare avvii duplicati della desktop app
- righe `Ydelse` / `PurcNo` cliccabili verso l’ordine di produzione figlio
- colonna `Salgspris/enhed` aggiunta nelle righe ordine vendita
- supporto `MultiOrdre` (`Ord.Gr4 = 3`) con badge `M` e tooltip `MultiOrdre`
- per i `MultiOrdre`, colonna **`NestMultiPris`** visibile solo in questi ordini
- logica `MultiOrdre` verificata: costo laser basato su **`kg forbrugt × media CstPr delle righe TrTp=5`**, calcolato **per singola `rute`** e poi aggregato su tutti i `nestingordre` collegati
- il popup laser può aggregare più `nestingordre`/`rute` dello stesso prodotto: per questo il prezzo unitario mostrato nel popup può differire da quello della riga principale se il medesimo totale viene ripartito su quantità diverse
- `R8200` è escluso dai costi/righe operazione; se una operazione `R*` ha `Færdigmeldt = 0`, l’app usa `Stykliste Minutter`, ricalcola i costi e mostra l’icona `🕒`
- i prodotti `R*` dentro `Produkt dele` (anche nei sottoordini) non devono essere mostrati né conteggiati

### Aggiornamenti recenti (`2026-04-22`)

- startup/warmup rivisto: la schermata rossa `loading.html` resta attiva finché il backend non segnala `ready=true` (aftercalc + margin warmup completati)
- endpoint `/warmup-status` esteso con `marginDone`, `marginTotal`, `combinedDone`, `combinedTotal`, `combinedPct`, `ready`
- rimosso il fallback timeout che bypassava il gate startup; l’ingresso avviene solo a warmup completo
- eliminato il prefetch su `mouseover` nella lista ordini per ridurre query inutili e carico DB
- logging cache aftercalc migliorato: eventi espliciti `AFTERCALC CACHE HIT`, `AFTERCALC IN-FLIGHT REUSE`, `AFTERCALC FRESH COMPUTE`
- route `/aftercalc/:ordno` allineata al percorso unico `getOrComputeAftercalc(...)` con fallback cache coerenti
- prevenuta doppia esecuzione warmup startup (de-duplicazione processo in background)
- revenue ordine aggiornata a: `Ord.InvoAm + Ord.DInvoIF` (importo fatturato + da fatturare)
- nuova sezione UI **Operation Oversigt** con toggle dedicato, raggruppamento per `R-kode`, quantità/minuti/costi e riepilogo totale
- in `Laseroversigt`, il totale non include più la vecchia voce “Samlet Operation kost” (spostata in `Operation Oversigt`)
- stato fatturazione ordine aggiunto nel banner:
   - `I produktion` se `InvoAm = 0`
   - `Delvist faktureret` se `InvoAm > 0` e `DInvoIF > 0`
   - `Komplet faktureret` se `DInvoIF = 0`
- per ordini `I produktion`, il banner mostra `Kost til dato (estimat)` e prognosi coerente con importi previsti
- `Kost til dato` impostato come somma dei `totalCost` dei `productionOrders` collegati
- chiarita la semantica di `Gr4` come **tipo ordine** (es. Multiordre) con rinomina variabili/UI note, senza modificare la logica business
- fix allocazione laser nel fallback aggregato: se il nesting totale è registrato su quantità maggiori della singola riga (es. 200 vs 100), il costo viene ripartito proporzionalmente evitando raddoppi su singolo articolo
- mantenuta e documentata la nota di divergenza prezzo unitario quando il totale laser viene redistribuito su quantità diverse (`allocation spread`)

### Lagerliste e VIA (`2026-09-15`)

- `Opfølgningsvarer` mostra il valore fisico e il residuo Lager dopo l’eventuale trasferimento di componenti acquistati a VIA
- le prenotazioni `Rsv` sono collegate solo a salgsordrer esistenti e aperte, con lotto `ShpBal` verificato; la sezione informativa non viene mai sommata due volte
- `Indkøbte dele til ordre` entra nel VIA con il consumo `NoFin` oppure, prima del consumo, quando ricezione, giacenza fisica e prenotazione verificata provano l’allocazione; il corrispondente valore FIFO esce dal Lager
- VIA, Lagerliste e CSV includono lo stesso valore dei componenti acquistati e il relativo dettaglio espandibile
- cache Lagerliste `v36`, cache VIA `v35` separata per database/data/perimetro; schema di valutazione `35` invariato

### VIA: rettifica del perimetro (`2026-09-16`)

- Con accesso Omsætning, la lista usa gli ordini con residuo `closing > 0,01 DKK` del modello Ordreflow, non tutti i candidati mensili. Tabella, KPI filtrato e CSV espongono il residuo, non il lordo.
- Lagerliste conserva il filtro storico e le regole di costo/allocazione. Costi mancanti, storico non riconciliato e residui minimi esclusi sono espliciti; il refresh singolo verifica l'identità dell'ordine.
- Verificati nel browser filtri/CSV, risposta errata, metadati, costi sconosciuti anche nel dashboard e recupero da errore. Dopo riavvio: 170/170 ordini e residui coincidenti con Ordreflow, totale 5.329.032,91 DKK; il residuo escluso di 0,00811 DKK spiega il centesimo rispetto alla totalizzazione Ordreflow. Refresh singolo 413646 verificato senza modifiche agli altri ordini. Suite Node e confronto numerico di Lagerliste non eseguiti; nessun rilascio incluso.

---

## 🚀 Avvio rapido

### Sviluppo

```bash
npm install
npm run desktop
```

### Solo server

```bash
npm start
```

### Build installer Windows

```bash
npm run build:win
```

Output in `dist/`.

---

## 🧭 Uso operativo

1. Avvia l’app desktop.
2. Inserisci il codice di accesso UI quando richiesto.
3. Usa `Søg` per aprire un ordine specifico.
4. Oppure usa la lista ordini con filtri per `Kunde` / `Bruger`.
5. Apri il dettaglio e controlla:
   - righe vendita
   - ordini di produzione
   - operazioni
   - `Delsum`, costo totale, ricavo e margine
6. Se disponibile, usa `Vis tegning` per aprire il PDF del disegno.

Per la guida completa vedi:

- [`docs/MANUALE_OPERATIVO_E_MANUTENZIONE.md`](docs/MANUALE_OPERATIVO_E_MANUTENZIONE.md) — contiene ora anche il capitolo completo **"Regole di calcolo complete (fonti, formule, manipolazioni)"**
- [`DESKTOP_DEPLOY.md`](DESKTOP_DEPLOY.md)
- [`AUTO_UPDATE_SETUP.md`](AUTO_UPDATE_SETUP.md)

---

## 🏗️ Struttura progetto

| Percorso | Scopo |
|---|---|
| `server.js` | bootstrap Express, UI HTML, cache orchestration |
| `electron-main.js` | contenitore desktop Electron |
| `db.js` | connessione SQL Server |
| `diskCache.js` | cache persistente su file |
| `routes/apiRoutes.js` | endpoint API principali |
| `services/aftercalcService.js` | logica calcolo aftercalc e production summary |
| `services/lagerlisteService.js` | Lagerliste 1 e snapshot storici (fonte dei valori FIFO) |
| `services/lagerliste2Service.js` | lettura route/nesting per la vista Beta, senza scritture DB |
| `services/drawingService.js` | ricerca/apertura disegni e immagini |
| `utils/productRules.js` | regole business prodotto |
| `utils/logger.js` | logging applicativo |
| `publish.ps1` | build + release automatizzata |

### Flusso applicativo

```text
Electron
  → avvia il server Express locale
  → attende il warmup di cache e margini
  → carica l'interfaccia operativa
  → API Express → servizi di dominio → SQL Server / cache su disco
```

### Stato dei test

Il progetto dispone di una prima suite automatica eseguibile con:

```bash
npm test
```

I test sono isolati e non si collegano a Visma: verificano sessioni, protezione scritture, BOM `readOnly`, PDF, selezione/stampa periodi Lagerliste, prezzi FIFO, migrazione GOH, Diverse, prenotazioni, priorità dei collegamenti ordine, ripartizione `NoPac`, eliminazione dei doppioni VIA/Færdige e componenti acquistati calcolati solo da `NoFin`. Le query SQL richiedono comunque una verifica con dati aziendali reali.

Questa è una rete di sicurezza iniziale, non ancora una copertura completa. Prima di refactor importanti è consigliato introdurre test di caratterizzazione con casi anonimizzati e risultati attesi, in particolare per:

- esclusioni `R1090`, `R8200` e righe `R*` nei componenti;
- fallback `NoFin` → `NoOrg`;
- ordini e sottoordini ricorsivi;
- `MultiOrdre`, laser e nesting;
- ricavo, costo totale e margine;
- equivalenza tra risultato fresco e risultato recuperato dalla cache.

---

## ⚠️ Nota importante

Il login crea una sessione server-side in memoria valida per 8 ore, disponibile sia come bearer token compatibile con la UI esistente sia come cookie `HttpOnly` same-origin. Le scritture critiche BOM/QMS, la modifica dei profili database e l’apertura locale dei PDF richiedono una sessione autenticata; il logout revoca la sessione.

L’applicazione resta un servizio interno: non tutti gli endpoint di lettura hanno autorizzazioni granulari e la sessione viene persa al riavvio del processo.

---

## 📌 Regole business da preservare

- `R1090` e `R8200` sono esclusi dai calcoli costo/operazioni rilevanti.
- `R6200` usa `NoOrg` come base minuti/costo effettivo.
- se una operazione `R*` ha `Færdigmeldt = 0` ma `NoOrg/Stykliste Minutter > 0`, il costo viene ricalcolato usando quel valore e la UI mostra `🕒`
- i prodotti `R*` dentro `Produkt dele` e nei sottoordini collegati vanno esclusi da vista e costi
- `R1100` con operatore `LASER EAGLE` e `ProdTp4=1` ha logica speciale di raddoppio.
- nelle viste laser aggregate, differenze tra prezzo unitario del popup e della riga principale possono dipendere dalla diversa quantità su cui viene ripartito lo stesso totale
- La risoluzione ricorsiva dei costi degli ordini figli è intenzionale e non va rimossa senza analisi funzionale.

---

## 🛠️ Release

Per pubblicare una nuova versione:

```powershell
.\publish.ps1
```

Lo script:
1. fa `git add/commit/push`
2. incrementa la versione patch
3. builda l’installer NSIS
4. pubblica la release GitHub

---

## 📄 Licenza / uso

Progetto interno Gantech.
