# Gantech Efterkalk

App desktop per **efterkalkulation** e analisi margini ordini, pensata per uso interno in ambiente produzione/fabbrica.

**Versione attuale:** `1.0.50`

---

## ✨ Funzioni principali

- elenco rapido degli ultimi ordini fatturati
- ricerca per `OrdNo`
- calcolo costi/ricavi/margine per ordine
- dettaglio ordini di produzione collegati
- apertura **tegning/PDF** con pulsante `Vis tegning`
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

- [`docs/MANUALE_OPERATIVO_E_MANUTENZIONE.md`](docs/MANUALE_OPERATIVO_E_MANUTENZIONE.md)
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
| `services/drawingService.js` | ricerca/apertura disegni e immagini |
| `utils/productRules.js` | regole business prodotto |
| `utils/logger.js` | logging applicativo |
| `publish.ps1` | build + release automatizzata |

---

## ⚠️ Nota importante

Il codice di accesso attuale è un **blocco lato client/UI**, non una vera sicurezza server-side.

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