# Gantech Efterkalk

App desktop per **efterkalkulation** e analisi margini ordini, pensata per uso interno in ambiente produzione/fabbrica.

**Versione attuale:** `1.0.36`

---

## ✨ Funzioni principali

- elenco rapido degli ultimi ordini fatturati
- ricerca per `OrdNo`
- calcolo costi/ricavi/margine per ordine
- dettaglio ordini di produzione collegati
- apertura **tegning/PDF** con pulsante `Vis tegning`
- cache locale + warmup automatico per velocizzare l’avvio
- pacchetto desktop Windows con aggiornamento automatico via GitHub Releases

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

- `R1090` è escluso dai calcoli costo/operazioni.
- `R6200` usa `NoOrg` come base minuti/costo effettivo.
- `R1100` con operatore `LASER EAGLE` e `ProdTp4=1` ha logica speciale di raddoppio.
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