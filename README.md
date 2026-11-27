# Gantech Efterkalk

App desktop per **efterkalkulation** e analisi margini ordini, pensata per uso interno in ambiente produzione/fabbrica.

Il prodotto è cresciuto fino a diventare un **hub operativo interno**: oltre al costing degli ordini riunisce dati commerciali e produttivi provenienti da Visma/SQL Server e comprende aree dedicate a BOM/preventivazione, magazzino, QMS, carico produttivo, fatturato e VIA.

**Versione attuale:** `1.1.124` (fonte: `package.json`)

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

### Kørselgodtgørelse

- Modulo nativo GOH, indipendente da Power Apps/SharePoint, disponibile dal menu e dal dashboard.
- In **Administration → Brugere** assegnare `Kørselgodtgørelse` (`mileage`) ai lavoratori; ai responsabili assegnare anche `Kørselgodtgørelse · Godkendelse og opsætning` (`mileageManage`). I superadmin hanno entrambi gli accessi. Solo i superadmin creano account GOH e modificano i relativi permessi; i responsabili collegano gli account esistenti ai dipendenti Visma.
- Importare una sola volta i quattro CSV di SharePoint dalla sezione amministrativa del modulo, verificando l'anteprima, quindi collegare ogni dipendente al proprio utente GOH. L'importazione conserva corse e dipendenti esistenti: associa i nomi identici ignorando maiuscole e spazi esterni e conserva i collegamenti GOH; nomi esistenti duplicati vanno prima risolti. I file originali e i dati personali non fanno parte del repository o dell'installer.
- All'apertura e all'aggiornamento il modulo sincronizza i dipendenti Visma (`Actor.EmpNo`, `Nm`, `Gr2`: 0 attivo, 99 non attivo) dal profilo aziendale unico; conserva targa e collegamento GOH esistenti. I record già presenti senza numero Visma vengono associati solo per nome esatto normalizzato e non ambiguo; eventuali corrispondenze ambigue sono segnalate e non associate automaticamente. I dipendenti Visma inattivi sono aggiornati se già riconosciuti, ma non vengono creati inutilmente se non presenti. Il primo profilo usato viene registrato: un profilo diverso è rifiutato per non mescolare aziende. I dipendenti manuali non riconosciuti da Visma non vengono rimossi o disattivati automaticamente.
- In **Medarbejdere og GOH-brugere**, un responsabile può associare facoltativamente ogni dipendente Visma al proprio utente Hub. Nelle nuove corse il dipendente associato è preselezionato, ma può essere cambiato. Senza associazione si può comunque scegliere un dipendente attivo. Se Visma non è raggiungibile, il modulo segnala che l'elenco conservato potrebbe non essere aggiornato.
- Lo storico Power Apps conserva i km già comprensivi del ritorno, le causali originali e gli altri dati esportati solo come archivio consultabile; le causali usano `Formaal` oppure `Titel` quando la prima colonna è vuota. Le corse con `Udbetalt` vero sono mostrate come pagate nell'archivio e nel filtro pagati, senza riapprovarle o contabilizzarle di nuovo; quelle non marcate restano tali. Non si inventano indirizzi separati né importi storici assenti. Le date mancanti restano documentate come tali e non bloccano nuove approvazioni.
- Ogni anno impostare esplicitamente le tariffe DKK/km fino a 20.000 km e oltre. L'anno è quello della trasferta; la franchigia residua considera solo i km delle corse operative già approvate/pagate per dipendente nello stesso anno, non lo storico Power Apps. L'assegnazione alle fasce avviene **nell'ordine di approvazione**, anche per trasferte retrodatate; le bozze mostrano solo una stima. Le tariffe non sono aggiornate automaticamente da SKAT.
- Chi ha accesso al modulo può creare una corsa per qualsiasi dipendente attivo; il creatore e il dipendente interessato possono vedere, modificare e inviare la bozza. Un responsabile può approvare anche una corsa propria o creata da lui, fissando importo, tariffe, fasce e audit; può respingere con motivazione solo corse di altri non create da lui. Il pagamento è una marcatura interna e non esegue bonifici. I valori già approvati non vengono ricalcolati dopo un cambio tariffa.
- Le corse storiche sono nascoste all'apertura e si mostrano con la casella **Vis historiske kørsler**. Le statistiche seguono i filtri selezionati e mostrano km e corse per mese, luoghi visitati e finalità; lo storico resta separato. Gli importi e le tariffe sono inviati e mostrati solo agli utenti con `mileageManage` o superadmin.
- Un responsabile può eliminare qualsiasi corsa con motivazione; il record rimane consultabile nel filtro amministrativo `Slettet` con la sua cronologia, ma è escluso da riepiloghi, PDF e calcoli successivi. Se vengono eliminati km operativi approvati/pagati, le corse approvate/pagate potenzialmente interessate sono marcate per controllo: i loro importi restano immutati e vanno verificati manualmente. Aggiungere, correggere o eliminare record d'archivio Power Apps non modifica le corse operative.
- PDF interno per singola trasferta e PDF annuale per dipendente. I PDF mostrano il nome completo del creatore, il nome azienda insieme all'indirizzo quando disponibile e le note in modo distinto; il PDF singolo usa inoltre una gerarchia visiva più chiara. Il PDF annuale include riepilogo e tabella delle corse; i record storici restano consultabili ma sono conteggiati separatamente e non entrano nei totali km o DKK. Gli importi compaiono solo nei PDF richiesti da chi ha `mileageManage` o è superadmin; bozze e richieste non approvate sono escluse dal totale DKK.
- Dati condivisi in `GOH.dbo.AppState`, chiave `mileage_v1`, con letture rigorose e scritture versionate atomiche; nessun fallback locale e nessun dato personale in localStorage. Il modulo non cambia per profilo Visma: riguarda il personale Gantech.
- **Percorsi opzionali:** un superadmin può inserire, sostituire o rimuovere la chiave in **Administration → Kørsel · Ruteservice**. La chiave è condivisa fra sessioni tramite `GOH.dbo.AppState` (`mileage_routing_settings_v1`), **in chiaro per scelta aziendale**: chi può leggere direttamente il database o i suoi backup può recuperarla. Limitare e proteggere tali accessi; l'interfaccia e le risposte API non restituiscono la chiave. Non inserire chiavi nel codice o in chat. La chiave salvata ha priorità su `GOH_ORS_API_KEY` nell'ambiente del processo; eliminandola, l'eventuale variabile resta attiva. Le modifiche sono disponibili alle altre sessioni alla successiva richiesta, senza riavvio.
- Il servizio usa gli endpoint HTTPS fissi `api.heigit.org/pelias/v1/search` per la ricerca indirizzi e `api.heigit.org/openrouteservice/v2/directions/driving-car` per i percorsi in auto. I vecchi endpoint `api.openrouteservice.org` sono stati disattivati da HeiGIT il 24 novembre 2026. Si usano la stessa chiave e la stessa quota ORS già configurate. Le richieste hanno timeout e limite locale di 30 chiamate/minuto per utente e processo. Solo un'azione esplicita dell'utente invia indirizzi/coordinate al servizio; nessuna richiesta esterna avviene senza chiave. Il ritorno viene calcolato attraverso le tappe in ordine inverso, non moltiplicando semplicemente l'andata. Errori o quote esaurite richiedono un nuovo tentativo oppure l'inserimento manuale esplicito, mai una distanza fittizia.
- **Find adresse online** invia la ricerca sia all'API ufficiale [Adressevælger di Klimadatastyrelsen](https://confluence.kds.dk/pages/viewpage.action?pageId=234782998) sia a openrouteservice e mostra i risultati di entrambi, con la relativa fonte. Per i risultati danesi con civico recupera le coordinate del registro DAR e le converte da ETRS89/UTM 32 a longitudine e latitudine WGS84 per ORS; non presenta risultati danesi con la sola via come indirizzi completi. La ricerca danese accetta al massimo 73 caratteri: per testi più lunghi viene usato solo openrouteservice con un avviso. Se un servizio fallisce, l'interfaccia lo segnala e mostra comunque i risultati dell'altro. I servizi possono proporre civici o località diversi da quelli cercati: verificare sempre il risultato prima di selezionarlo. Adressevælger usa il token pubblico raccomandato dal fornitore, che prevede in futuro una registrazione; verificare condizioni, privacy e aggiornamenti del servizio prima dell'uso aziendale. In caso di errore riprovare o inserire i km manualmente.
- Ogni lavoratore può cercare gli indirizzi condivisi e clienti/fornitori Visma per nome, numero o indirizzo; i suggerimenti vengono letti dalla tabella `Actor` con una query `SELECT` sui campi `CustNo`, `SupNo`, `Nm`, `Ad1`, `Ad2`, `PNo` e `PArea`. **Visma è di sola lettura: il modulo non scrive mai in Visma.** Solo i responsabili possono aggiungere, modificare o rimuovere indirizzi condivisi e causali dalla sezione di gestione. **Find adresse online** è esplicito. Selezionare un risultato online prima del salvataggio per conservarne le coordinate; modificando l'indirizzo salvato, le coordinate vengono rimosse per evitare una posizione errata.
- I mellemstop si possono riordinare trascinando la maniglia, facendo clic su di essa per scegliere una posizione, oppure usando le frecce della tastiera quando è selezionata; partenza e arrivo restano fissi. Durante il trascinamento una linea indica dove sarà inserito il mellemstop. Dopo lo spostamento occorre ricalcolare o inserire nuovamente la distanza. La ricerca online accetta anche nomi di luoghi e attività, eventualmente con la città, e mostra i risultati forniti dal servizio con via e numero quando disponibili. I dati dei luoghi non sono necessariamente completi: controllare il risultato prima di selezionarlo.
- Account, quota e condizioni del servizio devono essere verificati/accettati dall'azienda; nessun abbonamento viene attivato dall'app. Documentazione: [Directions](https://giscience.github.io/openrouteservice/api-reference/endpoints/directions/), [Geocoder](https://giscience.github.io/openrouteservice/api-reference/endpoints/geocoder/), [piani del fornitore](https://account.heigit.org/info/plans). I km manuali sono sempre disponibili senza API.

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

### Afterkalk: costo materiali senza ordine collegato (`2026-11-25`)

- Sugli ordini di vendita, le righe stock senza collegamento (`ProdTp4=0/2`) usano `CstPr`, come la colonna Visma `Kostpris`; nei sottoordini di produzione vale per i materiali `ProdTp4=2` senza acquisto collegato. Acquisti collegati e calcoli laser mantengono le proprie regole.

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

### Revisioni e archivio Økonomirapport

- **Hent rapport** blocca automaticamente il template alla prima lettura di ciascun rapporto/anno/periodo: righe, formule, impaginazione e selezione dei conti rimangono quelle, anche se la configurazione viene modificata successivamente. I mesi non ancora aperti usano il nuovo template. Gli importi della bozza vengono invece ricalcolati ad ogni lettura. La bozza resta salvabile per due ore, finché il server non viene riavviato.
- **Gem endelig rapport** conserva su GOH dati, conti di dettaglio, configurazione, revisione, autore, nota, versione dell'app e PDF. Il salvataggio avviene sulla copia server mostrata in anteprima, non su importi inviati dal browser.
- Ogni modifica alla configurazione crea una nuova revisione. Ogni nuovo definitivo dello stesso rapporto/anno/periodo crea **Rev. 2, Rev. 3, ...** e richiede una motivazione; non modifica le revisioni precedenti.
- **Gemte rapportrevisioner / Åbn gemt rapport** riapre la copia archiviata senza interrogare Visma. **Download PDF** restituisce gli stessi byte PDF salvati; **Download data (JSON)** esporta dati e configurazione.
- Il PDF usa A4 orizzontale per risultato e tabelle a 12 colonne, verticale per balance. Tabelle lunghe continuano su altre pagine con le intestazioni ripetute.
- Archivi e revisioni sono separati per profilo database e rapporto. Non esistono comandi per modificarli o eliminarli; GOH non disponibile o conflitti di salvataggio producono un errore esplicito.
- Le stampe precedenti all'introduzione dell'archivio non possono essere ricostruite retroattivamente: i primi definitivi vanno salvati dalla nuova funzione.

### VIA: rettifica del perimetro (`2026-09-16`)

- Con accesso Omsætning, la lista usa gli ordini con residuo `closing > 0,01 DKK` del modello Ordreflow, non tutti i candidati mensili. Tabella, KPI filtrato e CSV espongono il residuo, non il lordo.
- Lagerliste conserva il filtro storico e le regole di costo/allocazione. Costi mancanti, storico non riconciliato e residui minimi esclusi sono espliciti; il refresh singolo verifica l'identità dell'ordine.
- Verificati nel browser filtri/CSV, risposta errata, metadati, costi sconosciuti anche nel dashboard e recupero da errore. Dopo riavvio: 170/170 ordini e residui coincidenti con Ordreflow, t spiega il centesimo rispetto alla totalizzazione Ordreflow. Refresh singolo 413646 verificato senza modifiche agli altri ordini. Suite Node e confronto numerico di Lagerliste non eseguiti; nessun rilascio incluso.

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
