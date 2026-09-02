# Manuale operativo e di manutenzione — Gantech Efterkalk

## 1. Scopo dell’app

`Gantech Efterkalk` è una app desktop Windows basata su **Electron + Express + SQL Server** usata per:

- vedere gli ultimi ordini fatturati
- analizzare costi, ricavi e margini
- esplodere gli ordini di produzione collegati
- controllare operazioni, nesting e materiale laser
- aprire rapidamente i disegni PDF (`Vis tegning`)

Nel suo stato attuale l'app va considerata un **hub operativo interno**, non soltanto un calcolatore di margini. Riunisce informazioni commerciali e produttive di Visma/SQL Server e include anche moduli per BOM/preventivazione, magazzino, QMS, carico produttivo, fatturato e VIA.

**Versione applicativa documentata:** `1.1.47` (la fonte autorevole è il campo `version` di `package.json`).

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

I profili database possono essere marcati `readOnly`. Il backend controlla questo flag prima del comando BOM che crea prodotti in Visma e risponde `403` senza aprire una connessione o iniziare una transazione. Anteprima e verifica duplicati restano disponibili perché sono operazioni di lettura.

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

Il login crea una sessione server-side in memoria valida per 8 ore. Per compatibilità la pagina principale conserva il bearer token; un cookie `HttpOnly`, `SameSite=Strict` e same-origin permette alle pagine separate BOM/QMS di usare la stessa sessione. Il logout revoca entrambi.

Le scritture critiche BOM/QMS, le modifiche ai profili database e l’apertura locale dei PDF rifiutano richieste anonime. L’app resta destinata alla rete interna: non tutti gli endpoint di lettura hanno autorizzazioni granulari e le sessioni non sopravvivono al riavvio del processo.

### 4.2 Barra principale
Nella parte alta sono disponibili:

- `Søg` → apre il dettaglio ordine per numero
- `Opdater...` → azioni rapide su cache/lista/program
- `Skift marginberegning` → cambia formula di visualizzazione del margine
- `Skjul kundeliste` / `Vis kundeliste`
- `Ryd cache` → cancella la cache persistente
- filtro `Alle brugere`
- campo `Søg kunde i listen...`

### 4.3 Lagerliste 2 (Beta/Shadow)

`Lagerliste 2` è una pagina autonoma accessibile dalla dashboard e dal menu laterale. È stata separata intenzionalmente da Lagerliste 1: non modifica il calcolo esistente, non salva dati in Visma e usa soltanto endpoint di lettura protetti dal permesso `lagerliste`.

La pagina ha due viste:

1. **Route di nesting attuali**: raggruppa le righe `TrTp=5/7` per `nestingordre + route`, mostra lastre, prodotti (compresi i casi speciali `L2/L3`), ordini di produzione ricavati da `TrInf2`, ordine vendita ricavato dalle fonti `R4/Salgsref` o dalla gerarchia `OrdBasNo`, REST previsto, REST effettivamente registrato e stato della route.
2. **Confronto tra periodi**: usa gli stessi snapshot e gli stessi valori FIFO di Lagerliste 1, cerca contropartite note e separa i trasferimenti abbinati dalle righe ancora da controllare.

Per ridurre il numero di righe, la tabella route parte dal filtro `Ikke færdige`; le route concluse non vengono eliminate e restano disponibili scegliendo `Færdig` oppure `Alle`.

Stati mostrati:

- `⏳ Ikke startet`: nessun prodotto della route risulta completato;
- `◐ Delvist færdig`: almeno un prodotto è parziale oppure la route contiene righe con avanzamenti diversi;
- `✓ Færdig`: tutte le righe prodotto risultano completate;
- `? Ukendt`: la route non contiene righe prodotto sufficienti per determinarne lo stato.

Regole conservative della Beta:

- una lastra è riconosciuta con le condizioni storiche di Lagerliste 1 (`TrTp=5`, `ProdNo` valido che inizia per `3`, `Gr6=1` e relative esclusioni);
- qualsiasi riga prodotto `TrTp=7` appartiene alla route, senza esclusione basata sul suffisso: quindi anche `L2/L3` resta nel flusso VIA;
- per il riferimento alla vendita viene usata la precedenza conservativa `OrdLn.R4` → ultimo `ProdTr.R4` della stessa riga e dello stesso prodotto → `Ord.R4` → catena `TrInf2/OrdBasNo`; la fonte scelta è visibile accanto al numero SO;
- il filtro sullo stesso `ProdNo` è necessario perché una riga di nesting può essere riutilizzata: vecchie `ProdTr` della stessa `OrdLnNo` possono riferirsi a un prodotto e a una vendita precedenti;
- un prodotto senza R4 e senza collegamento gerarchico viene marcato `🏭 Ingen R4`: è un candidato per produzione interna o ordine lager, non un errore automaticamente confermato;
- la riga lastra negativa del nesting è mostrata come `↩ Forventet REST`: `NoOrg` rappresenta quantità e costo FIFO previsti; quando anche `NoFin` raggiunge `NoOrg`, il REST viene riconosciuto separatamente come `✓ REST færdigmeldt`;
- la registrazione effettiva `FreeInf1` (`FrInfTp=120`, `Gr7=1`) è mostrata come `✓ REST Plader`. Il collegamento usa prima la chiave esatta `nesting + ProdNo + codice SCR` (`OrdLn.TrInf2 = FreeInf1.Txt1`); solo quando tale codice non risolve il caso viene usato il collegamento meno specifico `nesting + ProdNo`;
- la færdigmelding del REST e la conclusione dei prodotti sono stati distinti: un REST può essere completamente færdigmeldt anche se un prodotto della stessa route è ancora parziale. In tal caso non viene mostrato un errore rosso; l'interfaccia segnala soltanto che la corrispondente riga non è ancora stata trovata nella lista REST `FreeInf1`. `⏳ REST ikke færdigmeldt endnu` è riservato alla riga REST il cui `NoFin` non ha ancora raggiunto `NoOrg`;
- `Materialeafvigelse` controlla la conservazione fisica/costo (`Plade FIFO - prodotti - REST previsto FIFO`). `REST-nedskrivning` mostra invece la differenza tra il costo FIFO del REST previsto e il valore di recupero realmente registrato: non viene colorata come errore materiale. Le differenze di quadratura entro 1 DKK sono trattate come arrotondamento visivo, ma il valore resta esposto;
- il REST entra nel pareggio tra periodi solo se è realmente registrato e collegabile senza ambiguità allo stesso nesting e alla stessa lastra;
- una riga lastra `TrTp=5` il cui `OrdLn.TrInf1` inizia con `Søg` viene marcata `♻ Søg-rest`: indica una rimanenza fisica non registrata in `Pladelager`/`REST`. Un suo ingresso in `VIA Plader` è spiegato come valore aggiunto da fonte non registrata, non come prelievo inventato dal magazzino. Nel confronto fra periodi il controllo viene eseguito sull'esatto `nestingordre + route + ProdNo`, anche se la route non rientra più nella finestra delle route attuali. I normali codici REST registrati (per esempio suffissi `_SCR0`) non rientrano in questa regola;
- i REST senza una sola route certa non vengono scartati: restano consultabili nella tabella espandibile `REST-rækker uden sikker routekobling`, con motivo e valore;
- un valore non spiegato resta visibile come residuo: l’algoritmo non inventa scarti e non forza il totale a zero;
- `VIA Tid` non proviene dal magazzino: un aumento viene classificato come valore di lavoro registrato e aggiunto all’ordine (`Registreret arbejdstid → VIA Tid`), con netto positivo; quando l’ordine termina può successivamente passare da VIA a `Færdige SO`;
- un aumento di `Pladelager` viene marcato `Indkøb → Pladelager` soltanto quando nel periodo esiste una corrispondente ricezione `ProdTr` con `TrTp=6`; l’entrata ha netto positivo;
- se la lastra acquistata viene ricevuta e immediatamente prelevata dallo stesso periodo, il movimento può apparire direttamente come `Indkøb → Pladelager → VIA Plader`, ma soltanto quando ricezione `TrTp=6`, prelievo `TrTp=5`, prodotto e nesting coincidono;
- una diminuzione di `Pladelager` viene collegata a `VIA Plader` soltanto quando `ProdTr` contiene un prelievo `TrTp=5` dello stesso prodotto nel periodo;
- le categorie `salgordreVia` e `finishedNotInvoiced` della Lagerliste storica possono contenere contemporaneamente lo stesso ordine. Lagerliste 2 non somma ciecamente entrambe: legge `NoPac` sulle righe prodotto principali (`ProdNo` che inizia per `1`) e ripartisce il costo tra VIA e `Færdige SO`. Quando `NoPac=NoFin`, `Færdige SO` diventa lo stato canonico e l'eventuale copia ancora presente in VIA viene esclusa; quando l'imballaggio è parziale, soltanto la quota non imballata resta in VIA. In questo modo il passaggio `VIA → Færdige SO` avviene una sola volta;
- la quota `NoPac / NoFin` è ponderata con `CCstPr` per non attribuire lo stesso peso a prodotti con costi diversi; righe pallet/trasporto come `510/520/521` non determinano la quota quando esistono righe prodotto `1…`;
- la scomparsa di un ordine da `Færdige SO` viene marcata `Færdige SO → Faktureret` soltanto se l’ordine possiede realmente `InvoNo`; se nello stesso intervallo l'ordine passa da VIA a completato e poi viene fatturato, viene mostrata la catena `VIA → Færdige SO → Faktureret` senza creare una contropartita positiva artificiale;
- l'uscita di un `Opfølgningsvare` viene collegata a `Færdige SO` soltanto quando nel periodo esiste una transazione negativa `ProdTr` (`TrTp=1`) dello stesso prodotto e `ProdTr.R4`/`OrdNo` identifica esattamente l'ordine vendita. Il prefisso del prodotto o la distinta base, da soli, non autorizzano l'abbinamento. Quando esiste un solo ordine destinatario certo, `ProdTr` prova il legame ma il suo `StcCst` corrente non limita il valore dello snapshot storico, perché il costo può essere stato ricalcolato successivamente;
- il residuo di una route `Plader VIA` completata viene collegato agli ordini di produzione/vendita risaliti da `TrInf2` e `OrdBasNo`, dopo avere prima sottratto l’eventuale REST registrato;
- un REST non viene mai abbinato alla diminuzione di una lastra per il solo `ProdNo`: deve comparire lo stesso nesting anche nel movimento `Plader VIA` del periodo;
- quando un REST registrato viene riutilizzato in un nesting successivo, l'ordine di origine e quello di consumo sono necessariamente diversi. Il trasferimento viene riconosciuto tramite il codice univoco `_SCR0` (`FreeInf1.Txt1` dell'origine = `OrdLn.TrInf1` della nuova lastra) e mostrato come `REST Plader → VIA Plader (genbrugt REST)` oppure `REST Plader → nesting` se la nuova route non è ancora iniziata;
- il REST è conservato al prezzo di recupero, mentre nel nuovo nesting può essere valorizzato al FIFO corrente della lastra. L'eventuale incremento è esposto separatamente come `Genindvundet værdi ved REST-forbrug → VIA Plader`, con netto positivo, invece di lasciare due falsi errori rosso/verde;
- un residuo di `Plader VIA` associato a una route non completata viene marcato `⏳ Åben route` invece di essere presentato come errore;
- i legami per prodotto/REST sono ad alta certezza; i legami basati soltanto sullo stesso ordine vendita sono mostrati con certezza media quando la ripartizione tra più route può essere ambigua.

La query route considera il mese corrente e i due mesi precedenti ed è mantenuta in cache per due minuti. Se uno snapshot storico non contiene `SalesOrdNo`, Lagerliste 2 prova a integrare il riferimento soltanto quando l'esatta `nestingordre + route` attuale porta a un unico ordine vendita; non sceglie tra più SO. `NoPac` non è storicizzato negli snapshot esistenti: la ripartizione delle sovrapposizioni usa lo stato Visma disponibile al momento del confronto e viene quindi presentata come riconciliazione operativa, non come ricostruzione contabile storica certificata.

`ShpBal` (`VareParti`) contiene informazioni utili sui lotti e sulle riserve, tra cui `RestBal`, `NoRsv`, `NoRsvInc`, `OrdNo`, costo e valore. Lagerliste 2 non usa ancora questi campi per pareggiare automaticamente gli ordini lager: `OrdNo` può rappresentare l'ordine di origine/ricezione del lotto e non dimostra da solo la successiva vendita destinataria. Finché il legame di prenotazione non è verificato, questi dati rimangono evidenza informativa e non una contropartita contabile.

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
| `Udelad kost` | checkbox UI persistita in GOH | se selezionata esclude il contributo costo della riga; `Salgspris` resta sempre incluso |
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

4. **Esclusione permanente del costo riga**
   - la checkbox `Udelad kost` non cancella la riga e non modifica Visma
   - il ricavo della riga (`DPrice × NoFin`) resta sempre incluso
   - viene sottratto dal costo ordine il contributo effettivo della riga (`EffectiveLineCost`; per un ordine di produzione collegato, `ProductionOrderTotalCost`)
   - ricavo, costo originale, costo escluso e costo rettificato restano visibili; vengono ricalcolati margine DKK, percentuale, lista ordini e `Rapport 2.0`
   - la scelta è salvata in GOH `dbo.AppState` con una chiave distinta per `OrdNo + LnNo` e resta valida fino alla rimozione del flag
   - se più righe condividono lo stesso ordine di produzione, il costo condiviso viene sottratto una volta sola e soltanto quando tutte quelle righe sono escluse; una selezione parziale resta segnalata senza creare un doppio storno
   - il salvataggio è accettato dalla UI soltanto dopo la conferma di GOH; in caso di errore la checkbox torna allo stato precedente

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

Con righe marcate `Udelad kost`:

```text
adjustedCost = totalCost - costo attribuibile alle righe escluse
```

Gli ordini di produzione condivisi vengono detratti una sola volta. Il fallback da minuti stykliste rimane separato e non viene nascosto implicitamente.

#### Ricavo totale ordine

```text
totalRevenue = Ord.InvoAm + Ord.DInvoIF
```

`Udelad kost` non modifica mai questa formula.

#### Margine in DKK

```text
margin = totalRevenue - adjustedCost
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
| `services/pdfOpenService.js` | validazione destinazione PDF e apertura senza shell |
| `services/authService.js` | utenti, sessioni bearer/cookie e guard di autenticazione |
| `services/aftercalcCostExclusionsService.js` | flag permanenti GOH per esclusione costo delle righe vendita |
| `services/omsaetningService.js` | riepilogo contabile Omsætning e dettaglio mensile fattura/ordine |
| `services/bomService.js` | letture BOM e creazione transazionale prodotti, con blocco `readOnly` |
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
| `GET /aftercalc-cost-exclusions/:ordno` | legge da GOH i flag costo permanenti dell’ordine |
| `POST /aftercalc-cost-exclusions/:ordno/:lineno` | salva/rimuove in GOH il flag della singola riga autenticata |
| `GET /omsaetning/month-detail` | dettaglio mensile AcTr, collegamenti fattura→ordine e settimane Ordreindgang |
| `GET /production-summary/:ordno` | riepilogo ordine di produzione |
| `GET /laser-route-metrics` | metriche laser/nesting |
| `GET /nesting-detail/:ordno/:prodno` | dettaglio nesting per prodotto |
| `POST /cache-refresh-order/:ordno` | refresh cache singolo ordine |
| `GET /cache-refresh-order-status/:ordno` | stato refresh ordine |
| `POST /cache-clear` | svuota cache persistente |
| `GET /cache-status` | elenco elementi in cache |
| `GET /warmup-status` | stato warmup iniziale |
| `POST /open-drawing` | apertura autenticata del PDF tramite viewer Windows predefinito |
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
- `dist/Gantech-Operations-Hub-Setup-<version>.exe`
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
- eseguire `npm test` (suite isolata: non usa Visma o database reali)
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

### Copertura dei test e rischio di regressione

Il progetto dispone di una prima suite automatica per autenticazione/sessioni, protezione delle scritture, comportamento BOM `readOnly` e apertura PDF. Usa esclusivamente simulazioni in memoria e non scrive su Visma o su database reali. Le verifiche operative su ordini reali rimangono necessarie perché la suite non copre ancora le formule di costing.

La priorità consigliata prima di ulteriori refactor è estendere i test di caratterizzazione con dati anonimizzati e risultati attesi per:

- esclusioni prodotto e operazione (`R1090`, `R8200`, componenti `R*`);
- fallback da `NoFin` a `NoOrg`;
- risoluzione ricorsiva dei costi degli ordini figli;
- calcolo e allocazione laser nei `MultiOrdre`;
- ordini con nesting multipli e fatturazione parziale;
- coerenza tra dati calcolati dal DB e dati letti dalla cache.

Questi test devono fissare il comportamento business approvato prima di separare o riscrivere le formule. I candidati principali per l'estrazione di funzioni di calcolo pure sono `services/aftercalcService.js` e `utils/productRules.js`.

### Priorità strutturali consigliate

1. Estendere la rete iniziale di test alle formule correnti con casi di caratterizzazione approvati.
2. Separare accesso SQL, normalizzazione dei dati e calcolo.
3. Suddividere `routes/apiRoutes.js` per dominio mantenendo invariati endpoint e payload.
4. Rendere tracciabile l'origine dei risultati (`memoria`, `disco`, `DB`), la data di calcolo e la versione dell'algoritmo.
5. Separare dal codice attivo snapshot, backup e artefatti operativi, senza eliminarli finché non esiste una politica di archiviazione approvata.

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
