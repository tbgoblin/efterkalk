# Manuale venditori — Gantech Operations Hub

## 1. Obiettivo

Questo manuale spiega come i venditori devono:

1. controllare i valori di `Omsætning` (fatturato)
2. controllare i valori di `Ordreindgang` (entrata ordini)
3. stampare report chiari e verificabili
4. evitare errori comuni prima di condividere i numeri

L'interfaccia del programma e i nomi dei pulsanti sono in danese.

---

## 2. Accesso rapido ai moduli

Dalla schermata principale:

1. `Åbn Omsætning` per analisi fatturato
2. `Åbn Ordreindgang` per analisi entrata ordini

In entrambi i moduli usare sempre prima `Opdater` per caricare dati aggiornati.

---

## 3. Omsætning — procedura operativa

## 3.1 Impostazioni base

1. scegliere il periodo (`Fra måned` / `Til måned`)
2. selezionare i conti (`Kontoer`)
3. opzionale: selezionare uno o più clienti (`Kunde`)
4. verificare soglie (`Tærskler`)
5. premere `Opdater`

## 3.2 Cosa controllare

1. KPI: `Omsætning (Mio)`, `Rækker`, `Perioder`
2. grafico stacked: composizione per conto
3. grafico trend: andamento totale mese per mese
4. tabella soglie (`Månedstabel med tærskler`): cliccare sul mese per aprire il pannello mensile a destra
5. nel pannello mensile controllare importo contabilizzato, ordini collegati e `Ordreindgang` delle settimane relative al mese
6. dettagli (`Måned/Kunde detaljer`) se aperti

Gli importi del pannello mensile derivano dalle stesse registrazioni contabili `AcTr` del totale Omsætning. Il collegamento all’ordine usa il numero fattura presente in `Ord`, con `CustTr` come storico. Cliccando sul numero ordine si apre l’Efterkalkulation. La colonna conto non viene ripetuta nella tabella compatta, perché i conti sono già indicati dal filtro. Movimenti senza corrispondenza univoca rimangono visibili come controllo e non vengono eliminati dal totale. La riga `I alt` somma Ordreindgang e Tilbud delle sole settimane visualizzate. Una settimana a cavallo tra due mesi è mostrata interamente, per mantenere lo stesso valore del modulo `Ordreindgang`.

## 3.3 Modalità confronto clienti

Se selezioni più clienti, il sistema entra in modalità confronto.

Controlli minimi:

1. confermare i clienti attivi
2. verificare che i trend siano coerenti col periodo
3. in stampa verificare la sezione `Kunde-sammenligning (aktiv)`

---

## 4. Ordreindgang — procedura operativa

## 4.1 Impostazioni base

1. impostare settimane (`Fra uge` / `Til uge` in formato `YYYYWW`)
2. scegliere se mostrare la linea `Tilbud` (checkbox)
3. premere `Opdater`

## 4.2 Cosa controllare

1. KPI: `Total Ordre`, `Total Tilbud`, `Gns. Ordre`, `Tilbud → Ordre`
2. grafico `Ugeudvikling`: la linea orizzontale viola mostra lo stesso `Gns. Ordre` del KPI per tutto il periodo selezionato; la linea arancione resta invece la media mobile di 3 settimane
3. `Ugetabel` (se aperta)
4. `Topkunder` (se aperta)

---

## 5. Stampa report (molto importante)

Ogni modulo ha:

1. selettore `Layout: Auto / Stående / Liggende`
2. pulsante `Print rapport`

## 5.1 Regole layout

1. `Auto`: il sistema decide orientamento in base al contenuto
2. `Stående`: forza verticale
3. `Liggende`: forza orizzontale

Nel report stampato sono presenti:

1. periodo
2. layout scelto e tipo scelta (`Auto` o `Manuel`)
3. stato filtri attivi

## 5.2 Sezioni chiuse e stampa

Il report stampa solo sezioni aperte nella UI.

Esempi:

1. `Ugetabel` chiusa -> non stampata
2. `Topkunder` chiusa -> non stampata
3. `Måned/Kunde detaljer` chiusa -> non stampata

Prima di stampare, aprire solo ciò che deve comparire nel PDF.

---

## 6. Checklist venditore prima di inviare un report

1. periodo corretto
2. clienti/conti corretti
3. `Opdater` eseguito
4. sezioni da includere aperte
5. sezioni non necessarie chiuse
6. layout (`Auto` o manuale) verificato
7. controllare intestazione report: modulo, periodo, data e layout

---

## 7. Errori tipici da evitare

1. stampare senza `Opdater`
2. periodo invertito o incompleto
3. stampare confronto clienti senza clienti selezionati
4. dimenticare una sezione chiusa che serve nel report
5. usare layout verticale con tabelle molto larghe (preferire `Liggende`)

---

## 8. Mini procedura standard consigliata

1. entra nel modulo
2. imposta filtri
3. `Opdater`
4. apri solo le sezioni da stampare
5. scegli `Layout: Auto` (o manuale)
6. `Print rapport`
7. verifica anteprima
8. salva/invia

---

## 9. Note per il deploy

Quando il team conferma che i controlli venditori sono corretti:

1. test rapido Omsætning
2. test rapido Ordreindgang
3. test stampa con `Auto`, `Stående`, `Liggende`
4. validazione finale con 1 caso reale per modulo

Dopo questi passaggi il rilascio puo essere eseguito in sicurezza.
