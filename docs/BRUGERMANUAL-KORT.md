# Brugermanual (kort)

Denne guide er lavet til daglig brug i produktion og administration.

## 1. Hvad kan jeg bruge systemet til?
- Få hurtigt overblik over ordrestatus, omsætning og ordreindgang.
- Sammenligne kunder, perioder og udvikling over tid.
- Slå en konkret ordre op og se kost, salg og margin.
- Udskrive visninger til møder og opfølgning.

## 2. Dashboard
- Start på dashboard og vælg det modul, du vil arbejde i.
- Brug knappen Ryd Efterkalk cache, hvis data virker forældede.
- Når cache ryddes, vises warmup-linjen, og data genindlæses i baggrunden.

## 3. Efterkalkulation (ordre-for-ordre)
Brug modulet når du vil forstå en enkelt ordre i dybden.

Du kan:
- Søge på ordrenummer.
- Åbne en ordre og se rapport med nøgletal.
- Se forskel mellem salg, kost og margin.
- Sætte flueben i `Udelad kost` på en salgsordrelinje, hvis netop den linjes kost ikke skal indgå i ordrekosten og marginen.
- Opdatere en enkelt ordre med knappen Opdater.

`Udelad kost` ændrer aldrig linjens salgspris: salget medregnes altid. Valget gemmes permanent i GOH og bruges også efter genstart og på andre arbejdsstationer, indtil fluebenet fjernes. Visma ændres ikke. Hvis GOH ikke kan bekræfte en ny ændring, bliver checkboxen ført tilbage, så en midlertidig rettelse ikke præsenteres som permanent.

Godt til:
- Opfølgning på ordre med lav margin.
- Kontrol før intern gennemgang eller kundedialog.

## 4. Omsætning (sammenlign kunder og perioder)
Brug modulet når du vil sammenligne performance.

Du kan:
- Filtrere på periode, kunde og andre relevante felter.
- Se totaler og udvikling på tværs af kunder.
- Klikke på en måned i `Månedstabel med tærskler` og se de bogførte beløb med tilknyttede efterkalkulationsordrer i panelet til højre.
- Se `Ordreindgang` for de virksomhedsuger, der har dage i den valgte måned.
- Identificere hvilke kunder der vokser eller falder.
- Udskrive resultatet efter opdatering.

Ordrenummeret i månedspanelet kan klikkes for at åbne Efterkalkulation. Omsætningsbeløbet kommer direkte fra de samme `AcTr`-bevægelser som månedstotalen; fakturanummeret bruges til koblingen. Kontokolonnen er udeladt fra den kompakte tabel, fordi de valgte konti allerede fremgår af filteret. En bevægelse uden en entydig ordre bliver stående som `Ikke koblet` og medregnes fortsat i totalen. Ugetabellens række `I alt` summerer både Ordreindgang og Tilbud for de viste uger. Uger, der går på tværs af to måneder, vises som hele uger, så `Ordreindgang` stemmer med det selvstændige modul.

Eksempel: Kunde-sammenligning
1. Vælg samme periode for alle kunder.
2. Filtrér først på kunde A og notér total omsætning.
3. Skift til kunde B og sammenlign.
4. Brug tallene til prioritering i salgsmøde.

## 5. Ordreindgang (fremadrettet overblik)
Brug modulet når du vil følge pipeline og fremdrift.

Du kan:
- Vælge Fra uge og Til uge.
- Se udvikling i ordre/tilbud over uger.
- Se periodens eksisterende `Gns. Ordre` som en lilla, vandret indikatorlinje i grafen. Værdien er den samme som KPI-feltet og skal ikke forveksles med den orange 3-ugers trendlinje.
- Få et hurtigt billede af kommende aktivitet.

Godt til:
- Ugeplanlægning.
- Kapacitetsdialog mellem salg og produktion.

## 6. Anbefalet daglig arbejdsgang
1. Start i dashboard og vælg modul efter opgaven.
2. Sæt filtre tydeligt (periode, kunde, uge).
3. Tryk Opdater.
4. Brug tallene til beslutning: sammenlign, prioriter, følg op.
5. Print ved behov til møder.

## 7. Hurtig hjælp ved tvivl
- Data ser gamle ud: brug Ryd Efterkalk cache og vent til warmup er færdig.
- Efterkalk kan ikke åbnes endnu: vent til warmup-linjen melder klar.
- Tal ser overraskende ud: opdater visningen og kontroller filtre igen.
