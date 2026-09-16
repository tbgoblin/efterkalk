# Brugermanual (kort)

Denne guide er lavet til daglig brug i produktion og administration.

## 1. Hvad kan jeg bruge systemet til?
- Få hurtigt overblik over ordrestatus, omsætning og ordreindgang.
- Sammenligne kunder, perioder og udvikling over tid.
- Slå en konkret ordre op og se kost, salg og margin.
- Kontrollere lager, igangværende arbejde, reservationer og købte dele til ordre.
- Udskrive visninger til møder og opfølgning.

## 2. Dashboard
- Dashboard og side-menu viser kun de moduler, din bruger har adgang til.
- Træk direkte i widgettens overskrift eller flyttehåndtag. Hjørnehåndtaget ændrer bredde og højde med musen. Ændringen gemmes automatisk i GOH, når musen slippes; andre widgets finder plads uden overlap. Begge håndtag understøtter piletaster. En ændret standardskabelon bliver til en personlig kopi.
- `Indret dashboard` samler flere ændringer i én redigering. Her bekræfter `Gem layout`, mens `Annuller layout` eller Escape fortryder. `Pak widgets tæt` udfylder ledig plads.
- Hver widget har et indstillingsikon ved siden af modul-pilen. Titel og lokalt filter kan tilpasses; lister har eget rækkeantal. Belastning har egen horisont på 1-90 dage, ressourcefilter og valg af rest-/aftenvisning. Laveste DB har en valgfri DB %-grænse, kommende VIA-leveringer har leveringshorisont, ansvarlige kan skjule DB %, og nøgletal kan skjule sekundære detaljer. `Gem indstillinger` gemmer i den personlige GOH-profil; `Annuller` ændrer intet. En standardskabelon bliver en personlig kopi.
- `Omsætning i kvartalet` viser total samt de tre måneder i kvartalet. Posteringer og kreditkorrektioner er de samme som i totalen. Måneder, der endnu ikke er begyndt, vises som afventende, ikke som nulomsætning. Månedsdetaljerne og kundeantallet kan slås fra i widgetindstillingerne.
- I Omsætningens stacked-graf står den samlede nettoomsætning for hver måned som en del af selve søjlen. Beløbet er i Mio DKK og følger de valgte konto-/kundekfiltre; det er ikke en ekstra tabel eller en ekstra summering. Samme SVG-label kommer med i `Print rapport`.
- `Kunder og aktivitet` viser en ringgraf med 3-6 største kunder og `Andre`. Vælg bogført omsætning eller antal fakturaordrer (kræver Efterkalk-adgang). Aktivitet tæller forskellige fakturerede salgsordrer i økonomiperioden, ikke fakturaer, nye ordrer eller regnskabslinjer. Klik et segment eller en kunde i forklaringen for samme udsnit i tabellen; klik igen for alle kunder. CSV følger udvalget og indeholder alle rækker i udsnittet.
- Ringgrafens arealer og procenttal beregnes kun af positive kundesaldi. Et negativt nettosaldo pr. kunde kan ikke være et positivt segment: netto, positive saldi og negative saldi vises særskilt. Klik `Negative kundesaldi` for de berørte kunder. Kreditposteringer modregnes fortsat i kundernes bogførte omsætning.
- `Seneste ordreindgang` viser de nyeste salgsordrer efter Visma-ordredato, derefter ordrenummer, uanset fakturering. Vælg 1-90 seneste dage og 5/10/20/50 rækker; beløb er aktuel ordreværdi efter Ordreindgangs valutaregel, ikke bogført omsætning eller et historisk snapshot. Tilbud, indkøbs-/produktionsordrer og kreditnotaer er udeladt. Widgetten kræver Ordreindgang og Omsætning; ordredetalje kræver Efterkalk. Visma-ordredato er ikke et teknisk oprettelsestidspunkt.
- `Kunder og aktivitet` og `Seneste ordreindgang` findes i skabelonen `Salg` og widgetbiblioteket under `Tilpas`. Eksisterende personlige dashboards ændres ikke automatisk. Indstillinger, valgte kundesegmenter og widgetvalg følger brugerens GOH-profil, også på andre arbejdsstationer.
- Lange tabeller og grafer ruller inde i den valgte widgetstørrelse. En høj Ordreflow-widget tvinger ikke nabowidgets til samme højde; ledig plads under en kort widget kan bruges af andre. Ordreflow bevarer måned, filtre, fokus og rulleposition under normale dataopdateringer.
- Dashboard-søgning i Belastning: ressourcenavn/-nummer filtreres lokalt, øvrig tekst søger kunde (fx `logi`), og et ordrenummer søger salgs-/produktionsordre. Kundefilteret følger Belastnings eksisterende SQL-logik. Kapacitet er fortsat ressourcernes tilgængelige kapacitet, ikke en beregnet kundeandel. Der ventes 450 ms efter indtastning; ens søgninger genbruger cache. Økonomiperioden skjules, når layoutet kun indeholder produktion og VIA.
- `Ordreflow / ordrebeholdning` findes i Økonomi, Salg og Ledelse samt under `Tilpas`. Widgetten har egen månedsvælger og viser Primo, Tilgang, Faktureret og Ultimo, opdelt i tidligere og nye ordrer. Klik en søjle for de tilhørende ordrer; vælg `Tidligere` eller `Nye`, søg lokalt og eksportér udsnittet med widgettens downloadikon. `Kun færdigfakturerede ordrer` viser ordrer, der er fuldt faktureret i måneden; vælg samtidig `Nye` for ind og ud i samme måned. Ordrenumre åbner Efterkalk, når du har adgang.
- Ordreflow er resterende **salgsværdi**, ikke produktionskost/VIA. Primo + tilgang - fakturering + eventuel regulering = ultimo. Indeværende måned slutter ved dagens dato; tidligere måneder ved månedens sidste dag. `Rest i dag, disse ordrer` følger de samme ordrer frem til nu, ikke ordrer modtaget efter den valgte måned.
- Ordreflow bruger alle omsætningskonti for at afstemme hele ordreværdien; Omsætnings kontofilter ændrer ikke dette. Kundefilteret respekteres i månedspanelet. Historik rekonstrueres fra finansposteringer og aktuelle ordreværdier, ikke gemte snapshots. Ændrede/annullerede ordrer kan derfor påvirke historikken. Ordrer uden entydig, afstemt fakturahistorik udelades med en synlig advarsel om delsummer. Negative ordreværdier indgår ikke som nye ordrer; korrektioner på almindelige ordrer bevarer fortegnet.
- Vælg `Økonomi`, `Produktion`, `Salg` eller `Ledelse`. Hver skabelon samler relevante widgets fra de tilladte moduler.
- Økonomiske dashboards og beløbswidgets kræver adgang til `Omsætning`, også når beløbene stammer fra Efterkalk eller VIA. Produktion kan bruges uden økonomiadgang; VIA-leveringslister viser da ingen beløb, heller ikke i CSV.
- `Ny dashboard` opretter en personlig kopi. Giv den et navn, vælg widgets, flyt dem med pilene og markér `Bred` for ekstra plads. `Gem dashboard` bekræfter ændringerne; `Annuller` bevarer den tidligere opsætning.
- `Tilpas` redigerer en personlig dashboard eller kopierer en standardskabelon. Du kan starte fra en anden skabelon eller en tom dashboard. Standardskabelonerne ændres aldrig.
- `Mine dashboards` skifter mellem dine gemte visninger. Der kan gemmes op til 12; en personlig dashboard kan omdøbes eller slettes fra editoren.
- Positioner, størrelser, widgetvalg, navne, valgt dashboard, økonomiperiode, antal rækker, søgning og Ordreflows måned/gruppe/søjle/filter gemmes i GOH-databasen, separat pr. loginbrugernavn og tilsluttet database. Ved næste adgang fra en anden postation hentes samme opsætning. Der gemmes ingen ordredata i profilen.
- `Gemt i GOH` vises først efter bekræftet lagring. Ved fejl vises en advarsel og mulighed for at prøve igen; en lokal kopi bevarer ændringerne, hvis browserlagring er tilgængelig. Hvis en anden postation har ændret profilen, overskrives den ikke automatisk. `Hent GOH-profil` beder om bekræftelse, før lokale, ikke-gemte ændringer erstattes.
- Tidligere lokale dashboards importeres ved første adgang, men kun hvis der ikke allerede findes en GOH-profil. En eksisterende GOH-profil har forrang. Genstart serveren og genindlæs siden efter installation af denne ændring.
- Omsætning-nøgletal, månedsgraf og kunder bruger bogførte Omsætning-data for den valgte økonomiperiode til indeværende måned. Standard er regnskabsåret fra juli; `Dette kalenderår` starter 1. januar og henter også januar-juni. Standardkontiene er 11012, 11015 og 11040; data er ikke begrænset af Efterkalks ordreliste. Måneder uden posteringer vises med nul.
- Økonomiperioden vælges som regnskabsår, indeværende måned, kvartal eller kalenderår. Kalenderår er ikke begrænset til regnskabsåret. Efterkalk-ordrer følger samme periode til i dag. Søgning filtrerer Omsætning på kunde, Efterkalk/VIA på kunde/ansvarlig/ordre og Belastning på ressource, kunde eller ordre. Vælg top/seneste 5 eller 10.
- DB-rangering, seneste Efterkalk-ordrer og ansvarlige bruger alle fakturerede salgsordrer (`TrTp = 1`) i den valgte periode, ikke den normale listes 150 ordrer. Indkøbs- og produktionsordrer medtages ikke. Ordren skal have fakturanummer og seneste fakturadato i perioden. Kreditnotaer udelades fra widgets, optællinger og CSV: negativt fakturabeløb eller markeringen `Kreditnota` i ordrenoten, også ved nulbeløb. Øvrige salgsordrer med nulbeløb eller negativt DB medtages stadig. Top/seneste 5 eller 10 begrænser kun visningen, ikke datagrundlaget. Bogført Omsætning beholder kreditposteringer.
- Åbnes en ordre fra Home/dashboard, vender `Luk` tilbage til Home. Åbnes den fra Efterkalkulation, vender `Luk` tilbage til Efterkalk-listen. Fra en rapport vender første `Luk` tilbage til ordren; næste `Luk` går til udgangspunktet.
- Opdateringsikonet ved CSV-knappen (`Opdater dashboard-data`) genindlæser de synlige widgets uden at ændre layout eller filtre. Knappen er låst under opdateringen; status viser tidspunkt eller fejlede kilder. Tidligere data bevares ved fejl. Ordrenoter genindlæses, og kost hentes via serverens cache med højst tre samtidige forespørgsler. VIA-reservationer og skjulte moduler genindlæses ikke.
- DB er fakturabeløb minus kendt ordrekost inkl. stykliste-tillæg, som i Efterkalk-listen. Kost hentes progressivt via eksisterende cache, højst tre ordrer ad gangen, når DB eller Kostgrundlag vises. Widgetten viser antal ordrer med kendt kost; manglende kost tæller aldrig som nul. Beregningen stopper, når dashboardet forlades, eller brugeren/databaseprofilen skifter.
- `Ansvarlige efter fakturabeløb` viser to linjer pr. ansvarlig: fakturabeløb og DB %. DB % er samlet DB divideret med samlet fakturabeløb for ordrer med kendt kost, ikke gennemsnittet af ordreprocenter. Ved manglende kost vises antal med kendt kost/antal ordrer; nul i beregningsgrundlag giver ingen procent. Sorteringen følger stadig fakturabeløb. DB % og kostdækning medtages i CSV.
- Produktion har fire Belastning-widgets: kapacitetsnøgletal, dagarbejde mod kapacitet pr. ressource, daglig plan og restarbejde før i dag. Standard er dagens dato + 20 dage, men hver widget kan vælge 1-90 dage og bruger samme datakilde som Belastning, omregnet fra minutter til timer. Samme horisonter genbruger én cache; forskellige horisonter hentes separat. Aften vises separat; restarbejde omfatter dag og aften. Ressourcebelastning sammenligner dagarbejde med kapacitet, med/uden rest efter indstillingen; nul kapacitet vises som `Ingen kapacitet`.
- Med Omsætning-adgang viser SalgOrdre VIA de åbne ordrer fra Ordreflow for indeværende måned til i dag, alle kunder og restsaldo over 0,01 DKK. Færdigfakturerede månedskandidater medtages ikke. MultiOrdre og ordrer uden produktion medtages, hvis restsaldoen er åben. Uafstemt fakturahistorik og små udeladte restbeløb markeres i status.
- Kolonnen `Restsalgsværdi`, nøgletallet `Ordrebeholdning` og CSV bruger de samme viste ordrer og restbeløb. Søgning filtrerer også totalen. Ukendt kost vises som `Ikke hentet`, ikke nul; manglende produktion kan give kendt nul i registreret kost. Ved opdateringsfejl bevares tidligere data med en advarsel. Enkeltordre-opdatering ændrer kun den valgte ordre; øvrige ordrer opdateres ved genindlæsning af listen.
- Uden Omsætning-adgang bevares det historiske produktionsudvalg uden restsaldo. Lagerlistes værdiansættelse bruger fortsat sit historiske produktionsfilter og eksisterende fordelingsregler, ikke VIA-skærmens kommercielle ordreudvalg. Serveren skal genstartes og siden genindlæses efter ændringen.
- Ordrelinjer åbner Efterkalk eller filtrerer VIA på ordren. Pilknappen i widgethovedet åbner det relevante modul. Downloadknappen eksporterer de viste widgetdata til CSV.
- Omsætning, periodens fakturaordrer og Belastning hentes kun for synlige, tilladte widgets. Data genbruges i 15 minutter, samtidige hentninger samles, og filtre ændrer ikke Visma. Et nyt widgetmodul kan udløse sin første hentning; søgning og omrokering genbruger data. Belastning-widgettens udsnit er uafhængigt af modulernes personlige filtre.
- Hvis dashboardet melder, at opdateringen afventer genstart, skal GOH-serveren genstartes og siden genindlæses. Det er ikke et tegn på manglende ordredata.
- Opdatér først det relevante modul ved gamle data; brug kun `Ryd Efterkalk cache` ved et kendt cacheproblem.
- Når cache ryddes, vises warmup-linjen, og data genindlæses i baggrunden.

## 3. Efterkalkulation (ordre-for-ordre)
Brug modulet når du vil forstå en enkelt ordre i dybden.

Du kan:
- Søge på ordrenummer eller åbne en ordre fra listen.
- Filtrere ordrelisten på bruger, kunde og minimumsbeløb.
- Åbne en ordre og se rapport med nøgletal, salgsordrelinjer og produktionsordrer.
- Se forskel mellem salg, kost og margin.
- Skifte marginberegning i visningen.
- Åbne underordrer, operationer, laser/nesting, `Ydelse` og `Underleverandør`.
- Bruge `Rapport 2.0` til samlet visning, udskrift eller PDF.
- Åbne `Vis tegning` og billedvisning, når dokumentation findes på produktet.
- Sætte flueben i `Udelad kost` på en salgsordrelinje, hvis netop den linjes kost ikke skal indgå i ordrekosten og marginen.
- Opdatere en enkelt ordre med knappen Opdater.

### Fakturaoversigt
Brug `Fakturaoversigt` til at vælge én kunde og et datointerval. Visningen samler kundens fakturaordrer og viser omsætning, beregnet kost og dækningsbidrag. Kundetrenden kan åbnes direkte fra oversigten.

### Månedens DB
Brug `Månedens DB` til at hente alle fakturaordrer for en valgt måned.

Du kan:
- Se totaler pr. kunde og de underliggende ordrer.
- Skifte mellem `Kun registreret` og `Inkl. manglende tid fra stykliste`.
- Se det særskilte stykliste-tillæg, når registreret produktionstid mangler.
- Eksportere både kunder og ordrer til CSV.
- Gemme den færdigberegnede månedsrevision i GOH, så den kan genåbnes og bruges til trend.

Måneden gemmes først, når ordreberegningerne er gennemført. En ordre uden komplet kost markeres i stedet for at få opfundet en værdi.

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
- Se det interaktive Ordreflow øverst i månedspanelet. Det følger den valgte måned og kunder; samme datagrundlag genbruges i Home-widgetten ved ens filtre.
- Se `Ordreindgang` for de virksomhedsuger, der har dage i den valgte måned.
- Identificere hvilke kunder der vokser eller falder.
- Udskrive resultatet efter opdatering.

Ordrenummeret i månedspanelet kan klikkes for at åbne Efterkalkulation. Omsætningsbeløbet kommer direkte fra de samme `AcTr`-bevægelser som månedstotalen; fakturanummeret bruges til koblingen. Kontokolonnen er udeladt fra den kompakte tabel, fordi de valgte konti allerede fremgår af filteret. En bevægelse uden en entydig ordre bliver stående som `Ikke koblet` og medregnes fortsat i totalen. Ugetabellens række `I alt` summerer både Ordreindgang og Tilbud for de viste uger. Uger, der går på tværs af to måneder, vises som hele uger, så `Ordreindgang` stemmer med det selvstændige modul.

Eksempel: Kunde-sammenligning
1. Vælg samme periode for alle kunder.
2. Filtrér først på kunde A og notér total omsætning.
3. Skift til kunde B og sammenlign.
4. Brug tallene til prioritering i salgsmøde.

## 5. SalgOrdre VIA
Brug modulet til at se værdien på åbne salgsordrer, der stadig er i produktion.

Du kan:
- Søge på ordre eller kunde og sortere alle kolonner.
- Se `Materiale`, `Stang`, `Indkøbte dele`, `Tid` og samlet kost pr. ordre.
- Åbne `Indkøbte dele til ordre` og sammenligne `Bestilt`, `Modtaget` og `Forbrugt`.
- Eksportere den viste liste til CSV.
- Se `Reserveret til ordre` som en separat lageroplysning.

Vigtigt:
- `Forbrugt` (`NoFin`) tæller med i VIA.
- En modtaget indkøbsdel tæller også med, når den er fysisk på lager og sikkert reserveret til den åbne ordre; samme værdi trækkes samtidig ud af Lager.
- `Bestilt` alene giver ikke VIA-værdi.
- Reservationstabellen lægges aldrig oven i totalen som en ekstra værdi.

## 6. Lagerliste
Brug modulet til den officielle lageroversigt, månedslukning og periodekontrol.

Du kan:
- Se `Aktuel` lagerliste eller åbne en gemt måned/dags-snapshot.
- Vælge standardpris eller FIFO for plader i den viste rapport.
- Slå et varenummer op med `Vareopslag` og se saldo, partier, reservationer og åbne ordrelinjer.
- Sammenligne `Periode A` og `Periode B` og udskrive den viste rapport som PDF.
- Åbne `Reserveret til ordre · info` efter Opfølgningsvarer.

`Opfølgningsvarer` viser både fysisk lagerværdi, `Flyttet til VIA` og `Lagerværdi efter VIA`. Modtagne indkøbsdele med sikker reservation flyttes til VIA; andre reservationer bliver i Lager. En reservation markeret `Ikke medregnet` mangler et sikkert parti eller en åben salgsordre og ændrer ingen totaler.

Kun Superadmin kan gemme/slette lukninger og snapshots, migrere lokale måneder til GOH eller oprette manuelle afstemninger. En manuel afstemning dokumenterer en forklaring; den ændrer aldrig Visma eller Lagerliste-totalen.

`Lagerliste 2 (Beta)` bruges til route/nesting og bevægelsesforklaring. Den er en kontrolvisning og erstatter ikke Lagerliste 1 som officiel lukning.

## 7. Ordreindgang (fremadrettet overblik)
Brug modulet når du vil følge pipeline og fremdrift.

Du kan:
- Vælge Fra uge og Til uge.
- Se udvikling i ordre/tilbud over uger.
- Se periodens eksisterende `Gns. Ordre` som en lilla, vandret indikatorlinje i grafen. Værdien er den samme som KPI-feltet og skal ikke forveksles med den orange 3-ugers trendlinje.
- Få et hurtigt billede af kommende aktivitet.

Godt til:
- Ugeplanlægning.
- Kapacitetsdialog mellem salg og produktion.

## 8. Ordreoversigt
Brug modulet til produktionsoversigten, når en ordre frigives.

Du kan:
- Hente en ordre og se varelinjer, save-/laserlister, rute og indkøbsdetaljer.
- Se leveringsdata, ressourcer, U-lev-ordrer og advarsler.
- Udskrive produktionspapirerne.

## 9. Belastning
Brug modulet til kapacitetsbelastning pr. ressource og dato.

Du kan:
- Vælge startdato, antal dage og ressourcegrupper.
- Filtrere på ordre eller kunde.
- Klikke på en søjle eller et kort for at se de underliggende ordrer.

## 10. BOMe+ Beregner
BOM-området samler styklister, komponenter, ressourcer, materialer, parametre og tilbudsberegning.

- De viste områder afhænger af dine BOM-rettigheder.
- Beregneren understøtter filanalyse, materiale, laser, buk og øvrige processer.
- Kontrollér altid preview før en godkendt Visma-skrivning.
- En databaseprofil markeret `readOnly` kan ikke skrive til Visma.

## 11. Personalehåndbog og QMS
- Personalehåndbogen åbnes fra dashboard eller side-menu og kan gennemsøges.
- QMS viser interne kvalitetssider og dokumenter, når funktionen er tilgængelig.
- Redigeringsfunktioner afhænger af brugerens adgang.

## 12. Administration (Superadmin)
Administration er opdelt i tre kort:

- `Brugere & adgang`: opret brugere, vælg visningsnavn, aktiv status og modulrettigheder.
- `Lagerliste · Diverse`: vedligehold månedlige manuelle værdier og skrotberegninger.
- `Omsætning · Arbejdsdage`: fastlæg arbejdsdage pr. måned til dagsmål og budget.

Indstillinger gemt i GOH gælder på tværs af arbejdsstationer.

## 13. Login og husket brugernavn
- Log ind med dit brugernavn og din kode.
- `Husk brugernavn på denne computer` gemmer kun brugernavnet, eksempelvis `MW`.
- Visningsnavnet, eksempelvis `Martin`, bruges kun i hilsner og brugeroversigter.

## 14. Anbefalet daglig arbejdsgang
1. Start i dashboard og vælg modul efter opgaven.
2. Sæt filtre tydeligt (periode, kunde, uge).
3. Tryk Opdater.
4. Brug tallene til beslutning: sammenlign, prioriter, følg op.
5. Print ved behov til møder.

## 15. Hurtig hjælp ved tvivl
- Data ser gamle ud: brug Ryd Efterkalk cache og vent til warmup er færdig.
- Efterkalk kan ikke åbnes endnu: vent til warmup-linjen melder klar.
- Tal ser overraskende ud: opdater visningen og kontroller filtre igen.
- Reservation mangler: kontrollér at den har parti og er knyttet til en åben salgsordre.
- Købt del har 0 DKK i VIA: kontrollér `Modtaget`, fysisk beholdning og sikker reservation; `Bestilt` alene er kun information.
