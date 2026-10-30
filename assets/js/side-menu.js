// ── Side menu / Brugermanual · client ────────────────────────────────────────
// Estratto verbatim dall'inline script di server.js. Usa i globali condivisi
// della pagina: pushModalStack/removeModalStack, applyModulePermissions,
// goToDashboard, openModule, forgetRememberedLogin, showAccessGate,
// submitAccessCode, sanitizeDisplayName, setLoggedUserDisplayName (dichiarati
// nell'inline script) e resetSalgordreVia (dichiarato in assets/js/via.js).
function openSideMenu() {
    const overlay = document.getElementById('sideMenuOverlay');
    if (!overlay) return;
    overlay.classList.add('open');
    sideMenuOpen = true;
    refreshSideMenuAuthState();
    pushModalStack('sideMenuOverlay');
    const input = document.getElementById('sideMenuLoginInput');
    if (!accessGranted && input) {
        setTimeout(() => input.focus(), 30);
    }
}

function toggleSideMenu() {
    if (sideMenuOpen) closeSideMenu();
    else openSideMenu();
}

function closeSideMenu(event) {
    if (event && event.target && event.target.id !== 'sideMenuOverlay') return;
    const overlay = document.getElementById('sideMenuOverlay');
    if (!overlay) return;
    overlay.classList.remove('open');
    sideMenuOpen = false;
    removeModalStack('sideMenuOverlay');
}

function refreshSideMenuAuthState() {
    const userInput = document.getElementById('sideMenuUserInput');
    const input = document.getElementById('sideMenuLoginInput');
    const loginBtn = document.getElementById('sideMenuLoginBtn');
    const status = document.getElementById('sideMenuAuthStatus');
    const logoutBtn = document.getElementById('sideMenuLogoutBtn');
    const administrationBtn = document.getElementById('sideMenuAdministrationBtn');
    if (!status) return;

    if (accessGranted) {
        status.textContent = 'Logget ind som ' + sanitizeDisplayName(loggedUserDisplayName) + '.';
        status.classList.add('ok');
        if (userInput) userInput.disabled = true;
        if (input) {
            input.value = '';
            input.disabled = true;
        }
        if (loginBtn) loginBtn.disabled = true;
        if (logoutBtn) logoutBtn.disabled = false;
        if (administrationBtn) administrationBtn.style.display = loggedUserRole === 'superadmin' ? '' : 'none';
    } else {
        status.textContent = 'Ikke logget ind.';
        status.classList.remove('ok');
        if (userInput) userInput.disabled = false;
        if (input) input.disabled = false;
        if (loginBtn) loginBtn.disabled = false;
        if (logoutBtn) logoutBtn.disabled = true;
        if (administrationBtn) administrationBtn.style.display = 'none';
    }
}

function submitAccessCodeFromSideMenu() {
    const sideUserInput = document.getElementById('sideMenuUserInput');
    const sideInput = document.getElementById('sideMenuLoginInput');
    const gateUserInput = document.getElementById('accessGateUserInput');
    const gateInput = document.getElementById('accessGateInput');
    if (sideUserInput && gateUserInput) gateUserInput.value = String(sideUserInput.value || '').trim();
    if (sideInput && gateInput) {
        gateInput.value = sideInput.value || '';
    }
    submitAccessCode();
}

function navigateFromSideMenu(target) {
    if (target === 'dashboard') {
        goToDashboard();
        closeSideMenu();
        return;
    }
    if (target === 'brugermanual') {
        openBrugermanual();
        closeSideMenu();
        return;
    }
    if (target === 'personalehåndbog') {
        openPersonalehåndbog();
        closeSideMenu();
        return;
    }
    openModule(target);
    closeSideMenu();
}

async function logoutFromSideMenu() {
    accessSessionRevision += 1;
    _accessLoginInProgress = true;
    const tokenToRevoke = authToken;
    accessGranted = false;
    authToken = null;
    loggedUserRole = 'user';
    loggedUserPermissions = {};
    loggedUsername = '';
    resetSalgordreVia();
    if (dashboardSourceCache) dashboardSourceCache.clear();
    if (window.GohOrderFlow) window.GohOrderFlow.reset();
    omsaetningSummaryCache.clear();
    omsaetningSummaryInFlight.clear();
    setLoggedUserDisplayName('Bruger');
    applyModulePermissions();
    const adminCard = document.getElementById('administrationDashCard');
    if (adminCard) adminCard.style.display = 'none';
    closeSideMenu();
    goToDashboard();
    showAccessGate();
    const gateInput = document.getElementById('accessGateInput');
    if (gateInput) gateInput.value = '';
    refreshSideMenuAuthState();
    const button = document.getElementById('accessGateBtn');
    if (button) button.disabled = true;
    try {
        await Promise.all([forgetRememberedLogin(), fetch('/auth/logout', {
            method: 'POST', cache: 'no-store', signal: AbortSignal.timeout(8000),
            headers: tokenToRevoke ? { 'Authorization': 'Bearer ' + tokenToRevoke } : {}
        }).then(response => { if (!response.ok) throw new Error('Logout fejlede'); })]);
    } catch (_) {
        document.getElementById('accessGateLoginStorageStatus').textContent = 'Sessionen kunne ikke lukkes på serveren. Prøv Log ud igen.';
    } finally {
        _accessLoginInProgress = false;
        if (button) button.disabled = false;
    }
}

function openBrugermanual() {
    const modal = document.getElementById('brugermanualModal');
    const body = document.getElementById('brugermanualBody');
    if (!modal || !body) return;
    body.innerHTML = ''
        + '<section class="manual-card">'
        + '<h4>1. Login, dashboard og navigation</h4>'
        + '<p>Log ind med brugernavn og kode. "Husk brugernavn" gemmer kun brugernavnet. En gyldig session genbruges ved genindlæsning og retur fra Ledelsesrapport (op til 8 timer, indtil Log ud eller servergenstart).</p>'
        + '<p>I Windows-desktopappen kan "Husk kode" gemme det senest valgte login krypteret med Windows DPAPI i din Windows-profil, særskilt pr. postation/RDS-client. Login genoprettes ved næste start. Brug kun dette på din egen Windows-konto. Log ud eller Glem gemt login sletter den gemte kode. Uden tilgængelig Windows-kryptering eller identificerbar RDS-client vises valget ikke. I en almindelig browser bruges browserens egen adgangskodeadministrator. Koder gemmes aldrig i GOH, fælles cache eller localStorage.</p>'
        + '<p>Ledelsesrapport har separate perioder: Omsætning og Største kunder pr. måned, Ordreindgang pr. uge, Belastning som aktuel plan mellem to datoer (inklusive rest før startdato). VIA kan afgrænses efter ordredato for de aktuelt åbne ordrer; kost er aktuel, ikke historisk. Hver sektion angiver sin egen periode. Alle valgte måneder i Omsætning og alle valgte uger i Ordreindgang bliver i ét diagram pr. modul.</p>'
        + '<p>Widgetindstillinger åbnes med indstillingsikonet ved modul-pilen: eget navn, lokalt filter, rækkeantal og relevante valg. Belastning kan have 1-90 dage pr. widget, med/uden rest og aftenvisning. Laveste DB har en DB %-grænse, næste leveringer en horisont, og kvartalsomsætning viser kvartalets tre måneder. Gem indstillinger gemmer i din GOH-profil; Annuller bevarer tidligere valg.</p>'
        + '<p>I Salg og widgetbiblioteket findes Kunder og aktivitet samt Seneste ordreindgang. Kundegrafen viser Top3-6+Andre efter bogført omsætning eller antal fakturaordrer. Klik et segment for tilhørende kunder; CSV følger udsnittet. Ringens areal bruger positive kundesaldi, negative saldi og netto vises særskilt. Ordreindgang følger Visma-ordredato, ikke fakturadato eller teknisk oprettelsestid, og viser aktuelle ordreværdier for valgte 1-90 dage. Alle widgetvalg og indstillinger er personlige og gemmes i GOH.</p>'
        + '<p>Træk direkte i widgettens overskrift eller flyttehåndtag; ændr bredde og højde med hjørnehåndtaget. Ændringer gemmes automatisk i GOH ved slip. Piletaster virker også. Indret dashboard samler flere ændringer: Gem layout bekræfter, Annuller eller Escape fortryder. Standarddashboards tilpasses på den samme fane, kun for dig. Pak widgets tæt udfylder plads; tabeller ruller inde i widgetten. Gemt i GOH vises først efter bekræftelse. Ved fejl bevares en lokal recovery-kopi, hvis browserlagring virker, og synkronisering kan prøves igen. En nyere profil fra en anden postation overskrives ikke automatisk.</p>'
        + '<p>Side-menuens moduler er alfabetisk sorteret. Nye personlige dashboards vises som faner ved siden af Økonomi, Produktion, Salg og Ledelse. Tilpas ændrer den valgte dashboard: navn, widgets, rækkefølge og størrelse. De fire standarddashboards kan også ændres pr. bruger, uden at oprette en ekstra fane. Nulstil min dashboard i Tilpas gendanner kun din standardvisning. De fire standardvisninger tæller ikke med i grænsen på 12 personlige dashboards. Alle ændringer gemmes i din GOH-profil pr. bruger og database.</p>'
        + '<ul><li>Dashboard og side-menu bruger de samme modulrettigheder. Økonomiske widgets kræver Omsætning-adgang; produktion kan vises uden beløb.</li><li>Vælg blandt de tilladte skabeloner: Økonomi, Produktion, Salg eller Ledelse.</li><li>Ny dashboard opretter en personlig kopi. Tilpas: navn, widgetvalg, rækkefølge med pile og bredde. Gem bekræfter; Annuller bevarer det tidligere layout.</li><li>Mine dashboards rummer op til 12 personlige visninger. Positioner, størrelser, widgetvalg, navne, sidste visning, periode, antal rækker, søgning og Ordreflow-filtre gemmes i GOH pr. brugernavn og tilsluttet database, uden ordredata. Samme profil hentes på andre postationer. Tidligere lokale layouts importeres kun, hvis GOH-profilen ikke findes.</li><li>Økonomiperioden styrer Omsætning, DB, ansvarlige og seneste Efterkalk-ordrer. Regnskabsår starter i juli; Dette kalenderår starter 1. januar og henter også januar-juni. Måned og kvartal følger kalenderen. Omsætning bruger standardkonti og beholder kreditposteringer. Kreditnotaer med negativt fakturabeløb eller kreditmarkering i ordrenoten udelades fra ordre-widgets, også ved nulbeløb. Øvrige salgsordrer med nulbeløb eller negativt DB bevares. Top 5/10 begrænser kun visningen.</li><li>Kost inkl. stykliste-tillæg hentes via cache med højst tre samtidige beregninger. Antal med kendt kost vises i widgetten; manglende kost tæller aldrig som nul.</li><li>Produktion viser Belastning: dagarbejde mod kapacitet, aftenarbejde, restarbejde og daglig plan i timer for i dag + 20 dage. VIA viser aktuelle leveringer.</li><li>Søgning filtrerer Omsætning på kunde, Efterkalk/VIA på kunde/ansvarlig/ordre og Belastning på ressource, kunde eller ordrenummer. Kunden ændrer ikke ressourcernes tilgængelige kapacitet. Økonomiperioden ændrer ikke produktionsperioden.</li><li>Åbn modulet eller eksportér viste data til CSV. Luk af en ordre åbnet fra Home vender tilbage til Home; fra Efterkalkulation vender Luk tilbage til ordrelisten. Tilladte widgetkilder indlæses efter behov og genbruges i 15 minutter; søgning og layout ændrer ikke Visma.</li><li>Ved besked om genstart: genstart GOH-serveren og genindlæs siden. Brug Ryd Efterkalk cache kun ved kendte cacheproblemer.</li></ul>'
        + '</section>'
        + '<section class="manual-card">'
        + '<h4>2. Efterkalkulation</h4>'
        + '<p>Ordreliste, kost, salg, margin og produktionsdetaljer pr. ordre.</p>'
        + '<ul><li>Søg på ordrenummer eller åbn en ordre fra listen; filtrér listen på bruger, kunde og minimumsbeløb.</li><li>Åbn underordrer, operationer, laser/nesting, Ydelse og Underleverandør.</li><li>Skift marginberegning, og brug Rapport 2.0 til samlet visning, udskrift eller PDF.</li><li>Åbn "Vis tegning" og billedvisning, når dokumentation findes på produktet.</li><li>"Udelad kost" fjerner kun den valgte kost fra marginen; salget bevares, og valget gemmes i GOH.</li><li>"Opdater" genberegner kun den valgte ordre.</li><li>Fakturaoversigt: vælg kunde og datointerval, og se fakturaordrer, omsætning, kost og dækningsbidrag.</li><li>Månedens DB: vælg måned og se alle kunder og ordrer med registreret eller korrigeret kost.</li><li>I Månedens DB kan du se kundetrend, eksportere kunde- og ordre-CSV og gemme den færdigberegnede måned i GOH.</li></ul>'
        + '</section>'
        + '<section class="manual-card">'
        + '<h4>3. SalgOrdre VIA</h4>'
        + '<p>Åbne salgsordrer med registreret værdi i produktionen.</p>'
        + '<ul><li>Se Materiale, Stang, Indkøbte dele, Tid og Total kost pr. ordre.</li><li>Åbn købte dele og kontrollér Bestilt, Modtaget, Forbrugt og Medregnet VIA.</li><li>Bestilt alene giver ingen VIA-værdi.</li><li>Modtagne, fysisk tilstedeværende og sikkert reserverede dele flyttes fra Lager til VIA uden dobbelt værdi.</li><li>Reservationstabellen er information og lægges ikke oven i totalen.</li></ul>'
        + '</section>'
        + '<section class="manual-card">'
        + '<h4>4. Omsætning</h4>'
        + '<p>Periode-, konto- og kundebaseret omsætningsanalyse.</p>'
        + '<p>Ordreflow findes i månedspanelet og som widget i Home: Primo, Tilgang, Faktureret og Ultimo. Klik søjler for ordrer; filtrér på Tidligere/Nye og Kun færdigfakturerede ordrer. Widgetten har egen måned, søgning og CSV. Rest er salgsværdi, ikke VIA-kost. Alle omsætningskonti indgår, uanset kontofilter; kunder afgrænses som valgt. Historik rekonstrueres fra bogføring og aktuelle ordreværdier. Uafstemt fakturahistorik giver en synlig advarsel om delsummer.</p>'
        + '<ul><li>Vælg regnskabsår, måneder, konti og eventuelle kunder.</li><li>Klik en måned for bogførte poster, tilknyttede ordrer og Ordreindgang.</li><li>Brug manuelle månedsmål eller dagsmål × arbejdsdage.</li><li>Udskriv den aktuelle visning efter opdatering.</li></ul>'
        + '</section>'
        + '<section class="manual-card">'
        + '<h4>5. Ordreindgang</h4>'
        + '<p>Ugevis overblik over ordrer, tilbud, budget og udvikling.</p>'
        + '<ul><li>Vælg Fra uge og Til uge.</li><li>Vis eller skjul tilbud, trendlinje og manuelt budget.</li><li>Ferieuger kan udelades fra trend, MA3 og periodemål.</li><li>Brug KPI, graf, ugetabel og topkunder til opfølgning.</li></ul>'
        + '</section>'
        + '<section class="manual-card">'
        + '<h4>6. Lagerliste</h4>'
        + '<p>Officiel lagerværdi, VIA, månedslukning og periodesammenligning.</p>'
        + '<ul><li>Vælg standardpris eller FIFO for plader.</li><li>Opfølgningsvarer viser fysisk lagerværdi minus købte dele, der er flyttet sikkert til VIA.</li><li>Reservationer, fysisk værdi, flyttet VIA og resterende lagerværdi vises særskilt.</li><li>Brug Vareopslag til saldo, partier, reservationer og åbne ordrelinjer.</li><li>Gem måned/snapshot, sammenlign perioder og udskriv PDF; skrivefunktioner kræver Superadmin.</li></ul>'
        + '</section>'
        + '<section class="manual-card">'
        + '<h4>7. Lagerliste 2 (Beta)</h4>'
        + '<p>Kontrolvisning til route/nesting og forklaring af bevægelser mellem perioder.</p>'
        + '<ul><li>Se åbne, delvise og afsluttede routes med produkter, plader og REST.</li><li>Sammenlign perioder og se dokumenterede bevægelser mellem Lager, VIA, Færdige SO og Faktureret.</li><li>Uforklarede differencer forbliver synlige; Beta ændrer ikke Visma eller Lagerliste 1.</li></ul>'
        + '</section>'
        + '<section class="manual-card">'
        + '<h4>8. Ordreoversigt</h4>'
        + '<p>Produktionsoversigt til frigivelse og udskrivning.</p>'
        + '<ul><li>Indtast ordrenummer og hent varelinjer, save-/laserlister, rute og indkøbsdetaljer.</li><li>Oversigten viser leveringsdata, ressourcer, U-lev-ordrer og advarsler.</li><li>Brug Print til produktionspapirer.</li></ul>'
        + '</section>'
        + '<section class="manual-card">'
        + '<h4>9. Belastning</h4>'
        + '<p>Kapacitetsbelastning pr. ressource og dato.</p>'
        + '<ul><li>Vælg startdato, antal dage og ressourcegrupper.</li><li>Filtrér på ordre eller kunde.</li><li>Klik en søjle eller et kort for at åbne de underliggende ordrer.</li></ul>'
        + '</section>'
        + '<section class="manual-card">'
        + '<h4>10. BOMe+ Beregner</h4>'
        + '<p>Styklister, komponenter, ressourcer, materialer, parametre og tilbudsberegning.</p>'
        + '<ul><li>Tilgængelige områder afhænger af dine BOM-rettigheder.</li><li>Beregneren understøtter filanalyse, materiale, laser, buk og øvrige processer.</li><li>Kontrollér altid preview før en godkendt Visma-skrivning; read-only profiler kan ikke skrive.</li></ul>'
        + '</section>'
        + '<section class="manual-card">'
        + '<h4>11. Personalehåndbog og QMS</h4>'
        + '<p>Søgning og visning af interne personale- og kvalitetsdokumenter.</p>'
        + '<ul><li>Personalehåndbogen åbnes fra dashboard eller side-menu.</li><li>Brug søgefeltet til at finde relevante sider.</li><li>QMS-funktioner og redigering vises kun, når de er tilgængelige for brugeren.</li></ul>'
        + '</section>'
        + '<section class="manual-card">'
        + '<h4>12. Administration (Superadmin)</h4>'
        + '<p>Fælles opsætning er opdelt i tre kort.</p>'
        + '<ul><li>Brugere & adgang: brugernavn, visningsnavn, status og modulrettigheder.</li><li>Lagerliste · Diverse: månedlige manuelle værdier og skrotberegninger.</li><li>Omsætning · Arbejdsdage: arbejdsdage pr. måned til dagsmål og budget.</li><li>Indstillinger gemt i GOH gælder på tværs af arbejdsstationer.</li></ul>'
        + '</section>'
        + '<div class="manual-meta">Tip: Brug side-menuen til hurtig navigation. Ser data gamle ud, opdatér først det aktuelle modul; ryd kun hele Efterkalk-cachen ved et kendt cacheproblem.</div>';
    modal.style.display = 'flex';
    pushModalStack('brugermanualModal');
}

function closeBrugermanual(event) {
    if (event && event.target && event.target.id !== 'brugermanualModal') return;
    const modal = document.getElementById('brugermanualModal');
    if (modal) modal.style.display = 'none';
    removeModalStack('brugermanualModal');
}
