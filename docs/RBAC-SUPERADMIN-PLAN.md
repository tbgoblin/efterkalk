# RBAC e pannello Superadmin — Piano implementazione

## 1. Obiettivo

Introdurre gestione utenti con permessi granulari per modulo e un pannello `Superadmin` per:

1. creare/disattivare utenti
2. assegnare accesso ai moduli
3. configurare impostazioni principali del programma
4. mantenere tracciabilita delle modifiche

Questo piano e pensato per essere implementato in modo incrementale, senza bloccare il deploy operativo.

---

## 2. Modello ruoli proposto

## 2.1 Ruoli base

1. `superadmin`
2. `admin`
3. `seller`
4. `viewer`

## 2.2 Permessi modulo (boolean)

1. `modules.efterkalk`
2. `modules.omsaetning`
3. `modules.ordreindgang`
4. `modules.belastning`

## 2.3 Permessi amministrativi

1. `admin.users.read`
2. `admin.users.write`
3. `admin.settings.read`
4. `admin.settings.write`
5. `admin.audit.read`

---

## 3. Struttura utente consigliata

File attuale: `users.json`

Struttura target:

```json
{
  "username": "mrossi",
  "displayName": "Mario Rossi",
  "sellerUsr": "mrossi",
  "passwordSalt": "...",
  "passwordHash": "...",
  "isActive": true,
  "role": "seller",
  "permissions": {
    "modules": {
      "efterkalk": true,
      "omsaetning": true,
      "ordreindgang": false,
      "belastning": false
    },
    "admin": {
      "users": { "read": false, "write": false },
      "settings": { "read": false, "write": false },
      "audit": { "read": false }
    }
  },
  "createdAt": "...",
  "updatedAt": "...",
  "updatedBy": "admin"
}
```

Nota:

1. `isSuperUser` puo essere mantenuto solo per retrocompatibilita
2. da nuovo schema il controllo deve passare da `role` + `permissions`

---

## 4. Pannello Superadmin (UI)

## 4.1 Sezioni

1. `Utenti`
2. `Permessi moduli`
3. `Impostazioni programma`
4. `Audit log`

## 4.2 Operazioni utente

1. crea utente
2. reset password
3. attiva/disattiva
4. assegna ruolo
5. override permessi modulo

## 4.3 Impostazioni programma (fase 2)

1. default layout stampa (`auto/portrait/landscape`)
2. soglie default Omsaetning
3. limiti cache/refresh
4. visibilita sezioni default (es. tabelle collassate)

---

## 5. API backend da introdurre

## 5.1 Auth/sessione

1. `POST /auth/login`
2. `POST /auth/logout`
3. `GET /auth/me`

## 5.2 Utenti (solo admin)

1. `GET /admin/users`
2. `POST /admin/users`
3. `PATCH /admin/users/:username`
4. `POST /admin/users/:username/reset-password`

## 5.3 Impostazioni

1. `GET /admin/settings`
2. `PATCH /admin/settings`

## 5.4 Audit

1. `GET /admin/audit`

---

## 6. Sicurezza minima necessaria

1. mai password in chiaro
2. hash `pbkdf2` (o `scrypt`) con salt univoco
3. sessione firmata HttpOnly
4. controllo permessi server-side su ogni endpoint
5. endpoint admin bloccati se non autorizzato

---

## 7. Rollout consigliato (3 fasi)

## Fase 1 (rapida, pre-deploy)

1. aggiungere modello permessi a utenti
2. nascondere moduli in dashboard in base ai permessi utente corrente
3. endpoint `GET /auth/me` con permessi effettivi

## Fase 2

1. login reale username/password
2. pannello Superadmin utenti + permessi
3. audit base modifiche utenti

## Fase 3

1. impostazioni globali moduli/program
2. audit avanzato e export
3. hardening sicurezza e policy password

---

## 8. Criteri di accettazione

1. un utente senza permesso non vede il modulo
2. accesso endpoint modulo negato anche via chiamata diretta
3. Superadmin puo modificare permessi senza restart
4. modifiche permessi tracciate in audit con `who/when/what`
5. nessuna regressione su Omsaetning/Ordreindgang/stampa

---

## 9. Prossimo step pratico

Implementare Fase 1 subito nel codice:

1. endpoint `GET /auth/me`
2. oggetto permessi utente nel frontend
3. filtro card dashboard + blocco `openModule()` in base permessi

Questo permette di andare in deploy con una base solida e preparare il pannello Superadmin nella release successiva.
