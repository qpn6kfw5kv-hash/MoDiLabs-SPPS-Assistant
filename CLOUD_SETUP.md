# Workspace cloud - fase 1

## Stato effettivo

Implementati: accesso con codice email, sessione Supabase, creazione idempotente
di workspace, selezione workspace e visualizzazione membri/ruoli. Il database
supporta modifica del profilo, rinomina con controllo versione, cambio ruolo,
revoca e trasferimento della proprieta tramite RPC; queste ultime operazioni
non hanno ancora controlli nell'interfaccia.

Non implementati: inviti, sintesi condivise, inventario cloud, allegati cloud,
Realtime, migrazione dei dati locali e sincronizzazione offline.
Nessuna collezione locale viene caricata automaticamente. Il cloud e disattivato
per default. La modalita locale resta disponibile anche con cloud configurato.
La modalita locale non e protetta dall'account cloud: su un computer condiviso
usare profili browser separati. Una revoca non cancella dati gia scaricati.

## Schema e autorizzazioni

| Tabella | Identita | Contenuto |
| --- | --- | --- |
| profiles | id UUID, FK auth.users | Nome visualizzato e timestamp |
| workspaces | id UUID | Nome, proprietario, versione e chiave di creazione idempotente |
| workspace_members | workspace_id + user_id | Ruolo e revoca |

RLS e attiva su tutte e tre le tabelle. Gli utenti anonimi non hanno accesso.
Gli autenticati leggono solo i propri workspace e i profili dei membri con cui
condividono un workspace attivo. Le scritture dirette sono vietate anche agli
owner. Le RPC controllano auth.uid() e le appartenenze correnti nel database,
non ruoli dichiarati dal browser o dai metadati del profilo.

- Owner: gestione ruoli, revoche, rinomina e trasferimento di proprieta.
- Admin: rinomina e gestione editor/viewer, senza promuovere admin o modificare owner/admin.
- Editor e viewer: in questa fase possono leggere i workspace di appartenenza.
- Ogni account puo creare un proprio workspace; diventa owner solo di quello.

Le implementazioni privilegiate risiedono in spps_private, con search_path fisso
e grant limitati; i wrapper pubblici sono SECURITY INVOKER. Vincoli differiti
garantiscono coerenza tra owner_id e l'unico membro owner attivo.

## Attivazione su progetto di prova

1. Creare un progetto Supabase di staging. Non usare subito dati di laboratorio reali.
2. Eseguire nel SQL Editor, una sola volta, il file
   `supabase/migrations/202609280001_workspace_foundation.sql`.
   La migrazione e transazionale e additiva: non modifica i file locali e non
   contiene DROP delle collezioni esistenti. Crea profili anche per utenti Auth
   gia presenti. Se esistono tabelle omonime, fermarsi e verificarne lo schema:
   non cancellarle per forzare l'applicazione della migrazione.
3. Lasciare `spps_private` fuori dagli schemi esposti alla Data API.
4. Abilitare Email Auth e mantenere disabilitato l'accesso anonimo. Configurare
   il template email per mostrare `{{ .Token }}`: questa UI accetta un codice,
   non gestisce il ritorno da magic link. Per destinatari esterni al team del
   progetto serve un SMTP configurato. Verificare anche limiti di invio e Site URL.
5. Inserire in `cloud-config.js` URL del progetto e chiave pubblicabile
   `sb_publishable_...` (oppure legacy anon), poi impostare `cloudEnabled: true`.
   Non inserire mai service_role, secret key o password database nel frontend.
   La validazione del client non rende sicuro pubblicare un segreto nel file.
6. Pubblicare via HTTPS tutti i file runtime del sito, inclusi i cinque nuovi
   script: cloud-config.js, supabase-client.js, auth-service.js,
   workspace-service.js, collaboration-ui.js. Aggiornare anche index.html,
   app.js, lab-ui.js e workspace.css. Il SQL va applicato al database, non
   eseguito da GitHub Pages. `.env.example` e solo un promemoria: questo sito
   statico non carica automaticamente variabili da un file .env.
7. Verificare login, scadenza codice, rinnovo sessione, reload, logout e secondo
   browser sul progetto reale prima di proseguire alla fase scientifica.

Il client carica supabase-js 2.117.2 da jsDelivr solo quando il cloud e attivo.
Serve connettivita; non e promessa una modalita cloud offline.

## Verifica degli accessi prima dell'uso reale

Creare account di prova distinti: owner, editor, viewer e un estraneo, con due
workspace separati. La UI non offre ancora inviti: eventuali membership di test
vanno predisposte dall'amministratore nel database di staging, mai tramite
una chiave privilegiata nel browser.

Provare anche richieste REST/RPC dirette con i token dei singoli account:
nessuna lettura tra workspace, nessuna scrittura diretta, nessuna escalation
di ruolo, revoca immediata sulle nuove richieste, rinomina con versione obsoleta
rifiutata e trasferimento owner atomico. Verificare propagazione logout tra tab.
Il logout rimuove la sessione di questo dispositivo; non equivale a invalidare
istantaneamente ogni access token gia emesso. Non abilitare accesso pubblico
al proxy AI esistente: non e stato trasformato in un backend autenticato.

## Test riproducibili

Con Node disponibile, dalla cartella del progetto:

```powershell
node tests/setup-db-test.mjs
node tests/workspace-rls.test.mjs
node tests/cloud-ui.test.cjs
```

Il primo comando scarica PGlite 0.5.8 e verifica SHA-512. La dipendenza rimane
in `.cloud-test-deps`, esclusa da Git e non caricata dal sito. I 55 controlli SQL
eseguono la migrazione su PostgreSQL embedded: il contesto Auth/JWT e simulato,
non il motore SQL o le policy. Il test browser usa Edge e Playwright; impostare
PLAYWRIGHT_MODULE al percorso del modulo se diverso dal runtime locale usato
qui. Il browser usa account sintetici e un SDK simulato: non certifica Auth,
SMTP, PostgREST o configurazione del progetto Supabase reale.

## Passi successivi e rollback

Il prossimo incremento introdurra il repository dati scientifici, senza
riscrivere le formule: identificatori locali preservati, importazione esplicita
con backup, anteprima, batch idempotenti e controllo concorrenza. Solo dopo
verifica verranno aggiunti allegati privati, inviti e aggiornamenti Realtime.

Per disattivare questa fase impostare `cloudEnabled: false` e ripubblicare la
configurazione. Non eliminare tabelle o account: disattivare la UI non cancella
il database, non revoca sessioni e non annulla eventuali workspace creati.
Non cancellare localStorage ne file condivisi durante il passaggio.

Riferimenti: [RLS Supabase](https://supabase.com/docs/guides/database/postgres/row-level-security),
[codici email](https://supabase.com/docs/guides/auth/auth-email-passwordless),
[SMTP](https://supabase.com/docs/guides/auth/auth-smtp).
