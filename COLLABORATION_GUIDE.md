# 🏢 SPPS Lab Assistant - Guida alla Collaborazione

Questa guida spiega come configurare l'applicazione per lavorare insieme ai tuoi colleghi usando una cartella condivisa (OneDrive o Teams).

## 1. Configurazione Iniziale (Solo la prima volta)

1. **Sposta la Cartella**: Assicurati che l'intera cartella del progetto (che contiene `index.html`, `app.js`, `style.css` e il tuo file **Excel dell'inventario**) sia salvata in una cartella condivisa su **OneDrive** o **Teams**.

## 🔒 Requisito Fondamentale: Contesto Sicuro (HTTPS)

Per motivi di sicurezza, i browser moderni (Chrome, Edge) permettono l'accesso alle cartelle locali (necessario per la sincronizzazione) **solo se il sito è servito tramite HTTPS** o `localhost`.

### 1. Attivazione GitHub Pages
Una volta caricati i file su GitHub, segui questi passaggi per rendere il sito attivo:
1.  Vai sul tuo repository su GitHub.com.
2.  Clicca sulla scheda **"Settings"** (Ingranaggio in alto).
3.  Nella colonna a sinistra, clicca su **"Pages"**.
4.  Sotto **"Build and deployment"** > **"Branch"**:
    - Seleziona **`main`** (o il nome del tuo ramo).
    - Assicurati che sia selezionata la cartella **`/(root)`**.
    - Clicca su **"Save"**.
5.  Attendi circa 1-2 minuti. In alto apparirà un banner con scritto: *"Your site is live at..."* seguito da un link tipo `https://tuo-nome.github.io/nome-progetto/`.

### Opzioni Gratuite al 100%:
1.  **GitHub Pages (Pubblico)**: Se rendi il tuo repository **Pubblico** (da Settings > General > Change visibility), GitHub Pages diventa completamente **gratuito**. Dato che il codice non contiene i tuoi dati della sintesi (che rimangono sul tuo PC/OneDrive), renderlo pubblico è sicuro.
2.  **Netlify (Consigliato per semplicità)**:
    - Vai su [Netlify Drop](https://app.netlify.com/drop).
    - Trascina la cartella del progetto (`spps-lab-assistant`) direttamente nel box.
    - Ti verrà dato un link `https://...` immediato e gratuito, senza configurazioni.
3.  **Server Locale (Estensione Chrome)**: Installate l'estensione "Web Server for Chrome". Create un server che punta alla cartella del progetto. Vi darà un indirizzo `http://127.0.0.1:8887` che funzionerà perfettamente.

## 📁 Configurazione Area di Lavoro Condivisa
1.  **Sincronizzazione OneDrive/Teams**: Assicuratevi che la cartella del progetto sia sincronizzata localmente sul vostro PC (deve apparire in Esplora File).
2.  **Attivazione**: Aprite l'app (via HTTPS), andate in Dashboard e cliccate su **"Attiva Area di Lavoro Condivisa"**.
3.  **Selezione**: Selezionate la cartella radice del progetto sincronizzata sul vostro PC.
4.  **Auto-Discovery**: Una volta attiva, l'app troverà automaticamente l'inventario Excel e sincronizzerà le sintesi con gli altri colleghi.

---

## 2. Come far partecipare i colleghi

Per far sì che i tuoi colleghi vedano le tue sintesi e l'inventario, devono seguire questi passaggi:

1. **Condividi il Link**: Invia ai tuoi colleghi il link `https://...` che hai generato (quello di GitHub Pages o Netlify).
2. **Accesso dei Colleghi**:
   - Ogni collega deve aprire il link dal proprio browser.
   - Deve andare nella Dashboard e cliccare su **"🔗 Attiva Area di Lavoro Condivisa"**.
   - Deve selezionare la cartella del progetto che ha **sincronizzato localmente sul proprio PC** (quella di OneDrive/Teams).
3. **Sincronizzazione Automatica**:
   - Una volta selezionata la cartella, il programma caricherà automaticamente tutte le sintesi, le cartelle e l'inventario Excel.
   - **Tutto è pronto!** Possono iniziare a lavorare e le loro modifiche saranno visibili a te e a tutti gli altri in tempo reale.

---

## 3. Note Importanti

- **Autorizzazioni Browser**: Ogni volta che si riapre il browser, quest'ultimo potrebbe chiedere il permesso di accedere alla cartella. Clicca su "Consenti" o "Allow" per permettere il salvataggio dei dati.
- **File Excel**: Se sposti o rinomini il file Excel dell'inventario, dovrai ricollegarlo una volta tramite il tasto "🔄 Riconnetti" per aggiornare il riferimento per tutti.
- **Backup**: Usa il tasto **"💾 Backup"** periodicamente per scaricare una copia di sicurezza di tutte le tue sintesi sul tuo computer locale.

## 💡 Differenza tra Codice e Dati

È importante capire cosa viene sincronizzato e dove:

1.  **Codice (GitHub/Netlify)**: Qui risiede il "programma" (`app.js`, `style.css`, `index.html`). 
    - **Cosa succede**: Se ricevi un aggiornamento o una correzione (come quella dell'ordine degli amminoacidi), devi ricaricare i file su GitHub per vedere le modifiche nel link `https://...`.
    - **Frequenza**: Solo quando vuoi aggiornare le funzioni del programma.

2.  **Dati (OneDrive/Teams)**: Qui risiedono le tue **sintesi** (`spps_data.json`) e l'**inventario** Excel.
    - **Cosa succede**: Ogni volta che aggiungi una sintesi o modifichi una cartella, l'app salva direttamente in questa cartella locale sul tuo PC. OneDrive sincronizza questo file con i tuoi colleghi in tempo reale.
    - **Frequenza**: Automatica ad ogni "Salva". **Non c'è bisogno di fare nulla su GitHub per i dati!**

---
*SPPS Lab Assistant - Sicuro, Privato e Collaborativo.*
