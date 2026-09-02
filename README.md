# Utility Mobilità

Raccolta di strumenti locali in HTML/JavaScript puro: ogni pagina funziona da sola nel browser, senza server e senza inviare file da nessuna parte. Non serve installare nulla: basta aprire i file `.html` (o pubblicarli su un hosting statico).

L'unica dipendenza esterna è la libreria [xlsx-js-style](https://github.com/gitbrent/xlsx-js-style) (fork di SheetJS con supporto alla scrittura degli stili delle celle), caricata da CDN in ogni pagina che legge o scrive file Excel. Al primo utilizzo serve quindi una connessione a internet; le pagine mostrano un avviso se la libreria non si carica.

## Struttura del progetto

| File | Cosa fa |
|---|---|
| `index.html` | Pagina di ingresso: un pulsante per ogni strumento. |
| `gestione-distribuzione.html` | Riconciliazione dei posti disponibili con quelli già assegnati. |
| `gestione-bandi.html` | Generatori di script SQL di inserimento per bandi, sedi in uscita e utenti. |

Quando si aggiunge una nuova pagina, va anche collegata da un nuovo pulsante in `index.html`.

---

## `index.html`

Griglia di "pulsantoni", uno per strumento, con titolo, breve descrizione e link alla pagina. Per aggiungere uno strumento basta un nuovo blocco `<a class="pulsante">` nella griglia.

---

## `gestione-distribuzione.html` — Gestione Distribuzione

Calcola il netto tra un file di posti disponibili e uno o più file di posti già assegnati da sottrarre.

### Flusso

1. **Posti disponibili** (file 1, obbligatorio): si carica l'Excel, si sceglie il foglio, la riga dei titoli e la prima riga di dati (con anteprima cliccabile), poi si indicano le **colonne chiave** (una o più, usate per abbinare le righe tra i file) e la **colonna del totale** di partenza.
2. **File da sottrarre** (facoltativi, se ne possono aggiungere quanti si vuole con "Aggiungi un file da sottrarre"): per ciascuno si carica l'Excel, si abbina ogni colonna chiave del file 1 alla colonna corrispondente in questo file, e si scelgono una o più **colonne da sommare**. Senza nessun file da sottrarre il netto coincide con i disponibili.
3. **Aggregazione dei codici** (facoltativa, per ogni file da sottrarre): regole per riscrivere uno o più codici in un codice diverso prima del confronto (utile quando i codici non coincidono esattamente tra i due file); mostra anche i codici del file che non trovano corrispondenza nel primo file, con il totale di posti che rappresentano.
4. **Calcolo**: per ogni riga del file 1, somma quanto trovato nei file da sottrarre sulla stessa chiave e calcola `netto = disponibili - sottratti`. I valori negativi si possono lasciare com'è o azzerare. Le chiavi presenti solo nei file da sottrarre (senza corrispondenza nel file 1) finiscono tra le "chiavi orfane".
5. **Esito**: tabella con filtri (tutte / solo variate / solo negative / solo non abbinate), esportazione in Excel e generazione di script SQL.

### Esportazione Excel

Foglio `RISULTATO`: si possono scegliere quali colonne originali del file 1 includere (le colonne chiave e quella del totale sono sempre incluse), più le colonne calcolate (dettaglio delle sottrazioni per file, `TOT_DA_SOTTRARRE`, netto, `ESITO`, `CONTROLLO`). Sulle righe con una sottrazione effettiva (`sub ≠ 0`) i posti in ingresso e il netto sono evidenziati in ambra; in rosso se il netto è negativo. `ESITO` vale `MODIFICATA` solo se c'è stata una sottrazione effettiva, altrimenti `NON MODIFICATA`.

Foglio `NON_ABBINATE`: le chiavi orfane con il relativo totale non sottratto.

Foglio `RIEPILOGO`: parametri usati, legenda colori, colonne incluse, totali.

### Script SQL

Due varianti (`SCRIPT_ORDINARIO` / `SCRIPT_104`, tabelle `MOBINT.SEDI_BANDO_ORDINARIO` / `MOBINT.SEDI_BANDO_104`): una `INSERT` per riga con `N_POSTI_ENTRATA` pari al netto appena calcolato, `ID_BANDO` parametrico, colonne del file 1 scelte per `CODICE_SEDE` e `ID_QUALIFICA`. Si può limitare l'export alle sole righe con almeno un posto positivo. Il pulsante "Copia SCRIPT_ORDINARIO" copia negli appunti lo stesso script che genererebbe il download, per chi lavora direttamente in un client SQL.

---

## `gestione-bandi.html` — Gestione Bandi

Tre generatori indipendenti di script SQL di inserimento (schema `MOBINT`), ciascuno con anteprima della prima `INSERT`, conteggio delle righe e un pulsante "Copia lo script" (oltre al download) prima dello scaricamento. Una barra di navigazione fissa in cima alla pagina permette di saltare direttamente a una delle tre sezioni.

### 1. Nuovo bando — `MOBINT.BANDI`

Form con anno, tipo bando, stato bando, descrizione, date (accettazione, apertura/chiusura, rettifica opzionale). Si possono mettere in coda più bandi (rimovibili singolarmente) e scaricare un unico script con tutte le `INSERT` più la `COMMIT` finale. L'`ID` non si inserisce mai a mano: è sempre `(SELECT NVL(MAX(ID), 0) + 1 FROM MOBINT.BANDI)`.

**Tipo bando** (`ID_TIPO_BANDO`): 1 L104/92, 2 Ordinaria, 3 104 Commissioni Territoriali, 4 104 Pnrr, 5 Ordinaria Commissioni Territoriali, 6 Ordinaria Pnrr.

**Stato bando** (`ID_STATO_BANDO`): 1 Aperto, 2 ValutazioneDomande, 3 ApertoPerAggiornamento, 4 ElaborazioneGraduatorie, 5 Concluso, 6 Chiuso (default), 7 PubblicazioneEsiti, 8 AperturaStraordinaria, 9 PubblicazioneGraduatoria.

Per ogni bando in coda il form chiede anche **ID_BANDO** e **nome del file** (di default `Mobilità.pdf`) del regolamento collegato: lo script genera, subito dopo ciascuna `INSERT` su `MOBINT.BANDI`, la relativa `INSERT INTO MOBINT.BANDI_REGOLAMENTI (ID, ID_BANDO, NOME)`, con lo stesso schema di `ID` a subquery (`NVL(MAX(ID), 0) + 1`) calcolato su `MOBINT.BANDI_REGOLAMENTI`.

### 2. Sedi in uscita — `MOBINT.BANDO_*_SEDI_USCITA`

Si carica un Excel, si sceglie la **tabella di destinazione**, le colonne per `CODICE_SEDE` e `N_MAX_USCITA` e si indica l'`ID_BANDO` (parametrico). Come per i bandi, l'`ID` è sempre `(SELECT NVL(MAX(ID), 0) + 1 FROM MOBINT.<tabella>)`. Le righe senza `CODICE_SEDE` sono saltate; quelle con `N_MAX_USCITA` a 0 sono escluse dallo script (entrambe contate a parte nell'anteprima).

Tabelle disponibili: `BANDO_OrdinarioCT_SEDI_USCITA`, `BANDO_104CT_SEDI_USCITA` (colonne base), `BANDO_OrdinarioPNRR_SEDI_USCITA`, `BANDO_104PNRR_SEDI_USCITA` (queste due hanno anche `NUM_DIP` e `DOTAZIONE`, prese da altre due colonne dell'Excel — i relativi selettori compaiono solo quando si sceglie una di queste due tabelle).

### 3. Utenti — `MOBINT.BANDO_*_UTENTI`

Si carica un Excel, si sceglie la tabella, le colonne per `MATRICOLA` e `CODICE_SEDE_USCITA` e si indica l'`ID_BANDO`. Stessa logica di `ID` a subquery delle sedi. Compare un elenco con tutte le righe valide (matricola e codice sede uscita non vuoti); ognuna ha una casella `CHECKCOPERTURA` modificabile singolarmente (di default tutte a 0), con due scorciatoie per segnarle o azzerarle tutte insieme. Le righe senza matricola o senza codice sede uscita sono saltate e contate a parte.

Tabelle disponibili: `BANDO_OrdinarioCT_UTENTI`, `BANDO_OrdinarioPNRR_UTENTI`, `BANDO_104PNRR_UTENTI`, `BANDO_104CT_UTENTI` (nessuna colonna extra per questa famiglia).

### Nomi dei file scaricati

Gli script di sedi e utenti includono nel nome del file un suffisso derivato dalla tabella scelta (es. `insert_sedi_uscita_104ct_20260101_1200.sql`), per distinguere gli script generati per tabelle diverse nella stessa sessione di lavoro.

---

## Interfaccia e accessibilità

Le tre pagine condividono lo stesso stile (dark, IBM Plex, accento ambra) e le stesse convenzioni: un pulsante "Copia" accanto a ogni script generato, per usarlo subito in un client SQL senza passare dal download, e gli elementi di stato/conteggio (es. `statoCalcolo`, `bandiConta`, `sqlConta`) marcati `aria-live` in modo che uno screen reader annunci i cambiamenti senza spostare il focus. Niente icone o emoji: lo stile è deliberatamente tipografico.

## Manutenzione di questo README

Questo file va tenuto aggiornato: ogni volta che si aggiunge, rimuove o modifica in modo sostanziale una funzionalità in una delle pagine (nuovo strumento, nuovo campo, nuova regola di validazione, nuova tabella, cambio di logica), va aggiornata anche la sezione corrispondente qui sopra.
