# Bandi Regione Veneto – app standalone + aggiornamento automatico

## Cosa c'è
| File | Funzione |
|---|---|
| `bandi_veneto.html` | App **standalone**: un solo file, dati incorporati, nessuna dipendenza esterna. Si apre con doppio clic (anche offline). |
| `aggiorna_bandi.py` | Scarica le fonti, aggiorna `data/bandi_veneto.json` e **`data/bandi_veneto.xlsx`** (Excel) e rigenera `bandi_veneto.html` (+ `docs/index.html`). |
| `standalone_template.html` | Modello dell'app (non modificare i dati qui: vengono inseriti dallo script). |
| `cache/detail_cache.json` | Schede di dettaglio già scaricate (evita di riscaricarle). |
| `.github/workflows/aggiorna.yml` | Aggiornamento automatico su GitHub (giornaliero o settimanale). |

## Fonti
1. **Portale Bandi Avvisi Concorsi** (`bandi.regione.veneto.it`): tutti i bandi + gli avvisi ancora pubblicati. Usa endpoint JSON interni, non documentati: possono cambiare senza preavviso.
2. **Calendario inviti FESR 2021-27** (`venetocoesione.regione.veneto.it`): bandi *previsti* (mese di apertura, dotazione).
3. **Calendari inviti FSE+** (stesso sito, CSV): bandi *previsti*.
4. **Foglio «Opportunità» di Veneto Coesione** (Google Sheet pubblico): atti FESR/INTERREG aperti non presenti nel portale.

Ogni fonte è indipendente: se una fallisce, le altre proseguono e la fonte compare come «non aggiornata».
Se il portale non è raggiungibile lo script esce con errore e **non** sovrascrive l'ultimo dataset buono.

Non integrate (da verificare): AVEPA (sito non raggiungibile dal mio ambiente di prova; è l'organismo pagatore, i bandi agricoli sono pubblicati dalle Direzioni regionali e in larga parte compaiono già nel portale), Veneto Innovazione (offre una REST API WordPress, ma pubblica soprattutto notizie su bandi regionali già presenti altrove), Veneto Sviluppo e Veneto Lavoro (nessuna sezione bandi di finanziamento strutturata).

## Avvio e aggiornamento semplificati
| File | Funzione |
|---|---|
| `Avvia_Bandi_Veneto.bat` (Windows), `Avvia_Bandi_Veneto.command` (macOS), `avvia.sh` (Linux) | Doppio clic: aggiorna i dati e apre l'app nel browser. |
| `avvia.py` | Il programma di avvio. Opzioni: `--veloce` (~20 s, solo liste e calendari), `--completo` (fino a 30 min, recupera lo storico), `--solo-apri`. |
| `Guida_rapida_Bandi_Veneto.pdf` | Micro-manuale (2 pagine) da condividere. |

**Caricamento automatico dei dati online:** quando l'app è aperta da un indirizzo web (es. GitHub Pages), a ogni visita legge `data/bandi_veneto.json` accanto alla pagina e, se è più recente di quello incorporato, lo usa. Da file locale usa i dati incorporati. Se i dati hanno più di 3 giorni compare un avviso.

## Aggiornare a mano
```bash
pip install -r requirements.txt
python aggiorna_bandi.py                       # giornaliero: liste, calendari, dettagli degli atti aperti, +150 chiusi
python aggiorna_bandi.py --max-details 400 --time-budget 1800   # passata più profonda
python aggiorna_bandi.py --no-refresh-open --max-details 0      # solo liste e calendari (circa 20 secondi)
```
Ogni dettaglio richiede 4-5 secondi: i bandi chiusi più datati si completano in più esecuzioni (150 per volta).

## Aggiornamento periodico (scegli una delle due)
**A) GitHub Actions + GitHub Pages (consigliato)**
1. Crea un repository e carica questa cartella.
2. Settings → Pages → «Deploy from a branch» → branch `main`, cartella `/docs`.
3. Il workflow `aggiorna.yml` gira ogni giorno alle 05:00 UTC. Per l'esecuzione **settimanale** commenta la riga giornaliera e attiva quella con `* * 1`.
4. L'app resta online su `https://<utente>.github.io/<repo>/`; il file `bandi_veneto.html` nel repository è la versione scaricabile.
Nota: non ho potuto verificare che i server GitHub raggiungano i siti regionali; se la prima esecuzione fallisce, lancia lo script dal tuo computer (opzione B).

**B) Sul tuo computer (cron / launchd)**
```cron
# ogni giorno alle 07:00
0 7 * * *  cd /percorso/bandi-veneto-app && /usr/bin/python3 aggiorna_bandi.py >> aggiorna.log 2>&1
# oppure ogni lunedì alle 07:00
0 7 * * 1  cd /percorso/bandi-veneto-app && /usr/bin/python3 aggiorna_bandi.py >> aggiorna.log 2>&1
```

## Excel e tema
- Nell'app, **«Scarica risultati (Excel)»** esporta i risultati filtrati in `.xlsx` (foglio *Risultati* con filtri e intestazione fissa, date e importi come valori reali, link cliccabili; foglio *Info* con data e fonti). Il file è generato nel browser, senza librerie esterne.
- Lo script produce anche l'Excel dell'intero dataset: `data/bandi_veneto.xlsx` (richiede `openpyxl`, già in `requirements.txt`).
- In alto a destra c'è il selettore **Auto / Giorno / Notte**: Auto segue il tema del sistema, la scelta manuale viene ricordata dal browser.

## Aggiornare l'HTML già scaricato senza rigenerarlo
Nell'app, «Carica dati aggiornati (.json)» accetta `data/bandi_veneto.json`.

## Limiti noti
- La classificazione Finanziamenti / Gare e affidamenti usa la categoria del portale; per gli atti senza dettaglio ricorre al titolo (può sbagliare).
- «Previsto» deriva dai calendari: un bando previsto già pubblicato viene rimosso confrontando il codice azione (FESR) o la somiglianza del titolo (FSE+), con possibili errori nei due sensi.
- La ricerca è lessicale con espansione per temi (nessun embedding).
- Il testo del portale non ha valore ufficiale: fa fede il BUR.
