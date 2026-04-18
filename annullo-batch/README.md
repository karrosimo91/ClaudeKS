# annullo-batch

Script Python per l'annullo massivo di documenti sul gestionale Archimede 2
(B2C Innovation / 24hassistance).

## Prerequisiti

- Python 3.10+
- Rete aziendale (VPN o on-site): i domini `*.24hassistance.com` non
  risolvono da cloud/sandbox.
- Dipendenze: `pip install -r requirements.txt`

## Setup

```bash
cp .env.example .env
# poi apri .env e compila ARCH_USERNAME, ARCH_OPERATORE, ecc.
```

`.env` è coperto dal `.gitignore` del repo. **Non committare credenziali.**

## Formato input

File `.xlsx` o `.csv` con queste colonne (header obbligatori):

| Colonna | Descrizione |
|---|---|
| `VerificaPreventivoID` | ID numerico (l'attributo `id` del `<tr>` in lista) |
| `CodicePreventivo` | Es. `MM25WEB47145199` |
| `CodicePolizza` | Es. `24h.70.361092` |

CSV accettato sia con separatore `,` sia `;`.

## Uso

Due modalità di input, mutuamente esclusive:

### A) Da file Excel/CSV

```bash
# Dry-run (default): nessuna chiamata reale di annullo
python annullo_batch.py --input documenti.xlsx

# Esecuzione vera
python annullo_batch.py --input documenti.xlsx --live
```

### B) Scrape dalla pagina lista `/RicercaDocumenti/Annullo` con filtri data

Richiede i cookie di sessione in `.env` (`ARCH_COOKIE_SESSION`, `ARCH_COOKIE_ASPXAUTH`, `ARCH_COOKIE_RVT`).

```bash
# Dry-run + dump del risultato del scraping per review
python annullo_batch.py --data-da 18/01/2026 --dump-docs trovati.xlsx

# Dry-run + esecuzione vera su finestra data
python annullo_batch.py --data-da 18/01/2026 --data-a 20/01/2026 --live
```

Il parser delle righe usa l'attributo `id` del `<tr class="preventivo">` per `VerificaPreventivoID` e regex per estrarre `CodicePreventivo` (pattern tipo `MM25WEB...`) e `CodicePolizza` (`24h.NN.NNNNNN`) dal testo della riga. Se i codici non vengono trovati la riga viene skippata con un warning — in quel caso manda il log con `--verbose` per capire il formato.

### Opzioni comuni

```bash
python annullo_batch.py ... --live --output esiti.xlsx --log run.log --verbose
```

Output Excel con colonne: `VerificaPreventivoID`, `CodicePreventivo`,
`CodicePolizza`, `Timestamp`, `Esito` (`success`/`error`/`dry-run`),
`HTTPStatus`, `Messaggio`.

## Comportamento

- **Dry-run di default** — serve `--live` per chiamare davvero la API.
- **Token JWT** rigenerato a ogni scadenza (~9 min) o su `401`.
- **Rate limit** ~2 req/s (`RATE_LIMIT_SECONDS = 0.5`).
- **Retry** fino a 3 tentativi su errori di rete / HTTP 5xx con backoff
  esponenziale + jitter.
- **Interruzione** con `Ctrl+C`: gli esiti già raccolti vengono comunque
  salvati.

## Endpoint usati

- `GET  https://backofficeapi.24hassistance.com/api/token?Username=...`
  → restituisce un JWT (stringa) valido 10 min.
- `POST https://backofficeapi.24hassistance.com/api/ricercadocumentiAnnullo/validate`
  header `token: <jwt>`, body `application/x-www-form-urlencoded`.

## Costanti di progetto (hard-coded nel payload)

- `MessaggioPredefinitoChiave = 1`
- `Risultati[0][DocumentoID] = 39`
- `Risultati[0][StatoValidazione] = 1` (OK annulla)
- `Risultati[0][VerificaEmissioneID] = 0`

Se dovessero cambiare per altre tipologie di documento, vanno esternalizzati
in config.

## Note

- I codici nella pagina lista vengono estratti via regex. Se la pagina ha un
  formato diverso (o i pattern `MM##XX...` / `24h.NN.NNNNNN` non matchano),
  serve adattare `RE_CODICE_PREVENTIVO` / `RE_CODICE_POLIZZA` all'inizio dello
  script.
- I cookie del browser hanno scadenza breve: se vedi 302/401 sulla pagina
  lista, rigenera i cookie dal browser e aggiorna `.env`.
