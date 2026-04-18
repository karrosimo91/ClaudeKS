# annullo-batch

Annullo massivo di documenti sul gestionale Archimede 2 (B2C Innovation / 24hassistance).

## Prerequisiti

- Python 3.10+
- Rete aziendale (VPN o on-site): i domini `*.24hassistance.com` non risolvono da cloud/sandbox.
- `pip install -r requirements.txt`

## Setup

```bash
cp .env.example .env
```

Poi apri `.env` e compila tutti i campi, **inclusi i cookie di sessione** (servono sempre, sia per la lista che per il dettaglio). I cookie si prendono dal browser: `F12` → Application → Cookies su `http://bo.24hassistance.com`, copiare i valori di:

- `ArchimedeMVC_SessionId` → `ARCH_COOKIE_SESSION`
- `.ASPXAUTH` → `ARCH_COOKIE_ASPXAUTH`
- `__RequestVerificationToken` → `ARCH_COOKIE_RVT`

I cookie scadono rapidamente: se vedi "Sessione MVC non valida", rigenerali.

## Come funziona

Per ciascun documento, lo script:

1. **Fetch dettaglio**: `GET /ricercadocumenti/details?id=<detail_id>&scopo=Annullo` per estrarre:
   - `CodicePolizza` (hidden input)
   - `DataAnnullamento` (pre-compilata per il documento)
   - Tutte le verifiche (`VerificaPreventivoID`, `DocumentoID`, `VerificaDescrizione`) — possono essere più di una per documento.
2. **POST validate**: `/api/ricercadocumentiAnnullo/validate` con `Risultati[0..N-1]` valorizzate.

Il `detail_id` corrisponde all'`id` del `<tr>` in lista e alla coda numerica del `CodicePreventivo` (es. `MM25WEB46437525` → `46437525`).

## Modalità di input

### A) Scrape lista con filtro data

```bash
# Dry-run: solo scraping + dettaglio, nessun annullo
python annullo_batch.py --data-da 17/04/2026 --dump-docs trovati.xlsx --verbose

# Esecuzione vera su tutta la finestra
python annullo_batch.py --data-da 17/04/2026 --data-a 20/04/2026 --live
```

### B) Da file Excel/CSV

Il file deve avere almeno la colonna `CodicePreventivo` (es. `MM25WEB46437525`). `CodicePolizza` è opzionale (se assente viene ricavata dal dettaglio).

```bash
python annullo_batch.py --input documenti.xlsx --live
```

## Opzioni

- `--live` — esegue davvero (default dry-run)
- `--output esiti.xlsx` — file di output con gli esiti
- `--log run.log` — file di log dettagliato
- `--verbose` — stampa DEBUG su stdout
- `--stato 0` — filtro stato per la lista (0=DaVerificare)

## Output

Excel con: `DetailID`, `CodicePreventivo`, `CodicePolizza`, `NVerifiche`, `Timestamp`, `Esito` (`success`/`error`/`dry-run`/`skipped`), `HTTPStatus`, `Messaggio`.

## Note

- **Rate limit** `~2 req/s`, **retry** 3x con backoff esponenziale su network/5xx.
- **Token JWT** rigenerato a scadenza (9 min) o su 401.
- **Ctrl+C**: salva gli esiti parziali e termina.
