#!/usr/bin/env python3
"""Annullo batch documenti Archimede 2 (B2C Innovation).

Flusso:
  1) Input: file Excel/CSV (colonna CodicePreventivo) oppure scrape della
     pagina lista /RicercaDocumenti/Annullo con filtri data.
  2) Per ogni documento, GET /ricercadocumenti/details?id=<detail_id>&scopo=Annullo
     per ricavare le N verifiche (VerificaPreventivoID, DocumentoID,
     VerificaDescrizione) e la DataAnnullamento.
  3) POST /api/ricercadocumentiAnnullo/validate con tutte le Risultati[N]
     valorizzate.

Dry-run di default; serve --live per eseguire davvero.
"""

from __future__ import annotations

import argparse
import csv
import logging
import os
import random
import re
import sys
import time
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Optional

import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook

API_BASE = "https://backofficeapi.24hassistance.com"
MVC_BASE = "http://bo.24hassistance.com"
TOKEN_TTL = timedelta(minutes=9)
RATE_LIMIT_SECONDS = 0.5
MAX_RETRIES = 3
REQUEST_TIMEOUT = 30

RE_TRAIL_DIGITS = re.compile(r"(\d+)\s*$")


@dataclass
class Verifica:
    verifica_preventivo_id: str
    documento_id: str
    verifica_descrizione: str


@dataclass
class DocumentoInput:
    detail_id: str
    codice_preventivo: str
    codice_polizza: str = ""


@dataclass
class DocumentoCompleto:
    input: DocumentoInput
    data_annullo: str
    verifiche: list[Verifica] = field(default_factory=list)


@dataclass
class RisultatoAnnullo:
    detail_id: str
    codice_preventivo: str
    codice_polizza: str
    n_verifiche: int
    timestamp: str
    esito: str
    http_status: Optional[int] = None
    messaggio: str = ""


class TokenManager:
    def __init__(self, username: str, session: requests.Session) -> None:
        self.username = username
        self.session = session
        self._token: Optional[str] = None
        self._expires_at: datetime = datetime.min.replace(tzinfo=timezone.utc)

    def get(self, force: bool = False) -> str:
        if force or self._token is None or datetime.now(timezone.utc) >= self._expires_at:
            self._refresh()
        assert self._token is not None
        return self._token

    def _refresh(self) -> None:
        logging.info("Richiesta nuovo token JWT (username=%s)", self.username)
        resp = self.session.get(
            f"{API_BASE}/api/token",
            params={"Username": self.username},
            timeout=REQUEST_TIMEOUT,
        )
        resp.raise_for_status()
        token = resp.text.strip()
        if len(token) >= 2 and token[0] == token[-1] == '"':
            token = token[1:-1]
        if not token:
            raise RuntimeError("Token vuoto nella risposta della API")
        self._token = token
        self._expires_at = datetime.now(timezone.utc) + TOKEN_TTL
        logging.debug("Token ottenuto, scade attorno a %s UTC", self._expires_at.isoformat())


def _mask(value: str) -> str:
    if not value:
        return "<VUOTO>"
    if len(value) <= 6:
        return "***"
    return f"{value[:3]}...{value[-3:]} (len={len(value)})"


def _derive_detail_id(codice_preventivo: str) -> str:
    """Da MM25WEB46437525 → 46437525. Prende la coda numerica."""
    m = RE_TRAIL_DIGITS.search(codice_preventivo or "")
    return m.group(1) if m else ""


def _detect_delim(path: Path) -> str:
    with path.open(encoding="utf-8-sig") as f:
        first = f.readline()
    return ";" if first.count(";") > first.count(",") else ","


def leggi_input(path: Path) -> list[DocumentoInput]:
    """Legge Excel/CSV. Colonne: CodicePreventivo (obbligatoria), CodicePolizza (facoltativa)."""
    suffix = path.suffix.lower()
    rows: list[DocumentoInput] = []

    def _append(cod_prev: str, cod_pol: str) -> None:
        cod_prev = (cod_prev or "").strip()
        cod_pol = (cod_pol or "").strip()
        if not cod_prev:
            return
        detail_id = _derive_detail_id(cod_prev)
        if not detail_id:
            logging.warning("Riga ignorata: impossibile ricavare detail_id da %r", cod_prev)
            return
        rows.append(DocumentoInput(detail_id=detail_id, codice_preventivo=cod_prev, codice_polizza=cod_pol))

    if suffix in (".xlsx", ".xlsm"):
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
        header = [str(c).strip() if c is not None else "" for c in header_row]
        if "CodicePreventivo" not in header:
            raise ValueError("Colonna CodicePreventivo mancante nel file Excel")
        i_prev = header.index("CodicePreventivo")
        i_pol = header.index("CodicePolizza") if "CodicePolizza" in header else -1
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row is None or all(v is None for v in row):
                continue
            _append(
                str(row[i_prev]) if row[i_prev] is not None else "",
                str(row[i_pol]) if i_pol >= 0 and row[i_pol] is not None else "",
            )
    elif suffix == ".csv":
        delim = _detect_delim(path)
        with path.open(newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f, delimiter=delim)
            if not reader.fieldnames or "CodicePreventivo" not in reader.fieldnames:
                raise ValueError("Colonna CodicePreventivo mancante nel CSV")
            for r in reader:
                _append(r.get("CodicePreventivo", ""), r.get("CodicePolizza", ""))
    else:
        raise ValueError(f"Formato non supportato: {suffix}")

    return rows


def _check_login_redirect(resp: requests.Response) -> None:
    if resp.status_code in (301, 302, 303, 307):
        loc = resp.headers.get("Location", "")
        if "login" in loc.lower():
            raise RuntimeError(
                f"Sessione MVC non valida: redirect a login ({loc}). "
                "Rigenera i cookie dal browser e aggiorna .env "
                "(ARCH_COOKIE_SESSION, ARCH_COOKIE_ASPXAUTH, ARCH_COOKIE_RVT)."
            )
        raise RuntimeError(f"Redirect inatteso verso {loc}")


def fetch_from_list(
    session: requests.Session,
    cookies: dict[str, str],
    data_da: str,
    data_a: str = "",
    stato: int = 0,
    max_pages: int = 100,
) -> list[DocumentoInput]:
    """Scrape /RicercaDocumenti/Annullo paginando."""
    url = f"{MVC_BASE}/RicercaDocumenti/Annullo"
    seen: set[str] = set()
    docs: list[DocumentoInput] = []

    last_page = 0
    for page in range(max_pages):
        last_page = page
        form = {
            "ControlloACampione": "False",
            "TipoUtente": "0",
            "DataDa": data_da,
            "DataA": data_a,
            "Stato": str(stato),
            "NumeroPagina": str(page),
            "Ordinamento": "ASC",
            "OrdinamentoNomeColonna": "DocumentoCaricato",
        }
        logging.info("Fetch lista pagina %d (DataDa=%s DataA=%s)", page, data_da, data_a or "-")
        if page == 0:
            logging.debug("Cookie inviati: session=%s aspxauth=%s rvt=%s",
                          _mask(cookies.get("ArchimedeMVC_SessionId", "")),
                          _mask(cookies.get(".ASPXAUTH", "")),
                          _mask(cookies.get("__RequestVerificationToken", "")))
        resp = session.post(url, data=form, cookies=cookies, timeout=REQUEST_TIMEOUT,
                            headers={"Content-Type": "application/x-www-form-urlencoded"},
                            allow_redirects=False)
        _check_login_redirect(resp)
        resp.raise_for_status()

        soup = BeautifulSoup(resp.text, "html.parser")
        rows = soup.select("tr.preventivo")
        logging.debug("Pagina %d: %d righe", page, len(rows))
        if not rows:
            break

        new_on_page = 0
        for tr in rows:
            detail_id = (tr.get("id") or "").strip()
            if not detail_id or detail_id in seen:
                continue
            seen.add(detail_id)

            td_prev = tr.select_one("td.codPreventivo")
            if not td_prev:
                logging.warning("Riga %s senza td.codPreventivo, salto", detail_id)
                continue
            cod_prev = td_prev.get_text(strip=True)
            td_pol = td_prev.find_next_sibling("td")
            cod_pol = td_pol.get_text(strip=True) if td_pol else ""

            docs.append(DocumentoInput(detail_id=detail_id, codice_preventivo=cod_prev, codice_polizza=cod_pol))
            new_on_page += 1

        if new_on_page == 0:
            break

    logging.info("Lista: raccolti %d documenti in %d pagine", len(docs), last_page + 1)
    return docs


def fetch_detail(
    session: requests.Session,
    cookies: dict[str, str],
    doc: DocumentoInput,
    data_annullo_override: str = "",
) -> DocumentoCompleto:
    """GET /ricercadocumenti/details?id=<detail_id>&scopo=Annullo → verifiche + DataAnnullamento."""
    url = f"{MVC_BASE}/ricercadocumenti/details"
    params = {"id": doc.detail_id, "scopo": "Annullo"}
    resp = session.get(url, params=params, cookies=cookies, timeout=REQUEST_TIMEOUT, allow_redirects=False)
    _check_login_redirect(resp)
    resp.raise_for_status()

    soup = BeautifulSoup(resp.text, "html.parser")

    # CodicePolizza: <input id="CodicePolizza" type="hidden" value="..."> (può avere spazi in coda)
    cod_pol = doc.codice_polizza
    inp_pol = soup.select_one('input[name="CodicePolizza"][type="hidden"]')
    if inp_pol and inp_pol.get("value"):
        cod_pol = inp_pol["value"].strip()

    # DataAnnullamento: <input id="DataAnnullamento" value="28/02/2026">
    data_annullo = data_annullo_override
    if not data_annullo:
        inp_data = soup.select_one('input[name="DataAnnullamento"]')
        if inp_data and inp_data.get("value"):
            data_annullo = inp_data["value"].strip()
    if not data_annullo:
        data_annullo = datetime.now().strftime("%d/%m/%Y")

    # Verifiche: una per ogni <div class="documento daValidare" data-id=... data-nome=...>
    verifiche: list[Verifica] = []
    for div in soup.select("div.documento.daValidare"):
        vp_id = (div.get("data-id") or "").strip()
        descr = (div.get("data-nome") or "").strip()
        doc_id = ""
        inp_doc = div.select_one('input[name="DocumentoID"][type="hidden"]')
        if inp_doc and inp_doc.get("value"):
            doc_id = inp_doc["value"].strip()
        if not vp_id or not doc_id:
            logging.warning("Dettaglio %s: verifica senza id/documento_id, salto (vp=%r doc=%r)",
                            doc.detail_id, vp_id, doc_id)
            continue
        verifiche.append(Verifica(verifica_preventivo_id=vp_id, documento_id=doc_id, verifica_descrizione=descr))

    doc.codice_polizza = cod_pol
    return DocumentoCompleto(input=doc, data_annullo=data_annullo, verifiche=verifiche)


def _backoff(attempt: int) -> None:
    delay = min(2 ** attempt, 30) + random.uniform(0, 0.5)
    logging.debug("Backoff %.2fs", delay)
    time.sleep(delay)


def annulla_documento(
    session: requests.Session,
    token_mgr: TokenManager,
    completo: DocumentoCompleto,
    operatore: str,
    dry_run: bool,
) -> RisultatoAnnullo:
    d = completo.input
    ts = datetime.now().isoformat(timespec="seconds")
    base = dict(
        detail_id=d.detail_id,
        codice_preventivo=d.codice_preventivo,
        codice_polizza=d.codice_polizza,
        n_verifiche=len(completo.verifiche),
        timestamp=ts,
    )

    if not completo.verifiche:
        logging.warning("Nessuna verifica trovata per %s, skip", d.detail_id)
        return RisultatoAnnullo(esito="skipped", messaggio="no verifiche", **base)

    if dry_run:
        logging.info("[DRY-RUN] annullerei %s (%s) con %d verifiche, DataAnnullo=%s",
                     d.detail_id, d.codice_preventivo, len(completo.verifiche), completo.data_annullo)
        for v in completo.verifiche:
            logging.debug("  - VerificaPreventivoID=%s DocumentoID=%s descr=%s",
                          v.verifica_preventivo_id, v.documento_id, v.verifica_descrizione)
        return RisultatoAnnullo(esito="dry-run", **base)

    payload: dict[str, str] = {
        "CodicePreventivo": d.codice_preventivo,
        "CodicePolizza": d.codice_polizza,
        "IdOperatore": operatore,
        "MessaggioPredefinitoChiave": "1",
        "MessaggioPredefinitoTesto": "",
        "DataAnnullo": completo.data_annullo,
    }
    for i, v in enumerate(completo.verifiche):
        payload[f"Risultati[{i}][VerificaEmissioneID]"] = "0"
        payload[f"Risultati[{i}][VerificaPreventivoID]"] = v.verifica_preventivo_id
        payload[f"Risultati[{i}][DocumentoID]"] = v.documento_id
        payload[f"Risultati[{i}][StatoValidazione]"] = "1"
        payload[f"Risultati[{i}][VerificaDescrizione]"] = v.verifica_descrizione

    for attempt in range(1, MAX_RETRIES + 1):
        headers = {
            "token": token_mgr.get(),
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        }
        try:
            resp = session.post(
                f"{API_BASE}/api/ricercadocumentiAnnullo/validate",
                headers=headers,
                data=payload,
                timeout=REQUEST_TIMEOUT,
            )
        except requests.RequestException as e:
            logging.warning("Errore di rete %s (tentativo %d/%d): %s",
                            d.detail_id, attempt, MAX_RETRIES, e)
            if attempt == MAX_RETRIES:
                return RisultatoAnnullo(esito="error", messaggio=f"network: {e}", **base)
            _backoff(attempt)
            continue

        if resp.status_code == 401 and attempt < MAX_RETRIES:
            logging.info("401 su %s, rigenero token e riprovo", d.detail_id)
            token_mgr.get(force=True)
            continue

        body = resp.text[:500].replace("\n", " ").strip()

        if 500 <= resp.status_code < 600 and attempt < MAX_RETRIES:
            logging.warning("HTTP %d su %s (tentativo %d/%d) body=%s",
                            resp.status_code, d.detail_id, attempt, MAX_RETRIES, body)
            _backoff(attempt)
            continue

        if 200 <= resp.status_code < 300:
            return RisultatoAnnullo(esito="success", http_status=resp.status_code, messaggio=body, **base)

        logging.error("HTTP %d su %s body=%s", resp.status_code, d.detail_id, body)
        return RisultatoAnnullo(esito="error", http_status=resp.status_code, messaggio=body, **base)

    return RisultatoAnnullo(esito="error", messaggio="max retries exceeded", **base)


def _dump_docs(docs: list[DocumentoInput], path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Documenti"
    ws.append(["DetailID", "CodicePreventivo", "CodicePolizza"])
    for d in docs:
        ws.append([d.detail_id, d.codice_preventivo, d.codice_polizza])
    wb.save(path)


def scrivi_risultati(risultati: list[RisultatoAnnullo], path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Esiti"
    ws.append(["DetailID", "CodicePreventivo", "CodicePolizza", "NVerifiche",
               "Timestamp", "Esito", "HTTPStatus", "Messaggio"])
    for r in risultati:
        ws.append([r.detail_id, r.codice_preventivo, r.codice_polizza, r.n_verifiche,
                   r.timestamp, r.esito, r.http_status, r.messaggio])
    wb.save(path)


def setup_logging(log_file: Path, verbose: bool) -> None:
    logging.root.handlers.clear()
    logging.root.setLevel(logging.DEBUG)
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    sh = logging.StreamHandler(sys.stdout)
    sh.setLevel(logging.DEBUG if verbose else logging.INFO)
    sh.setFormatter(logging.Formatter("%(levelname)s %(message)s"))
    logging.root.addHandler(fh)
    logging.root.addHandler(sh)


def main() -> int:
    parser = argparse.ArgumentParser(description="Annullo batch documenti Archimede 2")
    src = parser.add_mutually_exclusive_group(required=True)
    src.add_argument("--input", type=Path, help="Excel/CSV con colonna CodicePreventivo")
    src.add_argument("--data-da", type=str, help="Scrape della lista da questa data (dd/MM/yyyy)")
    parser.add_argument("--data-a", type=str, default="")
    parser.add_argument("--stato", type=int, default=0)
    parser.add_argument("--output", type=Path, default=Path(f"esiti_{datetime.now():%Y%m%d_%H%M%S}.xlsx"))
    parser.add_argument("--log", type=Path, default=Path(f"annullo_{datetime.now():%Y%m%d_%H%M%S}.log"))
    parser.add_argument("--live", action="store_true", help="Esegue la POST validate (default: dry-run)")
    parser.add_argument("--env", type=Path, default=Path(".env"))
    parser.add_argument("--verbose", action="store_true")
    parser.add_argument("--dump-docs", type=Path, help="Salva l'elenco documenti trovati dalla lista in un xlsx")
    args = parser.parse_args()

    setup_logging(args.log, args.verbose)
    load_dotenv(args.env, override=False)

    username = os.environ.get("ARCH_USERNAME")
    operatore = os.environ.get("ARCH_OPERATORE")
    data_annullo_override = os.environ.get("ARCH_DATA_ANNULLO", "").strip()

    if not username or not operatore:
        logging.error("ARCH_USERNAME e ARCH_OPERATORE devono essere impostati in %s", args.env)
        return 2

    cookies = {
        "ArchimedeMVC_SessionId": os.environ.get("ARCH_COOKIE_SESSION", ""),
        ".ASPXAUTH": os.environ.get("ARCH_COOKIE_ASPXAUTH", ""),
        "__RequestVerificationToken": os.environ.get("ARCH_COOKIE_RVT", ""),
    }
    if not cookies["ArchimedeMVC_SessionId"] or not cookies[".ASPXAUTH"]:
        logging.error("Cookie MVC non impostati in %s (ARCH_COOKIE_SESSION, ARCH_COOKIE_ASPXAUTH, ARCH_COOKIE_RVT)",
                      args.env)
        return 2

    session = requests.Session()

    if args.input:
        docs = leggi_input(args.input)
        logging.info("Letti %d documenti da %s", len(docs), args.input)
    else:
        docs = fetch_from_list(session, cookies,
                               data_da=args.data_da, data_a=args.data_a, stato=args.stato)
        if args.dump_docs:
            _dump_docs(docs, args.dump_docs)
            logging.info("Dump lista in %s", args.dump_docs)

    if not docs:
        logging.warning("Nessun documento da processare")
        return 0

    if not args.live:
        logging.warning("Modalità DRY-RUN: nessun annullo verrà eseguito. Usa --live per procedere.")

    token_mgr = TokenManager(username, session)

    risultati: list[RisultatoAnnullo] = []
    try:
        for i, d in enumerate(docs, 1):
            logging.info("[%d/%d] %s (%s)", i, len(docs), d.detail_id, d.codice_preventivo)
            try:
                completo = fetch_detail(session, cookies, d, data_annullo_override)
            except Exception as e:
                logging.error("Fetch dettaglio %s fallito: %s", d.detail_id, e)
                risultati.append(RisultatoAnnullo(
                    detail_id=d.detail_id, codice_preventivo=d.codice_preventivo,
                    codice_polizza=d.codice_polizza, n_verifiche=0,
                    timestamp=datetime.now().isoformat(timespec="seconds"),
                    esito="error", messaggio=f"fetch_detail: {e}",
                ))
                time.sleep(RATE_LIMIT_SECONDS)
                continue
            logging.info("  dettaglio: %d verifiche, DataAnnullo=%s, polizza=%s",
                         len(completo.verifiche), completo.data_annullo, completo.input.codice_polizza)
            r = annulla_documento(session, token_mgr, completo, operatore=operatore, dry_run=not args.live)
            risultati.append(r)
            time.sleep(RATE_LIMIT_SECONDS)
    except KeyboardInterrupt:
        logging.warning("Interrotto dall'utente dopo %d/%d", len(risultati), len(docs))
    finally:
        if risultati:
            scrivi_risultati(risultati, args.output)

    ok = sum(1 for r in risultati if r.esito == "success")
    err = sum(1 for r in risultati if r.esito == "error")
    dry = sum(1 for r in risultati if r.esito == "dry-run")
    skip = sum(1 for r in risultati if r.esito == "skipped")
    logging.info("Fine. success=%d error=%d dry-run=%d skipped=%d — esiti in %s",
                 ok, err, dry, skip, args.output)
    return 0 if err == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
