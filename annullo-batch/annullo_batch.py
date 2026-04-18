#!/usr/bin/env python3
"""Annullo batch documenti Archimede 2 (B2C Innovation).

Legge un file Excel/CSV con i documenti da annullare e chiama
l'endpoint di validate sulla API di backoffice, con gestione token,
rate limiting, retry e dry-run di default.
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
from dataclasses import dataclass
from datetime import datetime, timedelta
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

REQUIRED_COLS = ("VerificaPreventivoID", "CodicePreventivo", "CodicePolizza")

RE_CODICE_PREVENTIVO = re.compile(r"\b([A-Z]{2}\d{2}[A-Z]{2,}\d+)\b")
RE_CODICE_POLIZZA = re.compile(r"\b(24h\.\d+\.\d+)\b")


@dataclass
class DocumentoInput:
    verifica_preventivo_id: str
    codice_preventivo: str
    codice_polizza: str


@dataclass
class RisultatoAnnullo:
    verifica_preventivo_id: str
    codice_preventivo: str
    codice_polizza: str
    timestamp: str
    esito: str
    http_status: Optional[int] = None
    messaggio: str = ""


class TokenManager:
    def __init__(self, username: str, session: requests.Session) -> None:
        self.username = username
        self.session = session
        self._token: Optional[str] = None
        self._expires_at: datetime = datetime.min

    def get(self, force: bool = False) -> str:
        if force or self._token is None or datetime.utcnow() >= self._expires_at:
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
        self._expires_at = datetime.utcnow() + TOKEN_TTL
        logging.debug("Token ottenuto, scade attorno a %s UTC", self._expires_at.isoformat())


def _detect_delim(path: Path) -> str:
    with path.open(encoding="utf-8-sig") as f:
        first = f.readline()
    return ";" if first.count(";") > first.count(",") else ","


def leggi_input(path: Path) -> list[DocumentoInput]:
    suffix = path.suffix.lower()
    rows: list[DocumentoInput] = []

    if suffix in (".xlsx", ".xlsm"):
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
        header = [str(c).strip() if c is not None else "" for c in header_row]
        missing = [c for c in REQUIRED_COLS if c not in header]
        if missing:
            raise ValueError(f"Colonne mancanti nel file Excel: {missing}")
        idx = {c: header.index(c) for c in REQUIRED_COLS}
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row is None or all(v is None for v in row):
                continue
            rows.append(DocumentoInput(
                verifica_preventivo_id=str(row[idx["VerificaPreventivoID"]]).strip(),
                codice_preventivo=str(row[idx["CodicePreventivo"]]).strip(),
                codice_polizza=str(row[idx["CodicePolizza"]]).strip(),
            ))
    elif suffix == ".csv":
        delim = _detect_delim(path)
        with path.open(newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f, delimiter=delim)
            missing = [c for c in REQUIRED_COLS if c not in (reader.fieldnames or [])]
            if missing:
                raise ValueError(f"Colonne mancanti nel CSV: {missing}")
            for r in reader:
                rows.append(DocumentoInput(
                    verifica_preventivo_id=(r["VerificaPreventivoID"] or "").strip(),
                    codice_preventivo=(r["CodicePreventivo"] or "").strip(),
                    codice_polizza=(r["CodicePolizza"] or "").strip(),
                ))
    else:
        raise ValueError(f"Formato non supportato: {suffix}")

    return [r for r in rows if r.verifica_preventivo_id]


def fetch_from_list(
    session: requests.Session,
    cookies: dict[str, str],
    data_da: str,
    data_a: str = "",
    stato: int = 0,
    max_pages: int = 100,
) -> list[DocumentoInput]:
    """Scrape /RicercaDocumenti/Annullo paginando finché non trova righe nuove."""
    url = f"{MVC_BASE}/RicercaDocumenti/Annullo"
    seen: set[str] = set()
    docs: list[DocumentoInput] = []

    for page in range(max_pages):
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
        resp = session.post(url, data=form, cookies=cookies, timeout=REQUEST_TIMEOUT,
                            headers={"Content-Type": "application/x-www-form-urlencoded"})
        resp.raise_for_status()

        soup = BeautifulSoup(resp.text, "html.parser")
        rows = soup.select("tr.preventivo")
        logging.debug("Pagina %d: trovate %d righe", page, len(rows))
        if not rows:
            break

        new_on_page = 0
        for tr in rows:
            vp_id = (tr.get("id") or "").strip()
            if not vp_id or vp_id in seen:
                continue
            seen.add(vp_id)
            new_on_page += 1

            text = tr.get_text(" ", strip=True)
            m_prev = RE_CODICE_PREVENTIVO.search(text)
            m_pol = RE_CODICE_POLIZZA.search(text)
            if not m_prev or not m_pol:
                logging.warning(
                    "Riga %s: codici non trovati nel testo (prev=%s pol=%s). Testo: %s",
                    vp_id, bool(m_prev), bool(m_pol), text[:200],
                )
                continue
            docs.append(DocumentoInput(
                verifica_preventivo_id=vp_id,
                codice_preventivo=m_prev.group(1),
                codice_polizza=m_pol.group(1),
            ))

        if new_on_page == 0:
            break

    logging.info("Lista: raccolti %d documenti in %d pagine", len(docs), page + 1)
    return docs


def _backoff(attempt: int) -> None:
    delay = min(2 ** attempt, 30) + random.uniform(0, 0.5)
    logging.debug("Backoff %.2fs", delay)
    time.sleep(delay)


def annulla_documento(
    session: requests.Session,
    token_mgr: TokenManager,
    doc: DocumentoInput,
    operatore: str,
    data_annullo: str,
    verifica_descrizione: str,
    dry_run: bool,
) -> RisultatoAnnullo:
    ts = datetime.now().isoformat(timespec="seconds")
    base = dict(
        verifica_preventivo_id=doc.verifica_preventivo_id,
        codice_preventivo=doc.codice_preventivo,
        codice_polizza=doc.codice_polizza,
        timestamp=ts,
    )

    if dry_run:
        logging.info("[DRY-RUN] annullerei %s (%s)", doc.verifica_preventivo_id, doc.codice_preventivo)
        return RisultatoAnnullo(esito="dry-run", **base)

    payload = {
        "CodicePreventivo": doc.codice_preventivo,
        "CodicePolizza": doc.codice_polizza,
        "IdOperatore": operatore,
        "MessaggioPredefinitoChiave": "1",
        "MessaggioPredefinitoTesto": "",
        "Risultati[0][VerificaEmissioneID]": "0",
        "Risultati[0][VerificaPreventivoID]": doc.verifica_preventivo_id,
        "Risultati[0][DocumentoID]": "39",
        "Risultati[0][StatoValidazione]": "1",
        "Risultati[0][VerificaDescrizione]": verifica_descrizione,
        "DataAnnullo": data_annullo,
    }

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
                            doc.verifica_preventivo_id, attempt, MAX_RETRIES, e)
            if attempt == MAX_RETRIES:
                return RisultatoAnnullo(esito="error", messaggio=f"network: {e}", **base)
            _backoff(attempt)
            continue

        if resp.status_code == 401 and attempt < MAX_RETRIES:
            logging.info("401 su %s, rigenero token e riprovo", doc.verifica_preventivo_id)
            token_mgr.get(force=True)
            continue

        if 500 <= resp.status_code < 600 and attempt < MAX_RETRIES:
            logging.warning("HTTP %d su %s (tentativo %d/%d)",
                            resp.status_code, doc.verifica_preventivo_id, attempt, MAX_RETRIES)
            _backoff(attempt)
            continue

        body = resp.text[:500].replace("\n", " ").strip()
        esito = "success" if 200 <= resp.status_code < 300 else "error"
        return RisultatoAnnullo(esito=esito, http_status=resp.status_code, messaggio=body, **base)

    return RisultatoAnnullo(esito="error", messaggio="max retries exceeded", **base)


def _dump_docs(docs: list[DocumentoInput], path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Documenti"
    ws.append(list(REQUIRED_COLS))
    for d in docs:
        ws.append([d.verifica_preventivo_id, d.codice_preventivo, d.codice_polizza])
    wb.save(path)


def scrivi_risultati(risultati: list[RisultatoAnnullo], path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Esiti"
    ws.append(["VerificaPreventivoID", "CodicePreventivo", "CodicePolizza",
               "Timestamp", "Esito", "HTTPStatus", "Messaggio"])
    for r in risultati:
        ws.append([r.verifica_preventivo_id, r.codice_preventivo, r.codice_polizza,
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
    src.add_argument("--input", type=Path, help="File Excel (.xlsx) o CSV con documenti")
    src.add_argument("--data-da", type=str,
                     help="Prende i documenti dalla pagina lista filtrando da questa data (dd/MM/yyyy)")
    parser.add_argument("--data-a", type=str, default="",
                        help="Data A opzionale (dd/MM/yyyy)")
    parser.add_argument("--stato", type=int, default=0,
                        help="Filtro stato per la ricerca (default 0)")
    parser.add_argument("--output", type=Path,
                        default=Path(f"esiti_{datetime.now():%Y%m%d_%H%M%S}.xlsx"))
    parser.add_argument("--log", type=Path,
                        default=Path(f"annullo_{datetime.now():%Y%m%d_%H%M%S}.log"))
    parser.add_argument("--live", action="store_true",
                        help="Esegue realmente la chiamata (default: dry-run)")
    parser.add_argument("--env", type=Path, default=Path(".env"))
    parser.add_argument("--verbose", action="store_true")
    parser.add_argument("--dump-docs", type=Path,
                        help="Se indicato, scrive in questo xlsx i documenti trovati dalla lista (utile per review prima di --live)")
    args = parser.parse_args()

    setup_logging(args.log, args.verbose)
    load_dotenv(args.env, override=False)

    username = os.environ.get("ARCH_USERNAME")
    operatore = os.environ.get("ARCH_OPERATORE")
    data_annullo = os.environ.get("ARCH_DATA_ANNULLO") or datetime.now().strftime("%d/%m/%Y")
    verifica_descrizione = os.environ.get("ARCH_VERIFICA_DESCRIZIONE",
                                           "Copia della denuncia di furto")

    if not username or not operatore:
        logging.error("ARCH_USERNAME e ARCH_OPERATORE devono essere impostati nell'env (%s)", args.env)
        return 2

    session = requests.Session()

    if args.input:
        docs = leggi_input(args.input)
        logging.info("Letti %d documenti da %s", len(docs), args.input)
    else:
        cookies = {
            "ArchimedeMVC_SessionId": os.environ.get("ARCH_COOKIE_SESSION", ""),
            ".ASPXAUTH": os.environ.get("ARCH_COOKIE_ASPXAUTH", ""),
            "__RequestVerificationToken": os.environ.get("ARCH_COOKIE_RVT", ""),
        }
        if not cookies["ArchimedeMVC_SessionId"] or not cookies[".ASPXAUTH"]:
            logging.error("Per --data-da servono i cookie in env: "
                          "ARCH_COOKIE_SESSION, ARCH_COOKIE_ASPXAUTH, ARCH_COOKIE_RVT")
            return 2
        docs = fetch_from_list(
            session, cookies,
            data_da=args.data_da,
            data_a=args.data_a,
            stato=args.stato,
        )
        if args.dump_docs:
            _dump_docs(docs, args.dump_docs)
            logging.info("Documenti dumpati in %s", args.dump_docs)

    if not docs:
        logging.warning("Nessun documento da processare")
        return 0

    if not args.live:
        logging.warning("Modalità DRY-RUN: nessun annullo verrà eseguito. Usa --live per procedere.")

    token_mgr = TokenManager(username, session)

    risultati: list[RisultatoAnnullo] = []
    try:
        for i, doc in enumerate(docs, 1):
            logging.info("[%d/%d] %s", i, len(docs), doc.verifica_preventivo_id)
            r = annulla_documento(
                session, token_mgr, doc,
                operatore=operatore,
                data_annullo=data_annullo,
                verifica_descrizione=verifica_descrizione,
                dry_run=not args.live,
            )
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
    logging.info("Fine. success=%d error=%d dry-run=%d — esiti in %s",
                 ok, err, dry, args.output)
    return 0 if err == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
