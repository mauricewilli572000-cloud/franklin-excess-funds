"""
Franklin County, OH — Foreclosure Excess Funds (Excess Sales Proceeds) lead scraper.

Source: Franklin County Clerk of Courts publishes the money it is holding
from sheriff sales that sold for more than the judgment ("excess sales
proceeds") at

    https://clerk.franklincountyohio.gov/Public-Records/Public-Financial-Records/Excess-Sales-Proceeds

as a downloadable XLSX ("Print this Report") plus an embedded Power BI
report. The XLSX is the scrapable one, so that's what this reads.

Pipeline (one run):
  1. Load the Excess Sales Proceeds page and find the current XLSX link.
     The link is versioned by the CMS (".../v/1/...", ".../v/2/...") and
     changes when the Clerk uploads a new copy, so it is discovered every
     run instead of hardcoded. Falls back to Playwright if plain requests
     gets blocked.
  2. Parse the XLSX with openpyxl. The header row is found automatically
     and columns are mapped by name (see COLUMN_ALIASES), so small header
     changes don't break the run. Unknown columns are kept under "extra".
  3. Group rows by case number (one lead = one case, amounts summed).
  4. Enrich:
       a. Cross-reference the Franklin foreclosure scraper's own records
          (same case number) for property / mailing address + parcel.
       b. Franklin County Auditor parcel CSV: exact parcel match if the
          XLSX or the foreclosure record gave a parcel number, otherwise a
          *unique* owner-name match for a mailing address (flagged
          "Auditor name match - verify", since after a sheriff sale the
          former owner is usually no longer owner of record).
  5. Score 0-100, track first-seen / last-seen across runs, and mark
     cases that drop off the Clerk's list as "no longer listed" (usually
     means the money was claimed or disbursed).
  6. Write dashboard/excess_funds.json, dashboard/excess_funds_ghl_export.csv
     and data/history.json.

Usage:
    python scraper/excess_funds.py                  # full run
    python scraper/excess_funds.py --skip-parcels   # skip the big Auditor CSV
    python scraper/excess_funds.py --xlsx path.xlsx # parse a local copy
"""
from __future__ import annotations

import argparse
import csv
import io
import json
import logging
import os
import re
import sys
from dataclasses import dataclass, field, asdict
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Optional
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
import openpyxl

# --------------------------------------------------------------------------
# Config
# --------------------------------------------------------------------------

ROOT = Path(__file__).resolve().parent.parent
DASHBOARD_DIR = ROOT / "dashboard"
DATA_DIR = ROOT / "data"
CACHE_DIR = ROOT / ".cache"

FEED_PATH = DASHBOARD_DIR / "excess_funds.json"
GHL_CSV_PATH = DASHBOARD_DIR / "excess_funds_ghl_export.csv"
HISTORY_PATH = DATA_DIR / "history.json"
DEBUG_PATH = DATA_DIR / "last_run_debug.json"

SOURCE_PAGE_URL = os.environ.get(
    "EXCESS_PAGE_URL",
    "https://clerk.franklincountyohio.gov/Public-Records/Public-Financial-Records/Excess-Sales-Proceeds",
)
# Known-good link as of 2026-11 — only used if link discovery on the page fails.
FALLBACK_XLSX_URL = (
    "https://clerk.franklincountyohio.gov/files/assets/clerkofcourt/v/1/documents/"
    "public-financial-records/excesssalesproceeds.xlsx"
)
XLSX_URL_OVERRIDE = os.environ.get("EXCESS_XLSX_URL") or None

CASE_SEARCH_URL = "https://fcdcfcjs.co.franklin.oh.us/CaseInformationOnline/"

# The Franklin foreclosure scraper's records (same GitHub account). Matching
# case numbers there gives us property + mailing address for free.
FORECLOSURE_RECORD_URLS = [
    u.strip() for u in os.environ.get(
        "FORECLOSURE_RECORD_URLS",
        "https://raw.githubusercontent.com/mauricewilli572000-cloud/franklin-foreclosure-scraper/main/data/records.json,"
        "https://raw.githubusercontent.com/mauricewilli572000-cloud/franklin-foreclosure-scraper/main/data/pending_records.json",
    ).split(",") if u.strip()
]

PARCEL_CSV_URL = os.environ.get("PARCEL_CSV_URL") or None
PARCEL_CSV_INDEX_URL = "https://apps.franklincountyauditor.com/Parcel_CSV/"

REQUEST_TIMEOUT = 60
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)

# Safety valve: if a run parses far fewer cases than were active last run,
# something is wrong with the download/parse — don't mark everything as
# "no longer listed". Set EXCESS_ALLOW_SHRINK=1 to override once.
MIN_ACTIVE_RATIO = 0.3

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("excess_funds")

# --------------------------------------------------------------------------
# Column mapping
# --------------------------------------------------------------------------
# Header text is normalized (lowercase, non-alphanumerics -> single space)
# before matching. First alias that matches wins. Anything not mapped is
# kept on the record under "extra" so nothing in the Clerk's file is lost.
COLUMN_ALIASES: dict[str, tuple[str, ...]] = {
    "case_number": ("case number", "case no", "case num", "case", "case id", "case nbr"),
    "caption": ("case caption", "caption", "case name", "case title", "case style", "style",
                "description", "case description"),
    "plaintiff": ("plaintiff", "plaintiff name", "plaintiffs"),
    "defendant": ("defendant", "defendant name", "defendants", "debtor", "debtor name",
                  "owner", "owner name", "party", "party name", "payee", "name",
                  "former owner", "titleholder"),
    "amount": ("amount", "balance", "excess amount", "excess funds", "excess proceeds",
               "excess sales proceeds", "amount on deposit", "deposit amount",
               "current balance", "balance on deposit", "total", "amount held", "funds"),
    "deposit_date": ("deposit date", "date deposited", "date received", "receipt date",
                     "received date", "date", "deposited", "date of deposit",
                     "posting date", "posted date", "transaction date"),
    "sale_date": ("sale date", "sheriff sale date", "date of sale", "sold date"),
    "parcel_number": ("parcel", "parcel number", "parcel id", "parcel no", "ppn"),
    "property_address": ("property address", "address", "property", "site address",
                         "location"),
    "judge": ("judge", "assigned judge"),
}

HEADER_SCAN_ROWS = 25  # how far down the sheet to look for the header row


def norm_header(text: Any) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(text or "").lower()).strip()


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------

ENTITY_SUFFIXES = (
    " llc", " l.l.c", " inc", " inc.", " incorporated", " corp", " corp.",
    " corporation", " ltd", " ltd.", " lp", " l.p.", " co.", " company",
    " trust", " trustee", " bank", " na", " n.a.", " association", " fsb",
    " mortgage", " financial", " servicing", " lending", " credit union",
    " holdings", " properties", " investments", " partners",
)
GOVERNMENT_WORDS = (
    "treasurer", "state of ohio", "city of", "county of", "franklin county",
    "united states", "department of", "board of", "village of", "township",
    "internal revenue", "attorney general", "secretary of", "housing and urban",
)
ESTATE_RE = re.compile(
    r"\b(ESTATE OF|UNKNOWN HEIRS|HEIRS|DEVISEES|DECEASED|DEC'D|EXECUT(?:OR|RIX)|"
    r"ADMINISTRAT(?:OR|RIX)|FIDUCIARY|SURVIVING SPOUSE)\b",
    re.I,
)
VS_SPLIT_RE = re.compile(r"\s+(?:-\s*)?(?:VS?\.?|VERSUS)(?:\s*-)?\s+", re.I)
ETAL_RE = re.compile(r",?\s*\b(ET\s*AL\.?|ET\s*UX\.?|ET\s*VIR\.?|AND\s+ALL\s+OTHERS)\b.*$", re.I)


def clean_ws(s: Any) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip()


def is_entity(name: str) -> bool:
    low = f" {(name or '').lower().strip()} "
    return any(suf + " " in low or low.rstrip().endswith(suf) for suf in ENTITY_SUFFIXES)


def is_government(name: str) -> bool:
    low = (name or "").lower()
    return any(w in low for w in GOVERNMENT_WORDS)


def normalize_case_number(raw: Any) -> str:
    """'17CV009397', '17 CV 9397', '2017-CV-009397', '17CV-9397' -> '17CV009397'."""
    s = clean_ws(raw).upper()
    if not s:
        return ""
    m = re.search(r"\b(?:20)?(\d{2})\s*[- ]?\s*([A-Z]{2})\s*[- ]?\s*(\d{1,6})\b", s)
    if m:
        return f"{m.group(1)}{m.group(2)}{int(m.group(3)):06d}"
    return re.sub(r"[^A-Z0-9]", "", s)


def parse_money(v: Any) -> Optional[float]:
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return round(float(v), 2)
    s = str(v).strip()
    neg = s.startswith("(") and s.endswith(")")
    s = re.sub(r"[^\d.\-]", "", s)
    if not s or s in ("-", "."):
        return None
    try:
        val = float(s)
    except ValueError:
        return None
    return round(-val if neg else val, 2)


def parse_date(v: Any) -> str:
    if v is None or v == "":
        return ""
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, (int, float)) and 20000 < float(v) < 80000:  # Excel serial date
        return (date(1899, 12, 30) + timedelta(days=int(v))).isoformat()
    s = clean_ws(v)
    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%m-%d-%Y", "%b %d, %Y", "%B %d, %Y",
                "%Y-%m-%d %H:%M:%S", "%m/%d/%Y %I:%M:%S %p", "%m/%d/%Y %H:%M"):
        try:
            return datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            continue
    return s


def strip_party_noise(name: str) -> str:
    name = ETAL_RE.sub("", clean_ws(name))
    name = re.sub(r"\b(A\.?K\.?A\.?|F\.?K\.?A\.?|N\.?K\.?A\.?)\b.*$", "", name, flags=re.I)
    name = re.sub(r"^(UNKNOWN HEIRS|THE UNKNOWN HEIRS)[^A-Z]*(OF\s+)?", "", name, flags=re.I)
    name = re.sub(r"^(THE\s+)?ESTATE\s+OF\s+", "", name, flags=re.I)
    name = re.sub(r",?\s*(DECEASED|DEC'D)\b.*$", "", name, flags=re.I)
    return name.strip(" ,;-.")


def split_caption(caption: str) -> tuple[str, str]:
    """'US BANK NA VS JOHN DOE ET AL' -> ('US BANK NA', 'JOHN DOE ET AL')."""
    caption = clean_ws(caption)
    if not caption:
        return "", ""
    parts = VS_SPLIT_RE.split(caption, maxsplit=1)
    if len(parts) == 2:
        return parts[0].strip(" ,-"), parts[1].strip(" ,-")
    return "", caption


# --------------------------------------------------------------------------
# Download
# --------------------------------------------------------------------------

def make_session() -> requests.Session:
    s = requests.Session()
    s.headers.update({
        "User-Agent": USER_AGENT,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
    })
    return s


def find_xlsx_link(html: str, base_url: str) -> Optional[str]:
    soup = BeautifulSoup(html, "lxml")
    candidates = []
    for a in soup.find_all("a", href=True):
        href = a["href"]
        text = a.get_text(" ", strip=True).lower()
        if re.search(r"\.xlsx?(\?|$)", href, re.I):
            score = 0
            if "excess" in href.lower():
                score += 2
            if "report" in text or "excess" in text:
                score += 1
            candidates.append((score, urljoin(base_url, href)))
    if not candidates:
        return None
    candidates.sort(key=lambda t: -t[0])
    return candidates[0][1]


def looks_like_xlsx(content: bytes) -> bool:
    return content[:2] == b"PK"  # xlsx is a zip container


def download_with_requests(session: requests.Session) -> tuple[bytes, str]:
    xlsx_url = XLSX_URL_OVERRIDE
    if not xlsx_url:
        try:
            resp = session.get(SOURCE_PAGE_URL, timeout=REQUEST_TIMEOUT)
            resp.raise_for_status()
            xlsx_url = find_xlsx_link(resp.text, SOURCE_PAGE_URL)
            if xlsx_url:
                log.info("Found XLSX link on page: %s", xlsx_url)
            else:
                log.warning("No .xlsx link found on %s — using fallback URL", SOURCE_PAGE_URL)
        except requests.RequestException as exc:
            log.warning("Could not load source page (%s) — using fallback URL", exc)
        xlsx_url = xlsx_url or FALLBACK_XLSX_URL

    resp = session.get(xlsx_url, timeout=REQUEST_TIMEOUT, headers={"Referer": SOURCE_PAGE_URL})
    resp.raise_for_status()
    if not looks_like_xlsx(resp.content):
        raise RuntimeError(
            f"Response from {xlsx_url} is not an XLSX file "
            f"(content-type {resp.headers.get('content-type')!r}, {len(resp.content)} bytes) — "
            "probably a bot-protection page."
        )
    return resp.content, xlsx_url


def download_with_playwright() -> tuple[bytes, str]:
    """Fallback for when the site blocks plain HTTP clients: load the page
    in real Chromium and click the report link."""
    from playwright.sync_api import sync_playwright  # imported lazily: optional dependency

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        ctx = browser.new_context(user_agent=USER_AGENT, accept_downloads=True)
        page = ctx.new_page()
        page.goto(SOURCE_PAGE_URL, wait_until="domcontentloaded", timeout=90_000)
        link = page.locator("a[href*='.xlsx']").first
        href = link.get_attribute("href") or FALLBACK_XLSX_URL
        url = urljoin(SOURCE_PAGE_URL, href)
        # Try fetching through the browser context first (keeps cookies)…
        resp = ctx.request.get(url, headers={"Referer": SOURCE_PAGE_URL})
        body = resp.body()
        if not looks_like_xlsx(body):
            # …otherwise click it and capture the download.
            with page.expect_download(timeout=90_000) as dl_info:
                link.click()
            path = dl_info.value.path()
            body = Path(path).read_bytes()
        browser.close()
    if not looks_like_xlsx(body):
        raise RuntimeError("Playwright fallback also did not return an XLSX file.")
    return body, url


def download_xlsx() -> tuple[bytes, str]:
    session = make_session()
    try:
        return download_with_requests(session)
    except Exception as exc:  # noqa: BLE001
        log.warning("Plain download failed (%s); trying Playwright fallback", exc)
    return download_with_playwright()


# --------------------------------------------------------------------------
# Parse
# --------------------------------------------------------------------------

@dataclass
class SheetParse:
    sheet: str
    header_row: int
    headers: list[str]
    mapping: dict[str, int]          # field -> column index
    rows: list[dict]


def map_headers(headers: list[str]) -> dict[str, int]:
    normed = [norm_header(h) for h in headers]
    mapping: dict[str, int] = {}
    used: set[int] = set()
    # Pass 1: exact alias matches. Pass 2: alias contained in header.
    for exact in (True, False):
        for fld, aliases in COLUMN_ALIASES.items():
            if fld in mapping:
                continue
            for alias in aliases:
                hit = None
                for i, h in enumerate(normed):
                    if i in used or not h:
                        continue
                    if (h == alias) if exact else (re.search(rf"\b{re.escape(alias)}\b", h)):
                        hit = i
                        break
                if hit is not None:
                    mapping[fld] = hit
                    used.add(hit)
                    break
    return mapping


def parse_xlsx(content: bytes) -> SheetParse:
    wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
    best: Optional[SheetParse] = None
    for ws in wb.worksheets:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        if not rows:
            continue
        # Header row = the row in the first N whose cells map to the most fields
        # (must include a case number and an amount to count).
        header_idx, header_map = None, {}
        for i, row in enumerate(rows[:HEADER_SCAN_ROWS]):
            heads = [clean_ws(c) for c in row]
            m = map_headers(heads)
            if "case_number" in m and len(m) > len(header_map):
                header_idx, header_map = i, m
        if header_idx is None:
            continue
        headers = [clean_ws(c) for c in rows[header_idx]]
        data_rows = []
        for raw in rows[header_idx + 1:]:
            if raw is None or all(c in (None, "") for c in raw):
                continue
            rec = {"_cells": {headers[j] or f"col{j}": raw[j] for j in range(min(len(headers), len(raw)))}}
            for fld, j in header_map.items():
                rec[fld] = raw[j] if j < len(raw) else None
            data_rows.append(rec)
        parse = SheetParse(ws.title, header_idx + 1, headers, header_map, data_rows)
        if best is None or len(parse.rows) > len(best.rows):
            best = parse
    if best is None:
        raise RuntimeError(
            "Could not find a header row with a case-number column in any sheet. "
            "Run scraper/discover_excess_xlsx.py and update COLUMN_ALIASES."
        )
    log.info("Parsed sheet %r: header on row %d, %d data rows, mapped columns: %s",
             best.sheet, best.header_row, len(best.rows),
             {k: best.headers[v] for k, v in best.mapping.items()})
    unmapped = [h for i, h in enumerate(best.headers) if h and i not in best.mapping.values()]
    if unmapped:
        log.info("Unmapped columns (kept under 'extra'): %s", unmapped)
    return best


# --------------------------------------------------------------------------
# Lead model
# --------------------------------------------------------------------------

@dataclass
class ExcessLead:
    case_number: str
    caption: str = ""
    plaintiff: str = ""
    defendants_raw: str = ""
    owner: str = ""                      # best guess at the former owner / claimant
    amount: Optional[float] = None       # total excess on deposit for the case
    deposit_date: str = ""
    sale_date: str = ""
    row_count: int = 0
    parcel_number: str = ""
    prop_address: str = ""
    prop_city: str = ""
    prop_state: str = ""
    prop_zip: str = ""
    mail_address: str = ""
    mail_city: str = ""
    mail_state: str = ""
    mail_zip: str = ""
    address_source: str = ""             # where the address came from
    owner_type: str = ""                 # individual | entity | government | estate
    tax_foreclosure: bool = False
    flags: list[str] = field(default_factory=list)
    score: int = 0
    score_reasons: list[str] = field(default_factory=list)
    first_seen: str = ""
    last_seen: str = ""
    status: str = "active"               # active | no longer listed
    removed_date: str = ""
    amount_history: list[dict] = field(default_factory=list)
    extra: dict = field(default_factory=dict)
    source_url: str = SOURCE_PAGE_URL
    case_search_url: str = CASE_SEARCH_URL


def pick_owner(defendants: str) -> str:
    """First named defendant, minus 'et al', aka, 'unknown heirs of' etc.
    Skips government / lender co-defendants when a person is also named."""
    if not defendants:
        return ""
    pieces = re.split(r"\s*[;\n]\s*", defendants)
    names = [strip_party_noise(p) for p in pieces]
    names = [n for n in names if n and len(n) > 2]
    if not names:
        return strip_party_noise(defendants)
    for n in names:
        if not is_government(n) and not (is_entity(n) and any(not is_entity(x) for x in names)):
            return n
    return names[0]


def cell_str(v: Any) -> str:
    if isinstance(v, (datetime, date)):
        return parse_date(v)
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return clean_ws(v)


def build_leads(parse: SheetParse) -> dict[str, ExcessLead]:
    leads: dict[str, ExcessLead] = {}
    mapped_headers = {parse.headers[i] for i in parse.mapping.values()}
    skipped = 0
    for row in parse.rows:
        case = normalize_case_number(row.get("case_number"))
        if not case or len(case) < 6 or not re.search(r"\d", case):
            skipped += 1
            continue  # subtotal / footer / blank rows
        caption = cell_str(row.get("caption"))
        plaintiff = cell_str(row.get("plaintiff"))
        defendant = cell_str(row.get("defendant"))
        if caption and not (plaintiff and defendant):
            p2, d2 = split_caption(caption)
            plaintiff = plaintiff or p2
            defendant = defendant or d2

        lead = leads.get(case)
        if lead is None:
            lead = ExcessLead(case_number=case)
            leads[case] = lead
        lead.row_count += 1
        lead.caption = lead.caption or caption
        lead.plaintiff = lead.plaintiff or plaintiff
        if defendant and defendant not in lead.defendants_raw:
            lead.defendants_raw = "; ".join(x for x in (lead.defendants_raw, defendant) if x)

        amt = parse_money(row.get("amount"))
        if amt is not None:
            lead.amount = round((lead.amount or 0) + amt, 2)

        dd = parse_date(row.get("deposit_date"))
        if dd and (not lead.deposit_date or dd < lead.deposit_date):
            lead.deposit_date = dd
        sd = parse_date(row.get("sale_date"))
        lead.sale_date = lead.sale_date or sd
        lead.parcel_number = lead.parcel_number or cell_str(row.get("parcel_number"))
        addr = cell_str(row.get("property_address"))
        if addr and not lead.prop_address:
            lead.prop_address, lead.address_source = addr, "Clerk excess funds report"

        for k, v in row["_cells"].items():
            if k in mapped_headers or v in (None, ""):
                continue
            lead.extra.setdefault(k, cell_str(v))

    for lead in leads.values():
        lead.owner = pick_owner(lead.defendants_raw)
        all_parties = f"{lead.caption} {lead.defendants_raw}"
        if ESTATE_RE.search(all_parties):
            lead.owner_type = "estate"
        elif is_government(lead.owner):
            lead.owner_type = "government"
        elif is_entity(lead.owner):
            lead.owner_type = "entity"
        elif lead.owner:
            lead.owner_type = "individual"
        lead.tax_foreclosure = bool(re.search(r"treasurer", lead.plaintiff or lead.caption, re.I))

    log.info("Built %d case leads from %d rows (%d non-case rows skipped)",
             len(leads), len(parse.rows), skipped)
    return leads


# --------------------------------------------------------------------------
# Enrichment: foreclosure records (same case number)
# --------------------------------------------------------------------------

def load_foreclosure_index(session: requests.Session) -> dict[str, dict]:
    index: dict[str, dict] = {}
    for url in FORECLOSURE_RECORD_URLS:
        try:
            resp = session.get(url, timeout=REQUEST_TIMEOUT)
            resp.raise_for_status()
            data = resp.json()
        except Exception as exc:  # noqa: BLE001
            log.warning("Could not load foreclosure records from %s: %s", url, exc)
            continue
        recs = data.get("records", data) if isinstance(data, dict) else data
        if isinstance(recs, dict):
            recs = list(recs.values())
        for r in recs or []:
            if not isinstance(r, dict):
                continue
            cn = normalize_case_number(r.get("case_number"))
            if cn and (cn not in index or (r.get("prop_address") and not index[cn].get("prop_address"))):
                index[cn] = r
    log.info("Loaded %d foreclosure records keyed by case number", len(index))
    return index


def enrich_from_foreclosures(leads: dict[str, ExcessLead], fidx: dict[str, dict]) -> int:
    hits = 0
    for lead in leads.values():
        r = fidx.get(lead.case_number)
        if not r:
            continue
        hits += 1
        lead.parcel_number = lead.parcel_number or (r.get("parcel_number") or "")
        if not lead.prop_address and r.get("prop_address"):
            lead.prop_address = r.get("prop_address", "")
            lead.prop_city = r.get("prop_city", "")
            lead.prop_state = r.get("prop_state", "") or "OH"
            lead.prop_zip = r.get("prop_zip", "")
            lead.address_source = "Franklin foreclosure scraper (same case)"
        if not lead.mail_address and r.get("mail_address"):
            lead.mail_address = r.get("mail_address", "")
            lead.mail_city = r.get("mail_city", "")
            lead.mail_state = r.get("mail_state", "")
            lead.mail_zip = r.get("mail_zip", "")
        if not lead.owner and r.get("owner"):
            lead.owner = r["owner"]
        if "In foreclosure scraper" not in lead.flags:
            lead.flags.append("In foreclosure scraper")
    log.info("Foreclosure cross-reference matched %d cases", hits)
    return hits


# --------------------------------------------------------------------------
# Enrichment: Auditor parcel CSV
# --------------------------------------------------------------------------
# Same source and column names as the Franklin foreclosure scraper
# (confirmed there against a real 2026-08 Parcel.csv).
OWNER_COLUMNS = ("NAME1", "OWNER", "OWN1", "OWNERNAME", "OWNER_NAME", "OWNER1")
SITE_ADDR_COLUMNS = ("STADDR", "SITE_ADDR", "SITEADDR", "SITE_ADDRE", "PROPADDR", "SITEADDRESS")
SITE_CITY_COLUMNS = ("USPS_CITY", "SITE_CITY", "SITECITY", "PROPCITY", "SITECITYNAME")
SITE_ZIP_COLUMNS = ("ZIPCODE", "SITE_ZIP", "SITEZIP", "PROPZIP", "SITEZIPCODE")
MAIL_STREET_COLUMNS = ("MAILAD3", "OWNER_ADD1", "ADDR_1", "MAILADR1", "MAIL_ADDR1", "MAILADDR")
MAIL_CSZ_COLUMNS = ("MAILAD4", "OWNER_ADD2")
PARCEL_ID_COLUMNS = ("PARCEL ID", "PARCELID", "PARCEL_ID", "PARCELNO", "PARCEL NO")


def get_field(row: dict, candidates: tuple, upper_keys: dict[str, str]) -> str:
    for c in candidates:
        k = upper_keys.get(c.upper())
        if k is not None:
            v = row.get(k)
            if v not in (None, ""):
                return str(v).strip()
    return ""


def split_csz(text: str) -> tuple[str, str, str]:
    text = (text or "").strip()
    m = re.match(r"^(.*?)\s+([A-Z]{2})\s+(\d{5}(?:-\d{4})?)$", text)
    return (m.group(1).strip(), m.group(2), m.group(3)) if m else (text, "", "")


def digits(s: str) -> str:
    return re.sub(r"\D", "", s or "")


def _tokens(raw: str) -> tuple[list[str], bool]:
    n = re.sub(r"[^A-Z, ]", " ", (raw or "").upper())
    n = re.sub(r"\s+", " ", n).strip(" ,")
    n = re.sub(r"\b(JR|SR|II|III|IV)\b", "", n).strip(" ,")
    if "," in n:
        last, _, first = n.partition(",")
        return [t for t in [last.strip()] + first.split() if t], True
    return n.split(), False


def index_keys(owner: str) -> list[str]:
    """Keys for an Auditor NAME1 value, which is already 'LAST FIRST M'."""
    toks, _ = _tokens(owner)
    keys = [" ".join(toks)]
    if len(toks) > 2:
        keys.append(" ".join(toks[:2]))
    return [k for k in dict.fromkeys(keys) if k]


def lead_keys(owner: str) -> list[str]:
    """Keys for a court-caption name, most specific first. Captions are usually
    'FIRST M LAST' (or 'LAST, FIRST M'); convert to Auditor 'LAST FIRST M' order."""
    toks, had_comma = _tokens(owner)
    if not toks:
        return []
    aud = toks if had_comma or len(toks) < 2 else [toks[-1]] + toks[:-1]
    keys = [" ".join(aud), " ".join(toks), " ".join(aud[:2])]
    return [k for k in dict.fromkeys(keys) if k]


def name_keys(raw: str) -> set[str]:
    return set(index_keys(raw)) | set(lead_keys(raw))


@dataclass
class ParcelIndex:
    by_parcel: dict[str, dict]
    by_name: dict[str, Optional[dict]]   # None = ambiguous (2+ different mailing addresses)


def _list_dir(session: requests.Session, url: str) -> list[str]:
    resp = session.get(url, timeout=REQUEST_TIMEOUT)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "lxml")
    out = []
    for a in soup.find_all("a", href=True):
        name = a["href"].strip("/").split("/")[-1]
        if name and name not in ("..", "."):
            out.append(name)
    return out


def find_parcel_csv_url(session: requests.Session) -> str:
    if PARCEL_CSV_URL:
        return PARCEL_CSV_URL
    years = sorted((e for e in _list_dir(session, PARCEL_CSV_INDEX_URL) if re.fullmatch(r"\d{4}", e)),
                   reverse=True)
    for year in years[:2]:
        yurl = f"{PARCEL_CSV_INDEX_URL}{year}/"
        try:
            months = sorted((e for e in _list_dir(session, yurl) if re.fullmatch(r"\d{1,2}", e)),
                            key=int, reverse=True)
        except requests.RequestException:
            continue
        for month in months[:6]:
            murl = f"{yurl}{month}/"
            try:
                for name in _list_dir(session, murl):
                    if name.lower() == "parcel.csv":
                        return f"{murl}{name}"
            except requests.RequestException:
                continue
    raise RuntimeError(f"No Parcel.csv found under {PARCEL_CSV_INDEX_URL}; set PARCEL_CSV_URL.")


def load_parcel_index(session: requests.Session) -> Optional[ParcelIndex]:
    try:
        url = find_parcel_csv_url(session)
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        dest = CACHE_DIR / "Parcel.csv"
        log.info("Downloading Auditor parcel CSV: %s", url)
        with session.get(url, timeout=REQUEST_TIMEOUT * 6, stream=True) as resp:
            resp.raise_for_status()
            with dest.open("wb") as f:
                for chunk in resp.iter_content(1024 * 1024):
                    f.write(chunk)
        return build_parcel_index(dest)
    except Exception as exc:  # noqa: BLE001
        log.error("Auditor parcel data unavailable, continuing without it: %s", exc)
        return None


def build_parcel_index(path: Path) -> ParcelIndex:
    by_parcel: dict[str, dict] = {}
    by_name: dict[str, Optional[dict]] = {}
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        upper = {h.upper().strip(): h for h in (reader.fieldnames or [])}
        for row in reader:
            owner = get_field(row, OWNER_COLUMNS, upper)
            if not owner:
                continue
            mc, ms, mz = split_csz(get_field(row, MAIL_CSZ_COLUMNS, upper))
            rec = {
                "owner": owner,
                "prop_address": get_field(row, SITE_ADDR_COLUMNS, upper),
                "prop_city": get_field(row, SITE_CITY_COLUMNS, upper),
                "prop_zip": get_field(row, SITE_ZIP_COLUMNS, upper),
                "mail_address": get_field(row, MAIL_STREET_COLUMNS, upper),
                "mail_city": mc, "mail_state": ms or "OH", "mail_zip": mz,
            }
            pid = digits(get_field(row, PARCEL_ID_COLUMNS, upper))
            if pid:
                by_parcel[pid] = rec
            for k in index_keys(owner):
                prev = by_name.get(k, "missing")
                if prev == "missing":
                    by_name[k] = rec
                elif prev is not None and prev["mail_address"] != rec["mail_address"]:
                    by_name[k] = None  # ambiguous — never guess between two people
    log.info("Parcel index: %d parcel IDs, %d name keys", len(by_parcel), len(by_name))
    return ParcelIndex(by_parcel, by_name)


def enrich_from_parcels(leads: dict[str, ExcessLead], pidx: ParcelIndex) -> dict[str, int]:
    stats = {"parcel": 0, "name": 0}
    for lead in leads.values():
        if lead.parcel_number:
            pid = digits(lead.parcel_number)
            hit = pidx.by_parcel.get(pid) or (pidx.by_parcel.get(pid[:-2]) if pid.endswith("00") else None)
            if hit:
                stats["parcel"] += 1
                if not lead.prop_address:
                    lead.prop_address = hit["prop_address"]
                    lead.prop_city = hit["prop_city"]
                    lead.prop_state = "OH"
                    lead.prop_zip = hit["prop_zip"]
                    lead.address_source = lead.address_source or "Auditor parcel (parcel number)"
                # Mailing address from the parcel only if the Auditor still lists
                # our former owner — otherwise it's the sheriff-sale buyer's address.
                if not lead.mail_address and name_keys(hit["owner"]) & name_keys(lead.owner):
                    lead.mail_address, lead.mail_city = hit["mail_address"], hit["mail_city"]
                    lead.mail_state, lead.mail_zip = hit["mail_state"], hit["mail_zip"]
                elif not name_keys(hit["owner"]) & name_keys(lead.owner):
                    lead.extra.setdefault("Current owner of record (Auditor)", hit["owner"])
        if lead.mail_address or lead.owner_type not in ("individual", "estate") or not lead.owner:
            continue
        hit = None
        for k in lead_keys(lead.owner):       # most specific key first
            v = pidx.by_name.get(k, "missing")
            if v == "missing":
                continue
            hit = v                            # None = ambiguous: stop, don't go looser
            break
        if hit:
            stats["name"] += 1
            lead.mail_address, lead.mail_city = hit["mail_address"], hit["mail_city"]
            lead.mail_state, lead.mail_zip = hit["mail_state"], hit["mail_zip"]
            lead.extra.setdefault("Auditor name match parcel", hit["prop_address"])
            if "Auditor name match - verify" not in lead.flags:
                lead.flags.append("Auditor name match - verify")
            if not lead.prop_address:
                lead.address_source = "Auditor name match (another parcel they own)"
    log.info("Parcel enrichment: %d by parcel number, %d by unique owner name",
             stats["parcel"], stats["name"])
    return stats


# --------------------------------------------------------------------------
# Scoring
# --------------------------------------------------------------------------

def score_lead(lead: ExcessLead, today: date) -> None:
    pts, why = 0, []
    a = lead.amount or 0
    for threshold, p in ((50_000, 40), (25_000, 34), (10_000, 27), (5_000, 20), (2_500, 14), (1_000, 8)):
        if a >= threshold:
            pts += p
            why.append(f"${a:,.0f} on deposit (+{p})")
            break
    else:
        pts += 2
        why.append(f"Small balance ${a:,.0f} (+2)")

    if lead.owner_type == "individual":
        pts += 18; why.append("Individual former owner (+18)")
    elif lead.owner_type == "estate":
        pts += 22; why.append("Estate / heirs — owner likely unaware (+22)")
    elif lead.owner_type == "entity":
        pts += 5; why.append("LLC / company owner (+5)")
    elif lead.owner_type == "government":
        why.append("Government party (+0)")

    if lead.mail_address:
        pts += 15; why.append("Mailing address found (+15)")
    elif lead.prop_address:
        pts += 8; why.append("Property address found (+8)")

    if lead.deposit_date:
        try:
            age = (today - date.fromisoformat(lead.deposit_date)).days
            if age >= 365:
                pts += 12; why.append(f"Unclaimed {age // 30} months (+12)")
            elif age >= 120:
                pts += 7; why.append(f"Unclaimed {age // 30} months (+7)")
        except ValueError:
            pass
    if lead.tax_foreclosure:
        pts += 5; why.append("Tax foreclosure — owner often unrepresented (+5)")
    if lead.first_seen and (today - date.fromisoformat(lead.first_seen)).days <= 7:
        pts += 5; why.append("New on the list this week (+5)")
    if lead.status != "active":
        pts = min(pts, 20); why.append("No longer listed — likely claimed (capped at 20)")

    lead.score = max(0, min(100, pts))
    lead.score_reasons = why


def build_flags(lead: ExcessLead, today: date) -> None:
    keep = [f for f in lead.flags if f in ("In foreclosure scraper", "Auditor name match - verify")]
    f = []
    a = lead.amount or 0
    if a >= 25_000:
        f.append("$25k+ surplus")
    elif a >= 10_000:
        f.append("$10k+ surplus")
    f.append({"individual": "Individual owner", "estate": "Estate / heirs", "entity": "LLC / corp owner",
              "government": "Government party"}.get(lead.owner_type, "Owner unknown"))
    if lead.tax_foreclosure:
        f.append("Tax foreclosure")
    if lead.deposit_date:
        try:
            if (today - date.fromisoformat(lead.deposit_date)).days >= 365:
                f.append("Unclaimed 1yr+")
        except ValueError:
            pass
    if lead.mail_state and lead.mail_state != "OH":
        f.append("Out-of-state owner")
    if lead.first_seen and (today - date.fromisoformat(lead.first_seen)).days <= 7:
        f.append("New this week")
    if lead.status != "active":
        f.append("No longer listed")
    lead.flags = f + keep


# --------------------------------------------------------------------------
# History / outputs
# --------------------------------------------------------------------------

def load_history() -> dict[str, dict]:
    if HISTORY_PATH.exists():
        try:
            return json.loads(HISTORY_PATH.read_text(encoding="utf-8")).get("cases", {})
        except Exception as exc:  # noqa: BLE001
            log.warning("history.json unreadable (%s) — starting fresh", exc)
    return {}


def merge_history(current: dict[str, ExcessLead], history: dict[str, dict], today: str) -> list[ExcessLead]:
    merged: list[ExcessLead] = []
    known = set(ExcessLead.__dataclass_fields__)
    for cn, lead in current.items():
        prev = history.get(cn)
        lead.first_seen = (prev or {}).get("first_seen") or today
        lead.last_seen = today
        lead.status = "active"
        hist = list((prev or {}).get("amount_history") or [])
        if not hist or hist[-1].get("amount") != lead.amount:
            hist.append({"date": today, "amount": lead.amount})
        lead.amount_history = hist[-12:]
        # Keep an address found on an earlier run if this run's enrichment missed it
        if prev:
            for k in ("prop_address", "prop_city", "prop_state", "prop_zip", "mail_address",
                      "mail_city", "mail_state", "mail_zip", "parcel_number", "address_source"):
                if not getattr(lead, k) and prev.get(k):
                    setattr(lead, k, prev[k])
        merged.append(lead)
    for cn, prev in history.items():
        if cn in current:
            continue
        old = ExcessLead(**{k: v for k, v in prev.items() if k in known})
        if old.status == "active":
            old.status, old.removed_date = "no longer listed", today
        merged.append(old)
    return merged


def split_name(owner: str) -> tuple[str, str]:
    owner = (owner or "").strip()
    if not owner or is_entity(owner) or is_government(owner):
        return "", owner
    if "," in owner:
        last, _, first = owner.partition(",")
        return first.strip().title(), last.strip().title()
    parts = owner.split()
    if len(parts) == 1:
        return "", parts[0].title()
    return " ".join(p.title() for p in parts[:-1]), parts[-1].title()


# Same headers as the Franklin / Cuyahoga GHL exports so the same GHL import
# mapping works. "Amount/Debt Owed" carries the SURPLUS amount here, and
# "Date Filed" carries the date the excess was deposited.
GHL_HEADERS = [
    "First Name", "Last Name", "Mailing Address", "Mailing City", "Mailing State",
    "Mailing Zip", "Property Address", "Property City", "Property State", "Property Zip",
    "Lead Type", "Document Type", "Date Filed", "Document Number", "Amount/Debt Owed",
    "Seller Score", "Motivated Seller Flags", "Source", "Public Records URL",
]


def write_outputs(all_leads: list[ExcessLead], xlsx_url: str, debug: dict) -> dict:
    DASHBOARD_DIR.mkdir(parents=True, exist_ok=True)
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    now = datetime.now(timezone.utc).replace(microsecond=0).isoformat()

    active = sorted([l for l in all_leads if l.status == "active"], key=lambda l: (-l.score, -(l.amount or 0)))
    gone = [l for l in all_leads if l.status != "active"]
    today = date.today()
    new_count = sum(1 for l in active if l.first_seen == today.isoformat())

    feed = {
        "generated_at": now,
        "source": "Franklin County Clerk of Courts — Excess Sales Proceeds",
        "source_page": SOURCE_PAGE_URL,
        "source_file": xlsx_url,
        "total_leads": len(active),
        "total_surplus": round(sum(l.amount or 0 for l in active), 2),
        "leads_with_mailing_address": sum(1 for l in active if l.mail_address),
        "leads_with_property_address": sum(1 for l in active if l.prop_address),
        "new_this_run": new_count,
        "no_longer_listed": len(gone),
        "leads": [asdict(l) for l in active],
        "recently_removed": [asdict(l) for l in sorted(gone, key=lambda l: l.removed_date, reverse=True)[:100]],
    }
    FEED_PATH.write_text(json.dumps(feed, indent=2, default=str), encoding="utf-8")

    HISTORY_PATH.write_text(json.dumps(
        {"updated": now, "cases": {l.case_number: asdict(l) for l in all_leads}},
        indent=2, default=str), encoding="utf-8")

    with GHL_CSV_PATH.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(GHL_HEADERS)
        for l in active:
            if l.owner_type == "government":
                continue
            first, last = split_name(l.owner)
            w.writerow([
                first, last, l.mail_address, l.mail_city, l.mail_state, l.mail_zip,
                l.prop_address, l.prop_city, l.prop_state or ("OH" if l.prop_address else ""), l.prop_zip,
                "Excess Funds", "Excess Sales Proceeds", l.deposit_date, l.case_number,
                f"{l.amount:.2f}" if l.amount is not None else "",
                l.score, "; ".join(l.flags), "Franklin County Clerk of Courts", SOURCE_PAGE_URL,
            ])

    debug["output"] = {k: feed[k] for k in ("total_leads", "total_surplus", "leads_with_mailing_address",
                                             "leads_with_property_address", "new_this_run", "no_longer_listed")}
    DEBUG_PATH.write_text(json.dumps(debug, indent=2, default=str), encoding="utf-8")
    log.info("Wrote %s (%d active, $%s total, %d new, %d no longer listed)",
             FEED_PATH.name, len(active), f"{feed['total_surplus']:,.2f}", new_count, len(gone))
    return feed


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def run(xlsx_path: Optional[str], skip_parcels: bool, skip_foreclosures: bool) -> dict:
    today = date.today()
    debug: dict = {"started": datetime.now(timezone.utc).isoformat()}

    if xlsx_path:
        content, xlsx_url = Path(xlsx_path).read_bytes(), f"file:{xlsx_path}"
    else:
        content, xlsx_url = download_xlsx()
    log.info("Got XLSX (%d bytes) from %s", len(content), xlsx_url)
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    (CACHE_DIR / "excesssalesproceeds.xlsx").write_bytes(content)

    parse = parse_xlsx(content)
    debug["sheet"] = {"name": parse.sheet, "header_row": parse.header_row, "headers": parse.headers,
                      "mapping": {k: parse.headers[v] for k, v in parse.mapping.items()},
                      "rows": len(parse.rows)}
    leads = build_leads(parse)
    if not leads:
        raise RuntimeError("Parsed 0 cases from the XLSX — refusing to overwrite data.")

    history = load_history()
    prev_active = sum(1 for h in history.values() if h.get("status") == "active")
    if prev_active and len(leads) < prev_active * MIN_ACTIVE_RATIO and not os.environ.get("EXCESS_ALLOW_SHRINK"):
        raise RuntimeError(
            f"Only {len(leads)} cases parsed vs {prev_active} active last run — looks like a bad "
            "download/parse. Not updating. Set EXCESS_ALLOW_SHRINK=1 to accept."
        )

    session = make_session()
    if not skip_foreclosures:
        enrich_from_foreclosures(leads, load_foreclosure_index(session))

    # Carry addresses from history before Auditor matching so we don't redo work
    for cn, lead in leads.items():
        prev = history.get(cn) or {}
        if not lead.mail_address and prev.get("mail_address"):
            for k in ("mail_address", "mail_city", "mail_state", "mail_zip"):
                setattr(lead, k, prev.get(k, ""))

    if not skip_parcels:
        pidx = load_parcel_index(session)
        if pidx:
            debug["parcel_matches"] = enrich_from_parcels(leads, pidx)

    merged = merge_history(leads, history, today.isoformat())
    for lead in merged:
        build_flags(lead, today)
        score_lead(lead, today)
    return write_outputs(merged, xlsx_url, debug)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--xlsx", help="Parse a local XLSX instead of downloading")
    ap.add_argument("--skip-parcels", action="store_true", help="Skip Auditor parcel CSV enrichment")
    ap.add_argument("--skip-foreclosures", action="store_true", help="Skip foreclosure-records cross-reference")
    args = ap.parse_args()
    try:
        run(args.xlsx, args.skip_parcels, args.skip_foreclosures)
    except Exception as exc:  # noqa: BLE001
        log.error("Run failed: %s", exc)
        sys.exit(1)


if __name__ == "__main__":
    main()
