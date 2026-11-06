"""
Discovery helper — run this LOCALLY (Windows PowerShell is fine) once, then
send back the output file so the column mapping can be confirmed against the
real Clerk file.

    cd scraper
    pip install requests beautifulsoup4 lxml openpyxl
    python discover_excess_xlsx.py

It writes discover_excess_xlsx_output.json next to this script containing:
  - the XLSX link found on the Excess Sales Proceeds page
  - every sheet's first 15 rows (raw)
  - which header row / column mapping excess_funds.py would pick
  - the first 5 parsed leads
It also saves the downloaded file as excesssalesproceeds_sample.xlsx.
"""
import json
import sys
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import excess_funds as ef  # noqa: E402


def main() -> None:
    out: dict = {}
    session = ef.make_session()
    try:
        resp = session.get(ef.SOURCE_PAGE_URL, timeout=60)
        out["page_status"] = resp.status_code
        out["xlsx_link_found"] = ef.find_xlsx_link(resp.text, ef.SOURCE_PAGE_URL)
        out["all_links_with_xlsx_or_powerbi"] = [
            a for a in __import__("re").findall(r'href="([^"]+)"', resp.text)
            if ".xls" in a.lower() or "powerbi" in a.lower()
        ]
    except Exception as exc:  # noqa: BLE001
        out["page_error"] = repr(exc)

    content, url = ef.download_xlsx()
    out["downloaded_from"] = url
    out["bytes"] = len(content)
    (HERE / "excesssalesproceeds_sample.xlsx").write_bytes(content)

    import io
    wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
    out["sheets"] = {}
    for ws in wb.worksheets:
        rows = []
        for i, r in enumerate(ws.iter_rows(values_only=True)):
            if i >= 15:
                break
            rows.append([str(c) if c is not None else None for c in r])
        out["sheets"][ws.title] = {"first_rows": rows}

    try:
        parse = ef.parse_xlsx(content)
        out["chosen"] = {
            "sheet": parse.sheet, "header_row": parse.header_row, "headers": parse.headers,
            "mapping": {k: parse.headers[v] for k, v in parse.mapping.items()},
            "data_rows": len(parse.rows),
        }
        leads = ef.build_leads(parse)
        out["cases"] = len(leads)
        out["sample_leads"] = [ef.asdict(l) for l in list(leads.values())[:5]]
    except Exception as exc:  # noqa: BLE001
        out["parse_error"] = repr(exc)

    dest = HERE / "discover_excess_xlsx_output.json"
    dest.write_text(json.dumps(out, indent=2, default=str), encoding="utf-8")
    print(f"Done. Send back: {dest}")
    print(json.dumps(out.get("chosen", out.get("parse_error")), indent=2, default=str))


if __name__ == "__main__":
    main()
