# Franklin County, OH — Excess Funds (Excess Sales Proceeds) Lead Scraper

Daily pipeline that pulls the Franklin County Clerk of Courts **Excess Sales
Proceeds** report (money left over after a sheriff sale sold for more than
was owed), turns it into one lead per case, finds an address where it can,
scores each lead 0-100, and publishes:

- `dashboard/index.html` — standalone dashboard (GitHub Pages)
- `dashboard/excess_funds.json` — the feed (also read by the combined dashboard's **Franklin Excess Funds** tab)
- `dashboard/excess_funds_ghl_export.csv` — GoHighLevel import, same columns as the Franklin/Cuyahoga exports

Source page: https://clerk.franklincountyohio.gov/Public-Records/Public-Financial-Records/Excess-Sales-Proceeds

## Layout

```
scraper/excess_funds.py            # main scraper
scraper/discover_excess_xlsx.py    # one-time local check of the real XLSX columns
scraper/requirements.txt
dashboard/index.html               # standalone dashboard
dashboard/excess_funds.json        # feed (written each run)
dashboard/excess_funds_ghl_export.csv
data/history.json                  # every case ever seen: first/last seen, status, amount history
data/last_run_debug.json           # which sheet/columns were used, match counts
.github/workflows/scrape-excess.yml
combined-dashboard-update/index.html   # drop-in for franklin-foreclosure-scraper/dashboard/index.html
```

## How it works

1. **Download.** Loads the Excess Sales Proceeds page and finds the "Print this
   Report (XLSX)" link each run (the CMS puts a version number in the URL —
   `/v/1/`, `/v/2/` — that changes when the Clerk uploads a new file). If plain
   HTTP is blocked it retries in headless Chromium.
2. **Parse.** Finds the header row automatically and maps columns by name
   (`COLUMN_ALIASES` in `excess_funds.py`). Columns it doesn't recognize are kept
   and shown in the dashboard's detail view. Rows are grouped by case number and
   amounts summed. A caption like `US BANK NA VS JOHN DOE ET AL` is split into
   plaintiff and former owner ("et al", "aka", "unknown heirs of", "deceased" are
   stripped from the name).
3. **Enrich.**
   - **Same case in the foreclosure scraper** → property + mailing address and parcel.
   - **Auditor Parcel.csv** → exact parcel match if a parcel number is known.
     The mailing address is only taken from the parcel if the Auditor still
     lists the former owner (after a sheriff sale it's usually the buyer).
     Otherwise a **unique** owner-name match gives a mailing address, flagged
     `Auditor name match - verify`. If a name matches two or more different
     people, nothing is guessed.
4. **Score.** Amount on deposit (up to +40), individual (+18) or estate/heirs
   (+22) owner, address found (+15), unclaimed 4+ months / 1+ year (+7/+12),
   tax foreclosure (+5), new this week (+5). The dashboard shows the reasons.
5. **Track.** `data/history.json` remembers every case. A case that disappears
   from the Clerk's list is marked **no longer listed** (usually claimed or paid
   out), drops out of the feed and CSV, and appears under `recently_removed`.
   **Safety check:** if a run parses less than 30% of last run's active cases,
   it stops without changing anything (set `EXCESS_ALLOW_SHRINK=1` to accept it once).

Government parties (Treasurer, State of Ohio, etc.) are kept on the dashboard
but left out of the GHL CSV. In the CSV, **Amount/Debt Owed = surplus held** and
**Date Filed = deposit date**.

## First run checklist

1. Create a new **public** GitHub repo named `franklin-excess-funds` and push
   this folder to it (PowerShell steps below).
2. Repo **Settings → Pages → Source: GitHub Actions**.
3. **Actions → Franklin County Excess Funds Scrape → Run workflow.**
4. Confirm the run log line `Parsed sheet ... mapped columns: {...}` shows
   case number, amount and a name/caption column. If anything is missing, run
   `discover_excess_xlsx.py` locally and send back
   `discover_excess_xlsx_output.json`.
5. Copy `combined-dashboard-update/index.html` over
   `franklin-foreclosure-scraper/dashboard/index.html` to add the
   **Franklin Excess Funds** tab to the combined dashboard.

```powershell
# From the unzipped franklin-excess-funds folder
git init
git add .
git commit -m "Franklin County excess funds scraper"
git branch -M main
git remote add origin https://github.com/mauricewilli572000-cloud/franklin-excess-funds.git
git push -u origin main
```

## Running locally

```bash
pip install -r scraper/requirements.txt
python scraper/excess_funds.py                    # full run
python scraper/excess_funds.py --skip-parcels     # skip the large Auditor CSV
python scraper/excess_funds.py --xlsx file.xlsx   # parse a file you downloaded
python scraper/discover_excess_xlsx.py            # check real column layout
```

Environment overrides: `EXCESS_XLSX_URL` (direct file link), `PARCEL_CSV_URL`,
`FORECLOSURE_RECORD_URLS` (comma-separated JSON URLs).

## Notes

- Claims go through a motion and a hearing before a Magistrate — see
  https://clerk.franklincountyohio.gov/Resources/Foreclosure-Excess-Fund-Process.
  Ohio has rules on agreements to recover excess funds for someone else; check
  them before you contact owners about a fee arrangement.
- If the Clerk ever drops the XLSX and only keeps the Power BI report, the
  download step fails loudly and nothing is overwritten.
