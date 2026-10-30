# LinkedIn Profile Finder — Pampanga IT Alumni (Selenium, No API Keys)

Automates finding LinkedIn profiles for the names in **`list of names (1).xlsx`**.

Uses a real **Chrome browser** (via Selenium) to search Google — no API keys, no sign-ups, runs 100% locally.

Filters results by:

| Filter | Criteria |
|--------|----------|
| **School** | Don Honorio Ventura State University (DHVSU), Don Honorio, Pampanga State University |
| **OR Location** | Pampanga (any city or municipality) |
| **AND Field** | IT-related — Computer Science (CSS), Information Systems (IS), Information Technology (IT), developer, programmer, software engineer, etc. |

Matching profile URLs are written back to **column C** of the same Excel file.

---

## 1. Prerequisites

- Python 3.11+
- **Google Chrome** installed on your computer (any recent version)

That's it. The correct ChromeDriver is downloaded automatically.

---

## 2. Installation

```powershell
cd "C:\Users\ulap5\Downloads\auto linkedin"
pip install -r requirements.txt
```

---

## 3. Usage

### Run on all rows (browser runs hidden in background)
```powershell
python linkedin_finder.py
```

### Show the Chrome window while it runs
```powershell
python linkedin_finder.py --visible
```

### Dry run — search and log, but don't write to Excel
```powershell
python linkedin_finder.py --dry-run
```

### Resume from a specific row
```powershell
python linkedin_finder.py --start-row 50
```

### Test a single name
```powershell
python linkedin_finder.py --name "Reyes, John Vincent V."
python linkedin_finder.py --name "Reyes, John Vincent V." --visible
```

---

## 4. Output

- **Excel file** — Column C (`linkedin profile`) is updated in-place after every match found.
- **linkedin_finder.log** — Full log of every search query, match result, and skip.

---

## 5. Resume after stopping

Just run the script again — it automatically skips any row that already has a link:

```powershell
python linkedin_finder.py
```

Or check the log to see the last processed row:
```powershell
Get-Content linkedin_finder.log -Tail 20
```

---

## 6. How it works

```
For each name in column A (skips rows that already have a URL in column C):

  1. Convert "Lastname, Firstname M." → "Firstname Lastname"
  2. Open Google in Chrome and search 3 progressively broader queries:
       a. site:linkedin.com/in + name + DHVSU/Don Honorio/Pampanga State + IT field
       b. site:linkedin.com/in + name + Pampanga + developer/IT/software
       c. site:linkedin.com/in + name + Philippines
  3. For each LinkedIn URL in results:
       - Check title + snippet against school/location + IT keyword filters
       - First match that passes both filters → save URL to column C
  4. Wait 4–8 seconds between searches (avoids Google blocks)
  5. If Google shows a CAPTCHA → auto-pause 60s and retry
```

---

## 7. CAPTCHA handling

Google may occasionally show a CAPTCHA after many searches. The script:
- Detects it automatically
- Pauses 60 seconds and retries

If it keeps getting blocked, run with `--visible` to watch the browser, or increase the delay by editing `MIN_DELAY` / `MAX_DELAY` at the top of `linkedin_finder.py`.

---

## 8. Filter keywords (customisable)

Edit the three lists near the top of `linkedin_finder.py`:

- **`SCHOOL_KEYWORDS`** — `don honorio ventura`, `dhvsu`, `pampanga state university`, etc.
- **`LOCATION_KEYWORDS`** — `pampanga`, `angeles city`, `san fernando`, `clark`, + 13 more municipalities
- **`IT_KEYWORDS`** — `css`, `bsit`, `bsis`, `information technology`, `developer`, `programmer`, etc.
