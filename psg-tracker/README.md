# G2 PSG Tracker

A Google Sheets duty-roster and accountability tracker for a G2 Platoon Sergeant. It keeps five
DA Form 6 rosters per AR 220-45, tracks availability, appointments, equipment and taskings, and
puts it all on one dashboard, with a one-slide brief for senior leaders.

It has three parts:

| File | What it is |
|---|---|
| `build_tracker.py` | Generates the workbook (`.xlsx`). Formulas are written in Google Sheets dialect and come alive once the file is opened with Google Sheets. |
| `Code.gs` | Apps Script automation layer (menu, auto-fill, BDE-roster import, daily job, shared view, request form, one-slider, archive). |
| `seed_data_example.py` | Fictional example data with the same shape as the real roster. The real data lives in `seed_data_private.py`, which is git-ignored and is never committed. |

## Build

```bash
pip install openpyxl
python3 build_tracker.py G2_PSG_Tracker.xlsx   # uses seed_data_private.py if present, else the example
TZ=UTC node tests/parser.check.cjs             # BDE-roster importer check (fictional fixture)
```

## Install (about 5 minutes)

1. Upload `G2_PSG_Tracker.xlsx` to Drive, then choose **Open with ▸ Google Sheets** and **File ▸ Save as Google Sheets**.
2. In the new Google Sheet, go to **Extensions ▸ Apps Script**, replace `Code.gs` with this repo's `Code.gs`, save, and reload the sheet.
3. Run **G2 PSG ▸ Setup / repair (run once)** and approve the permissions. Setup installs the daily 0500 job, the formatting, the warning-only protections on computed cells, the mind map and the charts.
4. Run **G2 PSG ▸ Build / refresh one-slider brief** once. After that, the daily job keeps the slide current.

## Tabs

| Tab | Purpose |
|---|---|
| Dashboard | Main tab: KPI tiles, BLUF, next-up per DA 6 (with sparklines), this month by roster, unavailable today, upcoming duties, tracker registry, 7 charts and the mind map. |
| Duty Slots | One row per G2 slot (only slots labeled G2). The Check column flags conflicts; Suggested shows the next-up Soldier who is available. |
| Next Up | Who is next on each DA 6: highest number first, with each Soldier's availability today. |
| DA6 CQ NCO / RUNNER / SDNCO / FLAG / ROAD GUARD | DA Form 6 grids built automatically from Duty Slots and Availability. CQ Runner and SD Runner share one RUNNER roster, so no one is double tasked. |
| Personnel | Master roster: status, DA 6 exclusion, brought-forward numbers, computed eligibility and last/next detail. |
| Availability | Leave, pass, TDY, school, appointments and exemptions. The absence type maps to an AR 220-45 code (A, D or U). |
| Appointments, Equipment, Taskings | Starter trackers. Add more with **G2 PSG ▸ Add tracker table**. |
| Settings | Rank bands per roster, duty and absence types, holidays/DONSAs, rest window, tie-break, as-of date and file IDs. |
| Dash Data | Feeds for the dashboard and charts (protected). |

## DA 6 rules applied (AR 220-45)

- The number is the count of details made from the roster since the Soldier was last detailed.
- `0` = performed the duty. `A` = authorized absence; the number pauses. `D` = other duty; the number continues. `U` = misconduct; the number continues.
- A column is posted only on dates a detail was made.
- Names are alphabetical within grade, highest grade first.
- The highest number is next. If that Soldier is unavailable, the next highest goes.
- Weekend and holiday postings are shown in red.
- Brought-forward numbers come from the Personnel `BF #` columns. **Close & restart DA 6s** archives the rosters and carries the current numbers forward.

The regulation doesn't cover ties or rest between duties, so both are settings: **Tie-break** (roster order or junior first) and **Rest window**.

## Automation (G2 PSG menu)

| Item | What it does |
|---|---|
| Auto-fill open G2 slots | Writes each open slot's suggestion (next-up and available) into Soldier. |
| Import G2 slots from a BDE roster tab | Parses the brigade roster layout (SD/CQ tables, flag teams, road guard), keeps G2-labeled slots, fuzzy-matches names and adds them without duplicating. |
| Publish shared view (no PII) | Writes a separate file with upcoming G2 duties and next-up names only. Share this file, never the tracker. |
| Build / refresh one-slider brief | Builds a one-slide Google Slides brief for senior leaders: BLUF, six KPIs, the next 7 days of G2 duties and a linked live chart. The daily job refreshes it. Set **Brief marking** on Settings if your SOP requires a banner. |
| Create exemption request form | Creates a Google Form. Responses land on the Requests tab; tick Approve to post one to Availability. |
| Email me the brief now | Emails the BLUF, next 7 days and conflicts. The daily job sends it at 0500. |
| Add tracker table | Adds a styled table tab and registers it on the dashboard. |
| Close & restart DA 6s | Archives values-only copies, carries numbers forward and resets the FROM date. |

## Privacy

Only the PSG has access to the tracker. The shared view and the one-slider are separate files that carry
only what others need: no phones, no exemption reasons and no medical details. Real roster data is never
committed to this repo.
