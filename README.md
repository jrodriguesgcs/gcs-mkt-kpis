# HubSpot Traffic-Source Funnel Report

Fetches every HubSpot contact and its associated deals, applies four
portal-wide filters, and writes a single styled `.xlsx` to
`reports/funnel_report.xlsx`:

- **Funnel Report** — rows are **Original Traffic Source**, 3 levels deep
  (Source → Drill-Down 1 → Drill-Down 2). Columns are
  **Funnel** (New Contacts / New Deals / Existing Deals) × **Stage** (6
  stages each, 18 total) × **Time** (Month → Week → Day, full calendar
  year). Both axes use Excel's native **Group & Outline** (+/- buttons) —
  clicking + on a row expands its next drill-down level; clicking + on a
  Month column reveals its Weeks, and on a Week reveals its Days. Every
  rollup (row and time) is a real `=SUM()` formula referencing its
  children, so collapsing anything never changes a visible total.
- **Filters & Definitions** — a static reference sheet: every resolved
  property/pipeline/stage id, the four overall filters as applied, the
  funnel/stage definition table, the filter audit trail, and which
  months/weeks this run treated as in-progress vs. finished.

Run with: `python generate_funnel_report.py`

Deliberately **not** in scope for this pass: GitHub Actions / cron
scheduling and email delivery. The script is a single, non-interactive
entrypoint with clear exit codes and no hardcoded paths, so automation can
be layered on later without rework.

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
```

Get a HubSpot **Service Key** (Settings → Integrations → Service Keys,
currently public beta) or, if that's unavailable on this portal, a
**private app access token** — both are used identically as
`Authorization: Bearer <token>`, so no code change is needed either way.
Required scopes:

| Scope | Why |
|---|---|
| `crm.objects.contacts.read` | contact traffic-source/type/brand/lead-source properties |
| `crm.objects.deals.read` | deal pipeline/stage/proposal properties, associations |
| `crm.schemas.contacts.read` | Step 0 property resolution |
| `crm.schemas.deals.read` | Step 0 property resolution |
| `crm.pipelines.read` (or closest available) | resolving the Sales pipeline + Closed Lost stage id |

Paste the token into `.env` as `HUBSPOT_ACCESS_TOKEN`.

## Running it

```bash
python generate_funnel_report.py
```

Non-zero exit + a clear message on `stderr` on any failure (missing
token, an unresolvable property, a failed QA check). Safe to re-run —
`reports/funnel_report.xlsx` is cleanly overwritten each time, no
leftover state.

### What gets printed

1. **Step 0** — every resolved contact/deal property, pipeline id, and
   stage id, plus a `NOTE`/`WARNING` line for anything that diverged from
   this project's original guess (see **Terminology notes** below).
2. **Filter audit trail** — contact count remaining after each of the 4
   overall filters, in order, so an unexpectedly large or small drop is
   visible immediately.
3. **Row hierarchy** — how many Level-1/2/3 rows were built.
4. **QA** (see below) — printed as `[1]`..`[6]`, each `PASS`/`FAIL` with
   specifics on failure (which row, column, date).

If QA fails, the script still writes the workbook (so you can inspect it)
but exits non-zero — treat that build as unverified, not as delivered.

## Terminology notes (resolved vs. originally-guessed property names)

A few property names guessed when this report was scoped don't exist
verbatim on this portal; Step 0 resolves the actual ones by label match
and prints a `NOTE` each time this happens. As of the last verified run:

| Originally guessed | Actually used | Why |
|---|---|---|
| *(a property literally labeled "Brand")* | `hs_all_assigned_business_unit_ids` (label "Brands", HubSpot's Business Units feature) | no property is labeled exactly "Brand". An earlier resolution picked `client_x_brand` ("Client x Brand") purely because its label also contained "brand" — but live data showed only ~0.06% of contacts have it set to "Global Citizen Solutions" (a manual post-sale tag), vs. ~92.6% tagged to the "Global Citizen Solutions" business unit via `hs_all_assigned_business_unit_ids`. Step 0 now prefers this property by name when it's present and carries the required option. Since a contact can belong to more than one business unit, the match is "value is present" (`;`-separated token containment), not exact equality. |
| `sql_lost_reason` | `sql_lost` (label "SQL Lost Reason") | exact label match; the guessed internal name doesn't exist |
| `proposal_sent_date` | `deal_proposal_sent_datetime` (label "Deal Proposal Sent Date Time") | closest label match; the guessed name doesn't exist |
| `proposal_signed_date` | `deal_proposal_signed_datetime` (label "Deal Proposal Signed Date Time") | same reasoning |

The code never hardcodes any of the names above — it re-resolves them by
label at every run and will print an updated `NOTE`/`WARNING` if the
portal changes again.

## Design decisions worth knowing about

- **No separate "Year" column.** Month is the top always-visible time
  granularity per stage (outline level 0); there's no outline level left
  for a Year node once Week=1 and Day=2 are assigned. The report year is
  shown as a title in the sheet's corner cell instead.
- **Metric per stage-column**: "New Contacts" (only under Funnel: New
  Contacts) counts **distinct contacts**; every other stage-column counts
  **distinct deals**.
- **"This period"/"before this period" qualification is evaluated once,
  at Month granularity** (matching the spec's literal "this month"/
  "before this month" wording), using whichever date drives that stage
  (a deal's createdate for New Contacts/New Deals funnels, or the
  relevant milestone date for Existing Deals). Week and Day columns are a
  pure sub-split of that same already-qualified population by the
  stage's own date — **not** an independent re-evaluation at finer
  grain. This is what guarantees a Month total always equals the sum of
  its Weeks, and a Week total always equals the sum of its Days (verified
  by QA check 1/2 for every row, not just a sample).
- **Full calendar year of columns, all 12 months**, even future ones —
  a not-yet-elapsed day's cell is left blank (not zero); Week/Month
  totals still exist as real columns and sum whatever elapsed data
  exists beneath them.

## What's in the workbook

**Funnel Report**: column A holds the row label at whichever hierarchy
level is visible (indented per level). Row 1 is the Funnel band, row 2 the
Stage band, row 3 the Month/Week (`W1`, `W2`, ...) or Day (`YYYY-MM-DD`)
label. Each of the 18 stage-columns is colored from a 6-step HSL ramp
derived from its funnel's brand colour (Night Blue / Electric Blue /
Slate), lightest at "New Contacts" and darkest at "Proposal Signed" —
header text switches to white automatically once the background gets dark
enough to need it. The 6 combinations the spec marks not applicable (e.g.
"New Contacts" stage under the "New Deals" funnel) still get the full
Month→Week→Day column structure, just filled with the literal text `N/A`
in a neutral grey instead of a ramp colour. Numeric data cells themselves
stay plain white/neutral, per the "flat fills, sharp corners" rule —
color only ever marks a header band.

**Filters & Definitions**: static text, generated once per run — not a
live pivot, so it reads correctly even if reopened without recalculating.

## QA

Printed on every run, per the spec's required checks:

1. **Full click-path simulation** — for the richest Level-1 traffic
   source and the "Qualified Deals" stage under Funnel: New Contacts,
   confirms Month = sum(Weeks) and Week = sum(Days), for both a finished
   month and the current in-progress month.
2. **Row rollup integrity** — for every row (not a sample), Level 1 =
   sum(Level 2 children) and Level 2 = sum(Level 3 children), for every
   funnel/stage/month.
3. **Filter correctness spot-check** — for each of the 4 overall filters,
   confirms 2-3 sampled excluded contacts don't appear anywhere in the
   output, and 2-3 sampled blank-on-that-filter contacts (that survive
   all 4 filters) do.
4. **Funnel/Stage definition correctness** — prints each of the 18
   funnel×stage totals with a one-line description of its filter, for
   manual sanity-checking against the spec.
5. **To-date vs. completed logic** — prints which month/week this run
   treated as in-progress, and confirms only elapsed days/weeks were
   marked as having data.
6. **Column count sanity check** — prints the final column count and
   confirms it's comfortably under Excel's 16,384-column limit (typically
   ~7,900: 18 stage-groups × ~440 time columns each).

A failing check is described with the specific row/column/date needed to
debug it, not just "mismatch found".

## Recalculation

The workbook is saved with real `=SUM()` formulas. The script then tries
a headless LibreOffice round-trip (`soffice --headless --calc
--convert-to xlsx`) to force a proper recalculation pass; if LibreOffice
isn't available in the environment it falls back to writing the correct
cached value directly into each formula cell's XML (the formula itself is
untouched) so the file still opens showing real numbers rather than
stale/blank ones in viewers that don't recalculate on open. Excel itself
recalculates automatically on open regardless.

## Performance

The script fetches only contacts **created in the report year** (a
server-side `createdate >= Jan 1` filter via the Search API), plus every
deal associated with them. This was changed from an unfiltered full-portal
fetch specifically to cut runtime on large portals — see the trade-off
below before relying on it.

**Accepted trade-off:** the New Deals and Existing Deals funnels are
defined around contacts created *before* the period being evaluated. A
contact created in a prior year that still produced deal activity this
report year will **not** appear in either funnel under this filter — only
New Contacts (which requires `contact.createdate` this period anyway) is
unaffected. If your portal has a lot of deal activity on older contacts,
this will visibly undercount those two funnels; the original unfiltered
behavior (fetch every contact, regardless of year) is one line away if you
need it back — see `fetch_contacts()`'s docstring.

Also note: HubSpot's Search API caps total results at **10,000** per
query regardless of pagination. If a single report year's new contacts
exceed that on your portal, the excess will silently not be fetched —
compare the printed "Fetched N contacts" line against your own HubSpot
contact count filtered the same way if that's a realistic volume for you.
