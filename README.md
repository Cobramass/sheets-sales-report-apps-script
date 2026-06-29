# Sales Report Builder — Google Sheets / Apps Script demo

A one-click Sheets automation: it reads a raw **Transactions** tab and builds a clean
**Summary** tab (revenue by category and by rep, with counts and a timestamp). Portfolio demo,
but built to the bar that decides whether a non-technical buyer leaves five stars.

## Run it (as the buyer does)

No script editor required — that's the point.

1. Open the demo Sheet → menu **Sales Tools → Set up demo data** (creates a Transactions tab
   with realistic, deliberately-messy sample rows).
2. **Sales Tools → Build sales report.** The **Summary** tab appears. Re-run any time.

## The bar this demonstrates

| #1 review-killer in this lane | How it's defended here |
|---|---|
| *"Open the script editor and run…"* | A real **custom menu** + about dialog (`onOpen`). The buyer clicks a menu item, never sees code. |
| *Breaks when the client edits the sheet* | Columns are found **by header name** (reorder/recolour freely); rows by **dynamic last-row**. **Zero hardcoded A1 coordinates.** Missing a required column → a named, plain-English error. |
| *Silently stops / miscounts* | Blank rows and unparseable amounts are **skipped and counted** (shown on the Summary), never quietly treated as zero. Only Paid/Completed rows count — stated in the about dialog. |
| *Quota blowups on real data* | **Batched** `getValues`/`setValues` — one read per slice, two writes total. No cell-by-cell loops. |
| *Dies past 6 minutes, half-done* | **Resumable**: checkpoints progress to document properties and schedules a 1-minute continuation trigger, so a large sheet finishes across runs instead of half-finishing. |
| *Scary "unverified app" consent* | `@OnlyCurrentDoc` + narrow `oauthScopes` in [`appsscript.json`](appsscript.json) — the gentlest possible consent screen. |
| *A failed overnight run is invisible* | Every entry point is wrapped: a plain-English **toast** *and* an **email to the owner** with the reason. The notifier itself never throws. |
| *Hand-edited report gets clobbered* | The Summary tab is **protected (warning-only)** so edits there are intentional. |

Plus an in-sheet **Instructions** tab so the how-to travels with the file.

## Deploy it to your own Sheet

**Option A — clasp (one command):**
```bash
npm i -g @google/clasp && clasp login
cp .clasp.json.example .clasp.json   # paste your Script ID (Extensions → Apps Script → Project Settings)
clasp push
```

**Option B — manual:** in the Sheet, **Extensions → Apps Script**, paste `Code.gs` and `Setup.gs`
into files of the same names, set the manifest (`appsscript.json`) under Project Settings →
"Show appsscript.json", **Save**, reload the Sheet. The **Sales Tools** menu appears.

> **The one account-bound step:** publishing the live, public **view-only** demo Sheet has to be
> done from the owner's Google account (File → Share → "Anyone with the link: Viewer"). The code
> here reproduces that Sheet exactly via *Set up demo data*, so the published link is just this
> repo, deployed.

## Loom walkthrough script (60–90s — record on the live Sheet)

1. "Here's a raw sales export — messy: a blank row, a refund, an amount typed as `n/a`, a missing category."
2. Click **Sales Tools → Build sales report**. Summary appears instantly.
3. "Totals by category and rep, a grand total, and — importantly — it tells you it skipped 3 rows
   rather than silently miscounting them."
4. Reorder the Transactions columns, re-run → "still works, because it matches by header name, not position."
5. Rename a required column → show the plain-English error + the owner email. "It never fails silently."

## Files

```
Code.gs            menu, the report builder, batching, resumability, error handling
Setup.gs           demo-data + in-sheet instructions generator
appsscript.json    manifest: V8, @OnlyCurrentDoc, minimal scopes
.clasp.json.example clasp deploy pointer (copy to .clasp.json with your Script ID)
```

Built by Matthew Daly — Google Sheets / Apps Script automation, delivered click-to-run with a Loom.
