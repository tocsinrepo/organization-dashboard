# Organization Dashboard Styler

A tool that lets any organization set up their own logo, header text, color
scheme, fiscal-year display order, and chart axis minimums for their
Executive Director's Report dashboard — with a live preview, and a button
to apply it into their real Excel file.

Private for now. Will be made public once it's been tested.

## What it does

1. Pick your organization (or create a new one).
2. Upload your dashboard workbook in the sidebar — the preview then shows
   **your organization's real Income, Expense, and Data numbers**, styled with
   your branding, so you see exactly how your own report will look (optional:
   without a workbook the preview uses clearly-labeled sample numbers).
3. Set your branding using the tabbed controls: **Text & logo**, **Colors**
   (banner, accent, bar, and line colors update together — fine-tune any single
   one afterward), and **Chart options** (fiscal-year display order, axis
   minimums, and each chart's type — line or bar). Prefer working in Excel?
   Download a settings form instead, fill it in there, and upload it back —
   same result either way.
4. Watch the picture update instantly as you go.
5. Click "Save" so it remembers your settings next time.
6. Click "Apply branding" to get a branded copy of the workbook you uploaded —
   download it and you're done.

Each organization runs its own copy of this app on their own computer.
Nobody's logo, colors, or data ever leaves their machine.

**Note on fiscal-year display order:** "1st position" moves a year both to
the front of the legend *and* to the top of the drawn lines together —
Excel's file format doesn't allow reordering one without the other, so this
app doesn't pretend otherwise.

## How to run it (step by step)

**Step 1 — Get the code onto your computer.**
Open Terminal and type this, then press Enter:
```
git clone https://github.com/tocsinrepo/organization-dashboard.git
cd organization-dashboard
```

**Step 2 — Install what it needs.**
Type this and press Enter (only needed the first time):
```
pip install -r requirements.txt
```

**Step 3 — Start the app.**
Type this and press Enter:
```
streamlit run app.py
```

**Step 4 — A browser tab opens by itself.**
That's the app. If it doesn't open automatically, Terminal will show a web
address (something like `http://localhost:8501`) — copy that into your
browser.

**Step 5 — Upload your dashboard workbook (sidebar).**
In the left sidebar, upload a copy of your dashboard workbook. The preview now
shows *your* real Income, Expense, and Data numbers. (You can skip this and
still design your branding against sample numbers — but uploading lets you see
the real thing and is also the file that branding gets applied to.)

**Step 6 — Set up your organization.**
- In the sidebar, choose "+ New organization."
- In the **Text & logo** tab: type your organization's name and subtitle, and
  upload your logo.
- In the **Colors** tab: pick a color scheme — the picture on the right updates
  right away. Want to nudge just one color? Use the individual pickers.
- In the **Chart options** tab: set fiscal-year display order, axis minimums,
  and each chart's type (line or bar).

**Step 7 — Save it.**
Click "Save organization profile." Next time you open the app, your
organization will already be in the list.

**Step 8 — Apply it to your real dashboard file.**
- Click "Apply branding" (uses the workbook you uploaded in the sidebar).
- Click "Download branded workbook" and save it wherever you like. Your
  original upload is never changed.

## What "the standard template" means

This app expects the dashboard workbook to have:
- A **Dashboard** tab (the one-page visual summary — this is what gets branded),
- A **Summary** tab (feeds the charts),
- A **Raw** tab (the underlying P&L data).

The three charts on the Dashboard tab must be, in order: an Income line
chart, an Expense line chart, and a Data bar chart (Actual then
Budgeted). If your file doesn't match that shape, the app will tell you
rather than guessing and silently getting it wrong.

A ready-to-copy blank starter template (placeholder numbers, no real
organization's data) is on the roadmap — not built yet.

## Project layout

```
app.py                  Streamlit app (the screen you interact with)
lib/org_profile.py      the OrgProfile record; saves/loads it to disk
lib/storage.py          picks the storage backend: disk (local) or isolated
                         per-session (hosted multi-tenant, MULTI_TENANT=1)
lib/app_state.py        keeps the app's widgets, an OrgProfile, and the
                         chosen color scheme in sync (see its module
                         docstring if you're extending this file --
                         Streamlit widgets are stricter about this than
                         they first appear)
lib/dashboard_data.py   reads your real Income/Expense/Data numbers out of an
                         uploaded workbook so the preview can show them
lib/dashboard_preview.py draws the live on-screen preview
lib/excel_writer.py      writes the branding into a real workbook copy
lib/settings_form.py    builds/reads the downloadable Excel settings form
orgs/                    where each organization's saved settings live
                         (never committed to this repo — stays on your computer)
requirements.txt
```

## Roadmap

- **Now:** a live preview of your organization's **real Income / Expense / Data
  numbers** read from your uploaded workbook, in a 2×2 layout (Data top-left, an
  at-a-glance summary top-right, Income bottom-left, Expense bottom-right);
  logo, header text, a 5-preset color scheme (Purple & Orange, Gold & Teal,
  Navy & Silver, Blue & Sky, Slate & Crimson) that sets the banner/accent/bar/
  line colors together with the option to fine-tune any single one afterward,
  fiscal-year display order, axis minimums, **per-chart type switching (line or
  bar)**, and a downloadable/uploadable Excel settings form as an alternative to
  the on-screen controls — same presets and options there too.
- **Next:** editing number formats (e.g. `$#,##0,K`).
- **Later:** a blank starter template file so a brand-new organization doesn't
  need an existing sample file to start from.

### Note on chart-type switching

Changing a chart between line and bar updates both the live preview and the
workbook you apply to. The conversion is verified in this project's tooling
(LibreOffice), but a few Excel chart nuances can only be confirmed in Microsoft
Excel itself — so open a converted workbook in Excel and give it a look before
relying on it.

### Hosting for more than one organization

By default this app stores each organization's branding on disk (`orgs/`),
which is correct when every organization runs its **own local copy**. If you
host **one shared URL** (e.g. Streamlit Community Cloud) for several visitors,
set the environment variable `MULTI_TENANT=1` (or a Streamlit secret
`multi_tenant = true`) on that deployment. In that mode each browser session
gets its **own isolated, in-memory** set of organizations — no visitor can see
or overwrite another's — at the cost of those being ephemeral (they last for
the session, not across restarts). Leave it unset for local use and disk
persistence is unchanged.
