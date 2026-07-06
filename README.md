# Organization Dashboard Styler

A tool that lets any organization set up their own logo, header text, color
scheme, fiscal-year display order, and chart axis minimums for their
Executive Director's Report dashboard — with a live preview, and a button
to apply it into their real Excel file.

Private for now. Will be made public once it's been tested.

## What it does

1. Pick your organization (or create a new one).
2. Type your header text, upload your logo, pick your colors, choose which
   fiscal year displays first, and set axis minimums. Prefer working in
   Excel? Download a settings form instead, fill it in there, and upload it
   back — same result either way.
3. Watch the picture update instantly as you go.
4. Click "Save" so it remembers your settings next time.
5. Upload a copy of your real dashboard workbook and click "Apply branding"
   to get a branded version back — download it and you're done.

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

**Step 5 — Set up your organization.**
- On the left, choose "+ New organization" and type your organization's name.
- Type your header subtitle.
- Upload your logo.
- Pick your colors — the picture on the right updates right away.

**Step 6 — Save it.**
Click "Save organization profile." Next time you open the app, your
organization will already be in the list.

**Step 7 — Apply it to your real dashboard file.**
- Upload a copy of your dashboard workbook (it needs to be built from the
  standard template — same tabs and charts, just your own numbers).
- Click "Apply branding."
- Click "Download branded workbook" and save it wherever you like. Your
  original upload is never changed.

## What "the standard template" means

This app expects the dashboard workbook to have:
- A **Dashboard** tab (the one-page visual summary — this is what gets branded),
- A **Summary** tab (feeds the charts),
- A **Raw** tab (the underlying P&L data).

The three charts on the Dashboard tab must be, in order: an Income line
chart, an Expense line chart, and a Contributions bar chart (Actual then
Budgeted). If your file doesn't match that shape, the app will tell you
rather than guessing and silently getting it wrong.

A ready-to-copy blank starter template (placeholder numbers, no real
organization's data) is on the roadmap — not built yet.

## Project layout

```
app.py                  Streamlit app (the screen you interact with)
lib/org_profile.py      saves/loads each organization's settings
lib/dashboard_preview.py draws the live on-screen preview
lib/excel_writer.py      writes the branding into a real workbook copy
lib/settings_form.py    builds/reads the downloadable Excel settings form
orgs/                    where each organization's saved settings live
                         (never committed to this repo — stays on your computer)
requirements.txt
```

## Roadmap

- **Now:** logo, header text, color scheme (banners, bar chart, 4-series line
  charts), fiscal-year display order, axis minimums, and a downloadable/
  uploadable Excel settings form as an alternative to the on-screen controls.
  The settings form also has its own organization name/header subtitle
  fields and a 5-preset color-scheme dropdown (Purple & Orange, Cornerstones
  Classic, Navy & Silver, Forest & Amber, Slate & Crimson) that auto-fills
  every color cell -- and its own color swatch -- the moment you pick a
  scheme, so it can be filled out entirely in Excel without ever opening
  the web app.
- **Next:** switching a chart's type (e.g. bar to line).
- **Next:** editing number formats (e.g. `$#,##0,K`).
- **Later:** a blank starter template file so a brand-new organization doesn't
  need an existing Cornerstones-style file to start from.
