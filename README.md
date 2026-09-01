# CORDIS AI Funding Radar

![Weekly run](https://github.com/tuanasuslu/cordis-ai-funding-radar/actions/workflows/weekly_radar.yml/badge.svg)

**Live report:** https://tuanasuslu.github.io/cordis-ai-funding-radar/

A small automated pipeline that turns the EU's official CORDIS research-funding
dataset into a usable lead list: every SME in Germany, Austria, and Switzerland
that has received Horizon Europe / EIC Accelerator funding, tagged by how
central AI is to their funded project.

## What it does

1. Downloads the official CORDIS bulk export (Horizon Europe 2021-2027) from
   the EU Open Data Portal.
2. Works around a real quoting bug in CORDIS's own CSV export (see
   `load_table` in the script) that trips up pandas' built-in parsers.
3. Joins projects to their funded organisations, keeps only SMEs in the
   target countries, and splits results into two segments: **EIC
   Accelerator** and **Horizon Europe (general)**.
4. Tags every row `Core AI`, `AI-adjacent`, or `Not AI-related` based on
   whether AI/ML terms appear in the project title or objective - nothing is
   filtered out, since a company doesn't need to be an "AI company" to be a
   relevant lead.
5. Adds a ready-made Google search link for rows where CORDIS has no website
   on file (CORDIS's own contact data is frequently missing or stale - this
   script never guesses, it just makes manual verification one click away).
6. Publishes the result as both a formatted Excel workbook and a sortable,
   searchable static webpage (`docs/index.html`, served via GitHub Pages).
7. Runs automatically every Monday via GitHub Actions
   (`.github/workflows/weekly_radar.yml`), committing the refreshed output
   back to this repo.

## Running it yourself

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python cordis_ai_funding_radar.py --countries DE,AT,CH --output ai_funding_radar_dach.xlsx
```

## Known limitations

- Website/contact info is frequently missing or outdated in CORDIS - verify
  manually via the generated search link.
- Outside of EIC Accelerator, CORDIS doesn't break out the exact funding
  amount per company within a multi-partner project.
- CORDIS only covers EU-funded projects; AI companies funded by VC/private
  equity don't appear here.
- The "Scientific/Technical Domain" column is CORDIS's own EuroSciVoc
  classification, not a formal business sector (NACE/SIC) - CORDIS doesn't
  provide that data.
- CORDIS's own CSV export drops ~0.8% of rows due to a source-side quoting
  bug; the script reports the exact count on every run.
