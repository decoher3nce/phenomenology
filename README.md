# Phenomenology of remote sensing from space

Static site plus the workbook that mirrors it. Both are generated from `data.json` by `build.py`.

## Deploy on Render
Connected as a Blueprint: Render reads `render.yaml`, serves `public/` as a static site. No build step. Pushing to the default branch redeploys.

- `/` — the interactive grid, catalog, concept sheets, interaction reference, TRL scale
- `/workbook` — the Excel workbook (same data, same build)

## Regenerating
```
python3 build.py          # writes the HTML and XLSX from data.json
```
`build.py` writes to `/mnt/user-data/outputs/` and `/home/claude/` by default; adjust the two output paths at the bottom of the script, then copy into `public/`.

## Caveats carried by the data
TRL values are estimates against NASA definitions for use from space, from knowledge through mid-2026. Items flagged `verify` sit in a 2025–2027 pathfinder window. Novelty claims on concepts are the weakest part.
