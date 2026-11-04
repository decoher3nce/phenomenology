# Phenomenology of remote sensing from space

Static site plus the workbook that mirrors it. Both are generated from `data.json` by `build.py`.

## Hosting
Served on the family tailnet by the home Caddy gateway (`~/claude-code/gateway/Caddyfile`, not part of this repo). Caddy's `file_server` reads `public/` straight from this checkout, so there is no deploy step: a `build.py` run is live on the next request.

- `/phenomenology/` — the interactive grid, catalog, concept sheets, interaction reference, TRL scale, the evidence behind every TRL, and the references
- `/phenomenology/workbook` — the Excel workbook (same data, same build)

GitHub is the off-machine backup of the source; pushing does not publish anything.

## Regenerating
```
python3 build.py          # writes the HTML and XLSX from data.json
```
`build.py` writes into `public/` next to itself. Nothing else to copy.

## Sources
Every technique, concept and grid annotation cites its sources. `data.json` holds a `references` list (peer-reviewed papers with DOIs, agency and mission documents, patents; each record says how and when it was verified), text fields carry `{R012}` tokens that the site renders as linked markers and the workbook as `[R012]`, and each entry lists the evidence for its TRL: missions and instruments at TRL 9, programs and their qualification level at 7–8, the item tested and the test level at 6 and below. The site's Evidence and References sections and the workbook's Evidence and References sheets are built from the same records.

## Caveats carried by the data
TRL values are estimates against NASA definitions for use from space, from knowledge through mid-2026, checked against the literature and mission status in October and November 2026. Items flagged `verify` sit in a 2025–2027 pathfinder window. Novelty claims on concepts are the weakest part.
