# Remote-sensing phenomenology grid

A static reference site and a mirrored Excel workbook, both generated from one data file.

## Layout
- `data.json` — the single source of truth: sources, interactions, geometries, receivers, 80 techniques, 24 concepts, cell annotations, TRL scale, physical filters. Edit this, nothing downstream.
- `build.py` — reads `data.json`, writes `public/index.html` and `public/remote-sensing-phenomenology.xlsx`. No arguments.
- `public/` — what the gateway serves. Regenerated output; committed so the GitHub copy is complete on its own.
- `scripts/recalc.py` — recalculates workbook formulas with LibreOffice and reports errors; run after every build if `soffice` is installed. Needs `scripts/office/`.
- `scripts/apply_trl.py`, `scripts/add_interactions.py` — one-off migrations already applied. Provenance only.
- Hosting lives outside this repo: the home Caddy gateway (`~/claude-code/gateway/Caddyfile`, see `PUBLISHING.md` there) has a `handle_path /phenomenology/*` block that `file_server`s this checkout's `public/` and rewrites `/workbook` to the xlsx. Reload after editing it: `caddy reload --config ~/claude-code/gateway/Caddyfile --adapter caddyfile` (admin API, no sudo).

## Workflow
```
python3 build.py
python3 scripts/recalc.py public/remote-sensing-phenomenology.xlsx   # expect "status": "success", total_errors 0
git add -A && git commit && git push                                  # backup only; the site is live as soon as build.py finishes
```
Dependencies: Python 3.10+, `openpyxl`. LibreOffice only for recalc.

## Data conventions
- IDs: techniques `T01…`, concepts `C01…`. Keep them stable; the page links by ID.
- `src` is one source code for a technique, a list for a concept. `int` and `geom` are lists of codes. Codes are defined in the `sources`, `interactions`, `geometries`, `receivers` arrays.
- `trl` is an integer 1–9 judged for use *from space in the stated application*, per NASA definitions in `trlScale`. `trlNote` explains any non-obvious call. `verify: true` flags values that sit in a 2025–2027 pathfinder window.
- Concepts carry `dubious: true` when the physics does not close at orbital ranges; the grid excludes them from concept counts but still draws them (dashed).
- `cellNotes` keys are `source|interaction`; `kind` is `ruled`, `open`, `filled` or `concept`.
- Classification: reflection = geometric-optics regime; radar/lidar backscatter = elastic scattering; sounders count under emission and absorption; aperture synthesis is emission + correlation receiver, not phase.

## Things to keep true
- Both outputs come from one `build.py` run. Never hand-edit `public/`.
- The workbook's Grid sheet is formulas over Catalog/Concepts, ranges rows 2–400 and 2–200. Extend them if the sheets grow past that.
- Avoid putting the literal text of a spreadsheet error code in any cell; `recalc.py` matches it as a real error.
- Status tags and novelty claims in the data are estimates from mid-2026 knowledge. When adding or changing an entry, say in `trlNote` or `prior` what the basis is.

## Open items
- The clickable column headers (interaction drawer) were verified by parse only, not in a browser. Confirm once.
- The site is served from this working copy, so a half-finished edit to `public/` is visible immediately. Run `build.py` to completion; never hand-edit `public/`.
- The page uses only fragment links, which is what makes it safe under the `/phenomenology/` sub-path. Keep it that way (no absolute `/...` hrefs in `build.py`).
