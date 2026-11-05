# Space Remote Sensing: Phenomena, Techniques and Readiness

Short name "Space Remote Sensing", used for file names. The repo, the checkout folder and the `/phenomenology/` web path keep the old name so links stay stable.

A static reference site and a mirrored Excel workbook, both generated from one data file.

## Layout
- `data.json` — the single source of truth: sources, interactions, geometries, receivers, 80 techniques, 24 concepts, cell annotations, TRL scale, physical filters, and the `references` list every claim points to. Edit this, nothing downstream.
- `build.py` — reads `data.json`, writes `public/index.html` and `public/space-remote-sensing.xlsx`. No arguments. `PHEN_DATA=<file> PHEN_OUT=<dir>` build a test copy somewhere else without touching the live `public/`.
- `public/` — what the gateway serves. Regenerated output; committed so the GitHub copy is complete on its own.
- `scripts/recalc.py` — recalculates workbook formulas with LibreOffice and reports errors; run after every build if `soffice` is installed. Needs `scripts/office/`.
- `scripts/apply_trl.py`, `scripts/add_interactions.py` — one-off migrations already applied. Provenance only.
- `scripts/calc/` — the "first calculation" for a concept, one stdlib script each, assumptions at the top, verdict at the bottom. When a concept's TRL moves on a calculation, the script is the evidence and `trlNote` names it.
- Hosting lives outside this repo: the home Caddy gateway (`~/claude-code/gateway/Caddyfile`, see `PUBLISHING.md` there) has a `handle_path /phenomenology/*` block that `file_server`s this checkout's `public/` and rewrites `/workbook` to the xlsx. Reload after editing it: `caddy reload --config ~/claude-code/gateway/Caddyfile --adapter caddyfile` (admin API, no sudo).

## Workflow
```
python3 build.py
python3 scripts/recalc.py public/space-remote-sensing.xlsx   # expect "status": "success", total_errors 0
git add -A && git commit && git push                                  # backup only; the site is live as soon as build.py finishes
```
Dependencies: Python 3.10+, `openpyxl`. LibreOffice only for recalc. A local `.venv/` (gitignored) holds openpyxl: `python3 -m venv .venv && .venv/bin/pip install openpyxl`, then build with `.venv/bin/python build.py`.

## Data conventions
- IDs: techniques `T01…`, concepts `C01…`. Keep them stable; the page links by ID.
- `src` is one source code for a technique, a list for a concept. `int` and `geom` are lists of codes. Codes are defined in the `sources`, `interactions`, `geometries`, `receivers` arrays.
- `trl` is an integer 1–9 judged for use *from space in the stated application*, per NASA definitions in `trlScale`. `trlNote` explains any non-obvious call. `verify: true` flags values that sit in a 2025–2027 pathfinder window.
- Concepts carry `dubious: true` when the physics does not close at orbital ranges; the grid excludes them from concept counts but still draws them (dashed).
- `cellNotes` keys are `source|interaction`; `kind` is `ruled`, `open`, `filled` or `concept`.
- **References.** `references` is a flat list of records `{id: "R012", key, type, authors, org, title, venue, year, doi, url, accessed, verified, note}`. IDs are stable and never reused; append new records at the end. `type` is one of journal, conference, report, mission-page, patent, book, preprint, dataset, thesis, standard. Only authoritative sources are admitted: peer-reviewed papers with a DOI, agency and mission documents, patents, academic books; preprints only where no journal version exists. Every record carries how it was verified (`crossref`, `doi-redirect`, `url-fetched`) and the date. `title` is the source's own title verbatim (the paper's title, or the web page's headline without the site-name suffix, which belongs in `org`/`venue`); dates, status and other annotations go in `note`, never in `title`, so `refcheck --batch` can match every title mechanically. The one expected refcheck title failure is R134, a JSON search feed with no page title.
- **Citing.** Put `{R012}` tokens in any text field (`trlNote`, `prior`, `physics`, cell `text`, evidence `detail`) right after the clause the source supports; `build.py` turns them into linked markers on the site and `[R012]` in the workbook. Each technique, concept and annotated cell also carries `refs: [...]`, the union of everything it cites. A new claim without a reference is incomplete.
- **TRL evidence.** Each technique and concept carries `evidence: [...]`, one item per piece of evidence for its TRL: `{kind, name, instrument, operator, dates, level, detail, refs}`. `kind` is `mission` at TRL 9 (the mission and instrument in service), `program` at TRL 7–8 (what was flown or qualified and what remains), and `system`, `subsystem`, `component` or `technology` at TRL 6 and below (the item tested). `level` is the status, qualification or test level in words (operational; flight-demonstrated in space (TRL 7); airborne demonstration; laboratory proof of principle; ...). Every item needs at least one reference. The Evidence section and the workbook's Evidence sheet are built from this.
- `meta.trlRef` is the reference ID for the NASA TRL definitions.
- Classification: reflection = geometric-optics regime; radar/lidar backscatter = elastic scattering; sounders count under emission and absorption; aperture synthesis is emission + correlation receiver, not phase.

## Design and color
- The palette is colorblind-safe by construction. TRL bands are ordinal, so they share one hue (blue) stepped only in lightness, each marker prints its TRL digit, and concepts are circles while techniques are squares. Orange is the single accent, for verify flags and keyboard focus. No pair of colors differs mainly along red-green.
- The ramps were validated with the dataviz skill's ordinal checker in light and dark mode. Re-run it if any color changes: `node validate_palette.js "<b13,b46,b78,b9>" --ordinal --mode light|dark --surface <panel>`.
- Digits on the two lighter steps are dark ink and on the two darker steps white (swapped in dark mode). Keep every digit at 4.5:1 or better.
- Colors are CSS custom properties at the top of the template in `build.py`, defined three times: light, OS dark, and the `data-theme="dark"` toggle. Change all three together.
- The page has a sticky section nav, a theme toggle (auto, light, dark, remembered per browser), and a custom hover tooltip on grid markers. It must stay usable at phone width with no sideways page scroll.

## Things to keep true
- Both outputs come from one `build.py` run. Never hand-edit `public/`.
- The workbook's Grid sheet is formulas over Catalog/Concepts, ranges rows 2–400 and 2–200. Extend them if the sheets grow past that.
- Avoid putting the literal text of a spreadsheet error code in any cell; `recalc.py` matches it as a real error.
- Status tags and novelty claims in the data are estimates from mid-2026 knowledge. When adding or changing an entry, say in `trlNote` or `prior` what the basis is and cite it with a `{Rnnn}` token; add the record to `references` with its DOI or agency URL, verified.
- The workbook's Catalog and Concepts sheets gained Sources and TRL-evidence columns at the right-hand end; the Grid formulas reference columns C, D, G and J, so insert new columns after the existing ones, never between.

## Open items
- The site is served from this working copy, so a half-finished edit to `public/` is visible immediately. Run `build.py` to completion; never hand-edit `public/`.
- The page uses only fragment links, which is what makes it safe under the `/phenomenology/` sub-path. Keep it that way (no absolute `/...` hrefs in `build.py`).
