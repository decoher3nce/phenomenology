import json
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

import os, pathlib
ROOT = pathlib.Path(__file__).resolve().parent
OUT = ROOT / 'public'; OUT.mkdir(exist_ok=True)
D = json.load(open(ROOT / 'data.json'))
DATA_JS = json.dumps(D, ensure_ascii=False).replace('</', '<\\/')

HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Phenomenology of remote sensing from space</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Source+Serif+4:ital,opsz,wght@0,8..60,400;0,8..60,600;1,8..60,400&family=Source+Sans+3:wght@400;600&display=swap" rel="stylesheet">
<style>
:root{
  --bg:#F3F5F8; --panel:#FFFFFF; --ink:#1B2433; --muted:#5B6778; --rule:#D6DCE4; --rule-strong:#AAB4C2;
  --b9:#1F5FA8; --b78:#1B8A7A; --b46:#C77D14; --b13:#8447B8; --dub:#9AA3AF;
  --hatch:#C9D0DA; --sel:#1B2433; --link:#1F5FA8; --warn:#C77D14;
  --serif:"Source Serif 4", Georgia, "Times New Roman", serif;
  --sans:"Source Sans 3", "Helvetica Neue", Arial, sans-serif;
}
@media (prefers-color-scheme: dark){
  :root:not([data-theme="light"]){
    --bg:#121820; --panel:#1B2430; --ink:#E7ECF2; --muted:#9AA7B8; --rule:#2E3A48; --rule-strong:#4A586A;
    --b9:#6FA3E0; --b78:#4FBFAE; --b46:#E0A24A; --b13:#B58AE0; --dub:#6B7684; --hatch:#2E3A48; --sel:#E7ECF2; --link:#8DB8EE; --warn:#E0A24A;
  }
}
:root[data-theme="dark"]{
  --bg:#121820; --panel:#1B2430; --ink:#E7ECF2; --muted:#9AA7B8; --rule:#2E3A48; --rule-strong:#4A586A;
  --b9:#6FA3E0; --b78:#4FBFAE; --b46:#E0A24A; --b13:#B58AE0; --dub:#6B7684; --hatch:#2E3A48; --sel:#E7ECF2; --link:#8DB8EE; --warn:#E0A24A;
}
*{box-sizing:border-box}
html,body{margin:0}
body{background:var(--bg);color:var(--ink);font-family:var(--serif);font-size:17px;line-height:1.55;}
a{color:var(--link)}
.wrap{max-width:1180px;margin:0 auto;padding:2.5rem 1.25rem 5rem}
h1{font-size:clamp(1.9rem,4vw,2.9rem);line-height:1.1;font-weight:600;margin:0 0 .6rem;letter-spacing:-.01em;max-width:22ch}
.lede{font-size:1.15rem;max-width:62ch;margin:0 0 1rem}
.caveat{font-family:var(--sans);font-size:.95rem;color:var(--muted);max-width:74ch;border-left:3px solid var(--warn);padding-left:.8rem;margin:1.2rem 0 2.4rem}
h2{font-size:1.55rem;font-weight:600;margin:3.2rem 0 .6rem;letter-spacing:-.005em}
h3{font-size:1.12rem;font-weight:600;margin:1.4rem 0 .3rem}
p{max-width:72ch}
.sans{font-family:var(--sans)}
.muted{color:var(--muted)}

.legend{display:flex;flex-wrap:wrap;gap:.9rem 1.4rem;font-family:var(--sans);font-size:.92rem;margin:.6rem 0 1rem;align-items:center}
.legend span{display:inline-flex;align-items:center;gap:.4rem}
.sw{width:12px;height:12px;border-radius:2px;display:inline-block}
.sw.hatch{background:repeating-linear-gradient(45deg,var(--hatch) 0 2px,transparent 2px 5px);border:1px solid var(--rule-strong)}
.sw.open{border:1.5px dotted var(--rule-strong)}
.sw.dub{border:1.5px dashed var(--dub);background:transparent}
.filters{display:flex;flex-wrap:wrap;gap:.5rem 1.2rem;font-family:var(--sans);font-size:.92rem;margin:0 0 1rem;align-items:center}
.filters label{display:inline-flex;align-items:center;gap:.4rem}
select{font-family:var(--sans);font-size:.92rem;background:var(--panel);color:var(--ink);border:1px solid var(--rule-strong);border-radius:4px;padding:.25rem .4rem}
button.chip{font-family:var(--sans);font-size:.9rem;background:var(--panel);color:var(--ink);border:1px solid var(--rule-strong);border-radius:999px;padding:.2rem .7rem;cursor:pointer}
button.chip[aria-pressed="true"]{background:var(--sel);color:var(--bg);border-color:var(--sel)}
button.chip:focus-visible,td.cell:focus-visible,select:focus-visible,input:focus-visible{outline:2px solid var(--warn);outline-offset:2px}

.gridwrap{overflow-x:auto;border-top:2px solid var(--ink);border-bottom:2px solid var(--ink);padding:.2rem 0}
table.grid{border-collapse:separate;border-spacing:0;font-family:var(--sans);font-size:.85rem;min-width:900px;width:100%}
table.grid th{font-weight:600;text-align:left;padding:.5rem .4rem;vertical-align:bottom}
table.grid thead th{border-bottom:1px solid var(--rule-strong);font-size:.82rem;line-height:1.15;max-width:80px}
table.grid th.colhead{cursor:pointer;text-decoration:underline dotted var(--rule-strong);text-underline-offset:3px}
table.grid th.colhead:hover{background:color-mix(in srgb, var(--ink) 5%, transparent)}
table.grid th.colhead.selected{box-shadow:inset 0 -3px 0 var(--ink)}
.idl{margin:0}
.idl dt{font-family:var(--sans);font-size:.78rem;font-weight:600;color:var(--muted);margin-top:.5rem}
.idl dd{margin:0;line-height:1.5}
.ilist{list-style:none;margin:1rem 0 0;padding:0}
.ilist li{border-top:1px solid var(--rule);padding:1rem 0 1.1rem}
.ilist h3{margin:0 0 .15rem}
.ilist .note{font-family:var(--sans);color:var(--muted);font-size:.95rem;margin:0 0 .3rem}
.ilist .idl{max-width:76ch;font-family:var(--sans);font-size:.95rem}
.ilist .cols{font-family:var(--sans);font-size:.85rem;color:var(--muted);margin:.5rem 0 0}
table.grid th.rowhead{width:150px;border-right:1px solid var(--rule-strong);padding-right:.6rem;font-size:.9rem}
table.grid th.rowhead small{display:block;font-weight:400;color:var(--muted);font-size:.74rem;line-height:1.2;margin-top:.15rem}
td.cell{border-bottom:1px solid var(--rule);height:64px;min-width:72px;padding:.35rem .3rem;vertical-align:top;cursor:pointer;position:relative}
td.cell:hover{background:color-mix(in srgb, var(--ink) 5%, transparent)}
td.cell.selected{box-shadow:inset 0 0 0 2px var(--ink)}
td.cell.ruled{background:repeating-linear-gradient(45deg,var(--hatch) 0 2px,transparent 2px 6px)}
td.cell.open::after{content:"";position:absolute;inset:6px;border:1.5px dotted var(--rule-strong);border-radius:3px;pointer-events:none}
.dots{display:flex;flex-wrap:wrap;gap:3px;align-content:flex-start}
.dot{width:13px;height:13px;border-radius:2px;display:inline-flex;align-items:center;justify-content:center;font-size:8.5px;font-weight:600;color:#fff;line-height:1}
.dot.concept-item{border-radius:50%}
.dot.b9{background:var(--b9)} .dot.b78{background:var(--b78)} .dot.b46{background:var(--b46)} .dot.b13{background:var(--b13)}
.dot.dub{background:transparent;border:1.5px dashed var(--dub);color:var(--dub)}
.dot.dimmed{opacity:.18}

.drawer{background:var(--panel);border:1px solid var(--rule);border-radius:6px;padding:1rem 1.2rem;margin:1rem 0 0;font-family:var(--sans);font-size:.98rem}
.drawer h3{margin:.1rem 0 .4rem;font-family:var(--serif)}
.drawer .empty{color:var(--muted)}
.item{padding:.55rem 0;border-top:1px solid var(--rule)}
.item:first-of-type{border-top:0}
.item b{font-weight:600}
.tag{display:inline-block;font-size:.78rem;padding:.05rem .45rem;border-radius:3px;color:#fff;margin-right:.4rem;vertical-align:middle;font-family:var(--sans)}
.tag.b9{background:var(--b9)} .tag.b78{background:var(--b78)} .tag.b46{background:var(--b46)} .tag.b13{background:var(--b13)}
.tag.dub{background:transparent;color:var(--dub);border:1.5px dashed var(--dub)}
.tag.kind{background:transparent;color:var(--muted);border:1px solid var(--rule-strong)}
.tag.verify{background:transparent;color:var(--warn);border:1px solid var(--warn)}

.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:1rem;margin:1.2rem 0 0;font-family:var(--sans)}
.stat{border-top:1px solid var(--rule-strong);padding-top:.5rem}
.stat b{font-size:1.6rem;font-weight:600;font-family:var(--serif);display:block;line-height:1.1}
.stat span{color:var(--muted);font-size:.9rem}

.toolbar{display:flex;flex-wrap:wrap;gap:.6rem;align-items:center;margin:.6rem 0 .8rem;font-family:var(--sans)}
input[type=search]{font-family:var(--sans);font-size:.95rem;background:var(--panel);color:var(--ink);border:1px solid var(--rule-strong);border-radius:4px;padding:.35rem .6rem;min-width:240px}
.tblwrap{overflow-x:auto}
table.cat{border-collapse:collapse;width:100%;font-family:var(--sans);font-size:.9rem;min-width:820px}
table.cat th{text-align:left;font-weight:600;border-bottom:1px solid var(--ink);padding:.45rem .5rem;vertical-align:bottom}
table.cat td{border-bottom:1px solid var(--rule);padding:.5rem .5rem;vertical-align:top}
table.cat td.name{font-weight:600;min-width:180px}
table.cat td.sum{max-width:420px}
table.cat td.ex{color:var(--muted);max-width:220px}
table.cat td.axis{white-space:nowrap;color:var(--muted);font-size:.84rem}
table.cat td.trl{white-space:nowrap}
.trlnote{display:block;color:var(--muted);font-size:.8rem;font-weight:400;margin-top:.2rem;white-space:normal;max-width:260px}

.sheets{display:grid;grid-template-columns:repeat(auto-fill,minmax(340px,1fr));gap:1rem;margin-top:1rem}
.sheet{background:var(--panel);border:1px solid var(--rule);border-radius:6px;padding:1rem 1.1rem 1.1rem;display:flex;flex-direction:column;gap:.45rem}
.sheet h3{margin:0;font-size:1.15rem;line-height:1.2}
.sheet .meta{font-family:var(--sans);font-size:.85rem;color:var(--muted);display:flex;flex-wrap:wrap;gap:.35rem .6rem;align-items:center}
.sheet p{font-size:.95rem;margin:0;line-height:1.5}
.sheet dt{font-family:var(--sans);font-size:.78rem;font-weight:600;color:var(--muted);margin-top:.35rem}
.sheet dd{margin:0;font-size:.95rem;line-height:1.5}
.sheet dl{margin:0}
.sheet.dubious{border-style:dashed;opacity:.9}
.novelty{font-family:var(--sans);font-size:.78rem;border:1px solid var(--rule-strong);border-radius:3px;padding:.02rem .4rem}

.scale{font-family:var(--sans);font-size:.92rem;border-collapse:collapse;max-width:76ch;margin-top:.6rem}
.scale td{border-top:1px solid var(--rule);padding:.35rem .6rem .35rem 0;vertical-align:top}
.scale td:first-child{font-weight:600;white-space:nowrap;width:5.5rem}
.filterlist{margin:1rem 0 0;padding:0;list-style:none;max-width:76ch}
.filterlist li{border-top:1px solid var(--rule);padding:.8rem 0}
.filterlist li b{font-weight:600;display:block;margin-bottom:.15rem}

footer{margin-top:4rem;border-top:1px solid var(--rule-strong);padding-top:1rem;font-family:var(--sans);font-size:.88rem;color:var(--muted);max-width:76ch}
@media (max-width:720px){
  body{font-size:16px}
  .wrap{padding:1.5rem .9rem 4rem}
  td.cell{height:54px;min-width:60px}
  table.grid th.rowhead{width:110px}
}
@media (prefers-reduced-motion: no-preference){
  td.cell,button.chip{transition:background .12s ease}
}
</style>
</head>
<body>
<div class="wrap">
<header>
  <h1>Phenomenology of remote sensing from space</h1>
  <p class="lede">Every established technique placed on a grid of <em>source</em> × <em>physical interaction</em>, tagged by viewing geometry, receiver type and technology readiness level. Each marker carries its TRL. Hatched cells are ruled out by physics; dotted cells are open; circles are concepts that could fill them.</p>
  <p class="caveat" id="caveat"></p>
</header>

<section id="grid-section">
  <h2>The grid</h2>
  <div class="legend" aria-label="Legend" id="legend"></div>
  <div class="filters" role="group" aria-label="Filters">
    <label>Geometry <select id="f-geom"><option value="">any</option></select></label>
    <label>Receiver <select id="f-rec"><option value="">any</option></select></label>
    <label>TRL at most <select id="f-trl"><option value="9">9</option><option value="8">8</option><option value="7">7</option><option value="6">6</option><option value="5">5</option><option value="4">4</option><option value="3">3</option><option value="2">2</option><option value="1">1</option></select></label>
    <button class="chip" id="f-concepts" aria-pressed="true">Show concepts</button>
    <button class="chip" id="f-clear" aria-pressed="false">Clear selection</button>
  </div>
  <div class="gridwrap"><table class="grid" id="grid" aria-describedby="caveat"></table></div>
  <div class="drawer" id="drawer"><span class="empty">Select a cell to see what fills it, or why it is empty.</span></div>
  <div class="stats" id="stats"></div>
</section>

<section id="catalog-section">
  <h2>Established techniques</h2>
  <p class="muted sans" style="font-size:.95rem">One row per technique with its TRL for use from space. Interactions can be several; the grid counts each. Search matches names, summaries and example missions.</p>
  <div class="toolbar">
    <input type="search" id="q" placeholder="Search techniques" aria-label="Search techniques">
    <label class="sans" style="font-size:.92rem">Source <select id="c-src"><option value="">all</option></select></label>
    <label class="sans" style="font-size:.92rem">TRL band <select id="c-band"><option value="">all</option></select></label>
    <button class="chip" id="c-verify" aria-pressed="false">Only items to verify</button>
    <span class="muted sans" id="c-count" style="font-size:.9rem"></span>
  </div>
  <div class="tblwrap"><table class="cat" id="cat"></table></div>
</section>

<section id="concept-section">
  <h2>Concepts for the empty cells</h2>
  <p>Each sheet states the physics, the application, the failure mode most likely to kill it, and the one calculation to run first. TRL here is for the concept as a space application; several have TRL 9 components and a TRL 2 application. The novelty tag is my confidence that the concept has <em>not</em> flown.</p>
  <div class="toolbar">
    <label class="sans" style="font-size:.92rem">Sort <select id="s-sort"><option value="id">list order</option><option value="trl-desc">TRL, highest first</option><option value="trl-asc">TRL, lowest first</option><option value="novelty">novelty, high first</option></select></label>
    <button class="chip" id="s-dubious" aria-pressed="true">Include physics-dubious</button>
  </div>
  <div class="sheets" id="sheets"></div>
</section>

<section id="interaction-section">
  <h2>The twelve interactions</h2>
  <p>The column axis of the grid. Each is a distinct way a field or photon stream can be changed by the target, and each has its own observable, its own retrievable quantities and its own governing limit. Selecting a column header in the grid shows the same text with the techniques that use it.</p>
  <p class="muted sans" style="font-size:.92rem" id="inote"></p>
  <ul class="ilist" id="ilist"></ul>
</section>

<section id="trl-section">
  <h2>The TRL scale as applied here</h2>
  <p>NASA definitions. The judgment is for the <em>technique used from space in its stated application</em>, which is why ground-operational Raman lidar sits at TRL 3 and a calibration concept built entirely from flown components sits at TRL 2.</p>
  <table class="scale" id="scale"></table>
</section>

<section id="filters-section">
  <h2>Four physical filters</h2>
  <p>Run these before any concept study. They decide which empty cells are empty for good reason.</p>
  <ul class="filterlist" id="filterlist"></ul>
</section>

<section id="method-section">
  <h2>How to read and extend this</h2>
  <p>The grid is a generative device, not a taxonomy. A technique is a choice of source, one or more interactions, a viewing geometry and a receiver class; enumerating the product space and asking why each cell is empty produces the concept list. The receiver axis is where the field is thinnest: almost everything is incoherent intensity or coherent heterodyne, with SMOS and geodetic VLBI the only TRL 9 correlation receivers and Swarm's helium magnetometer the only TRL 9 quantum sensor in routine Earth service.</p>
  <p>The TRL column changes the reading of the concept list. Sorted by TRL, the top is atom-interferometer gradiometry (5), TIR polarimetry (5), and a cluster at 4 (Rydberg, vibrometry, chronometric geodesy, Brillouin, Rayleigh–Brillouin, Ku/Ka reflectometry); these are IRAD-shaped because the step to 6 is an airborne or hosted-payload demonstration, not new physics. The bottom (2) is where the cheapest verdicts live: C04 and C17 each move from 2 to 3 with one calculation and one field night.</p>
  <p>Two things deserve scrutiny before anyone quotes a number. First, every item marked verify sits in a 2025–2027 pathfinder window; a single launch moves it. Second, any claim that a signals-of-opportunity concept is unpublished.</p>
</section>

<footer id="foot"></footer>
</div>

<script id="data" type="application/json">__DATA__</script>
<script>
(function(){
  const D = JSON.parse(document.getElementById('data').textContent);
  const $ = s => document.querySelector(s);
  const esc = s => String(s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
  const byId = (arr, id) => arr.find(x => x.id === id);
  const band = it => it.dubious ? 'dub' : (it.trl >= 9 ? 'b9' : it.trl >= 7 ? 'b78' : it.trl >= 4 ? 'b46' : 'b13');
  const bandName = it => it.dubious ? 'physics-dubious' : byId(D.trlBands, band(it)).name;
  const NOV_RANK = {high:0, medium:1, low:2};
  const trlTag = it => `<span class="tag ${band(it)}" title="${esc(bandName(it))}">TRL ${it.trl}</span>${it.verify?'<span class="tag verify">verify</span>':''}`;

  document.getElementById('caveat').textContent = D.meta.caveat;
  document.getElementById('foot').textContent = D.meta.title + '. Compiled ' + D.meta.date + ', TRL revision ' + D.meta.updated + '. Working reference; extend the catalog and re-run the grid rather than trusting any single cell.';

  // legend
  let lg = '';
  D.trlBands.forEach(b => lg += `<span><i class="sw" style="background:var(--${b.id})"></i>${esc(b.name)}</span>`);
  lg += '<span><i class="sw dub"></i>Physics-dubious</span><span><i class="sw hatch"></i>Ruled out</span><span><i class="sw open"></i>Open cell</span><span>■ established &nbsp; ● concept</span>';
  $('#legend').innerHTML = lg;

  const fGeom = $('#f-geom'), fRec = $('#f-rec'), fTrl = $('#f-trl'), fConcepts = $('#f-concepts');
  D.geometries.forEach(g => fGeom.insertAdjacentHTML('beforeend', `<option value="${g.id}">${esc(g.name)}</option>`));
  D.receivers.forEach(r => fRec.insertAdjacentHTML('beforeend', `<option value="${r.id}">${esc(r.name)}</option>`));
  let state = {geom:'', rec:'', trl:9, concepts:true, sel:null, col:null};

  function itemMatches(it){
    if (state.geom && !it.geom.includes(state.geom)) return false;
    if (state.rec && it.rec !== state.rec) return false;
    if (it.trl > state.trl) return false;
    return true;
  }
  function cellItems(src, int){
    const t = D.techniques.filter(x => x.src === src && x.int.includes(int)).map(x => ({...x, kind:'technique'}));
    const c = state.concepts ? D.concepts.filter(x => x.src.includes(src) && x.int.includes(int)).map(x => ({...x, kind:'concept'})) : [];
    return t.concat(c);
  }

  function renderGrid(){
    const tbl = $('#grid');
    let h = '<thead><tr><th class="rowhead"></th>';
    D.interactions.forEach(i => h += `<th class="colhead${state.col===i.id?' selected':''}" tabindex="0" role="button" data-int="${i.id}" title="${esc(i.name)}: ${esc(i.note)}">${esc(i.short)}</th>`);
    h += '</tr></thead><tbody>';
    D.sources.forEach(s => {
      h += `<tr><th class="rowhead" scope="row">${esc(s.name)}<small>${esc(s.note)}</small></th>`;
      D.interactions.forEach(i => {
        const key = s.id + '|' + i.id;
        const items = cellItems(s.id, i.id);
        const note = D.cellNotes[key];
        const cls = ['cell'];
        if (note && note.kind === 'ruled' && items.length === 0) cls.push('ruled');
        if (items.length === 0 && (!note || note.kind === 'open')) cls.push('open');
        if (state.sel === key) cls.push('selected');
        let dots = '';
        items.sort((a,b)=>b.trl-a.trl).forEach(it => {
          const dim = itemMatches(it) ? '' : ' dimmed';
          dots += `<i class="dot ${band(it)}${it.kind==='concept'?' concept-item':''}${dim}" title="${esc(it.name)} (TRL ${it.trl})">${it.trl}</i>`;
        });
        h += `<td class="${cls.join(' ')}" tabindex="0" role="button" data-key="${key}" aria-label="${esc(s.name)} × ${esc(i.name)}: ${items.length} item${items.length===1?'':'s'}"><div class="dots">${dots}</div></td>`;
      });
      h += '</tr>';
    });
    h += '</tbody>';
    tbl.innerHTML = h;
    tbl.querySelectorAll('td.cell').forEach(td => {
      const open = () => { state.sel = td.dataset.key; state.col = null; renderGrid(); renderDrawer(); };
      td.addEventListener('click', open);
      td.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); open(); } });
    });
    tbl.querySelectorAll('th.colhead').forEach(th => {
      const open = () => { state.col = th.dataset.int; state.sel = null; renderGrid(); renderDrawer(); };
      th.addEventListener('click', open);
      th.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); open(); } });
    });
  }

  function interactionBlock(i, withItems){
    let h = `<dl class="idl"><dt>Mechanism</dt><dd>${esc(i.physics)}</dd><dt>What is measured</dt><dd>${esc(i.observable)}</dd><dt>What it yields</dt><dd>${esc(i.yields)}</dd><dt>Governing limit</dt><dd>${esc(i.limit)}</dd></dl>`;
    if (withItems){
      const t = D.techniques.filter(x => x.int.includes(i.id)).sort((a,b)=>b.trl-a.trl);
      const c = D.concepts.filter(x => x.int.includes(i.id)).sort((a,b)=>b.trl-a.trl);
      h += `<p class="muted" style="margin:.6rem 0 .2rem;font-size:.9rem">${t.length} technique${t.length===1?'':'s'} and ${c.length} concept${c.length===1?'':'s'} in this column</p>`;
      h += '<p style="font-size:.9rem;margin:0;line-height:1.7">' + t.map(x => `${trlTag(x)}${esc(x.name)}`).join('<br>') + (c.length? '<br>' + c.map(x => `${trlTag(x)}<a href="#${x.id}">${esc(x.name)}</a>`).join('<br>') : '') + '</p>';
    }
    return h;
  }

  function renderDrawer(){
    const d = $('#drawer');
    if (state.col){
      const i = byId(D.interactions, state.col);
      d.innerHTML = `<h3>${esc(i.name)}</h3><p class="muted" style="margin:.1rem 0 .5rem">${esc(i.note)}</p>` + interactionBlock(i, true);
      return;
    }
    if (!state.sel){ d.innerHTML = '<span class="empty">Select a cell to see what fills it, or why it is empty. Select a column header to read the physics of that interaction.</span>'; return; }
    const [sid, iid] = state.sel.split('|');
    const s = byId(D.sources, sid), i = byId(D.interactions, iid);
    const items = cellItems(sid, iid).filter(itemMatches);
    const note = D.cellNotes[state.sel];
    let h = `<h3>${esc(s.name)} × ${esc(i.name)}</h3>`;
    if (note) h += `<p class="muted" style="margin:.2rem 0 .6rem"><span class="tag kind">${esc(note.kind === 'ruled' ? 'ruled out' : note.kind === 'open' ? 'open' : note.kind)}</span>${esc(note.text)}</p>`;
    if (!items.length && !note) h += '<p class="muted">Nothing filed here and no annotation. Either an oversight or an unexplored combination.</p>';
    items.sort((a,b) => (a.kind===b.kind?0:(a.kind==='technique'?-1:1)) || b.trl-a.trl);
    items.forEach(it => {
      const geoms = it.geom.map(g => byId(D.geometries, g).name).join(', ');
      const rec = byId(D.receivers, it.rec).name;
      const noteLine = it.trlNote ? `<br><span class="muted" style="font-size:.86rem">TRL note: ${esc(it.trlNote)}</span>` : '';
      if (it.kind === 'technique'){
        h += `<div class="item">${trlTag(it)}<b>${esc(it.name)}</b> <span class="muted">${esc(it.id)}</span><br>${esc(it.summary)}<br><span class="muted" style="font-size:.86rem">${esc(it.examples)} · ${esc(it.band)} · ${esc(geoms)} · ${esc(rec)}</span>${noteLine}</div>`;
      } else {
        h += `<div class="item">${trlTag(it)}<b>${esc(it.name)}</b> <span class="muted">${esc(it.id)} · concept</span><br>${esc(it.physics)}${noteLine}<br><a href="#${it.id}">Open the concept sheet</a></div>`;
      }
    });
    d.innerHTML = h;
  }

  function renderStats(){
    const filled = new Set(), ruled = new Set(), open = new Set(), conceptOnly = new Set();
    D.sources.forEach(s => D.interactions.forEach(i => {
      const key = s.id+'|'+i.id;
      const anyT = D.techniques.filter(x => x.src===s.id && x.int.includes(i.id));
      const c = D.concepts.filter(x => x.src.includes(s.id) && x.int.includes(i.id) && !x.dubious);
      const note = D.cellNotes[key];
      if (anyT.length) filled.add(key);
      else if (c.length) conceptOnly.add(key);
      else if (note && note.kind==='ruled') ruled.add(key);
      else open.add(key);
    }));
    const total = D.sources.length * D.interactions.length;
    const n9 = D.techniques.filter(t=>t.trl>=9).length;
    const nVer = D.techniques.filter(t=>t.verify).length + D.concepts.filter(c=>c.verify).length;
    const medC = D.concepts.filter(c=>!c.dubious).map(c=>c.trl).sort((a,b)=>a-b);
    const med = medC[Math.floor(medC.length/2)];
    $('#stats').innerHTML = `
      <div class="stat"><b>${total}</b><span>cells in the grid</span></div>
      <div class="stat"><b>${filled.size}</b><span>cells with an established technique</span></div>
      <div class="stat"><b>${conceptOnly.size}</b><span>cells only a concept fills</span></div>
      <div class="stat"><b>${ruled.size}</b><span>ruled out by physics</span></div>
      <div class="stat"><b>${open.size}</b><span>open, no candidate yet</span></div>
      <div class="stat"><b>${D.techniques.length}</b><span>techniques catalogued, ${n9} at TRL 9</span></div>
      <div class="stat"><b>${D.concepts.length}</b><span>concepts, median TRL ${med}, ${D.concepts.filter(c=>c.dubious).length} flagged dubious</span></div>
      <div class="stat"><b>${nVer}</b><span>TRL values flagged to verify</span></div>`;
  }

  fGeom.addEventListener('change', e => { state.geom = e.target.value; renderGrid(); renderDrawer(); });
  fRec.addEventListener('change', e => { state.rec = e.target.value; renderGrid(); renderDrawer(); });
  fTrl.addEventListener('change', e => { state.trl = +e.target.value; renderGrid(); renderDrawer(); });
  fConcepts.addEventListener('click', () => { state.concepts = !state.concepts; fConcepts.setAttribute('aria-pressed', state.concepts); renderGrid(); renderDrawer(); });
  $('#f-clear').addEventListener('click', () => { state.sel = null; state.col = null; renderGrid(); renderDrawer(); });

  // catalog
  const cSrc = $('#c-src'), cBand = $('#c-band'), cVer = $('#c-verify'), q = $('#q');
  let onlyVerify = false;
  D.sources.forEach(s => cSrc.insertAdjacentHTML('beforeend', `<option value="${s.id}">${esc(s.name)}</option>`));
  D.trlBands.forEach(b => cBand.insertAdjacentHTML('beforeend', `<option value="${b.id}">${esc(b.name)}</option>`));
  function renderCat(){
    const term = q.value.trim().toLowerCase();
    const rows = D.techniques.filter(t => {
      if (cSrc.value && t.src !== cSrc.value) return false;
      if (cBand.value && band(t) !== cBand.value) return false;
      if (onlyVerify && !t.verify) return false;
      if (term && !(t.name+' '+t.summary+' '+t.examples+' '+t.band).toLowerCase().includes(term)) return false;
      return true;
    });
    let h = '<thead><tr><th>Technique</th><th>TRL</th><th>Source · interactions</th><th>Geometry · receiver</th><th>Summary</th><th>Examples · band</th></tr></thead><tbody>';
    rows.forEach(t => {
      h += `<tr><td class="name">${esc(t.name)}<br><span class="muted" style="font-weight:400;font-size:.8rem">${t.id}</span></td>
        <td class="trl">${trlTag(t)}${t.trlNote?`<span class="trlnote">${esc(t.trlNote)}</span>`:''}</td>
        <td class="axis">${esc(byId(D.sources,t.src).name)}<br>${t.int.map(i=>esc(byId(D.interactions,i).short)).join(', ')}</td>
        <td class="axis">${t.geom.map(g=>esc(byId(D.geometries,g).name)).join('; ')}<br>${esc(byId(D.receivers,t.rec).name)}</td>
        <td class="sum">${esc(t.summary)}</td>
        <td class="ex">${esc(t.examples)}<br>${esc(t.band)}</td></tr>`;
    });
    h += '</tbody>';
    $('#cat').innerHTML = h;
    $('#c-count').textContent = rows.length + ' of ' + D.techniques.length;
  }
  [cSrc, cBand].forEach(el => el.addEventListener('change', renderCat));
  cVer.addEventListener('click', () => { onlyVerify = !onlyVerify; cVer.setAttribute('aria-pressed', onlyVerify); renderCat(); });
  q.addEventListener('input', renderCat);

  // concept sheets
  const sSort = $('#s-sort'), sDub = $('#s-dubious');
  let showDub = true;
  function renderSheets(){
    let list = D.concepts.filter(c => showDub || !c.dubious);
    if (sSort.value === 'novelty') list = list.slice().sort((a,b)=>NOV_RANK[a.novelty]-NOV_RANK[b.novelty]);
    if (sSort.value === 'trl-desc') list = list.slice().sort((a,b)=>b.trl-a.trl);
    if (sSort.value === 'trl-asc') list = list.slice().sort((a,b)=>a.trl-b.trl);
    let h = '';
    list.forEach(c => {
      const srcs = c.src.map(s=>byId(D.sources,s).name).join(', ');
      const ints = c.int.map(i=>byId(D.interactions,i).name).join(', ');
      const geoms = c.geom.map(g=>byId(D.geometries,g).name).join(', ');
      h += `<article class="sheet ${c.dubious?'dubious':''}" id="${c.id}">
        <h3>${esc(c.name)}</h3>
        <div class="meta">${trlTag(c)}<span class="novelty">novelty ${esc(c.novelty)}</span><span>${esc(c.id)}</span></div>
        <div class="meta">${esc(srcs)} × ${esc(ints)} · ${esc(geoms)} · ${esc(byId(D.receivers,c.rec).name)}</div>
        <dl>
          ${c.trlNote?`<dt>TRL basis</dt><dd>${esc(c.trlNote)}</dd>`:''}
          <dt>Physics</dt><dd>${esc(c.physics)}</dd>
          <dt>Application</dt><dd>${esc(c.application)}</dd>
          <dt>What kills it</dt><dd>${esc(c.failure)}</dd>
          <dt>First calculation</dt><dd>${esc(c.first)}</dd>
          <dt>Prior art, as recalled</dt><dd>${esc(c.prior)}</dd>
        </dl></article>`;
    });
    $('#sheets').innerHTML = h;
  }
  sSort.addEventListener('change', renderSheets);
  sDub.addEventListener('click', () => { showDub = !showDub; sDub.setAttribute('aria-pressed', showDub); renderSheets(); });

  $('#scale').innerHTML = D.trlScale.map(r => `<tr><td>TRL ${r[0]}</td><td>${esc(r[1])}</td></tr>`).join('');
  $('#inote').textContent = D.interactionNote;
  $('#ilist').innerHTML = D.interactions.map(i => {
    const nT = D.techniques.filter(x => x.int.includes(i.id)).length;
    const nC = D.concepts.filter(x => x.int.includes(i.id)).length;
    const srcs = D.sources.filter(s => D.techniques.some(x => x.src===s.id && x.int.includes(i.id)) || D.concepts.some(x => x.src.includes(s.id) && x.int.includes(i.id))).map(s=>s.name).join(', ');
    return `<li id="int-${i.id}"><h3>${esc(i.name)}</h3><p class="note">${esc(i.note)}</p>${interactionBlock(i, false)}<p class="cols">${nT} technique${nT===1?'':'s'}, ${nC} concept${nC===1?'':'s'} · driven by: ${esc(srcs)}</p></li>`;
  }).join('');
  $('#filterlist').innerHTML = D.filters.map(f => `<li><b>${esc(f.name)}</b>${esc(f.text)}</li>`).join('');

  renderGrid(); renderDrawer(); renderStats(); renderCat(); renderSheets();
})();
</script>
</body>
</html>
"""
open(OUT / 'index.html', 'w').write(HTML.replace('__DATA__', DATA_JS))

# ---------------------------------------------------------------- XLSX
wb = Workbook()
FONT = 'Arial'
hdr_font = Font(name=FONT, bold=True, size=10)
body_font = Font(name=FONT, size=10)
blue_font = Font(name=FONT, size=10, color='0000FF')
hdr_fill = PatternFill('solid', fgColor='DCE3EC')
wrap = Alignment(wrap_text=True, vertical='top')

def sheet(ws, headers, rows, widths):
    ws.append(headers)
    for c in ws[1]:
        c.font = hdr_font; c.fill = hdr_fill; c.alignment = wrap
    for r in rows: ws.append(r)
    for row in ws.iter_rows(min_row=2):
        for c in row: c.font = body_font; c.alignment = wrap
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'B2'

name = {x['id']: x['name'] for x in D['sources'] + D['interactions'] + D['geometries'] + D['receivers']}

ws = wb.active; ws.title = 'README'
readme = [
    [D['meta']['title']],
    [D['meta']['subtitle']],
    ['Compiled ' + D['meta']['date'] + '; TRL revision ' + D['meta']['updated']],
    [''],
    [D['meta']['caveat']],
    [''],
    ['Sheets'],
    ['Catalog', 'One row per established technique with TRL (1–9), a note on the basis where it is not obvious, and a Verify flag. Blue cells are the ones to edit. Keep codes exactly as listed on Axes; write several interactions as |code1|code2| with pipes.'],
    ['Concepts', 'Novel concept sheets with TRL, physics, application, failure mode, first calculation, prior art and novelty confidence. Dubious = Y marks physics-dubious entries, which the Grid sheet excludes.'],
    ['Grid', 'Source × interaction, computed by formula from Catalog and Concepts: count of techniques, their maximum TRL, and count of concepts. Cell annotations follow.'],
    ['Axes', 'Code lists for source, interaction, geometry and receiver, and the NASA TRL definitions used.'],
    ['Filters', 'The four physical filters to run before any concept study.'],
    [''],
    ['Example row', 'Catalog row 2 (T01) shows the expected format for a new entry.'],
]
for r in readme: ws.append(r)
ws['A1'].font = Font(name=FONT, bold=True, size=14)
for row in ws.iter_rows(min_row=2):
    for c in row: c.font = body_font; c.alignment = wrap
ws['A7'].font = hdr_font
ws.column_dimensions['A'].width = 18; ws.column_dimensions['B'].width = 110

ws = wb.create_sheet('Axes')
rows = []
for kind, lst in [('source', D['sources']), ('interaction', D['interactions']), ('geometry', D['geometries']), ('receiver', D['receivers'])]:
    for x in lst: rows.append([kind, x['id'], x['name'], x.get('note', ''), x.get('physics', ''), x.get('observable', ''), x.get('yields', ''), x.get('limit', '')])
for lvl, defn in D['trlScale']:
    rows.append(['TRL', lvl, 'TRL %d' % lvl, defn, '', '', '', ''])
rows.append(['note', '', 'Classification conventions', D['interactionNote'], '', '', '', ''])
sheet(ws, ['Axis', 'Code', 'Name', 'Definition', 'Mechanism (interactions)', 'What is measured', 'What it yields', 'Governing limit'], rows, [12, 12, 36, 60, 70, 50, 60, 70])

ws = wb.create_sheet('Catalog')
rows = [[t['id'], t['name'], t['src'], '|'+'|'.join(t['int'])+'|', '|'+'|'.join(t['geom'])+'|', t['rec'], t['trl'], t['trlNote'], 'Y' if t['verify'] else '', t['band'], t['examples'], t['summary']] for t in D['techniques']]
sheet(ws, ['ID', 'Technique', 'Source code', 'Interaction codes (|a|b|)', 'Geometry codes (|a|b|)', 'Receiver code', 'TRL (1–9)', 'TRL basis', 'Verify', 'Band', 'Examples', 'Summary'], rows,
      [6, 34, 11, 20, 20, 11, 8, 44, 7, 22, 34, 70])
for row in ws.iter_rows(min_row=2, max_col=12):
    for c in row: c.font = blue_font
CAT_MAX = 400

ws = wb.create_sheet('Concepts')
rows = [[c['id'], c['name'], '|'+'|'.join(c['src'])+'|', '|'+'|'.join(c['int'])+'|', '|'+'|'.join(c['geom'])+'|', c['rec'], c['trl'], c['trlNote'], 'Y' if c['verify'] else '', 'Y' if c['dubious'] else '', c['novelty'], c['physics'], c['application'], c['failure'], c['first'], c['prior']] for c in D['concepts']]
sheet(ws, ['ID', 'Concept', 'Source codes (|a|b|)', 'Interaction codes (|a|b|)', 'Geometry codes (|a|b|)', 'Receiver code', 'TRL (1–9)', 'TRL basis', 'Verify', 'Dubious', 'Novelty confidence', 'Physics', 'Application', 'What kills it', 'First calculation', 'Prior art (as recalled)'], rows,
      [6, 34, 14, 18, 18, 10, 8, 40, 7, 8, 10, 60, 40, 46, 40, 40])
for row in ws.iter_rows(min_row=2, max_col=16):
    for c in row: c.font = blue_font
CON_MAX = 200

ws = wb.create_sheet('Grid')
def header_row(title):
    ws.append([title]); ws.cell(row=ws.max_row, column=1).font = hdr_font
    ws.append([''] + [i['short'] for i in D['interactions']])
    for c in ws[ws.max_row]: c.font = hdr_font; c.fill = hdr_fill; c.alignment = wrap

header_row('Established techniques per cell, counted from Catalog')
for s in D['sources']:
    ws.append([s['name']] + [f'=SUMPRODUCT((Catalog!$C$2:$C${CAT_MAX}="{s["id"]}")*ISNUMBER(SEARCH("|{i["id"]}|",Catalog!$D$2:$D${CAT_MAX})))' for i in D['interactions']])
ws.append([])
header_row('Maximum TRL of established techniques per cell (blank where none)')
for s in D['sources']:
    row = [s['name']]
    for i in D['interactions']:
        f = (f'=IF(SUMPRODUCT((Catalog!$C$2:$C${CAT_MAX}="{s["id"]}")*ISNUMBER(SEARCH("|{i["id"]}|",Catalog!$D$2:$D${CAT_MAX})))=0,"",'
             f'SUMPRODUCT(MAX((Catalog!$C$2:$C${CAT_MAX}="{s["id"]}")*ISNUMBER(SEARCH("|{i["id"]}|",Catalog!$D$2:$D${CAT_MAX}))*Catalog!$G$2:$G${CAT_MAX})))')
        row.append(f)
    ws.append(row)
ws.append([])
header_row('Concepts per cell (excluding physics-dubious), counted from Concepts')
for s in D['sources']:
    ws.append([s['name']] + [f'=SUMPRODUCT(ISNUMBER(SEARCH("|{s["id"]}|",Concepts!$C$2:$C${CON_MAX}))*ISNUMBER(SEARCH("|{i["id"]}|",Concepts!$D$2:$D${CON_MAX}))*(Concepts!$J$2:$J${CON_MAX}<>"Y"))' for i in D['interactions']])
ws.append([])
header_row('Maximum concept TRL per cell (excluding physics-dubious; blank where none)')
for s in D['sources']:
    row = [s['name']]
    for i in D['interactions']:
        cond = f'ISNUMBER(SEARCH("|{s["id"]}|",Concepts!$C$2:$C${CON_MAX}))*ISNUMBER(SEARCH("|{i["id"]}|",Concepts!$D$2:$D${CON_MAX}))*(Concepts!$J$2:$J${CON_MAX}<>"Y")'
        row.append(f'=IF(SUMPRODUCT({cond})=0,"",SUMPRODUCT(MAX({cond}*Concepts!$G$2:$G${CON_MAX})))')
    ws.append(row)
ws.append([])
ws.append(['Cell annotations']); ws.cell(row=ws.max_row, column=1).font = hdr_font
ws.append(['Source', 'Interaction', 'Kind', 'Note'])
for c in ws[ws.max_row]: c.font = hdr_font; c.fill = hdr_fill
for k, v in D['cellNotes'].items():
    a, b = k.split('|'); ws.append([name[a], name[b], v['kind'], v['text']])
for row in ws.iter_rows(min_row=3):
    for c in row:
        if c.font != hdr_font: c.font = body_font
        c.alignment = wrap
ws.column_dimensions['A'].width = 22
for i in range(2, len(D['interactions']) + 2): ws.column_dimensions[get_column_letter(i)].width = 9
ws.append([])
ws.append(['Blank TRL cells read as 0 in the MAX formulas; a non-numeric TRL entry will produce a formula error, so keep the TRL column to integers 1–9. Counting ranges cover Catalog rows 2–%d and Concepts rows 2–%d.' % (CAT_MAX, CON_MAX)])
ws.cell(row=ws.max_row, column=1).font = Font(name=FONT, italic=True, size=9)

ws = wb.create_sheet('Filters')
sheet(ws, ['Filter', 'What it decides'], [[f['name'], f['text']] for f in D['filters']], [26, 110])

wb.save(OUT / 'remote-sensing-phenomenology.xlsx')
print('built', len(D['techniques']), 'techniques,', len(D['concepts']), 'concepts ->', OUT)
