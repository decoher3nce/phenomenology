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
<script>try{var t=localStorage.getItem('phen-theme');if(t==='light'||t==='dark')document.documentElement.setAttribute('data-theme',t);}catch(e){}</script>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Source+Serif+4:ital,opsz,wght@0,8..60,400;0,8..60,600;0,8..60,700;1,8..60,400&family=Source+Sans+3:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
/* Palette: colorblind-safe by construction.
   TRL bands are ordinal, so they use ONE hue (blue) stepped in lightness only;
   lightness survives every color-vision deficiency, and each marker also prints
   its TRL digit. Orange is the single accent (verify flags, focus) and sits far
   from every blue under protan/deutan simulation. No red/green pairs anywhere.
   Ramps validated with the dataviz ordinal checker in both modes. */
:root{
  color-scheme: light;
  --bg:#F4F6F9; --panel:#FFFFFF; --panel-2:#F8FAFC; --cell:#F1F4F8;
  --ink:#0F1B2D; --ink-2:#334155; --muted:#5A6678; --rule:#E2E7EE; --rule-strong:#C3CCD8;
  --b9:#104281;  --b9-ink:#FFFFFF;
  --b78:#256ABF; --b78-ink:#FFFFFF;
  --b46:#5598E7; --b46-ink:#0F1B2D;
  --b13:#86B6EF; --b13-ink:#0F1B2D;
  --dub:#8A94A3; --hatch:#CDD4DE; --sel:#0F1B2D;
  --link:#1C5CAB; --flag:#C2540F; --flag-bg:#FDF1E9; --focus:#C2540F;
  --shadow:0 1px 2px rgba(15,27,45,.05),0 6px 20px -8px rgba(15,27,45,.12);
  --shadow-lift:0 2px 4px rgba(15,27,45,.06),0 14px 32px -12px rgba(15,27,45,.22);
  --nav-bg:rgba(244,246,249,.82);
  --serif:"Source Serif 4", Georgia, "Times New Roman", serif;
  --sans:"Source Sans 3", system-ui, -apple-system, "Segoe UI", sans-serif;
  --r:12px; --r-sm:8px;
}
@media (prefers-color-scheme: dark){
  :root:not([data-theme="light"]){
    color-scheme: dark;
    --bg:#10151C; --panel:#1A2029; --panel-2:#1F2631; --cell:#222A35;
    --ink:#E8EDF4; --ink-2:#C5CEDA; --muted:#9AA6B6; --rule:#2A3340; --rule-strong:#3F4B5B;
    --b9:#B7D3F6;  --b9-ink:#0F1B2D;
    --b78:#6DA7EC; --b78-ink:#0F1B2D;
    --b46:#3987E5; --b46-ink:#0F1B2D;
    --b13:#1C5CAB; --b13-ink:#FFFFFF;
    --dub:#7D8898; --hatch:#323C4A; --sel:#E8EDF4;
    --link:#86B6EF; --flag:#F08A4B; --flag-bg:#3A2619; --focus:#F08A4B;
    --shadow:0 1px 2px rgba(0,0,0,.3),0 8px 24px -10px rgba(0,0,0,.5);
    --shadow-lift:0 2px 6px rgba(0,0,0,.35),0 16px 36px -12px rgba(0,0,0,.6);
    --nav-bg:rgba(16,21,28,.82);
  }
}
:root[data-theme="dark"]{
  color-scheme: dark;
  --bg:#10151C; --panel:#1A2029; --panel-2:#1F2631; --cell:#222A35;
  --ink:#E8EDF4; --ink-2:#C5CEDA; --muted:#9AA6B6; --rule:#2A3340; --rule-strong:#3F4B5B;
  --b9:#B7D3F6;  --b9-ink:#0F1B2D;
  --b78:#6DA7EC; --b78-ink:#0F1B2D;
  --b46:#3987E5; --b46-ink:#0F1B2D;
  --b13:#1C5CAB; --b13-ink:#FFFFFF;
  --dub:#7D8898; --hatch:#323C4A; --sel:#E8EDF4;
  --link:#86B6EF; --flag:#F08A4B; --flag-bg:#3A2619; --focus:#F08A4B;
  --shadow:0 1px 2px rgba(0,0,0,.3),0 8px 24px -10px rgba(0,0,0,.5);
  --shadow-lift:0 2px 6px rgba(0,0,0,.35),0 16px 36px -12px rgba(0,0,0,.6);
  --nav-bg:rgba(16,21,28,.82);
}
*{box-sizing:border-box}
html{scroll-behavior:smooth;scroll-padding-top:72px}
html,body{margin:0}
body{background:var(--bg);color:var(--ink);font-family:var(--sans);font-size:16.5px;line-height:1.6;-webkit-font-smoothing:antialiased;text-rendering:optimizeLegibility}
a{color:var(--link);text-decoration-thickness:1px;text-underline-offset:2px}
a:hover{text-decoration-thickness:2px}
p{max-width:72ch}
.sans{font-family:var(--sans)}
.muted{color:var(--muted)}
:focus-visible{outline:2px solid var(--focus);outline-offset:2px;border-radius:4px}

/* top bar */
.topbar{position:sticky;top:0;z-index:30;background:var(--nav-bg);-webkit-backdrop-filter:saturate(1.4) blur(12px);backdrop-filter:saturate(1.4) blur(12px);border-bottom:1px solid var(--rule)}
.topbar .in{max-width:1240px;margin:0 auto;padding:.55rem 1.25rem;display:flex;align-items:center;gap:1rem}
.brand{font-family:var(--serif);font-weight:700;font-size:1.05rem;color:var(--ink);text-decoration:none;white-space:nowrap;display:inline-flex;align-items:center;gap:.55rem}
.brand .mark{width:22px;height:22px;border-radius:6px;background:linear-gradient(135deg,var(--b13),var(--b9));display:inline-block;box-shadow:inset 0 0 0 1px rgba(255,255,255,.15)}
.nav{display:flex;gap:.15rem;overflow-x:auto;scrollbar-width:none;flex:1;min-width:0}
.nav::-webkit-scrollbar{display:none}
.nav a{font-size:.9rem;font-weight:500;color:var(--ink-2);text-decoration:none;padding:.35rem .65rem;border-radius:999px;white-space:nowrap}
.nav a:hover{background:color-mix(in srgb,var(--ink) 7%,transparent);color:var(--ink)}
.nav a.active{background:var(--ink);color:var(--bg)}
.actions{display:flex;gap:.4rem;align-items:center}
.btn{font-family:var(--sans);font-size:.88rem;font-weight:600;display:inline-flex;align-items:center;gap:.4rem;background:var(--panel);color:var(--ink);border:1px solid var(--rule-strong);border-radius:999px;padding:.32rem .8rem;cursor:pointer;text-decoration:none;white-space:nowrap}
.btn:hover{border-color:var(--ink-2)}
.btn svg{width:15px;height:15px}

.wrap{max-width:1240px;margin:0 auto;padding:0 1.25rem 5rem}

/* hero */
.hero{padding:3.2rem 0 1.4rem}
.eyebrow{font-size:.8rem;font-weight:600;letter-spacing:.08em;text-transform:uppercase;color:var(--muted);margin:0 0 .8rem;display:flex;gap:.6rem;flex-wrap:wrap;align-items:center}
.eyebrow .pill{letter-spacing:.02em;text-transform:none;font-weight:600;background:var(--panel);border:1px solid var(--rule);border-radius:999px;padding:.12rem .6rem;color:var(--ink-2)}
h1{font-family:var(--serif);font-size:clamp(2.1rem,4.6vw,3.4rem);line-height:1.06;font-weight:700;margin:0 0 .9rem;letter-spacing:-.018em;max-width:24ch}
.lede{font-family:var(--serif);font-size:clamp(1.1rem,1.6vw,1.28rem);line-height:1.5;color:var(--ink-2);max-width:64ch;margin:0 0 1.4rem}
.callout{display:flex;gap:.8rem;align-items:flex-start;background:var(--flag-bg);border:1px solid color-mix(in srgb,var(--flag) 35%,transparent);border-radius:var(--r);padding:.85rem 1rem;max-width:84ch;font-size:.94rem;color:var(--ink-2);margin:0}
.callout svg{flex:none;width:20px;height:20px;color:var(--flag);margin-top:.1rem}
.callout p{margin:0;max-width:none}

h2{font-family:var(--serif);font-size:clamp(1.5rem,2.4vw,1.95rem);font-weight:700;margin:0 0 .4rem;letter-spacing:-.01em;line-height:1.15}
h3{font-family:var(--serif);font-size:1.12rem;font-weight:600;margin:1.2rem 0 .3rem;line-height:1.25}
section{margin-top:3.6rem}
.sec-head{display:flex;justify-content:space-between;align-items:flex-end;gap:1rem;flex-wrap:wrap;margin-bottom:.9rem}
.sec-head p{margin:.2rem 0 0;color:var(--muted)}
.kicker{font-size:.78rem;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:var(--link);margin:0 0 .35rem}
.card{background:var(--panel);border:1px solid var(--rule);border-radius:var(--r);box-shadow:var(--shadow)}

/* legend + filters */
.panel-head{display:flex;flex-wrap:wrap;gap:.8rem 1.4rem;justify-content:space-between;align-items:center;padding:.9rem 1.1rem;border-bottom:1px solid var(--rule)}
.legend{display:flex;flex-wrap:wrap;gap:.45rem .5rem;font-size:.86rem;align-items:center;margin:0}
.legend .lg{display:inline-flex;align-items:center;gap:.45rem;padding:.2rem .6rem .2rem .3rem;border:1px solid var(--rule);border-radius:999px;background:var(--panel-2);color:var(--ink-2);white-space:nowrap}
.sw{width:18px;height:18px;border-radius:5px;display:inline-flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;line-height:1;flex:none}
.sw.b9{background:var(--b9);color:var(--b9-ink)} .sw.b78{background:var(--b78);color:var(--b78-ink)} .sw.b46{background:var(--b46);color:var(--b46-ink)} .sw.b13{background:var(--b13);color:var(--b13-ink)}
.sw.hatch{background:repeating-linear-gradient(45deg,var(--hatch) 0 2px,transparent 2px 5px);border:1px solid var(--rule-strong)}
.sw.open{border:1.5px dashed var(--rule-strong)}
.sw.dub{border:1.5px dashed var(--dub);background:transparent}
.sw.round{border-radius:50%}
.sw.shape{background:var(--ink-2)}
.filters{display:flex;flex-wrap:wrap;gap:.5rem .8rem;font-size:.9rem;align-items:center;padding:.75rem 1.1rem;border-bottom:1px solid var(--rule);background:var(--panel-2)}
.filters label,.toolbar label{display:inline-flex;align-items:center;gap:.45rem;color:var(--muted);font-weight:500}
select,input[type=search]{font-family:var(--sans);font-size:.9rem;background:var(--panel);color:var(--ink);border:1px solid var(--rule-strong);border-radius:var(--r-sm);padding:.34rem .6rem;min-height:34px}
select{padding-right:1.8rem;appearance:none;-webkit-appearance:none;background-image:linear-gradient(45deg,transparent 50%,var(--muted) 50%),linear-gradient(135deg,var(--muted) 50%,transparent 50%);background-position:calc(100% - 14px) 52%,calc(100% - 9px) 52%;background-size:5px 5px;background-repeat:no-repeat}
input[type=search]{min-width:260px;padding-left:2rem;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%238A94A3' stroke-width='2.2' stroke-linecap='round'%3E%3Ccircle cx='11' cy='11' r='7'/%3E%3Cpath d='m20 20-3.5-3.5'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:.6rem 50%;background-size:15px}
button.chip{font-family:var(--sans);font-size:.88rem;font-weight:600;background:var(--panel);color:var(--ink-2);border:1px solid var(--rule-strong);border-radius:999px;padding:.32rem .8rem;cursor:pointer;min-height:34px;display:inline-flex;align-items:center;gap:.4rem}
button.chip::before{content:"";width:8px;height:8px;border-radius:50%;border:1.5px solid currentColor;opacity:.6}
button.chip[aria-pressed="true"]{background:var(--ink);color:var(--bg);border-color:var(--ink)}
button.chip[aria-pressed="true"]::before{background:currentColor;opacity:1}
button.chip.plain::before{display:none}
.spacer{flex:1}

/* grid */
.gridwrap{overflow-x:auto;padding:.5rem .6rem .7rem}
table.grid{border-collapse:separate;border-spacing:4px;font-size:.84rem;min-width:980px;width:100%}
table.grid th{font-weight:600;text-align:left;padding:.4rem .35rem;vertical-align:bottom}
table.grid thead th{font-size:.8rem;line-height:1.15;max-width:84px;color:var(--ink-2)}
table.grid th.colhead{cursor:pointer;border-radius:var(--r-sm)}
table.grid th.colhead span{border-bottom:1.5px dotted var(--rule-strong)}
table.grid th.colhead:hover{background:color-mix(in srgb,var(--ink) 6%,transparent);color:var(--ink)}
table.grid th.colhead.selected{background:var(--ink);color:var(--bg)}
table.grid th.colhead.selected span{border-color:transparent}
table.grid th.rowhead{width:158px;padding-right:.8rem;font-size:.9rem;position:sticky;left:0;background:var(--panel);z-index:2;color:var(--ink)}
table.grid th.rowhead small{display:block;font-weight:400;color:var(--muted);font-size:.74rem;line-height:1.25;margin-top:.2rem}
td.cell{background:var(--cell);border-radius:var(--r-sm);height:70px;min-width:74px;padding:.4rem .35rem;vertical-align:top;cursor:pointer;position:relative}
td.cell:hover{background:color-mix(in srgb,var(--ink) 9%,var(--cell))}
td.cell.selected{box-shadow:0 0 0 2px var(--sel)}
td.cell.ruled{background:repeating-linear-gradient(45deg,var(--hatch) 0 2px,transparent 2px 7px),var(--panel)}
td.cell.open{background:var(--panel)}
td.cell.open::after{content:"";position:absolute;inset:5px;border:1.5px dashed var(--rule-strong);border-radius:6px;pointer-events:none}
.dots{display:flex;flex-wrap:wrap;gap:3px;align-content:flex-start}
.dot{width:17px;height:17px;border-radius:4px;display:inline-flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;line-height:1;font-style:normal;box-shadow:0 0 0 1px color-mix(in srgb,var(--panel) 60%,transparent)}
.dot.concept-item{border-radius:50%}
.dot.b9{background:var(--b9);color:var(--b9-ink)} .dot.b78{background:var(--b78);color:var(--b78-ink)} .dot.b46{background:var(--b46);color:var(--b46-ink)} .dot.b13{background:var(--b13);color:var(--b13-ink)}
.dot.dub{background:transparent;border:1.5px dashed var(--dub);color:var(--dub);box-shadow:none}
.dot.dimmed{opacity:.16}
.tip{position:fixed;z-index:50;pointer-events:none;background:var(--ink);color:var(--bg);font-size:.82rem;line-height:1.35;padding:.45rem .65rem;border-radius:8px;max-width:280px;box-shadow:var(--shadow-lift);opacity:0;transform:translateY(4px)}
.tip.on{opacity:1;transform:none}
.tip b{display:block;font-weight:600}
.tip span{opacity:.75}

.drawer{margin:0;padding:1.1rem 1.3rem 1.2rem;border-top:1px solid var(--rule);font-size:.96rem;background:var(--panel)}
.drawer h3{margin:.1rem 0 .4rem;font-size:1.2rem}
.drawer .empty{color:var(--muted);display:flex;gap:.6rem;align-items:center}
.drawer .empty svg{width:18px;height:18px;flex:none}
.item{padding:.75rem 0;border-top:1px solid var(--rule)}
.item:first-of-type{border-top:0}
.item b{font-weight:600}
.tag{display:inline-flex;align-items:center;font-size:.76rem;font-weight:700;padding:.08rem .5rem;border-radius:999px;margin-right:.4rem;vertical-align:.08em;font-family:var(--sans);white-space:nowrap;line-height:1.5}
.tag.b9{background:var(--b9);color:var(--b9-ink)} .tag.b78{background:var(--b78);color:var(--b78-ink)} .tag.b46{background:var(--b46);color:var(--b46-ink)} .tag.b13{background:var(--b13);color:var(--b13-ink)}
.tag.dub{background:transparent;color:var(--muted);border:1.5px dashed var(--dub)}
.tag.kind{background:transparent;color:var(--muted);border:1px solid var(--rule-strong);font-weight:600}
.tag.verify{background:var(--flag-bg);color:var(--flag);border:1px solid color-mix(in srgb,var(--flag) 45%,transparent)}
.tag.verify::before{content:"\2691";margin-right:.25rem;font-size:.8em}

.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:.8rem;margin:1rem 0 0}
.stat{background:var(--panel);border:1px solid var(--rule);border-radius:var(--r);padding:.9rem 1rem;box-shadow:var(--shadow)}
.stat b{font-size:1.85rem;font-weight:700;font-family:var(--serif);display:block;line-height:1.05;letter-spacing:-.01em}
.stat span{color:var(--muted);font-size:.88rem;line-height:1.35;display:block;margin-top:.25rem}

/* catalog */
.toolbar{display:flex;flex-wrap:wrap;gap:.6rem .8rem;align-items:center;padding:.85rem 1.1rem;border-bottom:1px solid var(--rule);background:var(--panel-2);border-radius:var(--r) var(--r) 0 0}
.count{font-size:.86rem;color:var(--muted);font-variant-numeric:tabular-nums;margin-left:auto}
.tblwrap{overflow-x:auto}
table.cat{border-collapse:collapse;width:100%;font-size:.9rem;min-width:860px}
table.cat th{text-align:left;font-weight:700;font-size:.76rem;letter-spacing:.06em;text-transform:uppercase;color:var(--muted);border-bottom:1px solid var(--rule-strong);padding:.7rem .75rem;vertical-align:bottom;background:var(--panel)}
table.cat td{border-bottom:1px solid var(--rule);padding:.75rem .75rem;vertical-align:top}
table.cat tbody tr:hover td{background:var(--panel-2)}
table.cat tbody tr:last-child td{border-bottom:0}
table.cat td.name{font-weight:600;min-width:190px;color:var(--ink)}
table.cat td.sum{max-width:420px;color:var(--ink-2)}
table.cat td.ex{color:var(--muted);max-width:220px}
table.cat td.axis{white-space:nowrap;color:var(--muted);font-size:.84rem}
table.cat td.trl{white-space:nowrap}
.trlnote{display:block;color:var(--muted);font-size:.8rem;font-weight:400;margin-top:.35rem;white-space:normal;max-width:260px;line-height:1.4}
.id{font-size:.76rem;font-weight:600;color:var(--muted);font-variant-numeric:tabular-nums;letter-spacing:.02em}

/* concept sheets */
.sheets{display:grid;grid-template-columns:repeat(auto-fill,minmax(340px,1fr));gap:1rem;margin-top:1rem}
.sheet{background:var(--panel);border:1px solid var(--rule);border-radius:var(--r);padding:1.1rem 1.2rem 1.2rem;display:flex;flex-direction:column;gap:.5rem;box-shadow:var(--shadow);position:relative;overflow:hidden}
.sheet::before{content:"";position:absolute;inset:0 0 auto 0;height:4px;background:var(--band,var(--rule-strong))}
.sheet.b9{--band:var(--b9)} .sheet.b78{--band:var(--b78)} .sheet.b46{--band:var(--b46)} .sheet.b13{--band:var(--b13)}
.sheet:target{box-shadow:0 0 0 2px var(--focus),var(--shadow-lift)}
@media (hover:hover){.sheet{transition:box-shadow .18s ease,transform .18s ease}.sheet:hover{box-shadow:var(--shadow-lift);transform:translateY(-2px)}}
.sheet h3{margin:.2rem 0 0;font-size:1.18rem;line-height:1.22}
.sheet .meta{font-size:.84rem;color:var(--muted);display:flex;flex-wrap:wrap;gap:.35rem .5rem;align-items:center}
.sheet p{font-size:.95rem;margin:0;line-height:1.5}
.sheet dt{font-size:.72rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:var(--muted);margin-top:.6rem}
.sheet dd{margin:.1rem 0 0;font-size:.94rem;line-height:1.5;color:var(--ink-2)}
.sheet dl{margin:0}
.sheet.dubious{border-style:dashed;--band:repeating-linear-gradient(90deg,var(--dub) 0 8px,transparent 8px 14px)}
.sheet.dubious::before{background:var(--band)}
.novelty{font-size:.76rem;font-weight:600;border:1px solid var(--rule-strong);border-radius:999px;padding:.05rem .55rem;color:var(--ink-2)}

/* interactions */
.ilist{list-style:none;margin:1rem 0 0;padding:0;display:grid;grid-template-columns:repeat(auto-fill,minmax(440px,1fr));gap:1rem}
.ilist li{background:var(--panel);border:1px solid var(--rule);border-radius:var(--r);padding:1.1rem 1.2rem;box-shadow:var(--shadow)}
.ilist h3{margin:0 0 .15rem}
.ilist .note{color:var(--muted);font-size:.93rem;margin:0 0 .4rem}
.idl{margin:0}
.idl dt{font-size:.72rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:var(--muted);margin-top:.6rem}
.idl dd{margin:.1rem 0 0;line-height:1.5;color:var(--ink-2)}
.ilist .idl{font-size:.94rem}
.ilist .cols{font-size:.84rem;color:var(--muted);margin:.8rem 0 0;padding-top:.6rem;border-top:1px solid var(--rule)}

/* TRL scale + filters */
.split{display:grid;grid-template-columns:1fr 1fr;gap:1rem;align-items:start}
.scale{font-size:.92rem;border-collapse:collapse;width:100%}
.scale td{border-top:1px solid var(--rule);padding:.6rem .8rem;vertical-align:top}
.scale tr:first-child td{border-top:0}
.scale td:first-child{font-weight:700;white-space:nowrap;width:7rem}
.scale td:first-child .sw{margin-right:.5rem;vertical-align:-.25em}
.filterlist{margin:0;padding:0;list-style:none}
.filterlist li{border-top:1px solid var(--rule);padding:.85rem 1.1rem}
.filterlist li:first-child{border-top:0}
.filterlist li b{font-weight:700;display:block;margin-bottom:.15rem;font-family:var(--serif);font-size:1.02rem}
.prose{padding:1.3rem 1.5rem;font-family:var(--serif);font-size:1.05rem;line-height:1.65}
.prose p{margin:0 0 1rem;max-width:76ch}
.prose p:last-child{margin-bottom:0}

footer{margin-top:4rem;border-top:1px solid var(--rule);padding-top:1.2rem;font-size:.86rem;color:var(--muted);max-width:84ch}
@media (max-width:860px){ .split{grid-template-columns:1fr} .ilist{grid-template-columns:1fr} }
@media (max-width:720px){
  body{font-size:16px}
  .topbar .in,.wrap{padding-left:16px;padding-right:16px}
  .brand .txt{display:none}
  .btn .lbl{display:none}
  .btn{padding:.42rem .55rem}
  .topbar .in{gap:.6rem}
  .hero{padding-top:2rem}
  td.cell{height:58px;min-width:62px}
  table.grid th.rowhead{width:116px}
  .sheets{grid-template-columns:1fr}
  input[type=search]{min-width:0;width:100%}
  .count{margin-left:0}
}
@media (prefers-reduced-motion: no-preference){
  td.cell,button.chip,.nav a,.btn{transition:background .14s ease,color .14s ease,border-color .14s ease}
  .tip{transition:opacity .12s ease,transform .12s ease}
}
@media (forced-colors: active){ .dot,.sw,.tag{border:1px solid CanvasText} }
</style>
</head>
<body>
<div class="topbar"><div class="in">
  <a class="brand" href="#top"><i class="mark" aria-hidden="true"></i><span class="txt">Phenomenology</span></a>
  <nav class="nav" aria-label="Sections">
    <a href="#grid-section">Grid</a><a href="#catalog-section">Catalog</a><a href="#concept-section">Concepts</a><a href="#interaction-section">Interactions</a><a href="#trl-section">TRL</a><a href="#method-section">Method</a>
  </nav>
  <div class="actions">
    <a class="btn" href="remote-sensing-phenomenology.xlsx" download="remote-sensing-phenomenology.xlsx" aria-label="Download the workbook"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 3v12m0 0 4.5-4.5M12 15l-4.5-4.5M4 19h16"/></svg><span class="lbl">Workbook</span></a>
    <button class="btn" id="theme" type="button" aria-label="Color theme"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-hidden="true"><circle cx="12" cy="12" r="8"/><path d="M12 4a8 8 0 0 0 0 16z" fill="currentColor"/></svg><span class="lbl" id="theme-label">Auto</span></button>
  </div>
</div></div>
<div class="wrap" id="top">
<header class="hero">
  <p class="eyebrow"><span>Reference</span><span class="pill" id="updated"></span></p>
  <h1>Phenomenology of remote sensing from space</h1>
  <p class="lede">Every established technique placed on a grid of <em>source</em> × <em>physical interaction</em>, tagged by viewing geometry, receiver type and technology readiness level. Each marker carries its TRL. Hatched cells are ruled out by physics; dashed cells are open; circles are concepts that could fill them.</p>
  <div class="callout" role="note"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 3 2 20h20L12 3z"/><path d="M12 10v4m0 3h.01"/></svg><p id="caveat"></p></div>
</header>

<section id="grid-section">
  <div class="sec-head"><div><p class="kicker">Source × interaction</p><h2>The grid</h2><p>Select a cell for what fills it or why it is empty. Select a column header for the physics of that interaction.</p></div></div>
  <div class="card">
    <div class="panel-head"><div class="legend" aria-label="Legend" id="legend"></div></div>
    <div class="filters" role="group" aria-label="Filters">
      <label>Geometry <select id="f-geom"><option value="">any</option></select></label>
      <label>Receiver <select id="f-rec"><option value="">any</option></select></label>
      <label>TRL at most <select id="f-trl"><option value="9">9</option><option value="8">8</option><option value="7">7</option><option value="6">6</option><option value="5">5</option><option value="4">4</option><option value="3">3</option><option value="2">2</option><option value="1">1</option></select></label>
      <button class="chip" id="f-concepts" aria-pressed="true">Show concepts</button>
      <span class="spacer"></span>
      <button class="chip plain" id="f-clear" aria-pressed="false">Clear selection</button>
    </div>
    <div class="gridwrap"><table class="grid" id="grid" aria-describedby="caveat"></table></div>
    <div class="drawer" id="drawer" aria-live="polite"></div>
  </div>
  <div class="stats" id="stats"></div>
</section>

<section id="catalog-section">
  <div class="sec-head"><div><p class="kicker">Catalog</p><h2>Established techniques</h2><p>One row per technique with its TRL for use from space. Interactions can be several; the grid counts each. Search matches names, summaries and example missions.</p></div></div>
  <div class="card">
    <div class="toolbar">
      <input type="search" id="q" placeholder="Search techniques" aria-label="Search techniques">
      <label>Source <select id="c-src"><option value="">all</option></select></label>
      <label>TRL band <select id="c-band"><option value="">all</option></select></label>
      <button class="chip" id="c-verify" aria-pressed="false">Only items to verify</button>
      <span class="count" id="c-count"></span>
    </div>
    <div class="tblwrap"><table class="cat" id="cat"></table></div>
  </div>
</section>

<section id="concept-section">
  <div class="sec-head"><div><p class="kicker">Candidates</p><h2>Concepts for the empty cells</h2><p>Each sheet states the physics, the application, the failure mode most likely to kill it, and the one calculation to run first. TRL here is for the concept as a space application; several have TRL 9 components and a TRL 2 application. The novelty tag is my confidence that the concept has <em>not</em> flown.</p></div></div>
  <div class="card"><div class="toolbar" style="border-bottom:0;border-radius:var(--r)">
    <label>Sort <select id="s-sort"><option value="id">list order</option><option value="trl-desc">TRL, highest first</option><option value="trl-asc">TRL, lowest first</option><option value="novelty">novelty, high first</option></select></label>
    <button class="chip" id="s-dubious" aria-pressed="true">Include physics-dubious</button>
  </div></div>
  <div class="sheets" id="sheets"></div>
</section>

<section id="interaction-section">
  <div class="sec-head"><div><p class="kicker">Column axis</p><h2>The twelve interactions</h2><p>Each is a distinct way a field or photon stream can be changed by the target, and each has its own observable, its own retrievable quantities and its own governing limit. Selecting a column header in the grid shows the same text with the techniques that use it.</p></div></div>
  <p class="muted" style="font-size:.92rem" id="inote"></p>
  <ul class="ilist" id="ilist"></ul>
</section>

<section id="trl-section">
  <div class="split">
    <div>
      <div class="sec-head"><div><p class="kicker">Readiness</p><h2>The TRL scale as applied here</h2><p>NASA definitions. The judgment is for the <em>technique used from space in its stated application</em>, which is why ground-operational Raman lidar sits at TRL 3 and a calibration concept built entirely from flown components sits at TRL 2.</p></div></div>
      <div class="card"><table class="scale" id="scale"></table></div>
    </div>
    <div id="filters-section">
      <div class="sec-head"><div><p class="kicker">Before any study</p><h2>Four physical filters</h2><p>Run these before any concept study. They decide which empty cells are empty for good reason.</p></div></div>
      <ul class="filterlist card" id="filterlist"></ul>
    </div>
  </div>
</section>

<section id="method-section">
  <div class="sec-head"><div><p class="kicker">Method</p><h2>How to read and extend this</h2></div></div>
  <div class="card prose">
  <p>The grid is a generative device, not a taxonomy. A technique is a choice of source, one or more interactions, a viewing geometry and a receiver class; enumerating the product space and asking why each cell is empty produces the concept list. The receiver axis is where the field is thinnest: almost everything is incoherent intensity or coherent heterodyne, with SMOS and geodetic VLBI the only TRL 9 correlation receivers and Swarm's helium magnetometer the only TRL 9 quantum sensor in routine Earth service.</p>
  <p>The TRL column changes the reading of the concept list. Sorted by TRL, the top is atom-interferometer gradiometry (5), TIR polarimetry (5), and a cluster at 4 (Rydberg, vibrometry, chronometric geodesy, Brillouin, Rayleigh–Brillouin, Ku/Ka reflectometry); these are IRAD-shaped because the step to 6 is an airborne or hosted-payload demonstration, not new physics. The bottom (2) is where the cheapest verdicts live: C04 and C17 each move from 2 to 3 with one calculation and one field night.</p>
  <p>Two things deserve scrutiny before anyone quotes a number. First, every item marked verify sits in a 2025–2027 pathfinder window; a single launch moves it. Second, any claim that a signals-of-opportunity concept is unpublished.</p>
  </div>
</section>

<footer id="foot"></footer>
</div>
<div class="tip" id="tip" role="tooltip"></div>
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
  const BAND_DIGIT = {b9:'9', b78:'7', b46:'5', b13:'2'};
  let lg = '';
  D.trlBands.forEach(b => lg += `<span class="lg"><i class="sw ${b.id}">${BAND_DIGIT[b.id]}</i>${esc(b.name)}</span>`);
  lg += '<span class="lg"><i class="sw shape"></i>Established</span><span class="lg"><i class="sw shape round"></i>Concept</span><span class="lg"><i class="sw dub round"></i>Physics-dubious</span><span class="lg"><i class="sw hatch"></i>Ruled out</span><span class="lg"><i class="sw open"></i>Open cell</span>';
  $('#legend').innerHTML = lg;
  $('#updated').textContent = 'Updated ' + D.meta.updated;

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
    D.interactions.forEach(i => h += `<th class="colhead${state.col===i.id?' selected':''}" tabindex="0" role="button" data-int="${i.id}" data-tipname="${esc(i.name)}" data-tipsub="${esc(i.note)}"><span>${esc(i.short)}</span></th>`);
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
          dots += `<i class="dot ${band(it)}${it.kind==='concept'?' concept-item':''}${dim}" data-tipname="${esc(it.name)}" data-tipsub="TRL ${it.trl} · ${esc(bandName(it))} · ${it.kind}${it.verify?' · verify':''}">${it.trl}</i>`;
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
    if (!state.sel){ d.innerHTML = '<span class="empty"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 4l7 17 2.5-7.5L21 11z"/></svg>Select a cell to see what fills it, or why it is empty. Select a column header to read the physics of that interaction.</span>'; return; }
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
      h += `<tr><td class="name">${esc(t.name)}<br><span class="id">${t.id}</span></td>
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
      h += `<article class="sheet ${band(c)} ${c.dubious?'dubious':''}" id="${c.id}">
        <h3>${esc(c.name)}</h3>
        <div class="meta">${trlTag(c)}<span class="novelty">novelty ${esc(c.novelty)}</span><span class="id">${esc(c.id)}</span></div>
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

  $('#scale').innerHTML = D.trlScale.map(r => { const b = band({trl:+r[0]}); return `<tr><td><i class="sw ${b}">${r[0]}</i>TRL ${r[0]}</td><td>${esc(r[1])}</td></tr>`; }).join('');
  $('#inote').textContent = D.interactionNote;
  $('#ilist').innerHTML = D.interactions.map(i => {
    const nT = D.techniques.filter(x => x.int.includes(i.id)).length;
    const nC = D.concepts.filter(x => x.int.includes(i.id)).length;
    const srcs = D.sources.filter(s => D.techniques.some(x => x.src===s.id && x.int.includes(i.id)) || D.concepts.some(x => x.src.includes(s.id) && x.int.includes(i.id))).map(s=>s.name).join(', ');
    return `<li id="int-${i.id}"><h3>${esc(i.name)}</h3><p class="note">${esc(i.note)}</p>${interactionBlock(i, false)}<p class="cols">${nT} technique${nT===1?'':'s'}, ${nC} concept${nC===1?'':'s'} · driven by: ${esc(srcs)}</p></li>`;
  }).join('');
  $('#filterlist').innerHTML = D.filters.map(f => `<li><b>${esc(f.name)}</b>${esc(f.text)}</li>`).join('');

  // hover tooltip for markers and column heads (event delegation; grid re-renders)
  const tip = $('#tip');
  const grid = $('#grid');
  function showTip(el, x, y){
    tip.innerHTML = `<b>${el.dataset.tipname}</b><span>${el.dataset.tipsub}</span>`;
    tip.classList.add('on');
    const w = tip.offsetWidth, h = tip.offsetHeight;
    let left = x + 14, top = y + 16;
    if (left + w > innerWidth - 8) left = x - w - 14;
    if (top + h > innerHeight - 8) top = y - h - 12;
    tip.style.left = Math.max(8, left) + 'px'; tip.style.top = Math.max(8, top) + 'px';
  }
  grid.addEventListener('mousemove', e => {
    const el = e.target.closest('[data-tipname]');
    if (el && el.dataset.tipname) showTip(el, e.clientX, e.clientY); else tip.classList.remove('on');
  });
  grid.addEventListener('mouseleave', () => tip.classList.remove('on'));
  addEventListener('scroll', () => tip.classList.remove('on'), {passive:true});

  // theme toggle: auto -> light -> dark
  const root = document.documentElement, themeBtn = $('#theme'), themeLbl = $('#theme-label');
  const modes = ['auto','light','dark'];
  let mode = root.getAttribute('data-theme') || 'auto';
  const paintTheme = () => { themeLbl.textContent = mode[0].toUpperCase() + mode.slice(1); themeBtn.setAttribute('aria-label', 'Color theme: ' + mode); };
  paintTheme();
  themeBtn.addEventListener('click', () => {
    mode = modes[(modes.indexOf(mode) + 1) % modes.length];
    if (mode === 'auto') root.removeAttribute('data-theme'); else root.setAttribute('data-theme', mode);
    try { if (mode === 'auto') localStorage.removeItem('phen-theme'); else localStorage.setItem('phen-theme', mode); } catch(e){}
    paintTheme();
  });

  // highlight the section in view
  const navLinks = [...document.querySelectorAll('.nav a')];
  if ('IntersectionObserver' in window){
    const io = new IntersectionObserver(entries => entries.forEach(en => {
      if (en.isIntersecting) navLinks.forEach(a => a.classList.toggle('active', a.getAttribute('href') === '#' + en.target.id));
    }), {rootMargin:'-45% 0px -50% 0px'});
    navLinks.forEach(a => { const t = document.querySelector(a.getAttribute('href')); if (t) io.observe(t); });
  }

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
