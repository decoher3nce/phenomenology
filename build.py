import json
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

import os, pathlib
ROOT = pathlib.Path(__file__).resolve().parent
OUT = pathlib.Path(os.environ.get('PHEN_OUT', ROOT / 'public')); OUT.mkdir(exist_ok=True)
D = json.load(open(os.environ.get('PHEN_DATA', ROOT / 'data.json')))
D.setdefault('references', [])
for _e in D['techniques'] + D['concepts']:
    _e.setdefault('refs', []); _e.setdefault('evidence', [])
for _n in D['cellNotes'].values():
    _n.setdefault('refs', [])
import re as _re
def xl(text):
    """Citation tokens {R12} become [R12] in the workbook; the References sheet resolves them."""
    return _re.sub(r'\{(R\d+)\}', r'[\1]', text or '')
def ev_summary(e):
    parts = []
    for v in e.get('evidence', []):
        what = v.get('name', '') + (' / ' + v['instrument'] if v.get('instrument') else '')
        parts.append(f"{v.get('kind','')}: {what} ({v.get('level','')}{', ' + v['dates'] if v.get('dates') else ''}) [{' '.join(v.get('refs', []))}]")
    return '; '.join(parts)
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
.filters label,.toolbar label{display:inline-flex;align-items:center;gap:.45rem;color:var(--muted);font-weight:500;max-width:100%;min-width:0}
select,input[type=search]{max-width:100%;font-family:var(--sans);font-size:.9rem;background:var(--panel);color:var(--ink);border:1px solid var(--rule-strong);border-radius:var(--r-sm);padding:.34rem .6rem;min-height:34px}
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
.toolbar select{max-width:min(100%,240px)}
@media (max-width:720px){.toolbar select{max-width:min(100%,200px)}}
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

/* citations, evidence, references */
.cite{display:inline-block;font-family:var(--sans);font-size:.68rem;font-weight:700;line-height:1.35;padding:0 .32rem;margin:0 .1rem 0 .12rem;border-radius:4px;background:color-mix(in srgb,var(--link) 13%,transparent);color:var(--link);text-decoration:none;vertical-align:.2em;letter-spacing:.02em;white-space:nowrap}
.cite:hover,.cite:focus-visible{background:var(--link);color:var(--bg);outline:none}
.srcs{display:flex;flex-wrap:wrap;gap:.25rem .3rem;align-items:center;margin-top:.45rem}
.srcs .lbl{font-size:.7rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:var(--muted);margin-right:.15rem}
.srcs .cite{vertical-align:0;font-size:.72rem}
.evlist{list-style:none;margin:.25rem 0 0;padding:0;display:grid;gap:.4rem}
.evlist li{font-size:.86rem;line-height:1.45;color:var(--ink-2);padding-left:.95rem;position:relative}
.evlist li::before{content:"";position:absolute;left:0;top:.55em;width:.42rem;height:.42rem;border-radius:1px;background:var(--rule-strong)}
.evlist li.mission::before{background:var(--b9);border-radius:1px}
.evlist li.program::before{background:var(--b78)}
.evlist li.system::before,.evlist li.subsystem::before,.evlist li.component::before,.evlist li.technology::before{background:var(--b46);border-radius:50%}
.evlist b{font-weight:600;color:var(--ink)}
.lvl{display:inline-block;font-size:.72rem;font-weight:600;border:1px solid var(--rule-strong);border-radius:999px;padding:0 .48rem;margin-right:.3rem;color:var(--ink-2);background:var(--panel-2);white-space:nowrap;line-height:1.5}
.kindlbl{font-size:.7rem;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:var(--muted);margin-right:.3rem}
details.ev{margin-top:.45rem}
details.ev summary{cursor:pointer;font-size:.8rem;font-weight:600;color:var(--link);list-style:none;display:inline-flex;align-items:center;gap:.35rem;user-select:none}
details.ev summary::-webkit-details-marker{display:none}
details.ev summary::before{content:"";width:.38rem;height:.38rem;border-right:1.5px solid currentColor;border-bottom:1.5px solid currentColor;transform:rotate(-45deg);margin-left:.1rem}
details.ev[open] summary::before{transform:rotate(45deg)}
details.ev .evlist{margin-top:.5rem;max-width:420px}
table.cat tr:target td{background:color-mix(in srgb,var(--focus) 10%,var(--panel));box-shadow:inset 3px 0 0 var(--focus)}
/* evidence section */
.tiers{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:.8rem;margin:0 0 1rem}
.tier{background:var(--panel);border:1px solid var(--rule);border-radius:var(--r);padding:.85rem 1rem;box-shadow:var(--shadow);font-size:.9rem;color:var(--ink-2);display:flex;gap:.7rem;align-items:flex-start}
.tier .sw{margin-top:.2rem}
.tier b{display:block;color:var(--ink);font-weight:600;margin-bottom:.1rem}
details.trlgroup{background:var(--panel);border:1px solid var(--rule);border-radius:var(--r);box-shadow:var(--shadow);margin-top:.8rem;overflow:hidden}
details.trlgroup summary{cursor:pointer;list-style:none;display:flex;flex-wrap:wrap;align-items:center;gap:.6rem .9rem;padding:.9rem 1.1rem;user-select:none}
details.trlgroup summary::-webkit-details-marker{display:none}
details.trlgroup summary:hover{background:var(--panel-2)}
details.trlgroup summary .sw{width:26px;height:26px;font-size:13px;border-radius:7px}
details.trlgroup summary b{font-family:var(--serif);font-size:1.12rem;font-weight:600}
details.trlgroup summary .defn{color:var(--muted);font-size:.9rem;flex:1 1 260px}
details.trlgroup summary .count{font-size:.82rem;color:var(--muted);white-space:nowrap}
details.trlgroup summary::after{content:"";width:.5rem;height:.5rem;border-right:1.5px solid var(--muted);border-bottom:1.5px solid var(--muted);transform:rotate(45deg);margin:0 .3rem 0 .2rem;flex:none}
details.trlgroup[open] summary::after{transform:rotate(-135deg)}
details.trlgroup[open] summary{border-bottom:1px solid var(--rule)}
table.ev{min-width:1080px}
table.ev td{font-size:.86rem}
table.ev td.entry{font-weight:600;color:var(--ink);min-width:170px;vertical-align:top}
table.ev td.entry a{color:inherit;text-decoration:none}
table.ev td.entry a:hover{text-decoration:underline}
table.ev td.entry .id{display:block}
table.ev td.what b{font-weight:600;color:var(--ink)}
table.ev td.what{min-width:200px}
table.ev td.op,table.ev td.dates{white-space:nowrap;color:var(--muted);font-size:.82rem}
table.ev td.detail{color:var(--ink-2);max-width:380px}
table.ev td.src{white-space:nowrap}
/* references section */
.reflist{list-style:none;margin:0;padding:0}
.reflist li{display:grid;grid-template-columns:4.4rem 1fr;gap:.1rem .9rem;padding:.8rem 1.1rem;border-top:1px solid var(--rule);font-size:.92rem;line-height:1.5;scroll-margin-top:84px}
.reflist li:first-child{border-top:0}
.reflist li:target{background:color-mix(in srgb,var(--focus) 10%,var(--panel));box-shadow:inset 3px 0 0 var(--focus)}
.reflist .rid{font-family:var(--sans);font-weight:700;color:var(--link);font-size:.82rem;padding-top:.15rem;letter-spacing:.02em}
.reflist .ttl{font-weight:600;color:var(--ink)}
.reflist .ven{color:var(--ink-2)}
.reflist .links{display:flex;flex-wrap:wrap;gap:.25rem .9rem;font-size:.82rem;color:var(--muted);margin-top:.2rem;align-items:center}
.reflist .links a{color:var(--link);overflow-wrap:anywhere}
.reflist .type{font-size:.68rem;font-weight:700;letter-spacing:.06em;text-transform:uppercase;border:1px solid var(--rule-strong);border-radius:999px;padding:0 .45rem;color:var(--muted)}
.reflist .cited a{color:var(--ink-2);text-decoration:none;border-bottom:1px dotted var(--rule-strong);margin-right:.25rem}
.reflist .cited a:hover{color:var(--link);border-color:var(--link)}
.refnote{font-size:.9rem;color:var(--muted);margin:.6rem 0 0}
.ichip{display:inline-block;font-size:.78rem;font-weight:600;border:1px solid var(--rule-strong);border-radius:999px;padding:.02rem .55rem;margin:.1rem .25rem .1rem 0;color:var(--ink-2);text-decoration:none;background:var(--panel-2);white-space:nowrap}
.ichip:hover{border-color:var(--link);color:var(--link)}
table.cat td.ints{white-space:normal;min-width:150px}
.ilist li{scroll-margin-top:84px}
.ilist li:target{box-shadow:0 0 0 2px var(--focus),var(--shadow-lift)}
.ilist .members{margin-top:.8rem;padding-top:.6rem;border-top:1px solid var(--rule)}
.ilist .members summary{font-size:.86rem}
.ilist .memberlist{font-size:.9rem;line-height:1.75;margin:.4rem 0 0;max-width:none}
.ilist .memberlist a{color:var(--ink);text-decoration:none}
.ilist .memberlist a:hover{color:var(--link);text-decoration:underline}
table.grid th.colhead a{color:inherit;text-decoration:none;display:block}
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
  .reflist li{grid-template-columns:1fr}
  .reflist .rid{padding-top:0}
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
    <a href="#grid-section">Grid</a><a href="#catalog-section">Catalog</a><a href="#concept-section">Concepts</a><a href="#interaction-section">Interactions</a><a href="#trl-section">TRL</a><a href="#evidence-section">Evidence</a><a href="#references-section">References</a><a href="#method-section">Method</a>
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
  <div class="sec-head"><div><p class="kicker">Source × interaction</p><h2>The grid</h2><p>Select a cell for what fills it or why it is empty. Select a column header to go to that interaction's section, with every technique and concept in the family.</p></div></div>
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
  <div class="sec-head"><div><p class="kicker">Catalog</p><h2>Established techniques</h2><p>One row per technique with its TRL for use from space and the evidence behind it. Filter by source or interaction to see a whole family; a technique with several interactions appears under each. Search matches names, summaries, example missions and source IDs.</p></div></div>
  <div class="card">
    <div class="toolbar">
      <input type="search" id="q" placeholder="Search techniques" aria-label="Search techniques">
      <label>Source <select id="c-src"><option value="">all</option></select></label>
      <label>Interaction <select id="c-int"><option value="">all</option></select></label>
      <label>TRL band <select id="c-band"><option value="">all</option></select></label>
      <label>Sort <select id="c-sort"><option value="id">list order</option><option value="trl-desc">TRL, highest first</option><option value="trl-asc">TRL, lowest first</option><option value="int">interaction</option><option value="src">source</option><option value="name">name</option></select></label>
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
      <div class="sec-head"><div><p class="kicker">Readiness</p><h2>The TRL scale as applied here</h2><p>NASA definitions<span id="trlref"></span>. The judgment is for the <em>technique used from space in its stated application</em>, which is why ground-operational Raman lidar sits at TRL 3 and a calibration concept built entirely from flown components sits at TRL 2.</p></div></div>
      <div class="card"><table class="scale" id="scale"></table></div>
    </div>
    <div id="filters-section">
      <div class="sec-head"><div><p class="kicker">Before any study</p><h2>Four physical filters</h2><p>Run these before any concept study. They decide which empty cells are empty for good reason.</p></div></div>
      <ul class="filterlist card" id="filterlist"></ul>
    </div>
  </div>
</section>

<section id="evidence-section">
  <div class="sec-head"><div><p class="kicker">Readiness evidence</p><h2>What each TRL rests on</h2><p>Every technique and concept lists the evidence behind its number, with the source for each item. The kind of evidence changes with the level: at 9 it is a mission and an instrument in service; at 7 and 8 it is a program and how far it has been qualified or demonstrated; at 6 and below it is the system, subsystem, component or technology that was tested and the level of that test.</p></div></div>
  <div class="tiers">
    <div class="tier"><i class="sw b9">9</i><div><b>TRL 9: missions and instruments</b>The measurement is delivered operationally by named instruments on named missions.</div></div>
    <div class="tier"><i class="sw b78">7</i><div><b>TRL 7–8: programs and qualification</b>What was flown or qualified, where, and what remains before routine operation.</div></div>
    <div class="tier"><i class="sw b46">5</i><div><b>TRL 6 and below: tested item and test level</b>The system, subsystem, component or technology and whether it was tested in the laboratory, in the field, from an aircraft or balloon, or only in analysis.</div></div>
  </div>
  <div id="evidence"></div>
</section>

<section id="references-section">
  <div class="sec-head"><div><p class="kicker">Sources</p><h2>References</h2><p>Every numbered marker in the text, and every evidence item, points here. Only peer-reviewed papers, agency and mission documents, patents and academic books are admitted; each record was resolved against Crossref or fetched from the publisher on the date shown. Preprints are admitted only where no journal version exists and are labelled.</p></div></div>
  <div class="card">
    <div class="toolbar">
      <input type="search" id="rq" placeholder="Search references" aria-label="Search references">
      <label>Type <select id="r-type"><option value="">all</option></select></label>
      <span class="count" id="r-count"></span>
    </div>
    <ol class="reflist" id="reflist"></ol>
  </div>
  <p class="refnote" id="refnote"></p>
</section>

<section id="method-section">
  <div class="sec-head"><div><p class="kicker">Method</p><h2>How to read and extend this</h2></div></div>
  <div class="card prose">
  <p>The grid is a generative device, not a taxonomy. A technique is a choice of source, one or more interactions, a viewing geometry and a receiver class; enumerating the product space and asking why each cell is empty produces the concept list. The receiver axis is where the field is thinnest: almost everything is incoherent intensity or coherent heterodyne, with SMOS and geodetic VLBI the only TRL 9 correlation receivers and Swarm's helium magnetometer the only TRL 9 quantum sensor in routine Earth service.</p>
  <p>The TRL column changes the reading of the concept list. Sorted by TRL, the top is atom-interferometer gradiometry (5), TIR polarimetry (5), and a cluster at 4 (Rydberg, vibrometry, chronometric geodesy, Brillouin, Rayleigh–Brillouin, Ku/Ka reflectometry); these are IRAD-shaped because the step to 6 is an airborne or hosted-payload demonstration, not new physics. The bottom (2) is where the cheapest verdicts live. C04 and C17 were the first to be tried: each moved from 2 to 3 on a photon budget (the scripts are in the repository under scripts/calc), and each now needs a field night rather than a calculation.</p>
  <p>Two things deserved scrutiny before anyone quoted a number, and both were checked against the literature and mission status on 2026-10-17. First, the verify flags: of eleven, eight were settled, mostly upward, as Biomass, PREFIRE and the DQ-1 lidar reached routine operations, and three remain (Biomass tomographic products due 2027, the FY-4M sounder design, and the first ACES clock-comparison paper). Second, the novelty claims: of thirteen rated high, one survived. Most of the rest turned out to be funded programs or published proposals, and one, opportunistic laser occultation, had already been observed once by accident on a 2006 comm link. Every prior-art field now names what the search found and when.</p>
  <p>On 2026-11-04 the sourcing was made explicit. Each technique and concept now lists the evidence behind its TRL in the form the level calls for (missions and instruments at 9, programs and their qualification at 7 and 8, the item tested and the test level below that), and every claim in the notes carries a marker to a reference that was resolved against Crossref or fetched from the agency on that date. Only peer-reviewed papers, agency and mission documents, patents and academic books were admitted. Where the sources contradicted the text, the text was corrected and the Evidence table says so; where no admissible source could be found for a claim, the claim was softened or removed rather than left standing. The one TRL that moved was lightning A-band altitude, 6 to 5, because the retrieval it describes has never been published from orbit.</p>
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

  // references: {R12} tokens in text become linked markers; refs arrays become pill rows
  const REF = Object.fromEntries(D.references.map(r => [r.id, r]));
  const refWho = r => r.authors || r.org || '';
  const refShort = r => `${refWho(r).split(';')[0].split(',')[0]}${r.year ? ' ' + r.year : ''}`;
  const refTip = id => { const r = REF[id]; return r ? `data-tipname="${esc(r.title)}" data-tipsub="${esc(refWho(r))}${r.year?' · '+r.year:''}${r.venue?' · '+esc(r.venue):''}"` : ''; };
  const citeLink = id => REF[id] ? `<a class="cite" href="#${id}" ${refTip(id)} aria-label="Reference ${id}">${id}</a>` : `<span class="cite" title="missing reference">${esc(id)}</span>`;
  const cite = s => esc(s).replace(/\{(R\d+)\}/g, (m, id) => citeLink(id));
  const plain = s => String(s).replace(/\s?\{R\d+\}/g, '');
  const refPills = (ids, label) => ids && ids.length ? `<div class="srcs">${label === false ? '' : `<span class="lbl">${label || 'Sources'}</span>`}${ids.map(citeLink).join('')}</div>` : '';
  const LEVEL_KIND = {mission:'Mission', program:'Program', system:'System', subsystem:'Subsystem', component:'Component', technology:'Technology'};
  const evItem = v => `<li class="${esc(v.kind)}"><span class="kindlbl">${esc(LEVEL_KIND[v.kind] || v.kind)}</span><b>${esc(v.name)}</b>${v.instrument ? ` · ${esc(v.instrument)}` : ''}${v.operator ? ` <span class="muted">(${esc(v.operator)}${v.dates ? ', ' + esc(v.dates) : ''})</span>` : (v.dates ? ` <span class="muted">(${esc(v.dates)})</span>` : '')}<br><span class="lvl">${esc(v.level)}</span>${cite(v.detail || '')} ${(v.refs||[]).map(citeLink).join('')}</li>`;
  const evList = items => items && items.length ? `<ul class="evlist">${items.map(evItem).join('')}</ul>` : '';
  const evDetails = it => it.evidence && it.evidence.length ? `<details class="ev"><summary>Evidence: ${it.evidence.length} item${it.evidence.length===1?'':'s'}</summary>${evList(it.evidence)}</details>` : '';
  // who cites what
  const citedBy = {};
  const addCites = (id, obj, fields) => {
    const set = new Set(obj.refs || []);
    fields.forEach(f => String(obj[f] || '').replace(/\{(R\d+)\}/g, (m, r) => { set.add(r); return m; }));
    (obj.evidence || []).forEach(v => (v.refs || []).forEach(r => set.add(r)));
    set.forEach(r => { (citedBy[r] = citedBy[r] || []).push(id); });
  };
  D.techniques.forEach(t => addCites(t.id, t, ['trlNote', 'summary']));
  D.concepts.forEach(c => addCites(c.id, c, ['trlNote', 'prior', 'physics', 'application', 'failure', 'first']));
  Object.entries(D.cellNotes).forEach(([k, n]) => addCites('cell:' + k, n, ['text']));
  if (D.meta.trlRef) (citedBy[D.meta.trlRef] = citedBy[D.meta.trlRef] || []).push('TRL scale');

  document.getElementById('caveat').innerHTML = cite(D.meta.caveat);
  $('#trlref').innerHTML = D.meta.trlRef ? ' ' + citeLink(D.meta.trlRef) : '';
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
    D.interactions.forEach(i => h += `<th class="colhead" data-int="${i.id}" data-tipname="${esc(i.name)}" data-tipsub="${esc(i.note)} · open the section"><a href="#int-${i.id}" aria-label="${esc(i.name)}: open its section"><span>${esc(i.short)}</span></a></th>`);
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
    if (!state.sel){ d.innerHTML = '<span class="empty"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 4l7 17 2.5-7.5L21 11z"/></svg>Select a cell to see what fills it, or why it is empty. Select a column header to open that interaction\'s section.</span>'; return; }
    const [sid, iid] = state.sel.split('|');
    const s = byId(D.sources, sid), i = byId(D.interactions, iid);
    const items = cellItems(sid, iid).filter(itemMatches);
    const note = D.cellNotes[state.sel];
    let h = `<h3>${esc(s.name)} × ${esc(i.name)}</h3>`;
    if (note) h += `<p class="muted" style="margin:.2rem 0 .6rem"><span class="tag kind">${esc(note.kind === 'ruled' ? 'ruled out' : note.kind === 'open' ? 'open' : note.kind)}</span>${cite(note.text)}</p>${refPills(note.refs)}`;
    if (!items.length && !note) h += '<p class="muted">Nothing filed here and no annotation. Either an oversight or an unexplored combination.</p>';
    items.sort((a,b) => (a.kind===b.kind?0:(a.kind==='technique'?-1:1)) || b.trl-a.trl);
    items.forEach(it => {
      const geoms = it.geom.map(g => byId(D.geometries, g).name).join(', ');
      const rec = byId(D.receivers, it.rec).name;
      const noteLine = it.trlNote ? `<br><span class="muted" style="font-size:.86rem">TRL note: ${cite(it.trlNote)}</span>` : '';
      if (it.kind === 'technique'){
        h += `<div class="item">${trlTag(it)}<b>${esc(it.name)}</b> <span class="muted">${esc(it.id)}</span><br>${cite(it.summary)}<br><span class="muted" style="font-size:.86rem">${esc(it.examples)} · ${esc(it.band)} · ${esc(geoms)} · ${esc(rec)}</span>${noteLine}${evDetails(it)}${refPills(it.refs)}<a href="#${it.id}" style="font-size:.86rem">Open in the catalog</a></div>`;
      } else {
        h += `<div class="item">${trlTag(it)}<b>${esc(it.name)}</b> <span class="muted">${esc(it.id)} · concept</span><br>${cite(it.physics)}${noteLine}${evDetails(it)}${refPills(it.refs)}<a href="#${it.id}" style="font-size:.86rem">Open the concept sheet</a></div>`;
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
      <div class="stat"><b>${nVer}</b><span>TRL values flagged to verify</span></div>
      <div class="stat"><b>${D.references.length}</b><span>verified sources behind the numbers</span></div>`;
  }

  fGeom.addEventListener('change', e => { state.geom = e.target.value; renderGrid(); renderDrawer(); });
  fRec.addEventListener('change', e => { state.rec = e.target.value; renderGrid(); renderDrawer(); });
  fTrl.addEventListener('change', e => { state.trl = +e.target.value; renderGrid(); renderDrawer(); });
  fConcepts.addEventListener('click', () => { state.concepts = !state.concepts; fConcepts.setAttribute('aria-pressed', state.concepts); renderGrid(); renderDrawer(); });
  $('#f-clear').addEventListener('click', () => { state.sel = null; state.col = null; renderGrid(); renderDrawer(); });

  // catalog
  const cSrc = $('#c-src'), cBand = $('#c-band'), cVer = $('#c-verify'), q = $('#q'), cInt = $('#c-int'), cSort = $('#c-sort');
  let onlyVerify = false;
  D.sources.forEach(s => cSrc.insertAdjacentHTML('beforeend', `<option value="${s.id}">${esc(s.name)}</option>`));
  D.interactions.forEach(i => cInt.insertAdjacentHTML('beforeend', `<option value="${i.id}">${esc(i.name)}</option>`));
  const intOrder = Object.fromEntries(D.interactions.map((i, k) => [i.id, k]));
  const srcOrder = Object.fromEntries(D.sources.map((s, k) => [s.id, k]));
  D.trlBands.forEach(b => cBand.insertAdjacentHTML('beforeend', `<option value="${b.id}">${esc(b.name)}</option>`));
  function renderCat(){
    const term = q.value.trim().toLowerCase();
    let rows = D.techniques.filter(t => {
      if (cSrc.value && t.src !== cSrc.value) return false;
      if (cInt.value && !t.int.includes(cInt.value)) return false;
      if (cBand.value && band(t) !== cBand.value) return false;
      if (onlyVerify && !t.verify) return false;
      if (term && !(t.name+' '+t.summary+' '+t.examples+' '+t.band+' '+t.id+' '+(t.refs||[]).join(' ')+' '+(t.evidence||[]).map(v=>v.name+' '+v.instrument+' '+v.operator).join(' ')).toLowerCase().includes(term)) return false;
      return true;
    });
    const sortKey = cSort.value;
    if (sortKey === 'trl-desc') rows = rows.slice().sort((a,b) => b.trl - a.trl || a.id.localeCompare(b.id));
    if (sortKey === 'trl-asc') rows = rows.slice().sort((a,b) => a.trl - b.trl || a.id.localeCompare(b.id));
    if (sortKey === 'name') rows = rows.slice().sort((a,b) => a.name.localeCompare(b.name));
    if (sortKey === 'src') rows = rows.slice().sort((a,b) => srcOrder[a.src] - srcOrder[b.src] || b.trl - a.trl);
    if (sortKey === 'int') rows = rows.slice().sort((a,b) => Math.min(...a.int.map(i=>intOrder[i])) - Math.min(...b.int.map(i=>intOrder[i])) || b.trl - a.trl);
    let h = '<thead><tr><th>Technique</th><th>TRL · evidence</th><th>Source</th><th>Interactions</th><th>Geometry · receiver</th><th>Summary</th><th>Examples · band</th></tr></thead><tbody>';
    rows.forEach(t => {
      h += `<tr id="${t.id}"><td class="name">${esc(t.name)}<br><span class="id">${t.id}</span>${refPills(t.refs)}</td>
        <td class="trl">${trlTag(t)}${t.trlNote?`<span class="trlnote">${cite(t.trlNote)}</span>`:''}${evDetails(t)}</td>
        <td class="axis">${esc(byId(D.sources,t.src).name)}</td>
        <td class="ints">${t.int.map(i=>`<a class="ichip" href="#int-${i}" title="${esc(byId(D.interactions,i).name)}: open its section">${esc(byId(D.interactions,i).short)}</a>`).join('')}</td>
        <td class="axis">${t.geom.map(g=>esc(byId(D.geometries,g).name)).join('; ')}<br>${esc(byId(D.receivers,t.rec).name)}</td>
        <td class="sum">${cite(t.summary)}</td>
        <td class="ex">${esc(t.examples)}<br>${esc(t.band)}</td></tr>`;
    });
    h += '</tbody>';
    $('#cat').innerHTML = h;
    $('#c-count').textContent = rows.length + ' of ' + D.techniques.length;
  }
  [cSrc, cBand, cInt, cSort].forEach(el => el.addEventListener('change', renderCat));
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
          ${c.trlNote?`<dt>TRL basis</dt><dd>${cite(c.trlNote)}</dd>`:''}
          ${c.evidence&&c.evidence.length?`<dt>TRL evidence</dt><dd>${evList(c.evidence)}</dd>`:''}
          <dt>Physics</dt><dd>${cite(c.physics)}</dd>
          <dt>Application</dt><dd>${cite(c.application)}</dd>
          <dt>What kills it</dt><dd>${cite(c.failure)}</dd>
          <dt>First calculation</dt><dd>${cite(c.first)}</dd>
          <dt>Prior art</dt><dd>${cite(c.prior)}</dd>
        </dl>${refPills(c.refs)}</article>`;
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
    const tl = D.techniques.filter(x => x.int.includes(i.id)).sort((a,b)=>b.trl-a.trl);
    const cl = D.concepts.filter(x => x.int.includes(i.id)).sort((a,b)=>b.trl-a.trl);
    const members = tl.map(x => `${trlTag(x)}<a href="#${x.id}">${esc(x.name)}</a> <span class="id">${x.id}</span>`).concat(cl.map(x => `${trlTag(x)}<a href="#${x.id}">${esc(x.name)}</a> <span class="id">${x.id} · concept</span>`)).join('<br>');
    return `<li id="int-${i.id}"><h3>${esc(i.name)}</h3><p class="note">${esc(i.note)}</p>${interactionBlock(i, false)}<div class="members"><details class="ev" open><summary>${nT} technique${nT===1?'':'s'}, ${nC} concept${nC===1?'':'s'} in this family · driven by ${esc(srcs)}</summary><p class="memberlist">${members || '<span class="muted">Nothing filed here.</span>'}</p></details></div></li>`;
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
  document.addEventListener('mousemove', e => {
    const el = e.target.closest('[data-tipname]');
    if (el && el.dataset.tipname) showTip(el, e.clientX, e.clientY); else tip.classList.remove('on');
  });
  grid.addEventListener('mouseleave', () => tip.classList.remove('on'));
  document.addEventListener('focusin', e => { const el = e.target.closest('.cite[data-tipname]'); if (el){ const r = el.getBoundingClientRect(); showTip(el, r.left, r.bottom); } });
  document.addEventListener('focusout', () => tip.classList.remove('on'));

  // evidence by TRL level
  function renderEvidence(){
    const all = D.techniques.map(t => ({...t, kind2:'technique'})).concat(D.concepts.map(c => ({...c, kind2:'concept'})));
    const defn = Object.fromEntries(D.trlScale.map(r => [+r[0], r[1]]));
    let h = '';
    for (let L = 9; L >= 1; L--){
      const ents = all.filter(e => e.trl === L);
      if (!ents.length) continue;
      const nT = ents.filter(e => e.kind2==='technique').length, nC = ents.length - nT;
      const nEv = ents.reduce((a, e) => a + (e.evidence||[]).length, 0);
      const b = band({trl:L});
      const cols = L >= 9 ? ['Entry','Mission','Instrument','Operator','Dates','Status','Detail','Sources']
                 : L >= 7 ? ['Entry','Program','Instrument','Operator','Dates','Qualification','Detail','Sources']
                 : ['Entry','Kind · item tested','Instrument','Operator / lab','Dates','Test level','Detail','Sources'];
      let rows = '';
      ents.forEach(e => {
        const ev = e.evidence && e.evidence.length ? e.evidence : [null];
        ev.forEach((v, i) => {
          rows += '<tr>';
          if (i === 0) rows += `<td class="entry" rowspan="${ev.length}"><a href="#${e.id}">${esc(e.name)}</a><span class="id">${e.id} · ${e.kind2}${e.dubious?' · physics-dubious':''}${e.verify?' · verify':''}</span></td>`;
          if (!v){ rows += '<td colspan="7" class="muted">No evidence recorded yet.</td></tr>'; return; }
          rows += `<td class="what">${L < 7 ? `<span class="kindlbl">${esc(LEVEL_KIND[v.kind]||v.kind)}</span>` : ''}<b>${esc(v.name)}</b></td><td>${esc(v.instrument||'')}</td><td class="op">${esc(v.operator||'')}</td><td class="dates">${esc(v.dates||'')}</td><td><span class="lvl">${esc(v.level||'')}</span></td><td class="detail">${cite(v.detail||'')}</td><td class="src">${(v.refs||[]).map(citeLink).join('')}</td></tr>`;
        });
      });
      h += `<details class="trlgroup"><summary><i class="sw ${b}">${L}</i><b>TRL ${L}</b><span class="defn">${esc(defn[L]||'')}</span><span class="count">${nT} technique${nT===1?'':'s'}${nC?`, ${nC} concept${nC===1?'':'s'}`:''} · ${nEv} evidence item${nEv===1?'':'s'}</span></summary><div class="tblwrap"><table class="cat ev"><thead><tr>${cols.map(c=>`<th>${c}</th>`).join('')}</tr></thead><tbody>${rows}</tbody></table></div></details>`;
    }
    $('#evidence').innerHTML = h;
  }

  // references list
  const rq = $('#rq'), rType = $('#r-type');
  [...new Set(D.references.map(r => r.type))].sort().forEach(t => rType.insertAdjacentHTML('beforeend', `<option value="${esc(t)}">${esc(t)}</option>`));
  const host = u => { try { return new URL(u).hostname.replace(/^www\./,''); } catch(e){ return u; } };
  const anchorFor = id => id.startsWith('cell:') ? null : (id === 'TRL scale' ? '#trl-section' : '#' + id);
  const citedLabel = id => { if (!id.startsWith('cell:')) return id; const [a, b] = id.slice(5).split('|'); return `grid ${byId(D.sources,a).name} × ${byId(D.interactions,b).short}`; };
  function renderRefs(){
    const term = rq.value.trim().toLowerCase(), typ = rType.value;
    const list = D.references.filter(r => (!typ || r.type === typ) && (!term || [r.id, r.authors, r.org, r.title, r.venue, r.year, r.doi, (citedBy[r.id]||[]).join(' ')].join(' ').toLowerCase().includes(term)));
    $('#reflist').innerHTML = list.map(r => {
      const cited = (citedBy[r.id] || []);
      const link = r.doi ? `<a href="https://doi.org/${esc(r.doi)}" target="_blank" rel="noopener">doi:${esc(r.doi)}</a>` : (r.url ? `<a href="${esc(r.url)}" target="_blank" rel="noopener">${esc(host(r.url))}</a>` : '');
      return `<li id="${r.id}"><span class="rid">${r.id}</span><div><div class="ttl">${esc(r.title)}</div><div class="ven">${esc(refWho(r))}${r.year?` (${r.year})`:''}${r.venue?`. ${esc(r.venue)}`:''}${r.note?` <span class="muted">${esc(r.note)}</span>`:''}</div><div class="links"><span class="type">${esc(r.type)}</span>${link}${r.verified?`<span>verified ${esc(r.verified)}${r.accessed?', '+esc(r.accessed):''}</span>`:''}${cited.length?`<span class="cited">cited by ${cited.map(id => anchorFor(id) ? `<a href="${anchorFor(id)}">${esc(citedLabel(id))}</a>` : `<span>${esc(citedLabel(id))}</span>`).join('')}</span>`:''}</div></div></li>`;
    }).join('');
    $('#r-count').textContent = list.length + ' of ' + D.references.length;
  }
  rq.addEventListener('input', renderRefs); rType.addEventListener('change', renderRefs);
  const typeCounts = {}; D.references.forEach(r => typeCounts[r.type] = (typeCounts[r.type]||0) + 1);
  const nCitedEnt = D.techniques.filter(t => (citedBy && Object.values(citedBy).some(l => l.includes(t.id)))).length + D.concepts.filter(c => Object.values(citedBy).some(l => l.includes(c.id))).length;
  $('#refnote').textContent = D.references.length ? `${D.references.length} references: ${Object.entries(typeCounts).sort((a,b)=>b[1]-a[1]).map(([k,v])=>`${v} ${k}`).join(', ')}. ${nCitedEnt} of ${D.techniques.length + D.concepts.length} techniques and concepts carry at least one source.` : 'No references recorded yet.';
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

  renderGrid(); renderDrawer(); renderStats(); renderCat(); renderSheets(); renderEvidence(); renderRefs();
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
    ['Evidence', 'One row per piece of TRL evidence: the mission and instrument (TRL 9), the program and its qualification level (TRL 7–8), or the item tested and the test level (TRL 6 and below), each with its sources.'],
    ['References', 'Every source cited anywhere in the workbook, with DOI or URL, how it was verified and on what date, and which entries cite it. Markers like [R12] in any text cell point to the row with that ID.'],
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
rows = [[t['id'], t['name'], t['src'], '|'+'|'.join(t['int'])+'|', '|'+'|'.join(t['geom'])+'|', t['rec'], t['trl'], xl(t['trlNote']), 'Y' if t['verify'] else '', t['band'], t['examples'], xl(t['summary']), ' '.join(t['refs']), ev_summary(t)] for t in D['techniques']]
sheet(ws, ['ID', 'Technique', 'Source code', 'Interaction codes (|a|b|)', 'Geometry codes (|a|b|)', 'Receiver code', 'TRL (1–9)', 'TRL basis', 'Verify', 'Band', 'Examples', 'Summary', 'Sources (see References)', 'TRL evidence (see Evidence)'], rows,
      [6, 34, 11, 20, 20, 11, 8, 44, 7, 22, 34, 70, 18, 60])
for row in ws.iter_rows(min_row=2, max_col=12):
    for c in row: c.font = blue_font
CAT_MAX = 400

ws = wb.create_sheet('Concepts')
rows = [[c['id'], c['name'], '|'+'|'.join(c['src'])+'|', '|'+'|'.join(c['int'])+'|', '|'+'|'.join(c['geom'])+'|', c['rec'], c['trl'], xl(c['trlNote']), 'Y' if c['verify'] else '', 'Y' if c['dubious'] else '', c['novelty'], xl(c['physics']), xl(c['application']), xl(c['failure']), xl(c['first']), xl(c['prior']), ' '.join(c['refs']), ev_summary(c)] for c in D['concepts']]
sheet(ws, ['ID', 'Concept', 'Source codes (|a|b|)', 'Interaction codes (|a|b|)', 'Geometry codes (|a|b|)', 'Receiver code', 'TRL (1–9)', 'TRL basis', 'Verify', 'Dubious', 'Novelty confidence', 'Physics', 'Application', 'What kills it', 'First calculation', 'Prior art', 'Sources (see References)', 'TRL evidence (see Evidence)'], rows,
      [6, 34, 14, 18, 18, 10, 8, 40, 7, 8, 10, 60, 40, 46, 40, 40, 18, 60])
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
    a, b = k.split('|'); ws.append([name[a], name[b], v['kind'], xl(v['text']) + (' ' + ' '.join('[%s]' % r for r in v['refs']) if v['refs'] else '')])
for row in ws.iter_rows(min_row=3):
    for c in row:
        if c.font != hdr_font: c.font = body_font
        c.alignment = wrap
ws.column_dimensions['A'].width = 22
for i in range(2, len(D['interactions']) + 2): ws.column_dimensions[get_column_letter(i)].width = 9
ws.append([])
ws.append(['Blank TRL cells read as 0 in the MAX formulas; a non-numeric TRL entry will produce a formula error, so keep the TRL column to integers 1–9. Counting ranges cover Catalog rows 2–%d and Concepts rows 2–%d.' % (CAT_MAX, CON_MAX)])
ws.cell(row=ws.max_row, column=1).font = Font(name=FONT, italic=True, size=9)

ws = wb.create_sheet('Evidence')
rows = []
for e in D['techniques'] + D['concepts']:
    kind2 = 'technique' if e['id'].startswith('T') else 'concept'
    for v in e['evidence']:
        rows.append([e['id'], e['name'], kind2, e['trl'], v.get('kind', ''), v.get('name', ''), v.get('instrument', ''), v.get('operator', ''), v.get('dates', ''), v.get('level', ''), xl(v.get('detail', '')), ' '.join(v.get('refs', []))])
sheet(ws, ['Entry ID', 'Entry', 'Entry kind', 'TRL', 'Evidence kind', 'Mission / program / item tested', 'Instrument', 'Operator or lab', 'Dates', 'Status, qualification or test level', 'Detail', 'Sources (see References)'], rows,
      [8, 34, 10, 6, 12, 36, 26, 26, 14, 30, 70, 18])
ws.auto_filter.ref = ws.dimensions

ws = wb.create_sheet('References')
cited_by = {}
def _cite(eid, obj, fields):
    s = set(obj.get('refs', []))
    for f in fields:
        s.update(_re.findall(r'\{(R\d+)\}', obj.get(f, '') or ''))
    for v in obj.get('evidence', []): s.update(v.get('refs', []))
    for r in s: cited_by.setdefault(r, []).append(eid)
for t in D['techniques']: _cite(t['id'], t, ['trlNote', 'summary'])
for c in D['concepts']: _cite(c['id'], c, ['trlNote', 'prior', 'physics', 'application', 'failure', 'first'])
for k, v in D['cellNotes'].items(): _cite('grid ' + name[k.split('|')[0]] + ' × ' + name[k.split('|')[1]], v, ['text'])
if D['meta'].get('trlRef'): cited_by.setdefault(D['meta']['trlRef'], []).append('TRL scale')
rows = [[r['id'], r.get('type', ''), r.get('authors', ''), r.get('title', ''), r.get('venue', ''), r.get('year', ''), r.get('doi', ''), ('https://doi.org/' + r['doi']) if r.get('doi') else r.get('url', ''), r.get('org', ''), r.get('verified', ''), r.get('accessed', ''), r.get('note', ''), ', '.join(cited_by.get(r['id'], []))] for r in D['references']]
sheet(ws, ['ID', 'Type', 'Authors', 'Title', 'Venue', 'Year', 'DOI', 'Link', 'Organisation', 'Verified by', 'Accessed', 'Note', 'Cited by'], rows,
      [7, 12, 40, 60, 36, 7, 28, 44, 20, 12, 11, 30, 40])
for row in ws.iter_rows(min_row=2, min_col=8, max_col=8):
    for c in row:
        if c.value:
            c.hyperlink = c.value; c.font = Font(name=FONT, size=10, color='1C5CAB', underline='single')
ws.auto_filter.ref = ws.dimensions

ws = wb.create_sheet('Filters')
sheet(ws, ['Filter', 'What it decides'], [[f['name'], f['text']] for f in D['filters']], [26, 110])

wb.save(OUT / 'remote-sensing-phenomenology.xlsx')
print('built', len(D['techniques']), 'techniques,', len(D['concepts']), 'concepts,', len(D['references']), 'references ->', OUT)
