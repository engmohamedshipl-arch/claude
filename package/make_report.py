"""Planning package report (HTML for print to PDF). Args: report_data.json out.html"""
import sys, json, datetime, html

D = json.load(open(sys.argv[1]))
OUT = sys.argv[2]
E = html.escape
BLUE, GREY, INK, INK2, MUTED, GRID, SURF = '#2a78d6', '#c9c8c2', '#0b0b0b', '#52514e', '#8a8984', '#e6e5e0', '#fcfcfb'

def fmt(x, d=0):
    return f'{x:,.{d}f}'

# ----------------------------------------------------------------- charts (inline SVG)
def hbar(rows, width=640, bar=16, gap=8, label_w=250, val_fmt=lambda v: fmt(v), highlight=None, unit=''):
    vmax = max(v for _, v in rows) or 1
    plot = width - label_w - 70
    h = len(rows) * (bar + gap) + 10
    out = [f'<svg viewBox="0 0 {width} {h}" width="100%" role="img">']
    for i, (lab, v) in enumerate(rows):
        y = 5 + i * (bar + gap)
        w = max(2, plot * v / vmax)
        col = BLUE if (highlight is None or highlight == i) else GREY
        out.append(f'<text x="{label_w-8}" y="{y+bar*0.72}" text-anchor="end" font-size="11" fill="{INK2}">{E(lab)}</text>')
        out.append(f'<path d="M{label_w},{y} h{w-4} a4,4 0 0 1 4,4 v{bar-8} a4,4 0 0 1 -4,4 h-{w-4} z" fill="{col}"><title>{E(lab)}: {val_fmt(v)}{unit}</title></path>')
        out.append(f'<text x="{label_w+w+6}" y="{y+bar*0.72}" font-size="11" fill="{INK}">{val_fmt(v)}{unit}</text>')
    out.append('</svg>')
    return ''.join(out)

def gantt(fronts, width=680):
    t0 = datetime.date(2026, 12, 1); t1 = datetime.date(2027, 10, 1)
    label_w, bar, gap = 170, 16, 10
    plot = width - label_w - 70
    span = (t1 - t0).days
    x = lambda d: label_w + plot * (d - t0).days / span
    h = 30 + len(fronts) * (bar + gap)
    out = [f'<svg viewBox="0 0 {width} {h}" width="100%" role="img">']
    m = t0
    while m <= t1:
        xx = x(m)
        out.append(f'<line x1="{xx:.1f}" y1="18" x2="{xx:.1f}" y2="{h}" stroke="{GRID}" stroke-width="1"/>')
        out.append(f'<text x="{xx+2:.1f}" y="12" font-size="9.5" fill="{MUTED}">{m.strftime("%b %y")}</text>')
        m = (m.replace(day=28) + datetime.timedelta(days=4)).replace(day=1)
    for i, f in enumerate(fronts):
        y = 24 + i * (bar + gap)
        s = datetime.date.fromisoformat(f[8]); e = datetime.date.fromisoformat(f[9])
        SH = {'F1': 'Headworks & primary', 'F2': 'Bypass corridor', 'F3': 'Distribution & risers', 'F4': 'FCT effluent A-H', 'F5': 'FCT effluent I-L',
              'F6': 'Sludge A-D, G, H', 'F7': 'Sludge I-L, E, F', 'F8': 'Effluent & CCT links'}
        out.append(f'<text x="{label_w-8}" y="{y+bar*0.72}" text-anchor="end" font-size="11" fill="{INK2}">{E(f[0]+"  "+SH.get(f[0], f[1]))}</text>')
        w = x(e) - x(s)
        out.append(f'<rect x="{x(s):.1f}" y="{y}" width="{w:.1f}" height="{bar}" rx="4" fill="{BLUE}"><title>{f[0]}: {s:%d %b %Y} to {e:%d %b %Y}</title></rect>')
        out.append(f'<text x="{x(e)+5:.1f}" y="{y+bar*0.72}" font-size="10.5" fill="{INK}">{e:%d %b}</text>')
    out.append('</svg>')
    return ''.join(out)

def columns(months, vals, ref=None, width=680, height=180, unit=''):
    n = len(months); left, bottom, top = 34, 24, 14
    plot_w = width - left - 10; plot_h = height - bottom - top
    vmax = max(max(vals), ref or 0) * 1.15 or 1
    bw = plot_w / n
    out = [f'<svg viewBox="0 0 {width} {height}" width="100%" role="img">']
    for t in range(0, int(vmax) + 1, 5):
        y = top + plot_h * (1 - t / vmax)
        out.append(f'<line x1="{left}" y1="{y:.1f}" x2="{width-10}" y2="{y:.1f}" stroke="{GRID}"/>')
        out.append(f'<text x="{left-6}" y="{y+3:.1f}" text-anchor="end" font-size="9.5" fill="{MUTED}">{t}</text>')
    for i, (m, v) in enumerate(zip(months, vals)):
        xx = left + i * bw + 2
        hh = plot_h * v / vmax
        if v > 0:
            out.append(f'<path d="M{xx:.1f},{top+plot_h:.1f} v-{hh-4:.1f} a4,4 0 0 1 4,-4 h{bw-12:.1f} a4,4 0 0 1 4,4 v{hh-4:.1f} z" fill="{BLUE}"><title>{m}: {v:.1f}{unit}</title></path>')
            if hh > 18:
                out.append(f'<text x="{xx+(bw-4)/2:.1f}" y="{top+plot_h-hh+13:.1f}" text-anchor="middle" font-size="9.5" fill="#ffffff">{v:.1f}</text>')
            else:
                out.append(f'<text x="{xx+(bw-4)/2:.1f}" y="{top+plot_h-hh-4:.1f}" text-anchor="middle" font-size="9.5" fill="{INK}">{v:.1f}</text>')
        out.append(f'<text x="{xx+(bw-4)/2:.1f}" y="{height-8}" text-anchor="middle" font-size="9" fill="{MUTED}">{m}</text>')
    if ref:
        y = top + plot_h * (1 - ref / vmax)
        out.append(f'<line x1="{left}" y1="{y:.1f}" x2="{width-10}" y2="{y:.1f}" stroke="{INK2}" stroke-width="1.5" stroke-dasharray="5 4"/>')
        out.append(f'<text x="{width-12}" y="{y-5:.1f}" text-anchor="end" font-size="10" fill="{INK2}">{ref} front rigs in the programme</text>')
    out.append('</svg>')
    return ''.join(out)

# ----------------------------------------------------------------- content
kq = {r[0]: r[1] for r in D['keyq']}
rc = sorted([(t, v) for _, t, v in D['rc_by_bill']], key=lambda x: -x[1])
top = rc[:10]
top.append((f'Other {len(rc)-10} bills', sum(v for _, v in rc[10:])))
short = {'Underground pipework - 45 printed lines, 201 packages (DLU-002 Rev 03)': 'Underground pipework (piled slabs)',
         'Intermediate pumping station incl. shared IPS/RAS motor hall': 'IPS incl. shared motor hall'}
top = [(short.get(t, t), v) for t, v in top]
scen = [(s[0], s[1]) for s in D['scen']]
fr = D['fronts']
months = D['months'][:8]; rigs = D['hist']['Rigs'][:8]

REG = [
    ('PKG-00', 'Planning package report (this document)', 'Alex-West-Planning-Package-Report.pdf', 'PDF', 'Issued for information'),
    ('PKG-01', 'Pipe sequence and methodology workbook', '01-Sequence/Alex-West-Pipe-Sequence-Methodology.xlsx', 'Excel, live formulas', 'Planning draft'),
    ('PKG-02', 'Pipe works programme for Primavera P6', '02-Programme/Alex-West-Pipe-Works.xer', 'P6 XER', 'Planning draft'),
    ('PKG-03', 'Pipe works programme for MS Project', '02-Programme/Alex-West-Pipe-Works.xml', 'MSPDI XML', 'Planning draft'),
    ('PKG-04', 'Estimate bill of quantities, whole project', '03-BOQ/Alex-West-Estimate-BOQ.xlsx', 'Excel, live formulas', 'Quantities only, unpriced'),
    ('PKG-05', 'Project wiki (decisions, open items, method notes)', '04-Wiki/*.md', 'Markdown', 'Living record'),
    ('PKG-06', 'Build scripts and extracted model data', '05-Scripts-and-Data/', 'Python, JSON', 'For rebuild'),
    ('REF-01', 'Pipe execution packages v3 (input, 21 Oct 2026)', 'not included - project file', 'Excel', 'Input'),
    ('REF-02', 'Combined coordination model (input, 19 Oct 2026)', 'not included - project file (164 MB)', 'Blender', 'Input'),
]

def table(head, rows, cls=''):
    h = ''.join(f'<th>{E(x)}</th>' for x in head)
    b = ''.join('<tr>' + ''.join(f'<td>{c}</td>' for c in r) + '</tr>' for r in rows)
    return f'<table class="{cls}"><thead><tr>{h}</tr></thead><tbody>{b}</tbody></table>'

tiles = [
    ('9.6 months', 'Pipe works, 1 Dec 2026 to 18 Sep 2027 (best case, gates not yet entered)'),
    ('80%', 'Pipe length in sheet piles with groundwater at -1.50 (was 24%)'),
    ('145,800 m³', 'Reinforced concrete, whole project (gross)'),
    ('16,000', 'Piles 600 mm (model, allowance and pipe placeholder)'),
]
tile_html = ''.join(f'<div class="tile"><div class="big">{a}</div><div class="cap">{E(b)}</div></div>' for a, b in tiles)

METHOD = [
    ('1', 'Release', 'DLU-002 approved, pile design approved, structure ready (tie-ins), aeration zone piling complete (within 15 m), crossings identified'),
    ('2', 'Set-out and platform', 'Survey axis, piles, sheet lines; trial pits at crossings; granular platform at NG for rigs'),
    ('3', 'Pipe-support piling', 'From the platform before excavation; 600 mm assumed; integrity tests; preliminary load test'),
    ('4', 'Sheet piles', 'Where open cut does not fit or excavation is more than 2.5 m below GWL; press-in near structures'),
    ('5', 'Dewatering', 'Wellpoints, or deep wells beyond 4.5 m drawdown; at least 7 days drawdown; piezometers and settlement markers'),
    ('6', 'Excavate and strut', 'In layers, evenly around installed piles; deeper lines before shallower; shared trenches opened together'),
    ('7', 'Piled slab', 'Trim piles, blinding, 0.50 m slab with starter bars; straps cast in where flotation requires'),
    ('8', 'Lay and joint', 'Downstream first; GRP bell and spigot, HDPE butt fusion; tie-ins last as closure'),
    ('9', 'Surround and test', 'Granular haunch and surround; joint and line tests'),
    ('10', 'Backfill', 'Layers to NG, struts out progressively; dewatering kept on until the flotation check is met'),
    ('11', 'Off and extract', 'Dewatering off, sheets extracted slowly, voids grouted, surface reinstated'),
]

OPEN = [
    ('Geotechnical / temporary works', 'Confirm GWL -1.50, side slopes and the 2.5 m open-cut depth limit below GWL', 'Sheet-pile quantity swings from 24% to 99% of length'),
    ('Structural design', 'Pile and slab design for pipe supports; pile layouts for structures without them; thickener bridge sections', '5,788 placeholder and about 3,800 allowance piles; bridge concrete'),
    ('Master programme', 'Date each structure is ready for pipe tie-in', '71 tie-in packages'),
    ('Design management', 'Forecast approval of DLU-002 and of the pipe pile design', 'Gates on every package and all piling'),
    ('Survey / design', 'Identify crossing pipes P0114, P0117, P0279; surveyed site boundary and levels', '3 packages; provisional site raising'),
    ('Construction / QS', 'Productivity rates (all assumed except piling); unit rates for the BOQ', 'Durations; BOQ amounts'),
    ('Planning (Claude)', 'Refine structure excavation with a grid of each underside level (current method overstates PST, CCT, RAS/WAS)', 'BOQ excavation and backfill'),
]

css = f"""
@page {{ size: A4; margin: 16mm 15mm 16mm 15mm; }}
:root {{ --ink:{INK}; --ink2:{INK2}; --muted:{MUTED}; --blue:{BLUE}; --line:{GRID}; }}
body {{ font-family: Arial, Helvetica, sans-serif; color: var(--ink); background:#fff; font-size: 10.5pt; line-height: 1.42; margin:0; }}
h1 {{ font-size: 24pt; margin: 0 0 6px; color:#1f3864; }}
h2 {{ font-size: 14pt; color:#1f3864; margin: 0 0 8px; border-bottom: 2px solid #1f3864; padding-bottom: 4px; }}
h3 {{ font-size: 11pt; margin: 14px 0 6px; }}
p {{ margin: 0 0 8px; }}
.page {{ page-break-after: always; }}
.page:last-child {{ page-break-after: auto; }}
.cover {{ height: 255mm; display:flex; flex-direction:column; justify-content:space-between; }}
.band {{ background:#1f3864; color:#fff; padding: 18px 20px; border-radius: 6px; }}
.band h1 {{ color:#fff; }}
.sub {{ color: var(--ink2); font-size: 11pt; }}
.muted {{ color: var(--muted); font-size: 9pt; }}
table {{ width:100%; border-collapse: collapse; margin: 6px 0 10px; font-size: 9pt; }}
th {{ background:#1f3864; color:#fff; text-align:left; padding:5px 6px; font-weight:bold; }}
td {{ border-bottom: 1px solid var(--line); padding:4px 6px; vertical-align: top; }}
td.num, th.num {{ text-align:right; }}
.tiles {{ display:grid; grid-template-columns: repeat(4, 1fr); gap: 8px; margin: 10px 0 14px; }}
.tile {{ border:1px solid var(--line); border-radius:6px; padding:10px; }}
.big {{ font-size: 17pt; font-weight: bold; color:#1f3864; }}
.cap {{ font-size: 8.5pt; color: var(--ink2); }}
ul {{ margin: 0 0 8px 18px; padding:0; }}
li {{ margin-bottom: 4px; }}
.note {{ background:#f4f3ef; border-left: 3px solid #1f3864; padding: 8px 10px; margin: 8px 0; font-size: 9.5pt; }}
.fig {{ margin: 6px 0 4px; }}
.figcap {{ font-size: 8.5pt; color: var(--muted); margin-bottom: 10px; }}
"""

body = []
body.append(f"""
<section class="page cover">
  <div>
    <div class="band"><div style="font-size:10pt;opacity:.85">Alexandria West Wastewater Treatment Plant</div>
      <h1>Planning package</h1>
      <div style="font-size:12pt">Underground pipe works sequence, methodology and programme · Whole-project estimate BOQ</div></div>
    <p style="margin-top:18px" class="sub">Issue date 22 October 2026 · Prepared by Claude for Mohamed (Projects Controls Director)</p>
    <p class="sub">Status: issued for information and planning. Not a design, temporary-works design or construction approval.</p>
  </div>
  <div>
    <h3>Contents</h3>
    <ol style="margin-left:18px">
      <li>Executive summary</li><li>Document register</li><li>Basis and decisions</li><li>Groundwater and temporary works</li>
      <li>Construction methodology</li><li>Programme</li><li>Aeration piling status</li><li>Estimate BOQ</li><li>Open items, exclusions and how to use the files</li>
    </ol>
    <p class="muted">Inputs: pipe execution packages v3 (21 Oct 2026), drawing ALEXW-MWT-MEC-00-DLU-002 Rev 03 (Issued for Approval), combined coordination model (19 Oct 2026), decisions of 20-22 Oct 2026.</p>
  </div>
</section>""")

body.append(f"""
<section class="page">
<h2>1. Executive summary</h2>
<div class="tiles">{tile_html}</div>
<ul>
<li><b>Groundwater at -1.50 changes the pipe works.</b> 199 of 201 packages excavate below groundwater, by up to 6.1 m. With the assumed limit of 2.5 m below GWL for open cut, 6,562 m (80%) of the 8,151 m of pipe needs sheet piles, about 116,300 m² of sheets, against 1,972 m before groundwater was considered.</li>
<li><b>Every pipe sits on piles and a piled slab.</b> That is about 5,800 pipe-support piles on a placeholder layout. Piling is the longest activity on most fronts and peaks at about 13 rigs in December 2026.</li>
<li><b>Aeration rigs can hand over.</b> 8 rigs are completing aeration piling at a measured 5.4 piles per rig per day; the forecast is end November 2026, just as pipe piling starts.</li>
<li><b>Flotation.</b> 14 packages on UG-08, UG-09 and UG-10 (DN2600, shallow cover) can float even at final level and need permanent straps to the piled slab. A further 72 need dewatering kept on until the fill reaches finished level.</li>
<li><b>Programme.</b> Eight fronts and 2,211 logic-linked activities give 1 December 2026 to 18 September 2027. This is a best case: approval gates and structure readiness dates are not yet entered.</li>
<li><b>Whole-project quantities.</b> About 145,800 m³ reinforced concrete, 17,300 t reinforcement, 221,000 m² formwork, 16,000 piles and 8,294 m of pipe, all from the 3D model and the pipe workbooks. The aeration raft reconciles to the aeration quantity file within 0.4%. The BOQ is unpriced, and structure excavation is still overstated.</li>
</ul>
<div class="note"><b>Decisions needed to firm this up:</b> confirm the groundwater design value and the open-cut rule (geotechnical); pipe and structure pile designs; structure readiness dates for tie-ins; DLU-002 approval forecast; unit rates.</div>
</section>""")

body.append('<section class="page"><h2>2. Document register</h2><p>All files are in the package folder and on branch <code>claude/new-session-e5ojdh</code> of repository engmohamedshipl-arch/claude.</p>'
            + table(['No.', 'Title', 'File', 'Format', 'Status'], [[E(a), E(b), f'<code>{E(c)}</code>', E(d), E(e)] for a, b, c, d, e in REG])
            + """<h3>3. Basis and decisions</h3>""" + table(['Ref', 'Date', 'Decision', 'Owner'], [
                ['D-001', 'before 20 Oct', 'Excavation bottom = external pipe bottom - 0.50 m (piled slab)', 'Mohamed'],
                ['D-004', '20 Oct', 'NG -0.70, FL +1.00, minimum cover 1.0 m', 'Mohamed'],
                ['-', '21 Oct', 'Working space 1.0 m; no maximum package length; shoring where open cut does not fit', 'Mohamed'],
                ['-', '22 Oct', 'Groundwater -1.50; all pipes on piles and piled slab; shoring as needed', 'Mohamed'],
                ['-', '22 Oct', 'Thickener bridge carrying the mechanism is reinforced concrete', 'Mohamed'],
                ['-', '18 Oct', 'Aeration piling order: zones 1, 2, 3, then 5+4, then 6, 7, 8', 'Agreed timeline']])
            + '<p class="muted">Assumptions by Claude (slopes 1.5 and 2.0 H:V, open-cut limit 2.5 m below GWL, pile layout, rebar ratios, productivity) are listed on each workbook\'s Inputs sheet and in the wiki decisions log.</p></section>')

body.append(f"""
<section class="page">
<h2>4. Groundwater and temporary works</h2>
<p>Sheet-piled length of the 8,151 m pipe network under alternative rules. The highlighted bar is the default in the workbooks. The rule is the single largest driver of temporary-works cost.</p>
<div class="fig">{hbar(scen, highlight=3, unit=' m', label_w=270)}</div>
<div class="figcap">Figure 1 - Sheet-piled pipe length by rule (m). Source: sequence workbook, Sensitivity sheet.</div>
{table(['Topic', 'Result'], [
    ['Dewatering', '160 packages single-stage wellpoints; 39 deep wells or two-stage wellpoints (drawdown beyond 4.5 m); 2 not required'],
    ['Just-in-time pumping', 'Drawdown timed to finish when the excavation crew is free; median 74 pumping days per package'],
    ['Flotation, final', '14 packages (UG-08, UG-09, UG-10): permanent anchor straps to the piled slab'],
    ['Flotation, construction', '72 packages: keep dewatering until fill reaches FL, or strap to slab'],
    ['Sheet piles', 'About 116,300 m² (wall length x 2 x excavation depth); temporary, extracted after backfill']])}
</section>""")

body.append('<section class="page"><h2>5. Construction methodology (per package)</h2>'
            + table(['Step', 'Activity', 'Method and hold points'], [[a, f'<b>{E(b)}</b>', E(c)] for a, b, c in METHOD])
            + """<h3>Sequencing rules</h3><ul>
<li>Eight fronts by system. A shared-trench corridor that crosses systems is built by one front.</li>
<li>Within a front: deeper lines first, downstream first, shared-trench packages back to back.</li>
<li>Deeper-first links where packages are within 8 m and bottoms differ by more than 1 m. Shared trenches (bottoms within 1 m) are opened together and backfilled after all pipes are laid (6 holds are site coordination items).</li>
<li>Tie-ins follow the adjacent run, wait for structure readiness and close the line.</li>
<li>Packages within 15 m of open aeration piles wait for that zone. Pipe piling waits for pile design approval; nothing starts before DLU-002 approval.</li></ul></section>""")

fr_rows = [[f'<b>{f[0]}</b>', E(f[1]), f'<span class="muted">{E(f[2])}</span>', f'<span style="display:block;text-align:right">{fmt(f[4])}</span>',
            f'<span style="display:block;text-align:right">{fmt(f[6])}</span>', str(f[7]),
            datetime.date.fromisoformat(f[9]).strftime('%d %b %Y')] for f in fr]
body.append(f"""
<section class="page">
<h2>6. Programme</h2>
<div class="fig">{gantt(fr)}</div>
<div class="figcap">Figure 2 - Pipe works by front (start 1 Dec 2026, all fronts). Source: sequence workbook; same logic in the P6 and MS Project files.</div>
{table(['Front', 'Scope', 'Lines', 'Length (m)', 'Sheet-piled (m)', 'Crews', 'Finish'], fr_rows)}
<div class="fig">{columns(months, rigs, ref=10)}</div>
<div class="figcap">Figure 3 - Equivalent piling rigs needed per month for pipe-support piles (piles per month / 5 per rig per day / 26 days). Demand above the dashed line is tie-in piling, which no front rig covers while structure dates are blank.</div>
</section>""")

aer = [['1', '470', '6', '2', '0'], ['2', '518', '41', '2', '3'], ['3', '248', '236', '2', '0'], ['4', '221', '337', '3', '0'],
       ['5', '223', '137', '4', '0'], ['6', '134', '279', '5', '0'], ['7', '121', '296', '1', '0'], ['8', '144', '236', '0', '3'],
       ['<b>Total</b>', '<b>2,079</b>', '<b>1,568</b>', '<b>19</b>', '<b>6</b>']]
body.append('<section class="page"><h2>7. Aeration piling status (model, data date 17 Oct 2026)</h2>'
            + table(['Zone', 'Done', 'Pending', 'Defect', 'Unknown'], aer)
            + """<p>"Done" means recorded execution, not acceptance. Over the last 4 full weeks 8-9 rigs averaged 285 piles a week, about 5.4 piles per rig per working day. In the agreed zone order the 1,587 open piles finish around 29 November 2026. That forecast gates the pipe packages within 15 m of open aeration piles and frees the rigs for pipe piling.</p>
<h2 style="margin-top:18px">8. Estimate BOQ (whole project)</h2>"""
            + f'<div class="fig">{hbar(top, label_w=270, unit=" m³")}</div><div class="figcap">Figure 4 - Reinforced concrete by bill (m³). Source: estimate BOQ, measured on closed 3D model solids.</div></section>')

KEY = [('All RC', 'All reinforced concrete', 'm³'), ('REBAR', 'Reinforcement (ratio basis)', 't'), ('FORMWORK', 'Formwork', 'm²'), ('BLINDING', 'Blinding', 'm³'),
       ('PILE', 'Piles 600 mm (6,392 model, ~3,800 allowance, 5,788 pipe placeholder)', 'no.'), ('EXC', 'Excavation (structures overstated)', 'm³'),
       ('BACKFILL', 'Backfill', 'm³'), ('SHEET', 'Temporary sheet piles (pipes)', 'm²'), ('PIPE', 'Pipes', 'm'), ('FITTING', 'Pipe fittings', 'no.'),
       ('HANDRAIL', 'Handrails', 'm'), ('SITEFILL', 'Site raising NG to FL (provisional)', 'm³')]
body.append('<section class="page"><h3>Key quantities (gross, no deduction for work done)</h3>'
            + table(['Item', 'Quantity', 'Unit'], [[E(d), f'<span style="display:block;text-align:right">{fmt(kq[c])}</span>', u] for c, d, u in KEY])
            + """<ul>
<li><b>Cross-checks:</b> aeration raft 30,120 m³ in the model against 30,013 m³ in the aeration quantity file (0.4%); aeration piles 3,672, the same as the tracking sheet.</li>
<li><b>Confidence</b> is graded on each BOQ line: B = detailed model; C = civil-guide envelope; D = tender envelope or allowance (digesters, gas holders, sludge buildings, CHP, THP and others).</li>
<li><b>Weak points:</b> structure excavation is overstated; the thickener bridge is concrete by decision, but the model members are slender (21 m³ in total); only the raft rebar ratio is evidence-based; pile lengths are unknown.</li>
<li><b>Rates</b> are not included. Enter them in the yellow column and the summary prices itself.</li></ul>
<h2 style="margin-top:14px">9. Open items</h2>""" + table(['Owner', 'Item', 'Affects'], [[E(a), E(b), E(c)] for a, b, c in OPEN]) + '</section>')

body.append("""<section class="page"><h2>Exclusions and how to use the files</h2>
<h3>Not included</h3><ul>
<li>Preliminaries, contingency, M&amp;E and process equipment, architecture and finishes, roads, drainage, fencing and utilities.</li>
<li>Temporary works and dewatering for structures (pipe temporary works are included).</li>
<li>Existing structures and boundaries not in the model: existing PST, bypass and outfall chambers, treated effluent pump station, primary and thickened sludge pump stations, sludge flowmeter chambers.</li></ul>
<h3>Using the files</h3><ul>
<li><b>Workbooks:</b> yellow cells are inputs. Enter approval dates (Inputs), structure readiness (Structures), crews (Fronts) and rates (BOQ); everything recalculates.</li>
<li><b>P6:</b> File &gt; Import &gt; XER as a new project, then F9. Calendar Sat-Thu 8 h; WBS Project &gt; Front &gt; Line &gt; Package; SS+FF ladders kept.</li>
<li><b>MS Project:</b> open the XML. One link per task pair (the driving link is kept).</li>
<li><b>Keeping in step:</b> after changing inputs, re-export the programme files with <code>export_schedule.py</code> (see the folder README).</li>
<li><b>Wiki:</b> the decisions log and open items are the living record. Copy the Markdown pages into the repository's GitHub Wiki tab when it is enabled.</li></ul>
<div class="note">Common vertical datum not verified (R001). Levels are envelope levels, not inverts (R002), and are interpolated between drawing anchors (R003). DLU-002 Rev 03 is Issued for Approval. These holds apply to every pipe package.</div>
</section>""")

doc = f'<!doctype html><html lang="en"><head><meta charset="utf-8"><title>Alex West Planning Package</title><style>{css}</style></head><body>{"".join(body)}</body></html>'
open(OUT, 'w').write(doc)
print('written', OUT)
