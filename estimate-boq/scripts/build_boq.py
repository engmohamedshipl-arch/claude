"""Estimate BOQ for Alex West WWTP from the 3D coordination model and the pipe workbooks.
Args: objects.json hulls.json manifest.json package_workbook.xlsx sequence_workbook.xlsx out.xlsx
"""
import sys, re, json, math, collections, datetime
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as CL

OBJ, HULL, MAN, PKGWB, SEQWB, OUT = sys.argv[1:7]
objs = json.load(open(OBJ))
hulls = json.load(open(HULL))
manifest = {a['id']: a for a in json.load(open(MAN))['assets']}
NG = -0.7

# ----------------------------------------------------------------- classification of model elements
def akey(col):
    if col.startswith('01 Aeration') or col.startswith('02 Aeration'):
        return 'AW-AER'
    m = re.match(r'FST ([A-L])', col)
    if m:
        return 'AW-FST-' + m.group(1)
    if col.startswith('Pipes'):
        return 'PIPES'
    return col.split(' | ')[0]

def has(s, *ks):
    return any(k in s for k in ks)

def classify(name, col, cat):
    n = name.lower(); c = (cat or '').lower(); cl = col.lower()
    if c == 'water' or 'hydraulic reference' in n or re.search(r'\bwl \+', n):
        return 'EXCL'
    if has(n, 'display stub', ' pile') or '08 piles' in cl or c.startswith('08 piles'):
        return 'PILE'
    if 'steel truss' in cl:
        return 'RC_BRIDGE'   # Mohamed 22 Oct 2026: thickener bridge carrying the mechanism is reinforced concrete
    if has(n, 'blinding', 'screed', 'unspecified lower continuous layer'):
        return 'BLINDING'
    if has(n, 'rail', 'post') and not has(n, 'trail'):
        return 'HANDRAIL'
    if c.startswith('09 handrails') or c == 'rails' or 'indicative rails' in cl:
        return 'HANDRAIL'
    if has(n, 'foundation', 'stair foundation'):
        return 'RC_BASE'
    if has(n, 'stair', 'tread', 'ladder', 'riser') or (c == 'access' and not has(n, 'deck', 'ramp', 'landing support')):
        return 'STAIRS'
    if has(n, 'truss') or 'steel truss' in cl:
        return 'STEEL'
    if 'bridge deck' in cl or (has(n, 'bridge deck')):
        return 'GRATING'
    if has(n, 'crane', 'girder', 'runway', 'cabinet', 'blower-') and 'plinth' not in n or c in ('equipment', 'cranes', 'crane'):
        return 'EQUIP'
    if has(n, 'door', 'glazing') or c in ('doors', 'glazing'):
        return 'ARCH'
    if has(n, 'fill', 'floor slope'):
        return 'MASS_FILL'
    if has(n, 'plinth', 'skid pad') or c in ('plinth', 'plinths', 'pads'):
        return 'RC_PLINTH'
    if has(n, 'column', 'beam', 'pier') or re.match(r'^(aw-thk-[a-d] \| )?c\d', n):
        return 'RC_COLBEAM'
    if has(n, 'wall', 'divider', 'partition', 'parapet', 'return', 'stem', 'island') or c in ('walls', 'wall', 'front', 'upper', 'parapet') \
            or c.startswith('02 tapered') or 'tank wall' in cl:
        return 'RC_WALL'
    if has(n, 'raft', 'base', 'floor', 'foundation', 'ground slab') or c in ('raft', 'base', 'foundation', 'floor', 'lower') \
            or c.startswith('01 raft') or 'raft' in cl:
        return 'RC_BASE'
    if has(n, 'slab', 'deck', 'roof', 'walkway', 'cap', 'cover', 'landing', 'platform', 'spine', 'strip', 'bridge', 'ramp') \
            or c in ('deck', 'roof', 'walkway', 'cover') or c.startswith('07 walkways'):
        return 'RC_SLAB'
    return 'RC_MISC'

rows = []
for o in objs:
    cols = [c for c in o['cols'] if not c.startswith('AW-AER Status')]
    if not cols:
        continue
    col = cols[0]; k = akey(col)
    if k == 'PIPES':
        continue
    cat = o['props'].get('Category')
    cls = classify(o['name'], col, cat)
    rail = max(o['ldim']) if cls == 'HANDRAIL' and 'post' not in o['name'].lower() else 0.0
    zb = o['bb'][4]
    soffit = o['dn'] if cls in ('RC_SLAB', 'RC_COLBEAM', 'RC_MISC', 'RC_BRIDGE') and zb > NG + 0.3 else 0.0
    form = (o['vert'] + soffit) if cls.startswith('RC_') else 0.0
    rows.append(dict(asset=k, col=col, name=o['name'], cls=cls, vol=o['vol'], up=o['up'], form=form, rail=rail, zb=zb))
el = pd.DataFrame(rows)

STRUCTS = [  # type key, title, member assets, manifest id for maturity, confidence
    ('GRT', 'Grit removal and fine screens', ['AW-GRT'], 'AW-GRIT', 'C'),
    ('PST', 'Primary sedimentation tanks (8 basins)', ['AW-PST'], 'AW-PST', 'C'),
    ('IPS', 'Intermediate pumping station incl. shared IPS/RAS motor hall', ['AW-IPS', 'AW-IPS-SHARED'], 'AW-IPS', 'C'),
    ('AER', 'Aeration tanks', ['AW-AER'], 'AW-AER', 'B'),
    ('SDC', 'Secondary distribution chambers A-C (3 nr)', ['AW-SDC-A', 'AW-SDC-B', 'AW-SDC-C'], 'AW-SDC-A', 'C'),
    ('FST', 'Final sedimentation tanks A-L (12 nr)', [f'AW-FST-{c}' for c in 'ABCDEFGHIJKL'], 'AW-FST-A', 'B'),
    ('RAS', 'RAS/WAS pumping station', ['AW-RASWAS'], 'AW-RASWAS', 'C'),
    ('CCT', 'Chlorine contact tank', ['AW-CCT'], 'AW-CCT', 'C'),
    ('FIRE', 'Fire / service / cooling water structure', ['AW-FIRE'], 'AW-FIRE', 'C'),
    ('BLOW', 'Air blower building', ['AW-BLOW'], 'AW-BLOW', 'C'),
    ('THK', 'Primary sludge thickeners A-D (4 nr)', [f'AW-THK-{c}' for c in 'ABCD'], 'AW-THK-A', 'C'),
    ('DIG', 'Anaerobic digesters 01-08 (8 nr)', [f'AW-DIG-0{i}' for i in range(1, 9)], 'AW-DIG-01', 'D'),
    ('GHT', 'Gas holder plinths 01-04 (4 nr)', [f'AW-GHT-0{i}' for i in range(1, 5)], 'AW-GHT-01', 'D'),
    ('GHB', 'Gas holder blower house', ['AW-GHB'], 'AW-GHB', 'D'),
    ('DSB', 'Digested sludge buildings 01-02 (2 nr)', ['AW-DSB-01', 'AW-DSB-02'], 'AW-DSB-01', 'D'),
    ('GPS', 'Grease pump station', ['AW-GPS'], 'AW-GPS', 'D'),
    ('SCPS', 'Scum pump station', ['AW-SCPS'], 'AW-SCPS', 'D'),
    ('WPD', 'WAS pre-dewatering building', ['AW-WPD'], 'AW-WPD', 'D'),
    ('THP', 'THP system platform', ['AW-THP'], 'AW-THP', 'D'),
    ('FDW', 'Final dewatering building', ['AW-FDW'], 'AW-FDW', 'D'),
    ('CHP', 'CHP building', ['AW-CHP'], 'AW-CHP', 'D'),
    ('CHS', 'Chlorination storage building', ['AW-CHS'], 'AW-CHS', 'D'),
    ('ABPH', 'Air blowers power house', ['AW-ABPH'], 'AW-ABPH', 'D'),
]
type_of = {a: t for t, _, ms, _, _ in STRUCTS for a in ms}
el['type'] = el.asset.map(type_of)
assert el.type.notna().all(), el[el.type.isna()].asset.unique()

# ----------------------------------------------------------------- pipe data
seg = pd.read_excel(PKGWB, 'Segments', header=3)
seg = seg[seg['Segment ID'].astype(str).str.contains('-S')]
lines = pd.read_excel(PKGWB, 'Lines', header=3)
lines = lines[lines['Line'].astype(str).str.startswith('UG')]
pcs = pd.read_excel(PKGWB, 'Pieces', header=3)
pcs = pcs[pcs['Line'].astype(str).str.startswith('UG')]
seqk = pd.read_excel(SEQWB, 'Packages', header=4)
seqk = seqk[seqk['Package'].astype(str).str.startswith('UG')]
FIT = collections.OrderedDict([('Elbow 30', r'elbow 30'), ('Elbow 45', r'elbow 45'), ('Elbow 90', r'elbow 90'), ('Elbow (angle not stated)', r'elbow piece'),
                               ('Flange', r'flange'), ('Reducer', r'reducer'), ('Tee', r'tee'), ('Stub end', r'stub end')])
fit_line = collections.defaultdict(collections.Counter)
for _, r in seg.iterrows():
    for n, k in re.findall(r'(\d+)\s*x\s*([^,]+)', str(r['Fittings'])):
        k = k.strip().lower()
        for lab, pat in FIT.items():
            if re.match(pat, k):
                fit_line[r['Line']][lab] += int(n); break

# ----------------------------------------------------------------- workbook helpers
wb = Workbook()
F = 'Arial'
fnt = Font(name=F, size=10); bold = Font(name=F, size=10, bold=True)
title = Font(name=F, size=14, bold=True, color='1F3864'); sub = Font(name=F, size=10, italic=True, color='595959')
hdr_font = Font(name=F, size=10, bold=True, color='FFFFFF'); blue = Font(name=F, size=10, color='0000FF')
green = Font(name=F, size=10, color='008000')
hdr_fill = PatternFill('solid', fgColor='1F3864'); sec_fill = PatternFill('solid', fgColor='D9E1F2')
yel = PatternFill('solid', fgColor='FFFF00'); grey_fill = PatternFill('solid', fgColor='F2F2F2')
thin = Side(style='thin', color='BFBFBF'); box = Border(left=thin, right=thin, top=thin, bottom=thin)
wrap = Alignment(wrap_text=True, vertical='top'); center = Alignment(horizontal='center', vertical='center', wrap_text=True)
NUM = '#,##0;(#,##0);"-"'; NUM1 = '#,##0.0;(#,##0.0);"-"'

def header(ws, row, cols, widths):
    for j, h in enumerate(cols, 1):
        c = ws.cell(row=row, column=j, value=h); c.font = hdr_font; c.fill = hdr_fill; c.alignment = center; c.border = box
    ws.row_dimensions[row].height = 32
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[CL(j)].width = w

def put(ws, r, c, v, font=fnt, fmt=None, fill=None, align=None):
    cell = ws.cell(row=r, column=c, value=v); cell.font = font; cell.border = box
    if fmt: cell.number_format = fmt
    if fill: cell.fill = fill
    if align: cell.alignment = align
    return cell

def name(n, ref):
    wb.defined_names[n] = DefinedName(n, attr_text=ref)

wsR = wb.active; wsR.title = 'Read Me'
wsS = wb.create_sheet('Summary')
wsB = wb.create_sheet('BOQ')
wsI = wb.create_sheet('Inputs')
wsT = wb.create_sheet('Take-off')
wsX = wb.create_sheet('Structures')
wsP = wb.create_sheet('Pipes')
wsE = wb.create_sheet('Exclusions')

# ----------------------------------------------------------------- Inputs
wsI['A1'] = 'Inputs and assumptions'; wsI['A1'].font = title
wsI['A2'] = 'Yellow = edit. Quantities in the BOQ recalculate from these values. Rates are entered on the BOQ sheet.'; wsI['A2'].font = sub
header(wsI, 4, ['Name', 'Description', 'Value', 'Unit', 'Source / status'], [16, 64, 12, 10, 90])
INPUTS = [
    ('S', 'General'),
    ('CURRENCY', 'Currency of rates', 'EGP', '', 'Enter rates in this currency on the BOQ sheet.'),
    ('NG', 'Adopted natural ground level', -0.7, 'm', 'Package workbook D-004 (Mohamed, 20 Oct 2026).'),
    ('FL', 'Finished ground level', 1.0, 'm', 'Package workbook D-004.'),
    ('S', 'Earthworks to structures'),
    ('WS_STR', 'Working space around structure footprint at formation', 1.0, 'm', 'Same as agreed for pipes (21 Oct 2026).'),
    ('SLOPE_STR', 'Excavation side slope (H per 1 V); 0 = vertical (shored)', 1.5, 'H:V', 'ASSUMPTION - soft silty clay. Temporary works design governs.'),
    ('S', 'Piling'),
    ('PILE_DENS', 'Allowance where the model has no piles: Ø600 piles per m2 of footprint', 0.10, 'no./m2', 'Calibrated to aeration: 3,672 modelled piles / 34,995 m2 footprint = 0.105. ALLOWANCE only.'),
    ('S', 'Reinforcement ratios (kg per m3 of concrete) - ASSUMPTIONS unless noted'),
    ('RR_BASE', 'Rafts, bases and foundations', 107, 'kg/m3', 'Aeration raft estimate 3,203 t / 30,012 m3 (model text "AW-AER Tracking Documentation", incl. 10% laps and 10% extra bars).'),
    ('RR_WALL', 'Walls and parapets', 140, 'kg/m3', 'ASSUMPTION - water-retaining walls.'),
    ('RR_SLAB', 'Suspended slabs, decks, roofs, walkways', 130, 'kg/m3', 'ASSUMPTION.'),
    ('RR_COLBEAM', 'Columns, beams, piers', 180, 'kg/m3', 'ASSUMPTION.'),
    ('RR_MISC', 'Channels, baffles, weirs, chambers, hoppers', 130, 'kg/m3', 'ASSUMPTION.'),
    ('RR_PLINTH', 'Equipment plinths and pads', 80, 'kg/m3', 'ASSUMPTION.'),
    ('RR_BRIDGE', 'Concrete bridge members (thickeners)', 200, 'kg/m3', 'ASSUMPTION - slender bridge members. Bridge is concrete (Mohamed, 22 Oct 2026); model collection is labelled "Steel truss".'),
    ('RR_PIPESLAB', 'Piled slab under pipes', 110, 'kg/m3', 'ASSUMPTION - no slab design.'),
    ('S', 'Steel and metalwork'),
    ('STEEL_DENS', 'Density of steel', 7.85, 't/m3', 'Physical constant; applied to modelled steel member volume.'),
    ('RAILS_PER_RUN', 'Rail members per handrail run (top + knee rail)', 2, 'no.', 'ASSUMPTION - modelled rails are individual members.'),
    ('S', 'Underground pipework'),
    ('WS_PIPE', 'Working space each side of pipe (trench bottom width = OD + 2 x WS)', 1.0, 'm', 'AGREED 21 Oct 2026.'),
    ('SURR_ABOVE', 'Granular surround above pipe crown', 0.3, 'm', 'ASSUMPTION.'),
    ('S', 'Site-wide provisional'),
    ('SITE_AREA', 'Developed area for site raising (convex hull of all modelled assets)', round(hulls['_SITE_ALL']['area']), 'm2', 'Computed from the 3D model; site boundary not in the model - replace with surveyed area.'),
    ('SITE_DEDUCT', 'Deduct footprint of structures from site raising (1 = yes)', 1, '-', ''),
]
r = 5
for it in INPUTS:
    if it[0] == 'S':
        wsI.cell(row=r, column=1, value=it[1]).font = bold
        for c in range(1, 6): wsI.cell(row=r, column=c).fill = sec_fill
        r += 1; continue
    n, d, v, u, s = it
    put(wsI, r, 1, n, bold); put(wsI, r, 2, d, align=wrap)
    put(wsI, r, 3, v, blue, fill=yel, fmt='0.00' if isinstance(v, float) and v != int(v) else None)
    put(wsI, r, 4, u); put(wsI, r, 5, s, align=wrap)
    name(n, f'Inputs!$C${r}')
    r += 1
wsI.freeze_panes = 'A5'

# ----------------------------------------------------------------- Take-off (aggregated model data)
wsT['A1'] = 'Take-off from the 3D coordination model (Alex-West-Combined-Coordination.blend, 19 Oct 2026)'; wsT['A1'].font = title
wsT['A2'] = ('One row per structure / element class. Volumes and areas are measured on closed model solids (all 15,195 meshes are closed). '
             'Formwork = vertical faces + soffits of elevated elements. Blue = model data; the BOQ sums these rows.'); wsT['A2'].font = sub
th = ['Bill', 'Asset', 'Element class', 'Elements (no.)', 'Volume (m3)', 'Top area (m2)', 'Formwork area (m2)', 'Rail member length (m)', 'Lowest level (m)', 'Element names (sample)']
header(wsT, 4, th, [7, 14, 13, 10, 12, 12, 12, 12, 10, 90])
agg = el.groupby(['type', 'asset', 'cls']).agg(n=('vol', 'size'), vol=('vol', 'sum'), up=('up', 'sum'), form=('form', 'sum'),
                                                rail=('rail', 'sum'), zb=('zb', 'min'),
                                                names=('name', lambda s: '; '.join(sorted(set(re.sub(r'[\s_.]*\d+$', '', re.sub(r'^(FST-[A-L]|AW-[A-Z\-0-9]+) \| ', '', x)) for x in s))[:6]))).reset_index()
torder = {t: i for i, (t, *_ ) in enumerate(STRUCTS)}
agg['o'] = agg.type.map(torder)
agg = agg.sort_values(['o', 'asset', 'cls'])
T0 = 5
for i, rr_ in enumerate(agg.itertuples()):
    rr = T0 + i
    for j, v in enumerate([rr_.type, rr_.asset, rr_.cls, int(rr_.n), rr_.vol, rr_.up, rr_.form, rr_.rail, rr_.zb, rr_.names], 1):
        c = put(wsT, rr, j, v, blue)
        if j in (5, 6, 7, 8): c.number_format = NUM1
        if j == 9: c.number_format = '0.00'
TN = T0 + len(agg) - 1
wsT.freeze_panes = 'D5'; wsT.auto_filter.ref = f'A4:J{TN}'
def TK(col):
    return f"'Take-off'!${col}${T0}:${col}${TN}"

# ----------------------------------------------------------------- Structures (earthworks, piles)
wsX['A1'] = 'Structures - footprint, excavation, backfill and piles'; wsX['A1'].font = title
wsX['A2'] = ('Footprint = plan convex hull of elements that go below NG. Excavation from NG to the lowest element with working space and side slope '
             '(volume of an offset prism). Backfill = excavation - footprint x depth. Piles: modelled count where the model has piles, otherwise an allowance.'); wsX['A2'].font = sub
xh = ['Bill', 'Asset', 'Structure', 'Model stage (manifest 19 Oct 2026)', 'Footprint below NG (m2)', 'Footprint perimeter (m)', 'Lowest level (m)',
      'Depth below NG (m)', 'Excavation (m3)', 'Backfill (m3)', 'Base top area (m2)', 'Piles in model (no.)', 'Pile basis', 'Piles for BOQ (no.)']
header(wsX, 4, xh, [7, 14, 34, 40, 12, 11, 10, 10, 12, 12, 12, 11, 22, 11])
X0 = 5
xrow = {}
pile_model = el[el.cls == 'PILE'].groupby('asset').size().to_dict()
base_up = el[el.cls.isin(['RC_BASE', 'RC_PLINTH'])].groupby('asset').up.sum().to_dict()
zmin_all = el[~el.cls.isin(['EXCL', 'HANDRAIL', 'STAIRS', 'EQUIP', 'ARCH', 'PILE'])].groupby('asset').zb.min().to_dict()
i = 0
for t, tt, ms, mid, conf in STRUCTS:
    for a in ms:
        rr = X0 + i; xrow[a] = rr; i += 1
        h = hulls.get(a, dict(area=0, perim=0, zmin=zmin_all.get(a, 0)))
        man = manifest.get(a) or manifest.get(mid) or {}
        put(wsX, rr, 1, t, blue); put(wsX, rr, 2, a, blue)
        put(wsX, rr, 3, (man.get('name') if a in manifest else tt), blue)
        put(wsX, rr, 4, man.get('stage', 'Tender / provisional reconstruction'), blue, align=Alignment(wrap_text=False))
        put(wsX, rr, 5, round(h['area'], 1), blue, fmt=NUM); put(wsX, rr, 6, round(h['perim'], 1), blue, fmt=NUM)
        put(wsX, rr, 7, round(zmin_all.get(a, h['zmin']), 2), blue, fmt='0.00')
        put(wsX, rr, 8, f'=IF(E{rr}=0,0,MAX(0,NG-G{rr}))', fmt='0.00')
        put(wsX, rr, 9, f'=H{rr}*E{rr}+F{rr}*(WS_STR*H{rr}+SLOPE_STR*H{rr}^2/2)+PI()*(WS_STR^2*H{rr}+WS_STR*SLOPE_STR*H{rr}^2+SLOPE_STR^2*H{rr}^3/3)', fmt=NUM)
        put(wsX, rr, 10, f'=MAX(0,I{rr}-E{rr}*H{rr})', fmt=NUM)
        put(wsX, rr, 11, round(base_up.get(a, 0), 1), blue, fmt=NUM)
        pm = pile_model.get(a, 0)
        put(wsX, rr, 12, pm, blue, fmt=NUM)
        if pm:
            basis = 'Model (display stubs; FST stubs are placeholders)' if a.startswith('AW-FST') else 'Model (display stubs)'
            put(wsX, rr, 13, basis, blue); put(wsX, rr, 14, f'=L{rr}', fmt=NUM)
        else:
            put(wsX, rr, 13, 'Allowance: footprint x PILE_DENS', blue)
            put(wsX, rr, 14, f'=ROUNDUP(MAX(E{rr},K{rr})*PILE_DENS,0)', fmt=NUM)
XN = X0 + i - 1
put(wsX, XN + 1, 3, 'Total', bold)
for col in 'EIJLN':
    put(wsX, XN + 1, wsX[col + '1'].column, f'=SUM({col}{X0}:{col}{XN})', bold, fmt=NUM)
wsX.freeze_panes = 'D5'
def XR(col):
    return f'Structures!${col}${X0}:${col}${XN}'

# ----------------------------------------------------------------- Pipes (per line)
wsP['A1'] = 'Underground pipework - quantities per line'; wsP['A1'].font = title
wsP['A2'] = ('Lengths, sizes and fittings from the v3 package workbook (21 Oct 2026). Excavation, sheet piles, piled slab, piles, backfill, dewatering and flotation '
             'from the sequence workbook (22 Oct 2026: GWL -1.50, all pipes on piled slabs). Surround is a formula.'); wsP['A2'].font = sub
fit_cols = list(FIT.keys())
ph = ['Line', 'Material', 'Printed size', 'Planning OD (m)', '3D length (m)', 'Packages', 'Tie-in / connection ends'] + fit_cols + [
      'Excavation shored (m3)', 'Excavation open cut (m3)', 'Sheet pile area (m2)', 'Sheet pile wall (m)', 'Piled slab concrete (m3)', 'Pipe-support piles (no.)',
      'Backfill (m3)', 'Granular surround (m3)', 'Dewatered length (m)', 'Dewatering package-days', 'Pipe length needing anchor straps (m)', 'Crossings to locate (no. packages)']
header(wsP, 4, ph, [9, 8, 14, 9, 10, 9, 9] + [8] * len(fit_cols) + [11, 11, 11, 10, 10, 10, 11, 11, 10, 10, 11, 10])
P0 = 5
od = pcs.groupby('Line')['Planning dia (m)'].max().to_dict()
for i, lr in enumerate(lines.itertuples()):
    rr = P0 + i; l = lr.Line
    sk = seqk[seqk.Line == l]; sg = seg[seg.Line == l]
    ends = int((sk['Package type'] == 'Tie-in').sum() + 2 * (sk['Package type'] == 'Short connection').sum())
    vals = [l, lr._2, str(lr._3), round(od[l], 3), lr._11, len(sk), ends] + [fit_line[l][k] for k in fit_cols] + [
        sk['Excavation shored (m3)'].sum(), sk['Excavation open cut (m3)'].sum(), sk['Sheet pile area (m2)'].sum(), sk['Sheet pile wall (m)'].sum(),
        sk['Piled slab concrete (m3)'].sum(), sk['Pipe-support piles (no.)'].sum(), sk['Backfill (m3)'].sum(), None,
        sk.loc[sk['Drawdown required below GWL (m)'] > 0, 'Plan length (m)'].sum(), sk['Dewatering running (days)'].sum(),
        sk.loc[sk['Flotation action'].str.startswith('Anchor'), 'Plan length (m)'].sum(),
        int((sg['Crossings and parallel lines'].astype(str) != 'None.').sum())]
    for j, v in enumerate(vals, 1):
        c = put(wsP, rr, j, v, blue)
        if j >= 4: c.number_format = NUM1 if j in (4, 5) else NUM
    sc = 7 + len(fit_cols) + 8
    put(wsP, rr, sc, f'=E{rr}*((D{rr}+2*WS_PIPE)*(D{rr}+SURR_ABOVE)-PI()*D{rr}^2/4)', fmt=NUM)
PN = P0 + len(lines) - 1
put(wsP, PN + 1, 1, 'Total', bold)
for j in range(5, len(ph) + 1):
    put(wsP, PN + 1, j, f'=SUM({CL(j)}{P0}:{CL(j)}{PN})', bold, fmt=NUM)
wsP.freeze_panes = 'D5'
def PR(j):
    return f'Pipes!${CL(j)}${P0}:${CL(j)}${PN}'
PC = {h: j for j, h in enumerate(ph, 1)}

# ----------------------------------------------------------------- BOQ
wsB['A1'] = 'Alex West WWTP - estimate bill of quantities (whole project, from available information)'; wsB['A1'].font = title
wsB['A2'] = ('Quantities are formulas on the Take-off, Structures and Pipes sheets. Enter rates in the yellow column (currency on Inputs). '
             'Confidence: B = structural / detailed model; C = civil-guide or provisional envelope; D = tender envelope, placeholder or allowance.'); wsB['A2'].font = sub
bh = ['Bill', 'Item', 'Description', 'Unit', 'Quantity', 'Rate', 'Amount', 'Code', 'Basis / source', 'Confidence']
header(wsB, 4, bh, [6, 8, 70, 7, 13, 12, 15, 11, 60, 10])
B0 = 5
row = B0
bills = []

def bill_header(code, txt):
    global row
    for j in range(1, 11):
        c = wsB.cell(row=row, column=j); c.fill = sec_fill; c.border = box
    wsB.cell(row=row, column=1, value=code).font = bold
    wsB.cell(row=row, column=3, value=txt).font = bold
    bills.append((code, txt, row)); row += 1

def item(code, n, desc, unit, qty_f, ccode, basis, conf, info=False):
    global row
    put(wsB, row, 1, code); put(wsB, row, 2, f'{code}.{n:02d}')
    put(wsB, row, 3, desc, align=Alignment(wrap_text=True, vertical='top')); put(wsB, row, 4, unit)
    put(wsB, row, 5, qty_f, fmt=NUM if unit not in ('t',) else NUM1)
    if info:
        put(wsB, row, 6, 'excluded', Font(name=F, size=9, italic=True, color='808080')); put(wsB, row, 7, 0, fmt=NUM)
    else:
        put(wsB, row, 6, None, blue, fill=yel, fmt=NUM1); put(wsB, row, 7, f'=E{row}*F{row}', fmt=NUM)
    put(wsB, row, 8, ccode); put(wsB, row, 9, basis, align=Alignment(wrap_text=True, vertical='top')); put(wsB, row, 10, conf)
    row += 1

def tk(t, cls, col='E'):
    return f'SUMIFS({TK(col)},{TK("A")},"{t}",{TK("C")},"{cls}")'

RC = [('RC_BASE', 'Reinforced concrete rafts, bases, floors and foundations', 'RR_BASE'),
      ('RC_WALL', 'Reinforced concrete walls and parapets', 'RR_WALL'),
      ('RC_SLAB', 'Reinforced concrete suspended slabs, decks, roofs and walkways', 'RR_SLAB'),
      ('RC_COLBEAM', 'Reinforced concrete columns, beams and piers', 'RR_COLBEAM'),
      ('RC_MISC', 'Reinforced concrete channels, baffles, weirs, chambers, hoppers and secondary elements', 'RR_MISC'),
      ('RC_PLINTH', 'Reinforced concrete equipment plinths and pads', 'RR_PLINTH'),
      ('RC_BRIDGE', 'Reinforced concrete bridge carrying the thickener mechanism (beams, verticals, diagonals, bracing as modelled)', 'RR_BRIDGE')]
for bi, (t, tt, ms, mid, conf) in enumerate(STRUCTS, 1):
    code = f'{bi:02d}'
    bill_header(code, tt)
    sub_ = el[el.type == t]
    has_cls = set(sub_.cls)
    n = 1
    xsum = lambda col: f'SUMIFS({XR(col)},{XR("A")},"{t}")'
    item(code, n, 'Excavation below NG (-0.70) to underside of lowest element, incl. working space and side slopes; dispose / stockpile', 'm3',
         '=' + xsum('I'), 'EXC', 'Structures sheet: model footprint and lowest level; WS_STR, SLOPE_STR', conf); n += 1
    item(code, n, 'Backfill around structure to NG in compacted layers', 'm3', '=' + xsum('J'), 'BACKFILL', 'Structures sheet', conf); n += 1
    pm = sum(pile_model.get(a, 0) for a in ms)
    item(code, n, 'Piles 600 mm diameter (CFA / bored) - length to pile design' + ('' if pm else ' - ALLOWANCE (no pile layout in model)'), 'no.',
         '=' + xsum('N'), 'PILE', 'Model pile count' + (' (FST stubs are placeholders)' if t == 'FST' else '') if pm else 'Footprint x PILE_DENS',
         conf if pm else 'D'); n += 1
    if 'BLINDING' in has_cls:
        item(code, n, 'Plain concrete blinding / screed', 'm3', '=' + tk(t, 'BLINDING'), 'BLINDING', 'Model volume', conf); n += 1
    if 'MASS_FILL' in has_cls:
        item(code, n, 'Mass concrete fill / benching to sloped floors', 'm3', '=' + tk(t, 'MASS_FILL'), 'MASS_FILL', 'Model volume', conf); n += 1
    rc_present = [x for x in RC if x[0] in has_cls]
    for cls, desc, rr_name in rc_present:
        item(code, n, desc, 'm3', '=' + tk(t, cls), cls, 'Model volume', conf); n += 1
    if rc_present:
        item(code, n, 'Formwork to vertical faces and soffits of elevated elements', 'm2',
             '=' + '+'.join(tk(t, c, 'G') for c, _, _ in rc_present), 'FORMWORK', 'Model face areas', conf); n += 1
        item(code, n, 'High-yield steel reinforcement', 't',
             '=(' + '+'.join(f'{tk(t, c)}*{rn}' for c, _, rn in rc_present) + ')/1000', 'REBAR', 'Concrete volume x ratio (Inputs)', 'D' if conf != 'B' else 'C'); n += 1
    if 'STEEL' in has_cls:
        item(code, n, 'Structural steelwork', 't', '=' + tk(t, 'STEEL') + '*STEEL_DENS', 'STEEL', 'Model member volume x density', conf); n += 1
    if 'GRATING' in has_cls:
        item(code, n, '6 mm steel chequer-plate deck on thickener bridge', 'm2', '=' + tk(t, 'GRATING', 'F'), 'GRATING', 'Model top area', conf); n += 1
    if 'HANDRAIL' in has_cls:
        item(code, n, 'Handrails (run length)', 'm', '=' + tk(t, 'HANDRAIL', 'H') + '/RAILS_PER_RUN', 'HANDRAIL', 'Model rail members / RAILS_PER_RUN; indicative in model', 'D'); n += 1
    if 'STAIRS' in has_cls:
        item(code, n, 'Stairs, treads, landings and access steps (material to design) - modelled volume', 'm3', '=' + tk(t, 'STAIRS'), 'STAIRS',
             'Model volume; ' + str(int((sub_.cls == "STAIRS").sum())) + ' modelled elements', 'D'); n += 1
    if 'EQUIP' in has_cls:
        item(code, n, 'Equipment envelopes in model (blowers, cranes) - M&E supply, count for reference', 'no.', '=' + tk(t, 'EQUIP', 'D'), 'EQUIP',
             'Model element count', conf, info=True); n += 1
    if 'ARCH' in has_cls:
        item(code, n, 'Doors and glazing envelopes in model - architectural, count for reference', 'no.', '=' + tk(t, 'ARCH', 'D'), 'ARCH',
             'Model element count', conf, info=True); n += 1

# underground pipework bill
code = f'{len(STRUCTS) + 1:02d}'
bill_header(code, 'Underground pipework - 45 printed lines, 201 packages (DLU-002 Rev 03)')
n = 1
groups = lines.groupby(['Printed material', 'Printed size'])['3D length (m)'].sum().reset_index()
groups['od'] = groups.apply(lambda g: max(od[l] for l in lines[(lines['Printed material'] == g['Printed material']) & (lines['Printed size'] == g['Printed size'])].Line), axis=1)
groups = groups.sort_values(['Printed material', 'od'], ascending=[True, False])
for g in groups.itertuples():
    mat, size = g._1, str(g._2)
    item(code, n, f'Supply, lay and joint {mat} pipe {size}', 'm',
         f'=SUMIFS({PR(PC["3D length (m)"])},{PR(PC["Material"])},"{mat}",{PR(PC["Printed size"])},"{size}")', 'PIPE', 'Package workbook 3D length', 'B'); n += 1
for mat in ('GRP', 'HDPE'):
    for k in fit_cols:
        q = sum(fit_line[l][k] for l in lines[lines['Printed material'] == mat].Line)
        if q:
            item(code, n, f'{mat} fittings - {k.lower()} (all sizes)', 'no.', f'=SUMIFS({PR(PC[k])},{PR(PC["Material"])},"{mat}")', 'FITTING',
                 'Package workbook fittings per package', 'B'); n += 1
item(code, n, 'Connection of pipe to structure wall (puddle flange / grouting)', 'no.', f'=SUM({PR(PC["Tie-in / connection ends"])})', 'CONNECT',
     'Tie-in packages + 2 per short connection', 'C'); n += 1
item(code, n, 'Excavation in sheet-piled trench, below GWL', 'm3', f'=SUM({PR(PC["Excavation shored (m3)"])})', 'EXC', 'Sequence workbook (MAXOC 2.5 m rule)', 'C'); n += 1
item(code, n, 'Excavation in dewatered open cut', 'm3', f'=SUM({PR(PC["Excavation open cut (m3)"])})', 'EXC', 'Sequence workbook', 'C'); n += 1
item(code, n, 'Temporary steel sheet piles: supply / hire, install and extract (incl. walings and struts)', 'm2', f'=SUM({PR(PC["Sheet pile area (m2)"])})',
     'SHEET', 'Sequence workbook: wall length x depth x SHEET_RATIO', 'C'); n += 1
item(code, n, 'Dewatering: install wellpoints / deep wells along trench', 'm', f'=SUM({PR(PC["Dewatered length (m)"])})', 'DEWATER', 'Sequence workbook', 'C'); n += 1
item(code, n, 'Dewatering: operate and maintain (package-days of pumping)', 'pkg-day', f'=SUM({PR(PC["Dewatering package-days"])})', 'DEWATER',
     'Sequence workbook programme (just-in-time dewatering)', 'C'); n += 1
item(code, n, 'Piles 600 mm diameter under pipe slabs - PLACEHOLDER layout (rows at 3 m)', 'no.', f'=SUM({PR(PC["Pipe-support piles (no.)"])})', 'PILE',
     'Sequence workbook placeholder; no pile design', 'D'); n += 1
item(code, n, 'Reinforced concrete piled slab and blinding under pipes (0.50 m)', 'm3', f'=SUM({PR(PC["Piled slab concrete (m3)"])})', 'RC_BASE',
     'Sequence workbook', 'C'); n += 1
item(code, n, 'Reinforcement to piled slab under pipes', 't', f'=SUM({PR(PC["Piled slab concrete (m3)"])})*RR_PIPESLAB/1000', 'REBAR', 'Slab volume x RR_PIPESLAB', 'D'); n += 1
item(code, n, 'Granular bedding and surround to 0.3 m above crown', 'm3', f'=SUM({PR(PC["Granular surround (m3)"])})', 'SURROUND', 'Pipes sheet formula', 'C'); n += 1
item(code, n, 'Backfill to NG in compacted layers', 'm3', f'=SUM({PR(PC["Backfill (m3)"])})', 'BACKFILL', 'Sequence workbook', 'C'); n += 1
item(code, n, 'Anchor straps to piled slab where empty pipe can float at final level', 'm', f'=SUM({PR(PC["Pipe length needing anchor straps (m)"])})', 'ANCHOR',
     'Sequence workbook flotation check (UG-08, UG-09, UG-10)', 'C'); n += 1
item(code, n, 'Locate and protect crossing services (trial pits) at listed crossings', 'no.', f'=SUM({PR(PC["Crossings to locate (no. packages)"])})', 'XING',
     'Packages with crossings listed', 'C'); n += 1

# provisional / site-wide
code = f'{len(STRUCTS) + 2:02d}'
bill_header(code, 'Site-wide and provisional items')
n = 1
unas = [o for o in objs if any('Unresolved' in c for c in o['cols'])]
unas_len = sum(o['vol'] / (math.pi * float(o['props'].get('diameter')) ** 2 / 4) for o in unas
               if o['props'].get('tag') == 'UNASSIGNED' and float(o['props'].get('diameter') or 0) > 0)
item(code, n, 'Untagged model pipes without levels (90 objects, of which untagged length shown) - provisional, line identity to be resolved', 'm',
     round(unas_len, 1), 'PIPE', 'Model volume / pipe area; tagged unresolved pieces (about 335 m) are assumed already in the line lengths', 'D'); n += 1
item(code, n, 'Site raising from NG (-0.70) to FL (+1.00) over the developed area, less structure footprints - PROVISIONAL', 'm3',
     f'=(SITE_AREA-SITE_DEDUCT*SUM({XR("E")}))*(FL-NG)', 'SITEFILL', 'Convex hull of model; replace with surveyed site area and levels', 'D'); n += 1
BN = row - 1

# ----------------------------------------------------------------- Summary
wsS['A1'] = 'Alex West WWTP - estimate BOQ summary'; wsS['A1'].font = title
wsS['A2'] = 'Amounts fill in as rates are entered on the BOQ sheet. Key quantities are totals of the BOQ by code.'; wsS['A2'].font = sub
header(wsS, 4, ['Bill', 'Description', 'Amount', 'Share'], [7, 70, 18, 9])
for i, (bc, bt, _) in enumerate(bills):
    rr = 5 + i
    put(wsS, rr, 1, bc, bold); put(wsS, rr, 2, bt)
    put(wsS, rr, 3, f'=SUMIFS(BOQ!$G${B0}:$G${BN},BOQ!$A${B0}:$A${BN},A{rr})', fmt=NUM)
    put(wsS, rr, 4, f'=IF($C${5+len(bills)}=0,0,C{rr}/$C${5+len(bills)})', fmt='0.0%')
tr = 5 + len(bills)
put(wsS, tr, 2, 'Total (excluding preliminaries, M&E, contingency)', bold); put(wsS, tr, 3, f'=SUM(C5:C{tr-1})', bold, fmt=NUM)
kq = tr + 2
put(wsS, kq, 1, 'Key quantities', bold); kq += 1
for j, h in enumerate(['Code', 'Quantity', 'Unit', 'Description'], 1):
    c = wsS.cell(row=kq, column=j, value=h); c.font = hdr_font; c.fill = hdr_fill; c.border = box
kq += 1
KEYQ = [('RC_BASE', 'm3', 'RC rafts, bases, floors, foundations (incl. pipe slabs)'), ('RC_WALL', 'm3', 'RC walls'), ('RC_SLAB', 'm3', 'RC suspended slabs / decks / roofs'),
        ('RC_COLBEAM', 'm3', 'RC columns / beams'), ('RC_MISC', 'm3', 'RC channels, baffles, weirs, chambers'), ('RC_PLINTH', 'm3', 'RC plinths'),
        ('RC_BRIDGE', 'm3', 'RC thickener bridges'),
        ('BLINDING', 'm3', 'Blinding'), ('MASS_FILL', 'm3', 'Mass concrete fill'), ('FORMWORK', 'm2', 'Formwork'), ('REBAR', 't', 'Reinforcement'),
        ('STEEL', 't', 'Structural steel'), ('PILE', 'no.', 'Piles 600 mm (model + allowances + pipe placeholder)'), ('EXC', 'm3', 'Excavation'),
        ('BACKFILL', 'm3', 'Backfill'), ('SHEET', 'm2', 'Temporary sheet piles'), ('PIPE', 'm', 'Pipes'), ('FITTING', 'no.', 'Pipe fittings'),
        ('HANDRAIL', 'm', 'Handrails'), ('SITEFILL', 'm3', 'Site raising (provisional)')]
for c_, u, d in KEYQ:
    put(wsS, kq, 1, c_, bold); put(wsS, kq, 2, f'=SUMIFS(BOQ!$E${B0}:$E${BN},BOQ!$H${B0}:$H${BN},A{kq})', fmt=NUM); put(wsS, kq, 3, u); put(wsS, kq, 4, d)
    kq += 1
put(wsS, kq, 1, 'All RC', bold); put(wsS, kq, 2, f'=SUM(B{kq-len(KEYQ)}:B{kq-len(KEYQ)+6})', bold, fmt=NUM); put(wsS, kq, 3, 'm3'); put(wsS, kq, 4, 'All reinforced concrete')
kq += 2
put(wsS, kq, 1, 'Cross-checks', bold); kq += 1
checks = [
    ('Aeration raft: model vs aeration estimate', f'={tk("AER", "RC_BASE")}', 30012.624, 'Model RC_BASE for aeration vs 30,012.6 m3 net raft (AW-AER Tracking Documentation, 18 Oct 2026)'),
    ('Aeration piles in model vs tracking', f'=SUMIFS({XR("L")},{XR("A")},"AER")', 3672, '3,672 model piles; 2,079 recorded done, 19 defect, 1,568 pending (17 Oct 2026) - BOQ is gross'),
]
for lab, f_, ref, note in checks:
    put(wsS, kq, 1, lab, bold); wsS.merge_cells(start_row=kq, start_column=1, end_row=kq, end_column=2)
    put(wsS, kq, 3, f_, fmt=NUM); put(wsS, kq, 4, ref, blue, fmt=NUM); put(wsS, kq, 5, note); kq += 1
wsS.column_dimensions['E'].width = 90

# ----------------------------------------------------------------- Exclusions
wsE['A1'] = 'Exclusions, gaps and how to improve the estimate'; wsE['A1'].font = title
header(wsE, 3, ['Item', 'Status'], [50, 120])
EXC = [
    ('Preliminaries, site establishment, supervision, insurances, contingency', 'Not included.'),
    ('Mechanical, electrical, instrumentation and process equipment', 'Not included. Equipment envelopes in the model are listed as counts only.'),
    ('Architecture and finishes (doors, windows, cladding, finishes, MEP in buildings)', 'Not included; building walls are measured as concrete envelopes (material TBC).'),
    ('Pile lengths', 'No pile design lengths in the model (display stubs). Piles are counted, not measured in metres.'),
    ('Piles for structures without pile layouts', 'Allowance = footprint x PILE_DENS (calibrated on aeration). Replace with pile designs.'),
    ('Pipe-support piles', 'Placeholder layout (rows at 3 m, 2 per row for OD >= 1.0 m). Replace with design.'),
    ('Reinforcement', 'Ratios per element class (Inputs). Only the aeration raft ratio is evidence-based.'),
    ('Dewatering and temporary works for structures', 'Not quantified (no temporary works design); pipe dewatering and sheet piles are included.'),
    ('Existing structures and boundaries not in the model', 'Existing PST, bypass / outfall chambers, treated effluent pump station, inlet / outlet boundaries, primary and thickened sludge pump stations, sludge flowmeter chambers: not measured.'),
    ('Structures modelled from tender / provisional envelopes', 'Digesters, gas holders, sludge buildings, pump stations, CHP, THP, final dewatering: Confidence D.'),
    ('Typical structures', 'FST A-L and thickeners A-D are typical structures placed 12 and 4 times; tank-specific differences not captured.'),
    ('Work done', 'Gross quantities; no deduction for work executed (e.g. 2,079 aeration piles recorded done).'),
    ('Common vertical datum', 'Not verified (R001); depths and excavation depend on it.'),
    ('Site raising', 'Provisional: developed area is the convex hull of the model; levels NG -0.70 / FL +1.00 are single site values.'),
    ('Roads, drainage, landscaping, fencing, utilities, cable ducts', 'Not in the model - not included.'),
]
for i, (a, b) in enumerate(EXC):
    put(wsE, 4 + i, 1, a, bold, align=wrap); put(wsE, 4 + i, 2, b, align=wrap)

# ----------------------------------------------------------------- Read Me
wsR['A1'] = 'Alex West WWTP - estimate bill of quantities'; wsR['A1'].font = title
wsR['A2'] = 'Prepared 22 Oct 2026 by Claude for Mohamed. Quantity estimate from the available model and workbooks; not a priced or tender BOQ.'; wsR['A2'].font = sub
txt = [
    ('Sources', None),
    (None, '1. 3D coordination model Alex-West-Combined-Coordination.blend (19 Oct 2026): 15,195 closed meshes for 23 structure types (34 structures incl. typical copies); element volumes, face areas, levels, pile stubs.'),
    (None, '2. Pipe execution packages v3 (21 Oct 2026): 45 lines, lengths, sizes, fittings, crossings.'),
    (None, '3. Pipe sequence and methodology workbook (22 Oct 2026): GWL -1.50, sheet piles, dewatering, piled slabs, flotation.'),
    ('Method', None),
    (None, 'Every model element is classified by its name and category into a BOQ class (blinding, RC base / wall / slab / column-beam / misc / plinth, mass fill, steel, grating, handrails, stairs, piles, equipment, water). Water surfaces are excluded.'),
    (None, 'Concrete = closed-solid volume. Formwork = vertical faces + soffits of elevated slabs/beams/secondary elements. Reinforcement = volume x ratio. Excavation = offset prism from footprint below NG.'),
    ('Sheets', None),
    (None, 'Summary - totals per bill and key quantities; BOQ - priced bill (enter rates); Inputs - assumptions; Take-off - model data per structure and class; '
           'Structures - footprints, excavation, backfill, piles; Pipes - per-line pipe quantities; Exclusions - what is not covered.'),
    ('Confidence', None),
    (None, 'B = structural model or detailed design source (aeration Rev02, FST DD-004, pipe lines). C = civil-guide or provisional envelope. D = tender envelope, placeholder or allowance.'),
    ('Colour key', None),
    (None, 'Blue = source data; black = formula; yellow = enter / edit.'),
]
rr = 4
for h, t in txt:
    if h:
        wsR.cell(row=rr, column=1, value=h).font = bold
    else:
        c = wsR.cell(row=rr, column=1, value=t); c.font = fnt; c.alignment = Alignment(wrap_text=True, vertical='top')
    rr += 1
wsR.column_dimensions['A'].width = 160
for ws in wb.worksheets:
    ws.sheet_view.zoomScale = 90
wsB.freeze_panes = 'D5'
wb.save(OUT)
print('saved', OUT, 'BOQ rows', BN - B0 + 1, 'unclassified check:', el.cls.value_counts().to_dict())
el.to_csv(OUT.replace('.xlsx', '-elements.csv'), index=False)
