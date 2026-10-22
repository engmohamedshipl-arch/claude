# 3D model data

Source: `Alex-West-Combined-Coordination.blend`, 19 Oct 2026 (164 MB, uploaded 22 Oct 2026). Read with the Blender Python module (bpy 5.2.2) with scripts disabled. The web viewer link (n8n.a2zegypt.org) was not reachable from the cloud session, so the .blend file was used instead.

## Contents

- 15,195 mesh objects; every mesh is a closed solid, so volumes are geometrically reliable.
- Structures:
  - aeration tanks (structural Rev02)
  - 12 final sedimentation tanks (FST A–L: one typical structure placed 12 times, plus provisional attachments)
  - 4 primary sludge thickeners
  - primary sedimentation tanks (8 basins)
  - grit removal
  - IPS, plus the motor hall shared with RAS/WAS
  - RAS/WAS pumping station
  - 3 secondary distribution chambers
  - chlorine contact tank
  - fire / services
  - air blower building
  - 8 digesters, 4 gas holder plinths, gas holder blower house
  - 2 digested sludge buildings
  - grease and scum pump stations
  - WAS pre-dewatering, THP platform, final dewatering, CHP and chlorination storage buildings, air blowers power house
- Pipes: 770 objects (water line 532, sludge 107, bypass/outfall 41, unresolved elevation 90).
- Embedded texts: plant model manifest (asset stages and open gaps), site basis, source register, aeration tracking baseline, agreed timeline and methodology.
- Coordinates: model frame + (774450, 3450600) = site E/N. Checked against the pipe package workbook: 31 aeration piles within 15 m of UG-02-S03, 18 pending, nearest 11.7 m — an exact match.

## Aeration piling status (data date 17 Oct 2026)

| Zone | Done | Pending | Defect | Unknown |
|---|---|---|---|---|
| 1 | 470 | 6 | 2 | 0 |
| 2 | 518 | 41 | 2 | 3 |
| 3 | 248 | 236 | 2 | 0 |
| 4 | 221 | 337 | 3 | 0 |
| 5 | 223 | 137 | 4 | 0 |
| 6 | 134 | 279 | 5 | 0 |
| 7 | 121 | 296 | 1 | 0 |
| 8 | 144 | 236 | 0 | 3 |
| Total | 2,079 | 1,568 | 19 | 6 |

"Done" is recorded execution, not acceptance.

**Production:** 8–9 rigs. The last 4 full weeks averaged 285 piles/week (1,140 piles over the weeks starting 14 Sep – 5 Oct), about 5.4 piles per rig per working day. At that rate the 1,587 open piles finish around 29 Nov 2026 in the agreed zone order. This forecast gates pipe packages within 15 m of open aeration piles.

## Limits

- Pile objects are truncated display stubs: no pile lengths.
- FST pile stubs are placeholders stacked at each tank centre.
- Many sludge-line buildings are tender / provisional envelopes ("Construction approval not established").
- Common vertical datum not verified.
- 90 pipe objects have no centreline level (about 460 m equivalent length, mostly tagged to lines already measured).
