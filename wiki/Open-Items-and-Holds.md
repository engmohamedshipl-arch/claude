# Open items and holds

## Inputs needed from the project

| # | Item | Needed for | Owner |
|---|---|---|---|
| 1 | Confirm groundwater design value (-1.50), side slopes and the open-cut depth limit below GWL (assumed 2.5 m) | Sheet-pile quantity and method for every pipe package | Geotechnical / temporary works |
| 2 | Pile and slab design for pipe supports (spacing, piles per row, length) | 5,788 placeholder piles; piled-slab concrete | Structural design |
| 3 | Pile layouts for structures without them (PST, IPS, RAS/WAS, CCT, GRT, SDC, fire, buildings) | About 3,800 allowance piles in the BOQ | Structural design |
| 4 | Date each structure is ready for pipe tie-in | Tie-in packages (71) in the programme | Master programme |
| 5 | Forecast approval date of DLU-002 (Rev 03 is Issued for Approval) | Gate on every package | Design management |
| 6 | Pipe pile design approval date | Gate on all pipe piling | Design management |
| 7 | Identify unassigned crossing pipes P0114, P0117, P0279 | Excavation of UG-01-S02, UG-02-S04, UG-03-S03 | Survey / design |
| 8 | Productivity rates for excavation, slab, laying and backfill (all assumed except piling) | Durations | Planning / construction |
| 9 | Thickener bridge drawing (section sizes) | Concrete bridge quantity: the model members are slender, so 21.3 m³ is low | Structural design |
| 10 | Unit rates (currency on the BOQ Inputs sheet) | Pricing the BOQ | Cost / QS |
| 11 | Surveyed site boundary and levels | Provisional site raising (378,000 m³) | Survey |

## Open work items (Claude)

| # | Item | Status |
|---|---|---|
| A | Refine structure excavation with a 1 m grid of each structure's underside (current method uses the deepest point over the whole footprint and overstates PST, CCT and RAS/WAS) | Ready to run; stopped on 22 Oct 2026 pending go-ahead |
| B | Re-export the P6 / MS Project files after inputs change | On request (`scripts/export_schedule.py`) |
| C | Pipe works resource levelling to the rigs actually available (peak demand about 13 rigs in Dec 2026) | Not started |

## Holds that apply to every pipe package (from the package workbook)

- Common vertical datum not verified (R001).
- Levels are envelope (CL ± outside diameter), not internal invert (R002).
- Levels between drawing anchors are interpolated (R003).
- The 0.50 m piled-slab basis is a stated assumption.
- Structures not in the master model are not checked: existing PST, bypass and outfall chambers, treated effluent pump station.

## Site coordination items in the programme

Six shared-trench backfill holds can't be linked without creating a loop in the logic. They are marked "COORDINATE ON SITE" in the sequence workbook's Packages sheet (Interactions column): do not backfill beside the listed pipe until it is laid.
