# Alex West — underground pipe construction sequence and methodology

`Alex-West-Pipe-Sequence-Methodology.xlsx` is a planning workbook covering the 201 pipe packages (45 lines, 8,151 m) from the v3 package workbook. It applies three decisions made on 22 Oct 2026:

- groundwater level -1.50
- every pipe carried on piles and a piled slab
- shoring (sheet piles) where it is needed

It is not a design, a temporary-works design or a construction approval.

## Sheets

| Sheet | Content |
|---|---|
| Methodology | Step-by-step method, plant, hold points and the sequencing rules |
| Inputs | Levels, groundwater, slopes, pile layout placeholder, gates, productivity (yellow = edit) |
| Aeration | Pile status by zone from the 3D model (data date 17 Oct 2026), measured rig output, forecast completion of each zone |
| Structures | Date each structure is ready for tie-in (to be entered) |
| Fronts | 8 work fronts and their crew numbers |
| Packages | Method, quantities, dewatering, flotation check and dates for each package |
| Activities / Relationships | 2,211 activities with formula-driven dates, plus the logic as a P6 / MS Project import list |
| Programme | Monthly bar chart, activity counts, piles and rig demand |
| Sensitivity | Sheet-piled length under alternative ground rules |
| Piece check | Piece-level calculations behind Packages |

## Rebuild

```
python scripts/extract_aeration_piles.py <Alex-West-Combined-Coordination.blend> inputs/aer_piles.json   # needs the bpy module
python scripts/build.py <Alex-West-Pipe-Execution-Packages-All-Lines.xlsx> inputs/aer_piles.json inputs/aer_weekly.json inputs/manifest.json Alex-West-Pipe-Sequence-Methodology.xlsx
```

After building, recalculate the workbook in Excel or LibreOffice: openpyxl writes the formulas but no calculated values.
