# Alex West WWTP: planning package, 22 Oct 2026

**Transmittal:** issued for information and planning. Prepared by Claude for Mohamed (Projects Controls Director). Not a design, a temporary-works design or a construction approval.

Start with **`Alex-West-Planning-Package-Report.pdf`** (9 pages). It contains:

- executive summary
- document register
- basis and decisions
- groundwater and temporary works
- methodology
- programme
- aeration piling status
- estimate BOQ
- open items, exclusions and how to use the files

## Document register

| No. | Document | File | Status |
|---|---|---|---|
| PKG-00 | Planning package report | `Alex-West-Planning-Package-Report.pdf` | Issued for information |
| PKG-01 | Pipe sequence and methodology workbook (live formulas) | `01-Sequence/Alex-West-Pipe-Sequence-Methodology.xlsx` | Planning draft |
| PKG-02 | Pipe works programme for Primavera P6 | `02-Programme/Alex-West-Pipe-Works.xer` | Planning draft |
| PKG-03 | Pipe works programme for MS Project | `02-Programme/Alex-West-Pipe-Works.xml` | Planning draft |
| PKG-04 | Estimate bill of quantities, whole project (unpriced) | `03-BOQ/Alex-West-Estimate-BOQ.xlsx` | Quantities only |
| PKG-05 | Project wiki: decisions log, open items, method notes | `04-Wiki/*.md` (start at `Home.md`) | Living record |
| PKG-06 | Build scripts and extracted model data | `05-Scripts-and-Data/` | For rebuild |

Inputs not included (project files):

- REF-01: pipe execution packages v3 (21 Oct 2026)
- REF-02: combined coordination model (19 Oct 2026, 164 MB)

## Status of key results

| Topic | Result | Firmness |
|---|---|---|
| Pipe programme | 1 Dec 2026 to 18 Sep 2027, 8 fronts, 2,211 activities | Best case: approval gates and structure readiness not yet entered |
| Sheet piles (GWL -1.50) | 6,562 m of 8,151 m pipe (80%), about 116,300 m² | Depends on the assumed open-cut limit (2.5 m below GWL) |
| Reinforced concrete | about 145,800 m³ | From the 3D model; aeration raft checked within 0.4% |
| Piles 600 mm | about 16,000 no. | 6,392 modelled; about 3,800 allowance; 5,788 pipe placeholder |
| Structure excavation | about 511,000 m³ | **Overstated**: refinement pending |

## Next actions

See `04-Wiki/Open-Items-and-Holds.md`.

1. Confirm GWL and the open-cut rule.
2. Pile designs for the pipes and for the structures without pile layouts.
3. Structure readiness dates.
4. DLU-002 approval forecast.
5. Productivity and unit rates.
6. Refine structure excavation.

## Using the files

- **Workbooks:** yellow cells are inputs. Enter dates, crews and rates, then recalculate (Excel recalculates on open).
- **P6:** File > Import > XER as a new project, then F9.
- **MS Project:** open the XML file.
- **After changing inputs:**
  1. Re-export the programme: `python scripts/export_schedule.py <sequence workbook> <out.xer> <out.xml>` (needs Java and `pip install mpxj jpype1 openpyxl`).
  2. Rebuild the BOQ with `scripts/build_boq.py` from `data/objects.json` and `data/hulls.json`.

## Repository

Branch `claude/new-session-e5ojdh` of `engmohamedshipl-arch/claude`. The same files are in `pipe-sequence-methodology/`, `estimate-boq/`, `wiki/` and `package/`.
