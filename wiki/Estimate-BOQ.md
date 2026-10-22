# Estimate BOQ: whole project (22 Oct 2026)

Workbook: `estimate-boq/Alex-West-Estimate-BOQ.xlsx`. Quantities only: no price data was available, so the rate column is blank (currency on Inputs) and amounts fill in once rates are entered.

## Method

- **Classification:** every element of the 3D model is classified by name and category into a BOQ class:
  - blinding and mass fill
  - RC bases / walls / slabs / columns-beams / secondary elements / plinths / thickener bridge
  - steel deck, handrails, stairs
  - piles, equipment
  - water surfaces (excluded)
- **Concrete:** closed-solid volume.
- **Formwork:** vertical faces plus soffits of elevated elements.
- **Reinforcement:** volume x ratio, from Inputs.
- **Excavation:** footprint below NG with working space and slope.
- **Piles:** model counts where modelled; otherwise an allowance of 0.10 piles/m².
- **Pipework:** from the package and sequence workbooks.
- **Bills:** 23 structure bills, one underground pipework bill, one site-wide / provisional bill.
- **Confidence grade on each line:**
  - B = detailed model
  - C = civil-guide envelope
  - D = tender envelope or allowance

## Headline quantities (gross, no deduction for work done)

| Item | Quantity |
|---|---|
| Reinforced concrete | about 145,800 m³ (rafts and bases 88,700; walls 41,300; slabs 5,600; secondary elements 7,900; plinths 1,800; columns and beams 430; thickener bridges 21) |
| Reinforcement | about 17,300 t |
| Formwork | about 221,000 m² |
| Blinding | about 6,100 m³ |
| Piles 600 mm | about 16,000 no.: model 6,392 (aeration, thickeners, FST stubs), allowance about 3,800, pipe placeholder 5,788 |
| Excavation | about 511,000 m³ (**overstated**, see below) |
| Backfill | about 185,000 m³ |
| Temporary sheet piles (pipes) | about 116,300 m² |
| Pipes | 8,294 m, plus 474 fittings |
| Handrails | about 5,400 m |
| Site raising NG -0.70 to FL +1.00 | about 378,000 m³, provisional (developed area from the model hull, 31 ha) |

## Cross-checks

- Aeration raft: 30,120 m³ from the model against 30,013 m³ in the aeration quantity file (0.4%).
- Aeration piles: 3,672 in the model, the same as the tracking sheet.

## Weak points

- **Structure excavation is overstated.** It uses the deepest point over each footprint; for example, the PST inlet pit at -7.6 drives a raft at -3.2. A grid-based refinement is ready (open item A).
- **Thickener bridge:** reinforced concrete per the 22 Oct 2026 decision. The model has slender truss members (5.3 m³ per bridge), so the quantity is low until the bridge drawing is available.
- **Reinforcement:** only the raft ratio (107 kg/m³) is evidence-based.
- **Piles:** no lengths (display stubs only).
- **Not included:** M&E, architecture, roads and drainage, preliminaries, existing structures not in the model.
