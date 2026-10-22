# Pipe sequence and methodology (22 Oct 2026)

Workbook: `pipe-sequence-methodology/Alex-West-Pipe-Sequence-Methodology.xlsx`. It is formula-driven: change Inputs, Fronts, Structures or Aeration and everything recalculates.

## Basis

- The 201 packages from the [v3 package workbook](Pipe-Execution-Packages.md).
- Groundwater -1.50.
- All pipes on piles and a 0.50 m piled slab.
- Sheet piles where open cut does not work.
- The 3D model for aeration pile status and production (see [3D model data](3D-Model-Data.md)).

## Effect of groundwater at -1.50

- **Below groundwater:** 199 of 201 packages excavate below GWL, by up to 6.1 m (bottoms down to -7.6).
- **Sheet piling:** open cut is allowed only where it fits beside structures and piles **and** the excavation is no more than 2.5 m below GWL (assumption). Sheet-piled length rises from 1,972 m (24%) to **6,562 m (80%)**, about 116,300 m² of sheets.

| Scenario | Sheet-piled length | Share |
|---|---|---|
| Package workbook basis (no groundwater) | 1,972 m | 24% |
| Flatter slope below GWL only | 2,635 m | 32% |
| Depth limit 3.0 m below GWL | 4,543 m | 56% |
| **Depth limit 2.5 m below GWL (default)** | **6,562 m** | **80%** |
| Depth limit 2.0 m below GWL | 6,913 m | 85% |
| Everything below GWL sheet-piled | 8,093 m | 99% |

- **Dewatering:** single-stage wellpoints on 160 packages, deep wells or two-stage wellpoints on 39, not required on 2.
- **Flotation (empty pipe, pipe weight ignored):**
  - 14 packages on UG-08, UG-09 and UG-10 (DN2600, shallow cover) can float even at final level, so they need permanent straps to the piled slab.
  - 72 packages need dewatering kept running until fill reaches FL, or strapping.

## Method per package (11 activities)

1. Release: drawings, holds, structure readiness and aeration gates.
2. Set-out, crossing survey and working platform.
3. Pipe-support piling from the platform (600 mm assumed), before excavation.
4. Sheet piles on shored lengths, after piling.
5. Dewatering: install and draw down at least 7 days.
6. Excavate and strut to formation.
7. Trim piles, blinding, piled slab, cure.
8. Lay and joint pipe, from the downstream end; tie-ins last as the closure.
9. Surround and test.
10. Backfill to NG, removing struts progressively.
11. Dewatering off (after the flotation check), extract sheets, reinstate.

## Sequencing rules

- **Fronts:** eight, by system. A shared-trench corridor that crosses systems is built by one front (30 packages moved).
- **Order within a front:** deeper lines first and downstream first; shared-trench packages back to back along the corridor.
- **Deeper first:** packages within 8 m of each other with bottoms more than 1 m apart are linked deeper-first.
- **Shared trench:** where bottoms are within 1 m, excavate together and backfill only after all pipes are laid. 6 holds are site coordination items.
- **Set-out and dewatering timing:** set-out runs one package ahead of piling; dewatering is timed so drawdown completes when the excavation crew is free (just-in-time; median 74 pumping days per package).
- **Tie-ins:** follow the adjacent run, wait for structure readiness, and close the line.
- **Aeration:** packages within 15 m of open aeration piles wait for that zone (agreed zone order, forecast from actual production).
- **Gates:** pipe piling waits for pile design approval; nothing starts before DLU-002 approval.

## Fronts

| Front | Scope | Length | Start | Finish |
|---|---|---|---|---|
| F1 | Headworks and primary (UG-01 to UG-04) | 800 m | 1 Dec 2026 | 24 Jun 2027 |
| F2 | Bypass corridor (UG-05, UG-06, UG-24, UG-25, UG-26), 2 crews | 1,724 m | 1 Dec 2026 | 22 Jul 2027 |
| F3 | Distribution and FCT risers | 1,095 m | 1 Dec 2026 | 26 Jul 2027 |
| F4 | FCT effluent A-H to CCT | 551 m | 1 Dec 2026 | 22 Aug 2027 |
| F5 | FCT effluent I-L to CCT | 571 m | 1 Dec 2026 | 18 Sep 2027 |
| F6 | Sludge A-D, G, H to RAS/WAS + UG-23 | 1,164 m | 1 Dec 2026 | 8 Aug 2027 |
| F7 | Sludge I-L, E, F to RAS/WAS | 868 m | 1 Dec 2026 | 7 Aug 2027 |
| F8 | Effluent, outfall and CCT links, 2 crews | 1,378 m | 1 Dec 2026 | 4 Jul 2027 |

**Overall:** 1 Dec 2026 to 18 Sep 2027 (9.6 months). The approval gates are set equal to the start date and structure readiness is blank, so these dates are a logic-driven best case, not a forecast.

## Resources

- **Pipe-support piles:** about 5,800 on the placeholder layout. Demand peaks at about 13 equivalent rigs in Dec 2026, against 10 front rigs; the excess is tie-in piling, which no front rig covers while structure dates are blank.
- **Rig hand-over:** the 8 aeration rigs are forecast free around end Nov 2026, the natural hand-over to the pipe fronts.

## Checks

- 57,804 formulas, no errors.
- All 5,467 relationships hold on the calculated dates, and no crew works two packages at once.
- Total length reconciles to the package workbook.
