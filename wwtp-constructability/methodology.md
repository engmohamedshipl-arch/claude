# Area Estimation Methodology & Accuracy Notes

How footprint areas in `building-inventory.csv` are derived, and what limits accuracy.

## How areas are estimated from imagery

1. **Establish scale.** Use the image's scale bar (e.g. "100 m") or a structure
   of known dimension. This gives a metres-per-pixel ratio.
2. **Bound each structure.** For rectangular buildings/tanks: measure length × width.
   For circular tanks (clarifiers, digesters): measure diameter, area = π·(d/2)².
   For irregular shapes: decompose into rectangles/triangles and sum.
3. **Record footprint** (plan area at ground level). Height/stories are noted
   separately and are only a rough guess from shadows/context unless you supply them.

## Accuracy & caveats (important for constructability)

- **Imagery is approximate.** Expect roughly **±5–15%** on linear dimensions from
  a clean, well-scaled satellite image — worse for small or partially obscured
  structures, blurry images, or anything near the image edge (lens distortion).
- **Satellite imagery can be dated.** The map may not reflect recent demolition,
  new construction, or temporary structures. Always confirm against a recent visit.
- **No vertical data.** Areas are plan footprints. Heights, floor counts, and
  below-grade structures (tanks, galleries, basements) are NOT visible from above.
- **Roof ≠ footprint** for buildings with large eaves/canopies; tank water surface
  may read larger than the structural footprint.
- **Not a survey.** These figures support early constructability/planning only and
  must not be used for procurement, boundary, or contract quantities without a
  proper topographic/dimensional survey.

## Data that would sharpen the analysis (if you have it)

- Recent topographic / boundary survey, or site CAD/GIS (DWG, SHP, KML).
- Site plot plan or existing as-builts.
- Known dimension of any one structure (best single calibration aid).
- Recent ground or drone photos (condition, height, access).
- Geotech/borehole data, utility routes, and the new-plant layout footprint.

## Constructability lenses applied

For each structure and for the site overall, the report considers:

- **Access** — gate locations, internal road widths, turning, crane/heavy-haul routes.
- **Laydown / staging** — open yard area available for materials and plant.
- **Demolition / clearing** — what must come down, sequencing, live-plant constraints.
- **Live operations** — the plant likely stays in service during works; phasing matters.
- **Constraints** — wet/process areas, odour/H₂S zones, overhead/buried services,
  watercourses, neighbours, and site boundaries.
