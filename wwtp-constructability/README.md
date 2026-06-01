# Alexandria West WWTP — Constructability Site Analysis

Working folder for the constructability / site-analysis package for the
**Alexandria West Sewage Treatment Plant** project.

- **Site reference (shared link):** https://maps.app.goo.gl/ktDdK5Ao8yeuxDZ4A
- **Purpose:** Inventory existing structures, estimate footprint areas, and
  capture site context (access, laydown, constraints) to feed constructability
  planning.

## ⚠️ Status: awaiting imagery

This environment has **no live Google Maps access** (network egress is blocked,
and no Maps toolkit is connected). Building areas therefore cannot be auto-pulled
from the map. The analysis is driven from a **scaled satellite/aerial image** that
you provide — either uploaded in chat or placed in Google Drive.

**For accurate areas, the image must show the scale bar** (e.g. the "100 m"
ruler) or at least one structure of known real-world dimension for calibration.

## Contents

| File | Purpose |
|------|---------|
| `building-inventory.csv` | Structured inventory: one row per structure, with area, dimensions, type, condition, constructability notes. |
| `constructability-report.md` | Written analysis: site overview, structure summary, access, staging, demolition/clearing, constraints, risks. |
| `methodology.md` | How areas are estimated from imagery, accuracy caveats, and the data we still need. |

## Workflow

1. You provide a scaled satellite image (with scale bar).
2. Each distinct structure is labelled (Building A, B, C …) against the image.
3. Footprint areas + approximate dimensions are estimated and logged in the CSV.
4. The report and annotated findings are written from the inventory.
5. You verify against any survey/CAD data you hold and we refine.
