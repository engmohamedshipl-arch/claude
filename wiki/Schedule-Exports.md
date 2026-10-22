# Schedule exports: P6 and MS Project (22 Oct 2026)

Files in `pipe-sequence-methodology/`. Built by `scripts/export_schedule.py` from the recalculated sequence workbook using MPXJ 16.10.

| File | Open with | Content |
|---|---|---|
| `Alex-West-Pipe-Works.xer` | P6: File > Import > XER as a new project, then F9 | 2,163 activities, 5,395 relationships (SS + FF ladders kept) |
| `Alex-West-Pipe-Works.xml` | MS Project: File > Open, choose XML | 2,163 activities, 3,557 links |

Both files carry:

- WBS: Project > Front > Line > Package
- calendar Sat-Thu, 8 h/day, Friday off
- start-on-or-after constraints for the gates
- notes plus user-defined fields: front, line, package, activity code and quantity
- data date 1 Dec 2026

## Conversion rules

- **Lags:** converted from calendar days to working days (x 6/7). The just-in-time dewatering link is FS with negative lag equal to the dewatering duration.
- **Zero-duration placeholders** (48, e.g. sheet piling on open-cut packages) are removed; their logic is carried through.
- **MS Project:** allows one link per pair of tasks, so where the workbook has SS + FF, the link that drives the workbook dates is kept.

## Checks

- Both files were read back with MPXJ.
- A forward pass on the exported logic schedules all 2,163 activities with no loops and finishes 13 Sep 2027, against 18 Sep 2027 in the workbook (working-day rounding of lags).

## Keeping in step

The files are a snapshot. After changing workbook inputs (approval dates, structure readiness, crews, rates), recalculate the workbook and re-run the export script.
