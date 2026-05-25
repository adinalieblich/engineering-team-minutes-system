# Style Guide

The locked-in visual system. Use these exact values everywhere. Do not substitute, do not approximate.

## Priority palette

| Value | Background | Hex | Font | Font hex |
|---|---|---|---|---|
| Critical | Red | `#C00000` | White | `#FFFFFF` |
| High | Orange | `#ED7D31` | White | `#FFFFFF` |
| Medium | Yellow | `#FFC000` | Amber dark | `#7D5200` |
| Low | Sage green | `#A9D18E` | Forest green | `#375623` |

## Status palette

| Value | Background | Hex | Font | Font hex |
|---|---|---|---|---|
| Open | Pale blue | `#EEF4FA` | Medium blue | `#5C7FA1` |
| In Progress | Pale yellow | `#FFF2CC` | Medium amber | `#9C7700` |
| Done | Pale green | `#E5F1D9` | Medium green | `#5F8F37` |
| On Hold | Pale pink | `#FCE5EA` | Medium pink | `#A85871` |
| Waiting | Pale lavender | `#F1E8F7` | Medium purple | `#7E5BA4` |

## Row-type palette

| Row type | Background | Hex | Font | Font hex |
|---|---|---|---|---|
| Action | Pale navy | `#E1E9F3` | Dark navy | `#1F3864` |
| Action (overdue) | Pale red | `#FBE0E0` | Dark red | `#B01212` |
| Risk | Pale orange | `#FFF1DD` | Dark orange | `#B05E12` |
| Note | (no fill) | — | (default) | `#000000` |

## Banner colours

- Section header: `#1F3864` (dark navy), white text
- Project sub-header: `#2E75B6` (mid blue), white text
- Sub-section: `#5B9BD5` (lighter blue), white text

## Days Overdue gradient

| Days | Fill | Font |
|---|---|---|
| 1–7 | `#ED7D31` (orange) | White bold |
| 8+ | `#C00000` (red) | White bold |

## Risk Register CF

| Col | Value | Fill | Font |
|---|---|---|---|
| E/F | High | `#C00000` | `#FFFFFF` |
| E/F | Medium | `#FFC000` | `#7D5200` |
| E/F | Low | `#A9D18E` | `#375623` |
| I | Open | `#F2F2F2` | `#595959` |
| I | Monitored | `#FFF2CC` | `#7D5200` |
| I | Closed | `#E2EFDA` | `#375623` |

## Alternative theme (generic / commercial)

For non-government deployment, swap navy/blue for sage/eucalyptus:

| Role | Hex |
|---|---|
| Section header (deep) | `#505B40` |
| Sub-header (mid) | `#6B7A58` |
| Project header (soft) | `#9BA187` |
| Row ivory | `#FDFBF7` |

Keep priority/status/row-type palettes the same — they're palette-agnostic.

## Typography

- Body: Arial 10pt
- Banner text: Arial 10pt bold, white
- Header block: Arial 9pt, navy text
- Code/macros in docs: Courier New 9pt
