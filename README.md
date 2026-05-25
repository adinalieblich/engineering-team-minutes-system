# Engineering Team Meeting Minutes System

A meeting minutes and action/risk tracking system built in Excel + VBA for a local government engineering team. Solo build, AI-assisted (Claude/Anthropic), deployed in a locked-down corporate environment with SharePoint only — no plugins, no admin rights, no external tools.

Currently mid-migration to Notion. This repo documents the Excel build and the AI-assisted development methodology that came out of it.

---

## What's in here

| Folder | Contents |
|---|---|
| `docs/` | Project documentation — Living State file, How To Use guide, error catalogue, style guide |
| `src/vba/` | All VBA modules — macros, builders, helpers |
| `src/python/` | Python verification scripts (openpyxl) for comparing files against source of truth |
| `assets/` | Screenshots, palette references |
| `examples/` | Sanitised example workbook |
| `CLAUDE.md` | Instructions for Claude Code (or any AI assistant) picking up this project |

---

## The most interesting file

**[`docs/What-AI-Wont-Tell-You.html`](docs/What-AI-Wont-Tell-You.html)** — a practical field guide to AI-assisted Excel/VBA development. 23 confirmed errors with root causes and working fixes, plus 10 proven strategies. Written to brief the next AI session so the same mistakes don't repeat.

This is the meta-deliverable. The system is the artifact; the field guide is the methodology.

---

## Architecture (final state, "Option 3")

Single master `.xlsm` workbook, SharePoint-synced. Five sheets:

1. **Summary** — action counts by owner and status
2. **Meeting Minutes** — main working sheet, one row per agenda item, structured by section + project banners
3. **Risk Register** — auto-populated from rows tagged Risk
4. **How To Use** — embedded cheat sheet
5. **Changelog** — locked amendment record

### Key architectural decisions

- **Action Register is a filter view, not a separate sheet.** A macro filters Meeting Minutes where Type=Action. This killed an unreliable two-way sync that was breaking weekly.
- **Type is a dropdown, not checkboxes.** Form-control checkboxes drift, don't anchor properly, and accumulate as ghost shapes. Replaced 257 of them with a single dropdown column.
- **Project lifecycle = row position.** No Project Register sheet. Banner rows define structure. Manual cut/paste between lifecycle sections.
- **Conditional formatting uses `xlExpression` + `$Col` absolute references on contiguous sqrefs.** Fragmented sqrefs silently break new-row coverage — discovered the hard way (see Error 018).
- **Banners manually managed.** Auto-rebuild was too destructive. Removed.

### Column structure (Meeting Minutes)

| Col | Field | Notes |
|---|---|---|
| A | Type | Dropdown: Action / Risk / blank |
| B | Description | |
| C | Owner | Dropdown |
| D | Due | DD/MM/YYYY |
| E | Status | Dropdown: Open / In Progress / Done / On Hold / Waiting |
| F | Priority | Dropdown: Critical / High / Medium / Low |
| G | Days OD | Formula |
| H | Notes / Update | |
| I | Project | Hidden, used for AR filter |
| J | Date Added | Hidden, auto-stamped on row creation |

---

## How it was built — methodology

This is the part that matters more than the artifact.

### Mockup-first development
Every major design decision (colour theme, priority palette, AR formatting, Risk Register structure) was rendered as an interactive HTML mockup BEFORE any code was written. Stakeholder sign-off on the mockup. Build once. Documented in [`docs/What-AI-Wont-Tell-You.html`](docs/What-AI-Wont-Tell-You.html) — Strategy 4.

### File-upload-first diagnosis
AI's default pattern is to bucket issues into a plausible plan based on project docs. Roughly half the time that plan is wrong because the actual file state has drifted. Rule developed: never accept a plan that wasn't built from inspection of the current file. See Problem 7.

### Root-cause demand
When something failed twice, refused another patch attempt. This is how `bgColor` vs `fgColor` (the cardinal openpyxl CF bug), triple-quote escaping in CF formulas, and CF sqref fragmentation were all found. See Strategy 6.

### Living State file
Every session ends with a state archive. Every session starts by re-uploading it. The conversation IS the project memory. See [`docs/Living-State.html`](docs/Living-State.html).

### Pre-flight checks
Before running any macro: Sub/End Sub balance, With/End With balance, Dim placement, `xlSolid` before `.Interior.Color` on FormatConditions, `.HasTextFrame` guards. See Strategy 2.

### Never ship AI-built `.xlsx` files
openpyxl writes CF fills as `bgColor` which Excel renders invisible. Rule: AI writes VBA macros, the user runs them inside Excel. Files Excel touches stay valid; files Python touches break in invisible ways. See Error 001.

---

## Tech stack

- **Excel + VBA** — live system
- **Python + openpyxl** — verification and analysis only, never file writes that ship
- **Claude (Anthropic)** — build partner
- **HTML/CSS** — mockups, documentation, embedded help
- **SharePoint** — deployment

---

## Project status

Excel build: ~95% complete. In Phase 18 (data migration from old template) at time of last commit. Phases 19–21 pending (printable AR view, team rollout, cosmetic polish).

Currently in parallel migration to Notion — see `docs/Notion-Migration-Notes.md`.

---

## License

MIT — see [`LICENSE`](LICENSE).

---

## A note on AI-assisted development

This project was built with AI assistance from start to finish, but the skill being demonstrated is not "I used AI." It's the discipline of driving AI to produce working corporate-grade output: when to push back, when to demand diagnosis instead of patching, when to mockup before building, and when to refuse confident-sounding plans that weren't grounded in the actual file state.

The errors documented here are not hypothetical. Every one cost real time in a live environment.
