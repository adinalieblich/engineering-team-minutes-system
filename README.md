# Engineering Team Meeting Minutes System

[![Download v1.0 kit](https://img.shields.io/badge/download-v1.0_kit-A752AA?style=for-the-badge)](https://github.com/adinalieblich/engineering-team-minutes-system/releases)
[![Field guide PDF](https://img.shields.io/badge/field_guide-PDF_40pp-FF692A?style=for-the-badge)](https://github.com/adinalieblich/engineering-team-minutes-system/raw/main/docs/What-AI-Wont-Tell-You.pdf)
[![Docs site](https://img.shields.io/badge/docs-adinalieblich.github.io-2E1840?style=for-the-badge)](https://adinalieblich.github.io/engineering-team-minutes-system/)

> I'm a civil engineer working in AI enablement for AEC. This is a
> meeting-minutes and action/risk tracking system I built solo in
> Excel + VBA for a local government engineering team — deployed in
> a locked-down SharePoint-only environment with no plugins, no admin
> rights, no external tools. It runs the weekly meeting. The repo
> ships the build, a sanitised example, and a field guide of 23
> confirmed errors and 10 proven strategies from driving AI to produce
> corporate-grade VBA. If you work in AEC and want to run this at
> your team, the v1.0 kit is one download above.

By [Adina Lieblich](https://adinalieblich.com) · Perth, Australia · MIT licensed

---

## What's in here

| Folder | Contents |
|---|---|
| `docs/` | Documentation, **field guide PDF**, How To Use guide, architecture decisions, style guide, living state |
| `src/vba/` | All VBA modules — macros, builders, helpers (the actual build) |
| `src/python/` | Python verification scripts (openpyxl) — analysis only, never ship |
| `examples/` | Sanitised example workbook scaffolding |
| `assets/` | Screenshots, palette references |
| `CLAUDE.md` | Project instructions for AI assistants picking this up |

## Downloads

| Asset | Format | Size |
|---|---|---|
| [Field guide — What AI Won't Tell You](https://github.com/adinalieblich/engineering-team-minutes-system/raw/main/docs/What-AI-Wont-Tell-You.pdf) | PDF · 40pp | 1.4 MB |
| [How To Use the workbook](https://github.com/adinalieblich/engineering-team-minutes-system/raw/main/docs/How-To-Use.pdf) | PDF | 330 KB |
| [Macro Archive](https://github.com/adinalieblich/engineering-team-minutes-system/raw/main/docs/Macro-Archive.pdf) | PDF | 360 KB |
| [v1.0 kit (zip)](https://github.com/adinalieblich/engineering-team-minutes-system/releases/latest) | zip | ~3 MB |

The kit zip contains everything you need to run this at your own team: VBA modules, How To Use, Style Guide, field guide PDF, MIT license.

---

## The most interesting file

**[`docs/What-AI-Wont-Tell-You.html`](docs/What-AI-Wont-Tell-You.html)** ([PDF](docs/What-AI-Wont-Tell-You.pdf)) — a practical field guide to AI-assisted Excel/VBA development.

23 confirmed errors with root causes and working fixes. 10 strategies that worked. 9 VBA patterns. 8 AI failure modes. 7 design decisions. Written to brief the next AI session so the same mistakes don't repeat.

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
- **Type is a dropdown, not checkboxes.** Form-control checkboxes drift, don't anchor, and accumulate as ghost shapes. Replaced 257 of them with a single dropdown column.
- **Project lifecycle = row position.** No Project Register sheet. Banner rows define structure. Manual cut/paste between lifecycle sections.
- **Conditional formatting uses `xlExpression` + `$Col` absolute references on contiguous sqrefs.** Fragmented sqrefs silently break new-row coverage — discovered the hard way (Error 018).
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

This matters more than the artifact.

### Mockup-first development
Every major design decision (colour theme, priority palette, AR formatting, Risk Register structure) rendered as an interactive HTML mockup BEFORE any code was written. Stakeholder sign-off on the mockup. Build once. Field guide — Strategy 4.

### File-upload-first diagnosis
AI's default pattern is to bucket issues into a plausible plan based on project docs. Roughly half the time that plan is wrong because the actual file state has drifted. Rule: never accept a plan that wasn't built from inspection of the current file. See Problem 7.

### Root-cause demand
When something failed twice, refused another patch attempt. This is how `bgColor` vs `fgColor` (the cardinal openpyxl CF bug), triple-quote escaping in CF formulas, and CF sqref fragmentation were all found. Strategy 6.

### Living State file
Every session ends with a state archive. Every session starts by re-uploading it. The conversation IS the project memory. See [`docs/Living-State.html`](docs/Living-State.html).

### Pre-flight checks
Before running any macro: Sub/End Sub balance, With/End With balance, Dim placement, `xlSolid` before `.Interior.Color` on FormatConditions, `.HasTextFrame` guards. Strategy 2.

### Never ship AI-built `.xlsx` files
openpyxl writes CF fills as `bgColor` which Excel renders invisible. Rule: AI writes VBA macros, the user runs them inside Excel. Files Excel touches stay valid; files Python touches break in invisible ways. Error 001.

---

## Tech stack

- **Excel + VBA** — live system
- **Python + openpyxl** — verification and analysis only, never file writes that ship
- **Claude (Anthropic)** — build partner
- **HTML/CSS** — mockups, documentation, embedded help
- **SharePoint** — deployment

---

## Project status

Excel build: ~95% complete. Phase 18 (data migration from old template) in progress. Phases 19–21 pending (printable AR view, team rollout, cosmetic polish).

Parallel migration to Notion underway — see [`docs/Notion-Migration-Notes.md`](docs/Notion-Migration-Notes.md).

---

## A note on AI-assisted development

Built with AI assistance from start to finish. The skill being demonstrated is not "I used AI." It's the discipline of driving AI to produce working corporate-grade output: when to push back, when to demand diagnosis instead of patching, when to mockup before building, when to refuse confident-sounding plans that weren't grounded in the actual file state.

The errors documented here are not hypothetical. Every one cost real time in a live environment.

---

## License

MIT — see [`LICENSE`](LICENSE).

## Author

Adina Lieblich · Civil engineer · AI enablement in AEC · Perth, Australia
[adinalieblich.com](https://adinalieblich.com) ·
[LinkedIn](https://www.linkedin.com/in/adinalieblich) ·
[other repos](https://github.com/adinalieblich)
