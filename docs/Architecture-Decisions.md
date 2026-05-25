# Architecture Decisions

The system arrived at its current shape through three major architectural pivots. Each was a response to a class of failures that the previous architecture couldn't fix without rebuilding.

## Decision 1 — Action Register as filter view, not separate sheet

### Old architecture
Action Register was a separate sheet. A pair of macros (`SyncToAR`, `SyncStatusBack`) kept the AR in sync with Meeting Minutes whenever a checkbox was ticked or a status changed.

### What broke
- Sync ran on description + owner as the lookup key. If either was edited after sync, the row stopped matching and the link silently broke.
- Concurrent edits on SharePoint produced race conditions where AR rows would duplicate or vanish.
- The sync direction was ambiguous when both sides changed — last-write-wins, but "last" was confusing.

### New architecture
Action Register is a saved filter view of Meeting Minutes. The `ShowActionRegisterView` macro applies AutoFilter to MM where Type = "Action" and unhides the Project column (so the filtered actions show which project they belong to). `RestoreFullView` clears the filter and re-hides Project. There is no AR sheet.

### Why this is better
- No sync. No keys. No silent breaks.
- The data has one home. The view is computed.
- Status changes in AR-view are status changes in MM — by definition.

---

## Decision 2 — Type dropdown replaces checkbox columns

### Old architecture
Two columns of form-control checkboxes: "Action?" and "Risk?". Tick either to mark the row.

### What broke
- Form-control checkboxes are `Shape` objects, not cell values. They float over cells, but their anchor is not the cell — it's an offset position. Insert/delete rows above them and they drift.
- 257 of them accumulated in MM over time. The file got slow.
- New rows didn't inherit checkboxes — you had to copy them down manually, which sometimes copied the link to the wrong cell.
- Hiding rows didn't hide checkboxes properly.

### New architecture
A single Type column. Dropdown with values: Action, Risk, blank (default). One column instead of two. A typed value, not a Shape.

### Why this is better
- Type is a cell value, so it inherits properly on insert.
- No floating Shapes to anchor or drift.
- File size dropped, file speed up.
- CF rules can target the Type column with `xlExpression` and work on every row.

---

## Decision 3 — Project lifecycle by row position, no Project Register

### Old architecture
A Project Register sheet held one row per project with Lifecycle (Active / DLP / Closed / etc.), PM, budget, PC date, DLP end date. MM rows referenced their project via a hidden column. Banner rows in MM were rebuilt automatically by a `RebuildBanners` macro that read the Project Register.

### What broke
- `RebuildBanners` was destructive. It wiped and rewrote section/project banners, which sometimes wiped formatting, sometimes wiped notes that lived in banner rows.
- Editing the Project Register meant the banner positions could shift but the row contents wouldn't move — leaving rows under the wrong banner.
- The Project Register was almost never updated by anyone other than the system owner, so it stopped being a source of truth.

### New architecture
- Project Register sheet: deleted.
- Project lifecycle is determined by which section a row sits in within Meeting Minutes.
- Section banners are typed manually. Project sub-banners are typed manually.
- Moving a project's lifecycle = cut/paste the project's rows under a different section banner.

### Why this is better
- No automation that destroys data.
- The structure of the document IS the lifecycle map. No second source of truth to drift.
- Manual moves are infrequent (project lifecycle changes are rare events) and visible.

---

## Decision 4 — Mockup-first development

This isn't a system-architecture decision, it's a methodology decision. But it changed the project's economics more than any code decision.

### What changed
Before any significant code, build an interactive HTML mockup. Get sign-off. Then code.

### Examples that worked
- Colour theme (sage vs navy) — decision in 30 seconds from a side-by-side mockup
- Priority palette — four schemes mocked against real row colours, chosen before any CF rule was written
- AR row formatting (badge-only vs whole-row status colours) — decided before rebuilding the CF system, which saved a rebuild
- Risk & Decision Register design — five-tab interactive mockup, shared directly with stakeholder, governance decision resolved in one meeting instead of weeks

### Why it works
A human designer's mockup takes hours. AI's mockup takes minutes. The cost of "show before build" dropped to near zero, which means there's no reason not to.

### Why it works specifically with AI
AI can render a fully interactive, properly styled, navigable HTML mockup in minutes — using real data from the actual project. The mockup IS the decision document. Share it directly, get sign-off, build once.

---

## Decision 5 — Banners manually managed

Subset of Decision 3 but worth calling out: there is no `RebuildBanners` macro. There used to be. It was deleted because it was destructive.

The rule is: anything that touches user-typed content in cells (other than its own predictable output) is too dangerous to keep. If a macro can wipe a note someone typed, the macro gets deleted, even if 95% of the time it does the right thing.

---

## Decision 6 — Verification by compound key

Discovered during Phase 18 (data migration from old template into new).

### Problem
The migration script needed to compare the new file against an emergency-save source-of-truth. Initial comparison used Description as the key. Result: every duplicate description with different owners got collapsed into one match, hiding real divergences.

### Fix
Use compound key: Description + Owner. Tracks section/project context through merged banner cells. Field-by-field comparison: type, due, status, priority, notes.

### General principle
Whenever an AI proposes "compare on X", check whether X is unique. If not, the comparison will lie.
