# CLAUDE.md — Instructions for AI assistants working on this project

Read this file FIRST in every session. These rules exist because they were learned the hard way. Each one has a documented error or failure mode behind it.

---

## TL;DR

- Read the actual file before proposing fixes. Don't bucket from assumptions.
- AI writes VBA macros. The user runs them inside Excel. Never ship `.xlsx`/`.xlsm` files built with openpyxl.
- Short, granular, sequential responses. User has severe ADHD.
- When a fix fails twice, stop patching. Demand the root cause.
- Verify everything Claude claims about the file by re-inspecting it.

---

## The cardinal rule

**Never build or ship Excel files with openpyxl, python-docx, or any non-native library that touches a binary/ZIP-based format's internal XML.**

openpyxl writes conditional-formatting fills as `bgColor` which Excel renders as invisible. The file looks corrupt to the user but valid to openpyxl. The only fix is rebuilding the file via VBA inside a fresh blank workbook.

openpyxl IS safe for: reading values, reading structure for analysis, creating brand-new files BEFORE any Excel editing.

openpyxl is NOT safe for: any file that has been or will be touched by Excel. Use VBA instead.

This applies to Word (`.docx` → VBA not python-docx) and PowerPoint (`.pptx` → VBA not python-pptx) too.

---

## Response format rules

The user has severe ADHD. Responses must be:

- Short and succinct. Zero fluff.
- Granular and sequential. One step at a time when something is long.
- No multi-bucket plans before reading the actual file.
- No "let me explain my thinking" preambles.
- If a response feels long, split across multiple turns.

---

## Recurring failure modes (don't repeat these)

### 1. Confident wrong assessments
Claude has been wrong multiple times about XML column positions, file state, and what migration "will fix." If the user pushes back on a claim, open the file and verify. Don't argue from project docs.

### 2. Bucketing before reading the file
Don't propose a categorised plan ("Bucket A migration kills these / Bucket B fix these / Bucket C new features") before reading the current file. The categorisation looks structured but is built on assumptions. Once one bucket is wrong, the whole plan unravels.

### 3. "Migration / Notion / Tables will fix this"
Never claim a different system will fix something without naming the specific mechanism. "Tables auto-extend CF" is checkable. "Migration will sort it out" is not.

### 4. Patching instead of diagnosing
If the same area has needed three fixes, the fix isn't the answer — the architecture is wrong. Zoom out.

### 5. Forgetting the header block
Meeting Minutes header block is rows 1–14, cols A–J. Title, location, date, attendees, shortcuts panel. NEVER suggest column-delete operations — would destroy header. All data macros start at row 16 minimum (row 15 = column headers).

### 6. Suggesting new chats when things get complex
The conversation is the only persistent memory. A long messy chat is better than a clean restart. If context is getting lost, ask the user to re-upload `docs/Living-State.html`.

### 7. Misreading screenshots
Especially adjacent colours (pink/amber, sage/olive). Call out the actual hex, don't rely on a glance.

### 8. Forgetting bookmarked items
"Bookmarked" alone means it gets dropped within 5–10 turns. Real fix: write it into `docs/Living-State.html` Pending Tasks immediately, then re-export.

### 9. Not picking the user's stated best option
If the user says "go with the best", do that. Don't fall back to second-best because it's easier.

---

## VBA gotchas (confirmed in this project)

- Reserved word `any` cannot be used as a variable name.
- VBA line continuation limit: flatten nested arrays via helper subs.
- Dropdown syntax: `Formula1:="Active,DLP,Closed"` — plain quotes only. NO triple quotes. (Error 023)
- CF expression formula: `Formula1:="=$D2=""Action"""` — triple quotes here ARE correct because inner quotes need to survive into the formula. (Error 016)
- Use `xlExpression` with `$Col` absolute reference for CF on growing data. Never `cellIs equal to "value"`. (Error 018)
- CF `sqref` must be a single contiguous range like `A2:H1000`. Fragmented sqrefs silently break new-row CF coverage.
- Shape `.TextFrame.Characters.Text` can throw on shapes without text frames. Guard with `If shp.HasTextFrame Then`.
- `msoShapeTypeMax` is unreliable across Office versions — don't use it.
- Form-control checkboxes drift, accumulate, slow the file. Use a typed value (dropdown/Y-N/TRUE-FALSE). (Error 019)
- `.Interior.Color` on `FormatCondition` requires `.Interior.Pattern = xlSolid` to be set first, or the fill renders invisible.

---

## Workflow rules

### When the user reports a bug
1. Ask for the file. Inspect it. Don't propose fixes from project docs alone.
2. If the same area has failed before, check `docs/What-AI-Wont-Tell-You.html` error index.
3. Diagnose first. Propose the fix second.
4. After the fix, verify by re-inspecting the file.

### When the user asks for a new feature
1. Check `docs/Living-State.html` for locked decisions that might conflict.
2. If significant: build an interactive HTML mockup FIRST. Get sign-off. Then code.
3. Write VBA. User runs it inside Excel. Never deliver a pre-built `.xlsm`.

### When ending a session
1. Update `docs/Living-State.html` — pending tasks, locked decisions, recently completed.
2. Re-export it. The user re-uploads at the start of the next session.

### When the user pushes back
They're right. Open the file. Do the file-level check. Don't refer to project docs as evidence of current file state.

---

## Project state (as of last session)

See `docs/Living-State.html` for the canonical live status. Highlights:

- Phase 18 in progress: data migration from old template
- Active phases: 18 (data merge), 19 (printable AR view), 20 (team rollout), 21 (cosmetic polish + style guide PDF)
- Latest file: `Migration_v1_P18_complete.xlsm` (not in repo — sanitised example only)
- Architecture: "Option 3" — locked, do not propose alternatives
- Source of truth for data verification: `260505_EMERGENCY_SAVE.xlsx`
- Verification method: compound key (description + owner), not description alone

---

## File structure for context

```
docs/
  Living-State.html              ← read this every session
  What-AI-Wont-Tell-You.html     ← errors + strategies field guide
  How-To-Use.html                ← user-facing guide embedded in workbook
  Style-Guide.html               ← palette, hex codes, banner specs
  Notion-Migration-Notes.md      ← parallel migration in flight
src/
  vba/
    Module1.bas                  ← live macros (daily use)
    Module2-archive.bas          ← rare/old, do not run on current file without checking
  python/
    verify_against_source.py     ← compound-key verifier
    inspect_xml.py               ← CF/sqref/fill XML inspection
examples/
  example-workbook.xlsm          ← sanitised, no real data
```

---

## How to ask Claude to do things on this project

Examples that work well:

- "Read `docs/Living-State.html` first, then..."
- "Inspect `examples/example-workbook.xlsm` and report on CF sqref coverage. Don't propose fixes yet."
- "I'm seeing X. Diagnose root cause first, don't patch."
- "Mockup-first: I want to add a Decision Register. Build an interactive HTML mockup with 2-3 design options."

Examples that go badly:

- "Just fix it" (no file uploaded, no inspection — Claude will hallucinate a plan)
- "Continue from yesterday" (no Living State re-upload — context is gone)
- "Build me the file" (Claude will try to ship a .xlsm — don't accept it; demand VBA)
