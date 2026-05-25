# Contributing

This is a portfolio repo, but the methodology is reusable. If you're picking this up as a template for your own AI-assisted Excel/VBA work, here's how to use it.

## For human contributors

Reading order:
1. `README.md` — what's here and why
2. `docs/What-AI-Wont-Tell-You.html` — the field guide (most valuable file in the repo)
3. `docs/Architecture-Decisions.md` — the system's evolution
4. `docs/Living-State.html` — current project state
5. `src/vba/Module1.bas` — the live macros

If you want to adapt this for your own project:
- The architecture decisions are not portable. They're answers to specific failure modes that emerged in this codebase.
- The methodology IS portable. Mockup-first, file-upload-first, root-cause demand, Living State file — these work on any AI-assisted project.
- The error catalogue is partially portable. The VBA gotchas and openpyxl corruption signatures apply to anyone building Excel files with AI. The bucketing/patching/screenshot-misreading failure modes apply to any AI-assisted dev work.

## For AI contributors (Claude Code etc.)

Read `CLAUDE.md` in the repo root. That file is written for you. Follow it.

Specifically:
- Never write Excel files with openpyxl. The user runs your VBA inside Excel.
- Read the file before proposing fixes. Don't bucket from project docs.
- Short responses. The user has severe ADHD.
- When something fails twice, demand root cause. Don't patch.
- Update `docs/Living-State.html` at the end of each session.

## Adding a new entry to the error catalogue

If you encounter a new error worth documenting:
1. Confirm it actually breaks Excel (don't document hypothetical errors).
2. Identify the root cause, not just the symptom.
3. Write the working fix.
4. Add a card to `docs/What-AI-Wont-Tell-You.html` matching the existing format (Error number, symptom, root cause, fix, prevention rule).
5. If a diagnosis prompt was what unlocked it, include the prompt verbatim.

## Adding a new strategy

Same pattern, in the Strategies section. The bar is: this saved real time in a real session. Not theoretical.
