# Notion Migration Notes

Parallel migration of the Excel system to Notion is in progress. The Excel system continues to be the live system during migration.

## Why migrate

- The Excel system works, but every new row requires Ctrl+Shift+A or right-click + macro. Not discoverable for new team members.
- Conditional formatting on growing data is fragile (see What AI Won't Tell You — Error 018).
- Form-control checkboxes were a recurring failure mode; dropdowns fixed the symptom but the underlying constraint (Excel's row-inheritance model) keeps biting.
- The team is small enough that Notion's per-seat cost is workable.
- Notion's native database/views model maps cleanly onto how the team actually thinks about the work.

## Migration status

### Done
- Rename tab → "Minutes & Actions"

### In progress
- Rewrite "How To Use" for Notion context
- Build full shortcuts tab
- Add Changelog tab
- Team rollout

### Pending
- Module rewrite (Daily / Emergency / Setup / Diagnostic — with headers)
- Team workflow map in plain language

## Open UX issue

Add Row is unintuitive. Ctrl+Shift+A works in Excel but isn't discoverable. Options being considered:

- Per-section "+" button
- Floating context-aware button
- Right-click menu

Not yet resolved.

## What's NOT migrating

- The VBA macros — Notion's automations cover most of what they did, and the rest can be Notion buttons or formulas.
- The pre-flight check workflow — Notion doesn't have macro syntax to validate.
- The mockup-first methodology IS migrating, because that's about the AI partnership, not the platform.

## What might get worse in Notion

- Print layouts. The Excel system has a printable Action Register view (Phase 19). Notion's print fidelity is weaker.
- Offline access. SharePoint-synced Excel files work offline; Notion needs network.
- Permissions in a SharePoint-only IT environment. To be confirmed.
