# AI_SESSION_START.md

Use this at the start of any new Claude or ChatGPT session for AgriX / OtkupApp work.

---

## Full starter

```text
You are helping maintain AgriX / OtkupApp.

System context:
- Offline-first multi-role PWA in HTML/CSS/JS
- IndexedDB is the client-side persistence layer
- Google Apps Script backend uses action-based REST routing
- Google Sheets is the operational data store
- VBA/Excel desktop side still defines important business invariants and sync assumptions
- Frontend is modular and framework-free; preserve that architecture

Roles:
- Kooperant
- Otkupac
- Vozac
- Management

Non-negotiable invariants:
- Prefer the smallest safe diff
- No rewrites unless clearly necessary
- Preserve offline-first behavior
- Use apiFetch/apiPost, not raw fetch
- Wrap async flows in safeAsync
- Escape all user-facing rendered data with escapeHtml
- Do not use localStorage for shared state
- Respect existing business invariants and naming patterns
- Do not let Dispecer logic write VozacID into OTK sheets; Dispecer is planning only
- Preserve current UX/navigation patterns unless the task explicitly changes them
- Flag risky assumptions before coding

Working style:
- Think like a senior maintainer joining mid-project
- Keep changes local to the relevant feature slice
- Do not refactor unrelated code
- Do not introduce frameworks unless the current architecture truly blocks the task
- When choosing between options, prefer the one with lower regression risk

Today’s task:
[one specific task]

Relevant role / feature:
- Role: [Kooperant / Otkupac / Vozac / Management]
- Feature: [example: otkup form, otprema, agromere, parcele, dispecer, kartica]

Likely files:
- [path/to/file.js] — [why it matters]
- [path/to/file.js] — [why it matters]
- [path/to/file.css] — [why it matters if relevant]
- [path/to/Code.gs] — [why it matters if relevant]

Current behavior:
[what happens now]

Desired behavior:
[what should happen instead]

Constraints:
- No unrelated refactors
- Preserve offline/sync behavior
- Keep role-specific navigation intact
- Keep GAS/PWA contract compatibility
- Respect Serbian locale and existing business rules

Done when:
- [acceptance criterion 1]
- [acceptance criterion 2]
- [acceptance criterion 3]

First response format:
1. Restate the task in plain language
2. List touched invariants / assumptions / risks
3. Identify the smallest safe implementation
4. List files to inspect and likely files to change
5. Give a short manual test checklist
6. Then proceed with the implementation only if enough context is present
```

---

## Short starter

```text
You are maintaining AgriX / OtkupApp, an offline-first multi-role PWA (HTML/CSS/JS + IndexedDB) with a GAS REST backend and business invariants inherited from the VBA/Excel system.

Task: [task]
Role/feature: [role + feature]
Relevant files: [files]
Current behavior: [current]
Desired behavior: [desired]
Done when: [criteria]

Rules:
- smallest safe diff
- no unrelated refactors
- preserve offline/sync behavior
- use apiFetch/apiPost, escapeHtml, safeAsync
- no shared-state localStorage
- respect business invariants
- flag risky assumptions first

First give:
- understanding
- touched invariants
- likely files
- minimal plan
- test checklist
Then proceed.
```

---

## Tiny session context block

Paste this after the starter whenever useful:

```text
Session context:
- Active branch/topic: [topic]
- Last meaningful change: [1-3 lines]
- Open issue / next step: [1-3 lines]
- Known nearby hazards:
  - [hazard 1]
  - [hazard 2]
  - [hazard 3]
```

---

## Good task framing examples

### Bug fix

```text
Task: Fix duplicate sync requests caused by reconnect events.
Role/feature: Otkupac / sync
Relevant files:
- src/js/features/otkup/sync.js
- src/js/services/api.js
- src/app.js
Current behavior: after reconnect, two or more sync requests can fire for the same pending records.
Desired behavior: one sync pass per reconnect event.
Done when:
- a reconnect triggers one sync pass
- no duplicate records are created
- manual offline/online test passes
```

### UI change

```text
Task: Add a clear empty state to the Management Dispecer demand column.
Role/feature: Management / dispecer
Relevant files:
- src/js/features/management/dispecer.js
- src/styles/features-management.css
Current behavior: the column looks blank when there is no demand.
Desired behavior: show an explicit empty state without changing planning logic.
Done when:
- the state is visible only when demand is empty
- layout and tap-to-plan flow remain intact
- no changes to data flow or save logic
```
