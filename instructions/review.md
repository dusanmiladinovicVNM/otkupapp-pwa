# AI_REVIEW_STARTER.md

Use this when you want Claude or ChatGPT to review a proposed change, patch, diff, or plan.

---

## Full review starter

```text
Review this proposed AgriX / OtkupApp change like a strict senior maintainer.

System context:
- Offline-first multi-role PWA in HTML/CSS/JS
- IndexedDB for client persistence
- GAS action-based backend
- Google Sheets operational storage
- Business invariants inherited from the VBA/Excel side still matter

Review target:
- [patch / code block / diff / plan]

Feature area:
- Role: [Kooperant / Otkupac / Vozac / Management]
- Feature: [feature name]

Relevant files:
- [file]
- [file]
- [file]

Review criteria:
- architectural fit
- smallest-safe-diff discipline
- offline/sync correctness
- role-navigation safety
- GAS/PWA contract compatibility
- business-rule correctness
- DOM safety and escaping
- async/error-handling consistency
- regression risk

Non-negotiable checks:
- no raw fetch where apiFetch/apiPost should be used
- no unsafe user-facing rendering without escapeHtml
- no shared-state localStorage
- no logic that violates role boundaries or Dispecer planning rules
- no accidental break of offline-first behavior
- no unnecessary refactor outside the task slice

Return format:
1. Verdict: acceptable / risky / reject
2. Critical issues
3. Likely bugs or edge cases
4. Invariants potentially violated
5. Minimal fixes required
6. Manual test checklist
7. Safer alternative if the approach is too invasive
```

---

## Diff review version

```text
Review this diff for AgriX / OtkupApp.

Focus on:
- regression risks
- invariant violations
- offline/sync bugs
- role-specific navigation breakage
- GAS/backend contract mismatches
- unsafe DOM rendering
- localStorage misuse
- hidden coupling to other modules

Give me:
1. critical problems
2. medium-risk concerns
3. missed tests
4. the smallest changes needed before merge
```

---

## Bug-fix review version

```text
Review this bug-fix proposal for AgriX / OtkupApp.

Do not redesign it. Check whether the fix:
- actually addresses the root cause
- stays local to the feature slice
- preserves offline-first behavior
- keeps existing UX/navigation intact
- avoids new sync or data consistency issues

Return:
- root-cause fit
- risks
- edge cases
- minimal improvements
- test checklist
```

---

## Suggested add-on block

When the change is near sensitive areas, append this:

```text
Sensitive context:
- Nearby known hazards:
  - role-nav config mismatch risk
  - tabs.js guard issues
  - service worker asset coverage gaps
  - cache invalidation / duplicate sync behavior
Please explicitly check whether this change worsens any of those.
```
