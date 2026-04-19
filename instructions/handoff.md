# AI_HANDOFF.md

Use this at the end of every session so the next session starts fast.

---

## Handoff template

```text
Create a handoff for the next AgriX / OtkupApp session.

Include:
1. What changed
2. Files touched
3. Why those files changed
4. Invariants involved
5. Remaining issues / open questions
6. Regression risks
7. Exact next step
8. Copy-paste continuation prompt

Format it like this:

Summary:
- [2-5 bullets]

Files touched:
- [file] — [purpose]
- [file] — [purpose]

Invariants involved:
- [invariant]
- [invariant]

Open questions:
- [question]
- [question]

Regression risks:
- [risk]
- [risk]

Next step:
- [one concrete next step]

Continuation prompt:
Continuing work on AgriX / OtkupApp.

Handoff from previous session:
[paste summary here]

Today’s goal:
[new goal]

Before changing code:
- summarize where we are
- identify the smallest safe next step
- list files to inspect
- note regression risks
Then proceed.
```

---

## Short handoff version

```text
Create a concise handoff with:
- what changed
- files touched
- invariants involved
- remaining issue
- next step
- continuation prompt
```

---

## Maintainer note

A good handoff should make the next session possible without re-explaining the whole app. It should describe only the feature slice that was touched, not the entire system.
