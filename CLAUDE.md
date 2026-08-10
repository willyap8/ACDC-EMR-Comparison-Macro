# CLAUDE.md — EMRComparisonMacro

## Project context

Part of the **ACDC Data Migration Project**: migrating clinical scheduling data from legacy **Bookwise** and **iPM** systems into **Oracle Millennium EMR** across three hospital sites — Box Hill, Maroondah, Yarra Ranges.

This repo is one of three sibling tools in the macro suite:
- `EMRComparisonMacro` — this repo (`ACDC-EMR-Comparison-Macro`)
- `ACDC-Bookwise-Comparison-Macro` (separate repo)
- `ACDC-iPM-Comparison-Macro` (separate repo)

Deployed as an Excel Add-In (`.xlam`), launched via a Quick Access Toolbar (QAT) button, used by **non-technical clinical and booking staff** in an M365/SharePoint environment.

## What this macro does

One-way accuracy check: compares EMR appointments against a source-of-truth master, selecting **either**:
- a consolidated **Bookwise** master, or
- a consolidated **iPM** report

...depending on each row's `LOCATION` value. The user chooses full or Bookwise-only mode up front; every selected source is opened read-only; each EMR row is routed to one source (or to manual review) based on its `LOCATION` and IDs.

Current version: **v5**.

## Non-negotiable rules — data integrity

These override any other instruction or convenience shortcut:

- Never modify source files. Both the Bookwise master and the iPM report are opened **read-only** and closed without saving.
- Hard stops fire *before* any changes begin if pre-conditions aren't met (see Hard Stops below) — never partially process then fail.
- Prefer local copies over running directly against live SharePoint-hosted files (sync conflict risk).
- Never use `CDate` to parse a date/time **string**. Parse strings explicitly, Australian day-first order (see `TryParseDate`). `CDate`/`CDbl` are only ever applied to values already typed as `Date` or to numeric serials — never to text.
- Dates/times written to review sheets follow the canonical-text pattern (`NormaliseDateText` / `NormaliseTimeText`): pre-format the destination cell as `@` (text), then write the canonical string. This prevents US-locale serial-number reinterpretation when the file is opened on a US-locale machine.

## Add-in architecture rules

- Always reference `ActiveSheet.Parent`, **never** `ThisWorkbook` — the macro runs from the add-in, not the target workbook, so `ThisWorkbook` points to the wrong file.
- Sub names must stay stable across point releases. The QAT button binds to a specific sub name (`EMRComparisonMacro`); renaming it silently breaks the entry point for every user who already has the add-in installed.
- Re-runs must be idempotent: column-insertion logic reuses/renames legacy columns rather than duplicating them; output/review sheets are cleared and fully rebuilt each run, not appended to.

## v5 current behaviour

- `Match Status` column (renamed from `Manual Review Status` in v4 — the macro renames a legacy-named column in place; if you see the old name in older code paths, treat it as legacy).
- Clean-match labels: `"Ok - appt match bookwise"` / `"Ok - appt match iPM"`.
- Non-compared columns are grey-shaded (visual cue for reviewers — don't remove without a reason). Compared fields, the ID columns, `STATUS` and `Match Status` are never greyed.
- Output sheets:
  - `Bookwise Mismatches to Review`
  - `iPM Mismatches to Review`
  - `EMR Appt Unkn ID`
- A protected `Reconcile Audit Log` sheet records each run (one appended row per run — this is the one sheet that is *not* rebuilt from scratch). Keep it protected; don't add write paths that bypass it.

## Hard stops (all fire before the EMR extract is touched)

- Active sheet is not named `EMR Extract`.
- Any picker required by the selected mode is cancelled (Bookwise master, or the iPM report in full mode).
- Missing required Bookwise / EMR columns, or no data rows; in full mode, missing required iPM columns or data rows.
- Duplicate `Book No.` in the Bookwise master.
- In full mode, cancelled appointment(s) in the iPM report (`Attend Status = Cancelled`).
- In full mode, duplicate `i.PM Schedules ID` in the iPM report.
- Bookwise `Location` blank or not one of Box Hill / Maroondah / Yarra.

## Domain terms

`ACDC_BOOKING_ID`, `IPM_SCHEDULE_ID`, `Match Status`, `EMR Appt Unkn ID`, `Bookwise Mismatches to Review`, `iPM Mismatches to Review`, `Reconcile Audit Log`.

## Working conventions

- **Plan before building.** Resolve requirements and open questions before writing code. For any non-trivial change, state the plan first.
- **Ask before assuming.** If a request is ambiguous, ask a short clarifying question rather than guessing — especially anything touching hard-stop conditions, the read-only source guarantee, or the audit-log/output-sheet logic.
- **Inspect real files before proposing fixes.** When storage types, formats, or column positions matter, confirm them against real **de-identified** sample files (e.g. with Python/openpyxl) rather than assuming structure from memory or from the code alone. In Claude Code web/CLI sessions the container has no access to local machines — upload de-identified samples into the session when this is needed.
- **Version incrementally.** Each version should be tested and confirmed working before the next change is made. Wait for a precise failure description before diagnosing.
- **VBA modules are version-controlled as plain-text `.bas`/`.cls` exports** — no native binary VBA committed to git. `EMRComparisonMacro.bas` is the source of record; keep it in sync with the actual `.xlam` after every change.

## Open items

- Log any noted-but-not-yet-actioned enhancements here as they come up, so they aren't lost between sessions.
