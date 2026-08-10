# ACDC EMR Comparison Macro

A single-file Excel VBA macro that validates an **EMR (Oracle Millennium)**
appointment extract against a consolidated **Bookwise** master and, optionally, a
consolidated **iPM** report during the ACDC data migration:

- a consolidated **Bookwise** master workbook, and
- a consolidated **iPM** report.

Each EMR row is routed to the correct source based on its `LOCATION`, then its
fields are compared. The macro flags field discrepancies, unknown / duplicate /
missing / both-present IDs, site mismatches and rows that require manual review,
greys out non-compared columns, and writes an append-only audit log — all
without ever modifying the Bookwise or iPM source files.

- **Module:** [`EMRComparisonMacro.bas`](EMRComparisonMacro.bas)
- **Entry point:** `Public Sub EMRComparisonMacro`
- **Version:** v5

---

## What it does

The macro performs a **one-way check**: the EMR extract is validated *against*
the Bookwise master and the iPM report. Bookings that exist in Bookwise / iPM
but are missing from the EMR are **not** flagged, because migrations are
progressive (the source systems are ahead of the EMR).

For each EMR row it decides one of four outcomes:

1. **Compared** — a confirmed, single-ID row whose `LOCATION` maps to a Bookwise
   or an iPM source is matched to its source record and each field is compared.
2. **Manual review** — the row is set aside (not compared) because of a missing
   ID, duplicate ID, both IDs present, unknown location, or a non-confirmed
   status.
3. **Skipped** — in Bookwise-only mode, an otherwise-comparable iPM-location
   row receives exactly `SKIPPED - NOT COMPARED` and is not compared, highlighted,
   or copied to an output sheet.
4. **Unknown ID** — a compared row whose booking / schedule ID does not exist in
   the corresponding source.

### Routing (per EMR row)

Routing is driven by the EMR `LOCATION` value:

- **Bookwise-location** rows → compared against the Bookwise master (keyed on
  `ACDC_BOOKING_ID` → `Book No.`).
- **iPM-location** rows → compared against the iPM report (keyed on
  `IPM_SCHEDULE_ID` → `i.PM Schedules ID`).
- **Unknown location** → routed to manual review, compared against neither.

A row that has **both** `ACDC_BOOKING_ID` *and* `IPM_SCHEDULE_ID` populated is
sent to manual review and compared against neither source.

### Field comparisons

**Bookwise-location rows**

| EMR field | Bookwise field | Notes |
|-----------|----------------|-------|
| `MRN` | `UR Number` | Trimmed, case-insensitive |
| `APPOINTMENT_DATE` | `Date` | Australian `dd/mm`; 2-digit years map to 2000s (`26` → 2026) |
| `APPOINTMENT_TIME` | `Start Time` | Normalised to `hh:mm` |
| `LOCATION` (implied site) | `Location` | EMR location implies an expected site (Box Hill / Maroondah / Yarra), verified against the Bookwise `Location` column |

**iPM-location rows**

| EMR field | iPM field | Notes |
|-----------|-----------|-------|
| `MRN` | `Patient Id` | Trimmed, case-insensitive |
| `APPOINTMENT_DATE` | `Appointment Date` | Australian `dd/mm`; 2-digit years map to 2000s |
| `APPOINTMENT_TIME` | `Appointment Time` | Normalised to `hh:mm` |
| `LOCATION` (implied clinic) | `Clinic Location` | EMR location implies an expected iPM clinic location, verified against `Clinic Location` |
| — | `Session Code` | A blank `Session Code` on a matched iPM row adds a manual-review note |

---

## Requirements

### EMR extract (the active workbook when you run the macro)

- The active sheet **must be named exactly** `EMR Extract`.
- Headers in **row 1**, data from **row 2**.
- Required headers (row 1, resolved by name, case-insensitive):
  `ACDC_BOOKING_ID`, `IPM_SCHEDULE_ID`, `MRN`, `LOCATION`, `APPOINTMENT_DATE`,
  `APPOINTMENT_TIME`, `STATUS`.
- A `Match Status` column is ensured at column 1 automatically. If a legacy
  `Manual Review Status` column is present (from earlier versions) it is
  **renamed in place** to `Match Status`; otherwise a new column is inserted.
  The operation is re-run safe / idempotent.

### Bookwise master (selected in every run)

- Consolidated, multi-site layout: headers in **row 1**, data from **row 2**
  (no search-parameter or separator rows).
- Required headers (row 1): `Book No.`, `UR Number`, `Date`, `Start Time`,
  `Location`.
- The `Location` column must be **fully populated** on every row that carries a
  `Book No.`, and every value must be one of: **Box Hill**, **Maroondah**,
  **Yarra**.

### iPM report (selected only in full-comparison mode)

- Headers in **row 1**, data from **row 2**.
- Required headers (row 1): `i.PM Schedules ID`, `Patient Id`,
  `Clinic Location`, `Appointment Date`, `Appointment Time`, `Session Code`,
  `Attend Status`.

Every selected source file is opened **read-only** and **closed without saving** —
it is never modified. Bookwise-only mode does not request, open, read, validate, or
close an iPM workbook.

---

## How to run

The macro is designed to be deployed as an **Excel Add-In (`.xlam`)** and run
from a Quick Access Toolbar button.

1. Import `EMRComparisonMacro.bas` into a VBA project (or into an add-in
   workbook) via the VBA editor (`Alt`+`F11` → *File → Import File…*).
2. Save the add-in as `.xlam` and load it (*File → Options → Add-ins → Manage:
   Excel Add-ins → Browse…*), then add a button for `EMRComparisonMacro` to the
   Quick Access Toolbar.
3. Open the EMR extract, ensure the active sheet is named `EMR Extract`, and
   click the button.
4. At the opening **Yes / No / Cancel** prompt, choose whether iPM appointments
   should be compared:
   - **Yes — full comparison:** select the consolidated **Bookwise master** at
     **Step 1 of 2**, then the consolidated **iPM report** at **Step 2 of 2**.
   - **No — Bookwise only:** select only the consolidated **Bookwise master**;
     its picker title has no two-step wording.
   - **Cancel:** exit immediately without showing a file picker or changing a
     workbook.

> The macro uses `ActiveSheet.Parent` (not `ThisWorkbook`) throughout, so all
> output sheets are created in **your** EMR workbook, never in the add-in.

All file pickers required for the selected mode are shown **before** any change is
made. Cancelling any displayed picker aborts cleanly with nothing modified.

---

## Safety gates (checked before any change is made)

The macro validates everything up front and aborts cleanly if a precondition
fails, leaving all workbooks untouched:

1. **Wrong sheet** — active sheet is not named `EMR Extract`.
2. **File picker cancelled** — no Bookwise master, or (in full mode) no iPM
   report selected.
3. **Missing Bookwise columns** — `Book No.`, `UR Number`, `Date`,
   `Start Time`, or `Location` not found in row 1; or no Bookwise data rows.
4. **Full mode only — missing iPM columns** — any of `i.PM Schedules ID`, `Patient Id`,
   `Clinic Location`, `Appointment Date`, `Appointment Time`, `Session Code`,
   `Attend Status` not found in row 1; or no iPM data rows.
5. **Hard stop — duplicate `Book No.`** in the Bookwise master.
6. **Full mode only — hard stop for a cancelled iPM appointment** — any row with
   `Attend Status = Cancelled` (remove it before running).
7. **Full mode only — hard stop for duplicate `i.PM Schedules ID`** in the iPM report.
8. **Hard stop — bad Bookwise `Location`** — blank, or not
   Box Hill / Maroondah / Yarra.
9. **Missing EMR columns** or no data rows.

All applicable hard stops fire **before** the EMR extract is touched, so a failed
precondition never leaves the workbook half-processed. In Bookwise-only mode, all
iPM source gates are bypassed while Bookwise and EMR requirements remain in force.

---

## Outputs

Running the macro (re)builds the following in the EMR workbook. Each run starts
from a fresh slate: prior highlights and `Match Status` values across the data
region are cleared first, and the generated sheets are rebuilt.

### On the `EMR Extract` sheet

- **`Match Status`** column (column 1) — populated with:
  - A combined, semicolon-separated **manual-review label** for any flagged row,
    e.g. `Missing ID - Manual Rv; Status - Manual Rv`. Possible flags:
    `Both IDs Present - Manual Rv`, `Missing ID - Manual Rv`,
    `Duplicate ID - Manual Rv`, `Status - Manual Rv`,
    `Unkn Location - Manual Rv`, `iPM Session Code blank - Manual Rv`.
  - `SKIPPED - NOT COMPARED` for an otherwise-comparable iPM-location row in a
    Bookwise-only run. Existing manual-review flags are evaluated first and always
    take precedence over skipping.
  - A **clean-match label** for a compared row with no discrepancies:
    `Ok - appt match bookwise` or `Ok - appt match iPM`. (A clean iPM row with a
    blank `Session Code` shows the note only, no `Ok`.)
- **Cell highlights:**
  - 🟧 **Orange** — ID / location issues (missing ID, duplicate ID, both IDs
    present, unknown location, or a compared ID not found in its source).
  - 🟨 **Yellow** — a field mismatch on a compared row (MRN, date, time, site).
  - ⬜ **Grey** — **non-compared** columns (header + all data rows) are shaded
    light grey. Compared fields, the ID columns, `STATUS` and `Match Status` are
    never greyed.

### Generated sheets

| Sheet | Contents |
|-------|----------|
| **`Bookwise Mismatches to Review`** | The matching Bookwise row (all columns) for every compared Bookwise row with a field/site mismatch. Differing `UR Number` / `Date` / `Start Time` / `Location` cells highlighted yellow, plus a `Dt/Tm Added by Macro` timestamp. Non-compared columns greyed. |
| **`iPM Mismatches to Review`** | In full mode, the normal iPM mismatch table, highlights, timestamp, and greying. In Bookwise-only mode, the cleared sheet contains exactly one populated cell: A1 = `iPM appointments not compared this run by macro`. |
| **`EMR Appt Unkn ID`** | The EMR row (all data columns except `Match Status`) for every compared row — Bookwise **or** iPM — whose ID is not found in its source. ID cell highlighted orange, plus a timestamp. Non-compared columns greyed. |
| **`Reconcile Audit Log`** | Visible, **protected** append-only log — one row per run. |

### Audit log columns

`Run Date/Time` · `Run By (Username)` · `Bookwise File Path` · `iPM File Path` ·
`Total Processed` · `Bookwise Compared` · `iPM Compared` · `Clean Matches` ·
`Mismatches` · `Unknown IDs` · `Manual Review Total` · `Missing ID` ·
`Duplicate ID` · `Both IDs Present` · `Unknown Location` · `Status` ·
`iPM Rows Skipped`

A summary message box with the same counts is shown at the end of each run,
including `iPM rows skipped`. Full mode records zero skipped; Bookwise-only mode
records zero iPM compared. The audit `iPM File Path` contains the selected path in
full mode or exactly `NOT SELECTED - iPM comparison disabled` in Bookwise-only
mode. Existing v4 audit sheets are upgraded in place by appending the new header;
historical rows are not shifted or cleared.

---

## Location handling

The EMR `LOCATION` value determines how a row is treated and which source it is
compared against.

**Bookwise locations** → compared against Bookwise, mapped to an expected site:

| EMR `LOCATION` | Expected Bookwise site |
|----------------|------------------------|
| `4.3 BHH` | Box Hill |
| `Alexandra Day Oncology BHH` | Box Hill |
| `DAYCHEMO MAR` | Maroondah |
| `YR ONC YRS` | Yarra |

**iPM locations** → compared against iPM, mapped to an expected clinic location:

| EMR `LOCATION` | Expected iPM `Clinic Location` |
|----------------|--------------------------------|
| `A 4.2 BHH` | BHH Bldg A 4.2 Oncology |
| `ADH Clinic` | Alexandra District Health Oncology |
| `BCC Ground Floor MAR` | EH Breast and Cancer Centre |
| `OP YRS` | YRH OPD Clinic |
| `Outpatients HEA` | HDH OPD Clinics |

**Anything else** → flagged as an unknown location and routed to manual review.

The location lists live in one place (`LocationCategory`), with the site /
clinic mappings in `ExpectedSite` and `ExpectedIPMLocation`.

---

## Status gate

A row is only compared when its `STATUS` matches `Booked(Confirmed)`
(case-insensitive). Any other status routes the row to manual review.

---

## Notes

- Unparseable dates or times are treated as **not equal**, so they surface for
  manual review rather than silently passing.
- Duplicate EMR IDs do **not** stop the macro — they are flagged for manual
  review. Duplicate Bookwise `Book No.` values **do** stop every run. Duplicate iPM
  `i.PM Schedules ID` values and cancelled iPM appointments stop full-mode runs;
  those iPM validations do not run in Bookwise-only mode.
- `Option Explicit` is on; all comparisons trim and are case-insensitive.

---

## Version history

- **v5** — Added a Yes/No/Cancel run-mode prompt, conditional iPM selection and
  validation, Bookwise-only skipping for otherwise-comparable iPM rows, conditional
  iPM mismatch output, and skipped-row summary/audit reporting with an idempotent
  v4 audit-schema upgrade.
- **v4** — `Manual Review Status` column renamed to `Match Status` (re-run safe;
  reuses/renames the legacy column). Clean matches labelled
  `Ok - appt match bookwise` / `Ok - appt match iPM`. Non-compared columns
  shaded light grey on the EMR extract and all output sheets.
- **v3** — Added the second file picker and the full iPM comparison path, iPM
  hard stops (missing columns, cancelled Attend Status, duplicate schedule ID),
  the `iPM Mismatches to Review` sheet, and the both-IDs flag.
- **v2** — Date/time copy-out written as canonical **text** into Text-formatted
  cells; locale-safe time normalisation.
- **v1** — Initial one-way EMR → Bookwise check.
