# ACDC EMR → Bookwise Migration Check Macro

A single-file Excel VBA macro that validates an **EMR (Oracle Millennium)**
appointment extract against a consolidated **Bookwise** master workbook during
the ACDC data migration. It flags field discrepancies, unknown/duplicate/missing
booking IDs, site mismatches and rows that require manual review, then writes an
audit log — all without ever modifying the Bookwise file.

- **Module:** [`CheckEMRtoBookwiseMigration.bas`](CheckEMRtoBookwiseMigration.bas)
- **Entry point:** `Public Sub CheckEMRtoBookwiseMigration`

---

## What it does

The macro performs a **one-way check**: the EMR extract is validated *against*
the Bookwise master. Bookwise bookings that are missing from the EMR are **not**
flagged, because migrations are progressive (Bookwise is ahead of the EMR).

For each EMR row it decides one of three outcomes:

1. **Compared** — a confirmed, Bookwise-location booking with a single, known ID
   is matched to its Bookwise record and each field is compared.
2. **Manual review** — the row is set aside (not compared) because of a missing
   ID, duplicate ID, unknown location, iPM location, or a non-confirmed status.
3. **Unknown ID** — a compared row whose booking ID does not exist in the
   Bookwise master.

### Field comparisons (compared rows)

| EMR field | Bookwise field | Notes |
|-----------|----------------|-------|
| `MRN` | `UR Number` | Trimmed, case-insensitive |
| `APPOINTMENT_DATE` | `Date` | Australian `dd/mm`; 2-digit years map to 2000s (`26` → 2026) |
| `APPOINTMENT_TIME` | `Start Time` | Normalised to `hh:mm` |
| `LOCATION` (implied site) | `Location` | EMR location implies an expected site, verified against the Bookwise `Location` column |

The booking key linking the two sides is the EMR `ACDC_BOOKING_ID` matched to
the Bookwise `Book No.`.

---

## Requirements

### EMR extract (the active workbook when you run the macro)

- The active sheet **must be named exactly** `EMR Extract`.
- Headers in **row 1**, data from **row 2**.
- Required headers (row 1, resolved by name, case-insensitive):
  `ACDC_BOOKING_ID`, `MRN`, `LOCATION`, `APPOINTMENT_DATE`, `APPOINTMENT_TIME`,
  `STATUS`.
- A `Manual Review Status` column is inserted at column 1 automatically if it
  isn't already present (the operation is idempotent).

### Bookwise master (selected via the file picker)

- Consolidated, multi-site layout: headers in **row 1**, data from **row 2**
  (no search-parameter or separator rows).
- Required headers (row 1): `Book No.`, `UR Number`, `Date`, `Start Time`,
  `Location`.
- The `Location` column must be **fully populated** on every row that carries a
  `Book No.`, and every value must be one of: **Box Hill**, **Maroondah**,
  **Yarra**.

---

## How to run

The macro is designed to be deployed as an **Excel Add-In (`.xlam`)** and run
from a Quick Access Toolbar button.

1. Import `CheckEMRtoBookwiseMigration.bas` into a VBA project (or into an
   add-in workbook) via the VBA editor (`Alt`+`F11` → *File → Import File…*).
2. Save the add-in as `.xlam` and load it (*File → Options → Add-ins → Manage:
   Excel Add-ins → Browse…*), then add a button for
   `CheckEMRtoBookwiseMigration` to the Quick Access Toolbar.
3. Open the EMR extract, ensure the active sheet is named `EMR Extract`, and
   click the button.
4. When prompted, select the consolidated Bookwise master to compare against.

> The macro uses `ActiveSheet.Parent` (not `ThisWorkbook`) throughout, so all
> output sheets are created in **your** EMR workbook, never in the add-in.

The Bookwise file is opened **read-only** and **closed without saving** — it is
never modified.

---

## Safety gates (checked before any change is made)

The macro validates everything up front and aborts cleanly if a precondition
fails, leaving both workbooks untouched:

1. **Wrong sheet** — active sheet is not named `EMR Extract`.
2. **File picker cancelled** — no Bookwise master selected.
3. **Missing Bookwise columns** — `Book No.`, `UR Number`, `Date`,
   `Start Time`, or `Location` not found in row 1.
4. **Hard stop — duplicate `Book No.`** in the Bookwise master.
5. **Hard stop — bad `Location`** — blank, or not Box Hill / Maroondah / Yarra.
6. **Missing EMR columns** or no data rows.

---

## Outputs

Running the macro (re)builds the following in the EMR workbook. Each run starts
from a fresh slate: prior highlights and `Manual Review Status` values across the
data region are cleared first.

### On the `EMR Extract` sheet

- **`Manual Review Status`** column (column 1) — populated with a combined,
  semicolon-separated label for any flagged row, e.g.
  `Missing ID - Manual Rv; Status - Manual Rv`.
- **Cell highlights:**
  - 🟧 **Orange** — ID / location issues (missing ID, duplicate ID, unknown
    location, or a compared ID not found in Bookwise).
  - 🟨 **Yellow** — a field mismatch on a compared row (MRN, date, time, site).

### Generated sheets

| Sheet | Contents |
|-------|----------|
| **`Bookwise Mismatches to Review`** | The matching Bookwise row (all columns) for every compared row with a field/site mismatch. Differing `UR Number` / `Date` / `Start Time` / `Location` cells highlighted yellow, plus a `Dt/Tm Added by Macro` timestamp. |
| **`EMR Appt Unkn ID`** | The EMR row (all data columns except `Manual Review Status`) for every compared row whose ID is not in Bookwise. ID cell highlighted orange, plus a timestamp. |
| **`Reconcile Audit Log`** | Visible, **protected** append-only log — one row per run. |

### Audit log columns

`Run Date/Time` · `Run By (Username)` · `Bookwise File Path` · `Total Processed`
· `Compared` · `Clean Matches` · `Mismatches` · `Unknown IDs` ·
`Manual Review Total` · `Missing ID` · `Duplicate ID` · `Unknown Location` ·
`iPM Location` · `Status`

A summary message box with the same counts is shown at the end of each run.

---

## Location handling

The EMR `LOCATION` value determines how a row is treated:

- **Bookwise locations** (`4.3 BHH`, `Alexandra Day Oncology BHH`,
  `DAYCHEMO MAR`, `YR ONC YRS`) → compared, and mapped to an expected Bookwise
  site (Box Hill / Maroondah / Yarra) for the site check.
- **iPM locations** (`A 4.2 BHH`, `ADH Clinic`, `BCC Ground Floor MAR`,
  `OP YRS`, `Outpatients HEA`) → routed to manual review.
- **Anything else** → flagged as an unknown location for manual review.

### Phase 2 hook

iPM-location rows are currently routed to manual review via
`LocationCategory()`. To automate the iPM comparison later, add a second
read-only picker and replace **only** the iPM branch in the main loop; the iPM
location list lives in one place (`LocationCategory`).

---

## Status gate

A row is only compared when its `STATUS` matches `Booked(Confirmed)`
(case-insensitive). Any other status routes the row to manual review.

---

## Notes

- Unparseable dates or times are treated as **not equal**, so they surface for
  manual review rather than silently passing.
- Duplicate EMR booking IDs do **not** stop the macro — they are flagged for
  manual review. Duplicate Bookwise `Book No.` values **do** stop it (hard stop).
- `Option Explicit` is on; all comparisons trim and are case-insensitive.
