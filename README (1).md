# XML → SharePoint / Dataverse Import Utility

Extracts licensing application data from InfoPath XML checklists on the network
share and produces a SharePoint/Dataverse-ready CSV and XLSX, plus a full audit
trail.

**Business Intelligence Unit, Consumer Services Operations Division, MPBSDP**
Last updated: 2026-08-06 · Current script version: **v8**

---

## 1. Quick start

```powershell
python "z:\MIPO\OPBISP\Business Intelligence\09 Work In Progress\Ahmed Balfaqih\Python XML to SharePoint CSV\xml_to_sharepoint_v8.py"
```

Requires Python 3.9+ with `pandas` and `openpyxl`. No template file needed — the
140 column names are embedded in the script.

Optional: override paths without editing the file.

```powershell
python xml_to_sharepoint_v8.py "C:\source\folder" "C:\output\folder"
```

Confirm you are running the right version. v8 prints this first:

```
Columns from: embedded column list (140 fields)  (+ 'Missing_ID')
  + retained legacy column: ApplicationFolderID
  + retained legacy column: XMLFileName
```

If you instead see `ERROR: Template file not found`, you are running the
obsolete `xml_to_sharepoint_utility.py`. Delete or rename it — VS Code will
re-run whatever file was last open.

---

## 2. Files

| File | Purpose |
|---|---|
| `xml_to_sharepoint_v8.py` | **Main script.** Scans XMLs, applies all transformation rules, writes the export. |
| `inspect_import_errors.py` | **Post-import diagnostic.** Run only if SharePoint/Dataverse rejects rows. |
| `README.md` | This document. |

Superseded versions (v1–v7) can be archived. v8 is a superset of all of them.

---

## 3. Configuration

All settings are at the top of the script, under `# CONFIGURATION`.

| Setting | Default | Notes |
|---|---|---|
| `INPUT_DIRECTORY` | `...\Licensing Unit - Licensing Applications` | Source share |
| `OUTPUT_DIR` | `...\Downloads\sharepoint_export` | Must **not** be inside `INPUT_DIRECTORY`; the script refuses |
| `TEMPLATE_FILE` | `""` | Empty = use embedded columns. Point at an `.xlsx`/`.csv` to override |
| `APPLY_VALUE_RULES` | `True` | Master switch for all §6 transformations |
| `RETAIN_LEGACY_COLUMNS` | `True` | Keeps `ApplicationFolderID` + `XMLFileName` |
| `YESNONA_UNCONFIRMED` | `["SubmittedFee"]` | See §6.2 — set to `[]` to revert |
| `SORT_OUTPUT` | `True` | Deterministic row order by application number |
| `FLATTEN_NEWLINES` | `True` | Required for CSV import; see §7 |

---

## 4. Output files

Written to `OUTPUT_DIR`:

| File | What it's for |
|---|---|
| `sharepoint_upload_ready.csv` | **The deliverable.** 143 columns, UTF-8 BOM, CRLF, all fields quoted |
| `sharepoint_upload_ready.xlsx` | Same data as Excel (avoids the CSV format warning) |
| `processing_summary_report.txt` | **Read this every run.** Reconciliation, transformation counts, warnings |
| `folder_coverage.csv` | One row per application folder + whether it has a checklist. Empty folders sorted first |
| `failed_files.csv` | Every XML that could not be read, and why |
| `unmapped_fields.csv` | XML field values with no matching column — nothing is silently discarded |
| `xml_tags_not_in_template.csv` | Tally of XML tags with no column |

### Column layout

140 template columns + 2 legacy + `Missing_ID` = **143**.

`Missing_ID` is the extra key column, e.g. `12151-Checklist`. Where a file name
has no leading number (`Master Checklist.xml`), the folder's application number
is prefixed, giving `13049-Master Checklist`, so the value stays unique.

---

## 5. Reading the report

Three checks, in order, every run.

**1. Balance check** — proves nothing was dropped.

```
BALANCE CHECK: 2880 + 0 = 2880  (must equal 2880)
```

**2. Choice-field pre-import check** — added in v8 to stop the Dataverse
rejection loop. If any column still holds raw `1`/`2`/`3` you will see:

```
!! ACTION REQUIRED BEFORE IMPORT !!
  Priority    values=['1', '2', '3']  3 cells
```

Fix by adding the column to `YESNONA_COLUMNS` and rerunning. When clean it
reads `CHOICE-FIELD PRE-IMPORT CHECK: PASSED`.

**3. Column type warnings** — lists every column exceeding 255 characters.
Create these in SharePoint as *Multiple lines of text*, not *Single line of
text*. Currently `ActivityLog` (max 10,329), `Comment` (6,756),
`EmailMessage` (1,417).

---

## 6. Transformation rules

Source: Rachelle's `Import_File_-_Rachelle_Copy` notes sheet, plus follow-up
decisions. All are logged with counts in the report.

### 6.1 Booleans → `True` / `False`
`SendEmail`, `ManagerDecision`. Accepts Yes/No, true/false, 1/0.

### 6.2 `1`/`2`/`3` → `Yes` / `No` / `N/A`
Applied to **54 columns** (Rachelle's column H list, plus `AppOffDirBankrupt`
and `AppOffDirConvicted`).

Idempotent — values already reading Yes/No/N/A pass through unchanged, which
matters because the real data mixes both formats. Unrecognised values are left
alone and flagged rather than guessed at.

> **Unconfirmed:** `SubmittedFee` is also converted, but was **not** on
> Rachelle's list. Evidence: 222 filled cells containing only 1/2/3, with zero
> decimal or dollar values despite the name. Distribution `1`×205, `3`×13,
> `2`×4. **Confirm the meaning with Rachelle**, then either leave as-is or set
> `YESNONA_UNCONFIRMED = []`.

### 6.3 Person names
Applied to 11 columns: `ComplianceOfficerName`, `RegistrationClerk`,
`ManagerEmail`, `ComplianceOfficerEmail`, `ManagerName`, `ManagerSendTo`,
`RegistrarSendTo`, `AssignedTo`, `CreatedBy`, `AssignedBy`, `UserName`.

`Last, First (pronoun) (dept)` → `First Last`, but **only** for the 7 confirmed
active staff. Everyone else — former staff, unrecognised formats — is blanked,
per the "names in red must be blank" instruction.

```
Chris Pittens · Richard Ross · Jocelyne Ng Kam Man · Priya Ajmani
Anthony Fimiani · Kate Kotsopoulos · Joy Theodoulis
```

The allowlist was derived automatically from every (source, target) pair in the
mapping sheet, so it is not hand-typed. To add someone, add one line to
`PERSON_ALLOWLIST` and rerun — nothing else changes. The report lists every
distinct blanked name with a count so you can audit who was affected.

**Two field-specific exceptions:**

- `RegistrarName` — only `Chris Pittens` or `Richard Ross` survive. Anyone
  else, including otherwise-valid staff, is blanked.
- `RegistrarSendTo` — blanks Chris Pittens and Richard Ross *specifically*
  (a registrar forwarding to themselves), per the red marking on those cells.

### 6.4 Other value rules

| Rule | Detail |
|---|---|
| `AppOffDirBankrupt` → `zzAppOffDirBankrupt` | Copied after the Yes/No/N/A conversion. `AppOffDirBankrupt` is authoritative; differing pre-existing zz values are overwritten and counted in the report |
| `CPICRequired` | `No Nit` → `No Hit` (exact match only) |
| `ApplicationType` | Strips stray trailing year: `-2022` → blank, `Renewal-2023` → `Renewal`. Runs *after* the folder-name fallback so a cleaned blank isn't re-polluted |
| Dates | `2026-09-01T00:00:00` → `2026-09-01`. Matches the exact ISO pattern only, so it cannot misfire on free text |
| `ApplicationID` | Trailing `.0` stripped. **The column type must still be set to text in the import mapping** — a CSV cannot enforce that |

### 6.5 Misspelled form fields

The InfoPath form writes some field names incorrectly. These are remapped:

| XML writes | Mapped to | Files |
|---|---|---|
| `FiancialStatement` | `FinancialStatement` | 543 |
| `MangerEndDate` | `ManagerEndDate` | 102 |
| `zzSubmittedFee` | `SubmittedFee` | 1 |

An alias never fires if its source is itself a template column, so adding
columns later cannot cause the same value to land in two places.

---

## 7. Why the script is built the way it is

Each of these fixed a real failure, not a hypothetical one.

- **Windows long paths.** All file opens use the `\\?\` prefix. The source path
  is ~80 characters before nesting; deep folders exceed the 260-char MAX_PATH
  and fail silently otherwise.
- **`os.walk` error trapping.** Without an `onerror` handler, an unreadable
  folder is skipped with no message. Verified: unreadable folders are now named
  in the report.
- **XML repair.** Bare `&` (common in business names — "Smith & Jones") is a
  fatal parse error. The script escapes it and retries, along with stripping
  illegal control characters, BOMs, and junk before the prolog. HTML files
  masquerading as XML are reported separately rather than counted as corrupt.
- **Rich text vs containers.** `<my:group>` wraps other fields; concatenating
  it would dump every child into one cell. Wrappers contribute only their own
  direct text, while rich-text fields (`<div>`, `<b>` children) still
  concatenate properly.
- **Newline flattening.** `ActivityLog` contains embedded newlines in ~2,600
  cells. SharePoint's CSV importer reads a newline inside a quoted cell as a
  new record — the old 2,750-record export was 35,708 physical lines. Newlines
  become ` | `.
- **Output isolation.** The script refuses to write into the folder it scans.
  Earlier exports were left inside the source share and got re-scanned.
- **Sniff skip.** Known binary extensions are not opened to check whether they
  are secretly XML. This cut a 22-minute run substantially, since ~20,300 of
  the 20,447 non-XML files are PDFs, PNGs and MSGs.

---

## 8. Known data issues

Not script bugs — real conditions in the source data.

### Folders with no checklist (~317)
The tree has ~3,192 application folders but ~2,880 XMLs. The difference is
folders that contain no checklist at all. `folder_coverage.csv` lists them,
empty ones first. This is the worklist for chasing missing paperwork.

### Duplicate `Missing_ID` (4 groups, 10 rows)
Will collide on import.

- `checklist` × 4 — `RachelleTest2`, `RachelleTest3`,
  `Matttest20260227-resumereview1`, `International Customer Care Services Inc`
- `10000-Master Checklist` × 2 — `Application Folder Template` and
  `Application Folder Template_old`
- `20753-Checklist` × 2 and `23283-Checklist` × 2 — the same checklist exists
  both inside its numbered folder and as a loose copy

### Test and template folders (8 rows)
`Application Folder Template`, `Application Folder Template_old`, three folders
whose checklist is a stray `Master Checklist.xml` (14292, 14301, 14387),
`Matttest20260227-resumereview1`, `RachelleTest2`, `RachelleTest3`. These are
not real applications and arguably should be excluded before import. Removing
them takes the count to 2,872.

> `International Customer Care Services Inc` looks like a *real* business with a
> badly named file. Do not exclude it as test data.

### `AssignedTo2` (45 files)
Deliberately left unmapped — unclear whether it is a second assignee or a
duplicate of `AssignedTo`. Values are preserved in `unmapped_fields.csv`.
A commented-out alias exists in the config if it turns out to be the same field.

---

## 9. Explaining a changed row count

**The count is expected to grow between runs.** This is a live folder; new
applications arrive continuously.

Worked example (2026-08-04): a previous export had 2,875 records, a new one had
2,880. All five extra rows were genuine new renewals — `34781`, `34786`,
`34793`, `34803`, `34814`. Evidence that it was real intake, not a scanning
change:

1. Zero rows disappeared; all 2,875 originals were still present.
2. All five IDs were above the previous maximum (34751).
3. All five sat in folders that did not previously exist.
4. All five had 13 of 143 fields filled, against a median of 50 — the profile
   of a freshly created checklist.

**Do not force the count back to match an older file.** Doing so omits real
applications from the system. Explain the delta instead: *"Your file is from
Thursday, mine is from today; five renewals arrived between them, here they are
by ID."*

To investigate a delta yourself: diff `folder_coverage.csv` between runs
(folders flipping `NO XML FOUND` → `OK` are newly added checklists), or sort the
export by `CreatedOn` / `ModifiedOn` and inspect the newest rows.

---

## 10. When an import rejects rows

Dataverse and SharePoint report row numbers but not the offending column.

**Step 1.** Check the report's choice-field section — a raw 1/2/3 column is the
most common cause.

**Step 2.** Run the diagnostic. Edit `ERROR_ROWS` at the top of
`inspect_import_errors.py` with the reported row numbers, then:

```powershell
python inspect_import_errors.py "C:\...\sharepoint_export\sharepoint_upload_ready.csv"
```

It profiles the rows that succeeded and reports what is abnormal in the ones
that failed, so it does not need to know the target schema. It only reports a
condition when it *distinguishes* a failing row from the passing ones —
`ActivityLog` exceeds 255 characters in thousands of rows, so length is not why
11 specific rows failed. It also tests both row alignments, since the importer
does not say whether it counts the header.

Outputs `import_error_diagnosis.txt`, `import_error_rows.csv` (the rejects
alone, for manual entry) and `sharepoint_upload_safe.csv`.

### Recurring causes seen so far

| Symptom | Cause | Fix |
|---|---|---|
| `The value 2 is not a valid value for the ... choice field` | Raw 1/2/3 in a Choice column | Add the column to `YESNONA_COLUMNS` |
| `a person is missing from the address book` | Person/Group column with an unresolvable name | Change the column to *Single line of text*, or correct the name |
| Row count wildly inflated | Embedded newlines in a quoted cell | Keep `FLATTEN_NEWLINES = True` |
| Text truncated or row rejected | Cell over 255 chars in a single-line column | Change to *Multiple lines of text* |

**Recommendation:** create the columns as text initially. Person, Date, Number
and Choice columns all reject rows on any mismatch, and this data spans years of
historical applications with staff who have since left. Convert column types
after the data is in and cleaned.

---

## 11. Reconciling with a Dataverse dataflow

A dataflow run on 2026-08-04 reported 2,543 created + 25 failed = 2,568 rows,
against a 2,880-row export. The table was named `cleaned_nullAppFolder`,
implying a filter on non-null AppFolder that drops ~312 rows upstream. That is
close to the ~317 folders with no checklist and may be the same population seen
from the other end. **Confirm the filter is intentional** — otherwise real
records are being excluded before Dataverse ever sees them.
