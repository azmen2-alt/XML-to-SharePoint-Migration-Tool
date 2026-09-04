NOTE: SCROLL DOWN: 
Activity Log Fixed PY includes fixing the issue where some Activity logs are missing.
Activity Log and override fixed PY includes the above + fixing columns with worong mapping.

# Python-XML-to-CSV
extract, read, and edit XMLs

## Run it

```powershell
python "z:\user\xml_to_sharepoint_v8.py"
```

Requires Python 3.9+ with `pandas` and `openpyxl`. No template file needed.

Confirm v8 starts by seeing `Columns from: embedded column list (140 fields)`. If you see a template error, delete the old `xml_to_sharepoint_v8_Final.py`.

## What you get

**`sharepoint_upload_ready.csv`** — the deliverable. 143 columns, UTF-8 BOM, CRLF, all fields quoted.

**`processing_summary_report.txt`** — read this every run. Shows the balance check (proves nothing was dropped), the choice-field pre-import check (catches raw 1/2/3 before Dataverse rejects it), and which columns exceed 255 characters.

**`folder_coverage.csv`** — one row per application folder. Empty folders (no checklist) sorted first — that's your worklist.

**`failed_files.csv`** — every XML that couldn't be read and why. Should be zero.

# Licensing Applications → SharePoint Export

Two Python scripts that read the checklist XML in every application folder on the
Licensing Unit shared drive and write **one row per application** into a single CSV,
ready to upload to SharePoint / Dataverse.

| File | Use it when |
|---|---|
| `Activity_Log_Fixed_here.py` | You want the standard conversion only. Line breaks stay as plain line breaks. |
| `Activity_Log_and_override_fixed_here.py` | You want the standard conversion **plus** the four column corrections, and line breaks written as `<br>` for PowerApps rich-text fields. |

Both scripts are otherwise identical: same source folder, same template, same output
file name, same checks. Only run **one** of them at a time — they write to the same
output file and the second run overwrites the first.

> The comment block at the top of each file still says
> `python xml_to_sharepoint_v19.py`. That was the old filename. Use the real filenames
> above.

---

## Which one should I run?

Run `Activity_Log_and_override_fixed_here.py` if the destination fields are
**rich text** in PowerApps *and* the business has confirmed the four corrections below.

Run `Activity_Log_Fixed_here.py` if the destination fields are **plain text**, or if you
want the raw form values converted the same way for every checklist column with no
exceptions.

### The three differences, exactly

**1. Four columns get a different Yes/No/N/A rule**

The form stores `1`, `2`, `3`. Normally that means `1 = Yes`, `2 = No`, `3 = N/A`.
For these fields the form displays something different from what it saves, so the
override version flips them:

| Column | Standard version | Override version |
|---|---|---|
| `RiskProfileReport` | 1→Yes, 2→No, 3→N/A | **1→N/A**, 2→No, **3→Yes** |
| `LARSubmitted` | 1→Yes, 2→No, 3→N/A | **1→N/A**, 2→No, **3→Yes** |
| `FinancialStatement` | 1→Yes, 2→No, 3→N/A, `N/A` stays `N/A` | 1→Yes, 2→No, 3→N/A, **`N/A`→Yes** |
| `FinStmtReceived` | no rule | no effect — see "Known issues" |

**2. Line breaks**

- Standard version: a line break inside a comment box stays a line break.
- Override version: it becomes the text `<br>`, which PowerApps rich-text fields render
  as a new line.

**3. Nothing else.** No extra columns, no extra checks, no different file name.

---

# Part A — Activity Log Fixed and Activity Log and override fixed .py -  For the non-technical reader

## What the tool does

1. Looks inside the Licensing Applications folder on the shared drive.
2. In each application folder, finds the checklist file (the saved form, an `.xml` file).
3. Pulls the answers out of it.
4. Puts them into one spreadsheet row, using the column names from the template file
   the business provided.
5. Repeats for every folder, then writes one CSV.

## What it never does

- It **never changes, moves or deletes anything** on the shared drive. It only reads.
- It **never fixes the drive**. If two copies of a checklist exist, the tool skips the
  older one in the CSV — the stray copy is still sitting on the drive.
- It does not upload anything. You do the upload.

## One-time setup

You need three things on the machine that runs it:

1. **Python** installed.
2. **Two Python add-ons**: `pandas` and `openpyxl`. Install with this command:
   ```
   pip install pandas openpyxl
   ```
3. **Three paths set inside the script.** Open either file in Notepad and find the
   `PART 1 - SETTINGS` block near the top. Change only the text between the quotes:

   ```python
   INPUT_DIRECTORY = r"C:\Users\...\Licensing Unit - Licensing Applications"
   OUTPUT_DIR      = r"C:\Users\...\Downloads\sharepoint_export"
   TEMPLATE_PATH   = r"C:\Users\...\Downloads\template.xlsx"
   ```

   Keep the `r` and both quote marks. `INPUT_DIRECTORY` is where the applications are.
   `OUTPUT_DIR` is where the CSV goes — it must **not** be inside `INPUT_DIRECTORY`.
   `TEMPLATE_PATH` is the file the column names come from.

If the business sends a new template, either drop it in at the same path or point
`TEMPLATE_PATH` at the new file. Nothing else needs changing.

## Running it

Open a command prompt in the folder holding the script and type:

```
python Activity_Log_and_override_fixed_here.py
```

To keep a record of the run, add `> run_log.txt`:

```
python Activity_Log_and_override_fixed_here.py > run_log.txt
```

A full run over the whole drive takes a while. Leave it alone until the summary prints.

## What you get

One file: **`sharepoint_upload_ready.csv`** in the output folder.

Everything else appears on screen as an **EXPORT SUMMARY**. Read it before you upload.

## Reading the summary — the four things to check

**1. The BALANCE line.** This is the main safety check.

```
BALANCE    1200+3+2=1205 of 1205 xml  (these must be equal)
```

Rows exported + rows dropped as duplicates + failed files must equal the number of XML
files found. If the two numbers differ, something was lost. Do not upload; investigate.

**2. `Choice-field check  PASSED`.** If instead you see
`!! DO NOT UPLOAD. These columns still hold raw 1/2/3` — stop. SharePoint will reject
those rows. The message names the columns; a developer adds them to the conversion list
and you rerun.

**3. Any line starting with `!!`.** These are problems, not notes.

**4. Rows exported.** Compare it to roughly what you expect. A number that **rose**
since the last run is normal — new applications keep arriving. Never force the total to
match an older file; that leaves real applications out of the system.

## The messages, in plain English

| Message | What it means | What to do |
|---|---|---|
| `TRANSFORMATION RULE(S) DISABLED - column absent` | The template no longer has a column a rule needs, so that conversion did not run. | Check with the business whether the column was dropped on purpose. |
| `DO NOT UPLOAD. These columns still hold raw 1/2/3` | Untranslated form codes are in the file. SharePoint will reject them. | Stop. Send the column names to whoever maintains the script. |
| `FILE(S) COULD NOT BE READ` | A checklist file is empty, is actually an HTML file, or is damaged. Those applications are **not** in the CSV. | Look at the named files on the drive. |
| `Create these as 'Multiple lines of text'` | Some answers are longer than 255 characters. | Set those SharePoint columns to multi-line text before uploading. |
| `duplicate row(s) dropped, newest file kept` | Two checklists for the same application. The newer one was used. | Fine for the upload. Someone should still tidy the drive. |
| `XML FIELD(S) HAVE NO COLUMN` | The form holds answers the template has no column for, so they were left out. | Ask the business whether any are needed. |
| `Values left as-is because they were not recognised` | A value was neither 1/2/3 nor Yes/No/N/A, so it was passed through untouched. | Usually a typo in the original form. |
| `folder(s) contain no checklist` | An application folder has no form in it. | Missing paperwork on the drive, not a fault in the tool. |
| `UNREADABLE FOLDER(S)` | Permissions or a network drop. | Reconnect the drive and rerun. |

## Common questions

**Can I run it twice?** Yes. It rewrites the CSV from scratch each time.

**Can it damage the shared drive?** No. It only opens files for reading.

**The CSV looks odd in Excel.** Excel sometimes reformats long numbers and dates on
opening. Upload the file as produced; don't save over it from Excel.

**I see `<br>` in the text.** You ran the override version. That is intended for
rich-text fields. If the destination field is plain text, run the standard version
instead.

---

# Part B — Activity Log Fixed and Activity Log and override fixed .py - For the technical reader

## Requirements

- Python 3.7+ (`from __future__ import annotations` covers older type-hint syntax).
- `pandas`, plus `openpyxl` if the template is `.xlsx` / `.xlsm` / `.xltx`.
- Windows is handled specially (`\\?\` long-path prefix) but the script also runs on
  Linux/macOS.

## Settings

| Name | Purpose |
|---|---|
| `INPUT_DIRECTORY` | Root scanned recursively for `*.xml`. |
| `OUTPUT_DIR` | Created if missing. Rejected if it resolves inside `INPUT_DIRECTORY`. |
| `TEMPLATE_PATH` | Header row read with `nrows=0`. This is the only source of column names. `.xlsx/.xlsm/.xltx` → `read_excel`; anything else → `read_csv` with `utf-8-sig`. |
| `ID_COLUMN` = `Missing_ID` | Extra trailing column holding the folder-number + filename stem fallback identifier. |
| `DATE_FMT` = `%Y-%m-%d` | Append `" %H:%M:%S"` for times. |
| `SP_TEXT_LIMIT` = `255` | Threshold for the "make this multi-line text" warning. |
| `DROP_COLUMNS` = `["group", "test"]` | Template columns confirmed blank in a full export, excluded from output. |
| `LEGACY_COLUMNS` | `ApplicationFolderID`, `ApplicationFolderName`, `XMLFileName` — appended if the template no longer carries them. |
| `TS_COLUMNS` | `XML_Created`, `XML_Modified`, taken from the file system, not the XML. |

Output column order: template columns (minus `DROP_COLUMNS`) → missing legacy columns →
`Missing_ID` → `XML_Created` → `XML_Modified`.

## Command line

```
python Activity_Log_Fixed_here.py [SOURCE_DIR] [OUTPUT_DIR]
```

Positional arguments override `INPUT_DIRECTORY` and `OUTPUT_DIR`. `TEMPLATE_PATH` has no
argument — edit it in the file.

## Pipeline

1. **Validate** — source exists; output is not nested inside source; `mkdir -p` output.
2. **Load template** — dedupe headers, strip blanks and `Unnamed:` columns. Exits if the
   header row is unusable.
3. **Rule pre-flight** — every column any rule touches is compared against the template.
   Missing ones are reported as `rules_off` and their rules silently do not run.
4. **Walk** — `os.walk` with `onerror` collected, not raised. Counts top-level folders,
   loose files, extensions, and XML per top-level folder.
5. **Per file**: parse (with repair) → flatten to a lookup → map onto columns →
   transformations → file-system timestamps → append row.
6. **DataFrame** — `fillna("").astype(str)`, so every cell is a string.
7. **Sort** — numeric prefix of `ApplicationFolderID`, then the raw ID, then `Missing_ID`.
8. **Dedupe** on `ApplicationID`.
9. **Write CSV** — `utf-8-sig`, `QUOTE_ALL`, `\r\n` line endings.
10. **Summary** to stdout.

## Column matching

`key()` normalises both sides before comparing: NFKD normalise, strip every
non-alphanumeric character, lowercase. So `Application Received`, `application_received`
and `ApplicationReceived` all match the same template column.

`flatten()` walks every element:

- Namespaces are stripped (`{ns}tag` → `tag`).
- An element is treated as **structural** (a wrapper) only if its own name matches no
  template column *and* it has a non-formatting child. Structural elements contribute
  `el.text` only; everything else contributes all descendant text joined. This is what
  keeps rich-text content intact without absorbing sibling fields.
- `HTML_NOISE` is the set of formatting tags used for that decision.
- Attributes are collected too, and merged **before** element values, so an element value
  wins on a name collision.
- Repeated tags are joined with `; ` after `dict.fromkeys` dedupe.
- `TAG_ALIASES` maps form misspellings to real columns (`FiancialStatement` → 543 files,
  `MangerEndDate` → 102 files, `zzSubmittedFee`). An alias is only applied when the
  misspelling is not itself a template column and the correct column is still empty.

## Transformation order

Order matters; each step assumes the previous one ran.

1. Strip a trailing `.0` from `ApplicationID`.
2. Truncate ISO datetimes to date only, for any column whose value matches `ISO_DT`.
3. **Yes/No/N/A** across `YESNONA_COLUMNS` (55 columns). Override version consults
   `CUSTOM_MAPPINGS` first. `as_yesnona` also tolerates `1.0/2.0/3.0` and any casing of
   `yes/no/n/a/na`. Unrecognised values are recorded in `odd_values` and passed through.
4. Mirror `AppOffDirBankrupt` into `zzAppOffDirBankrupt`, counting overwrites.
5. Booleans (`SendEmail`, `ManagerDecision`) → `True` / `False`.
6. `CPICRequired`: `"no nit"` → `"No Hit"`.
7. Backfill from the folder and file name — `ApplicationID`, `ApplicationFolderID`,
   `ApplicationFolderName`, `BusinessName`, `ApplicationType`, `XMLFileName` — **only
   where the XML left the cell empty**.
8. Strip a trailing year from `ApplicationType`.
9. Set `Missing_ID`.

`folder_parts()` parses `12151 - Some Business Name - New - 2019` into number, business
name, type, year. Two passes: split on ` - ` with a four-digit tail, then the
`FOLDER_TAIL` regex fallback.

`clean()` removes zero-width spaces (`\u200b`) and converts non-breaking spaces to normal
spaces — both crashed PowerApps — collapses runs of spaces/tabs, then handles line breaks
(the one line that differs between the two scripts).

## The override version in detail

```python
CUSTOM_MAPPINGS = {
    "RiskProfileReport": {"1": "N/A", "2": "No", "3": "Yes"},
    "LARSubmitted":      {"1": "N/A", "2": "No", "3": "Yes"},
    "FinancialStatement":{"1": "Yes", "2": "No", "3": "N/A", "N/A": "Yes"},
    "FinStmtReceived":   {"1": "Yes", "2": "No", "3": "N/A", "N/A": "Yes"},
}
```

The lookup happens inside the `YESNONA_COLUMNS` loop:

```python
if c in CUSTOM_MAPPINGS and row[c] in CUSTOM_MAPPINGS[c]:
    new = CUSTOM_MAPPINGS[c][row[c]]
else:
    new = as_yesnona(row[c])
```

Three things to know about this design:

- **It is exact-match only.** A value of `"3.0"` does not match the key `"3"`, so it
  falls through to `as_yesnona` and becomes `N/A` — the opposite of the intended `Yes` for
  `RiskProfileReport` and `LARSubmitted`. `RAW_CHOICE` shows `1.0/2.0/3.0` do occur in
  this data. If the override matters, add the `.0` keys or strip the suffix before the
  lookup.
- **`FinStmtReceived` has no effect.** The loop iterates `YESNONA_COLUMNS`, and
  `FinStmtReceived` is not in that list. Add it there if the rule is wanted.
- **`FinancialStatement`'s 1/2/3 mapping is identical to the default.** Its only real
  effect is `N/A` → `Yes`.

The `<br>` change in `clean()` applies to **every** column, not only rich-text ones. A
line break in a plain-text destination field will show the literal characters `<br>`.

## XML repair

`parse_xml()` tries a clean `ET.parse` first. On `ParseError` it re-reads the bytes and:

1. Raises a clear error if the file is empty or is actually HTML.
2. Decodes as `utf-8-sig` with `errors="replace"`, strips characters illegal in XML 1.0,
   and cuts anything before the first `<`.
3. Tries again, then tries once more with bare `&` escaped to `&amp;`.
4. Re-raises the original error if all of that fails.

Repaired files are counted as `Repaired while reading`. Anything unrecoverable lands in
`failures` with its exception type and is **excluded from the CSV** — which is why it
appears in the BALANCE line.

`long_path()` / `show()` add and strip the Windows `\\?\` (and `\\?\UNC\`) prefix so paths
over 260 characters work, while printed paths stay readable.

## Dedupe

Key: `ApplicationID`. Rows with a blank ID are always kept. Where an ID repeats, the row
with the most recent `XML_Modified` survives (stable sort). Dropped rows are listed in the
summary with ID, filename, folder and timestamp.

## Output format

`utf-8-sig` (BOM, so Excel reads accents correctly), `QUOTE_ALL`, `\r\n`. Every field is a
quoted string; no type inference on the destination side.

## Safety guards

- Exits if the source folder does not exist.
- Exits if `OUTPUT_DIR` resolves inside `INPUT_DIRECTORY` — prevents writing into the
  shared drive.
- Exits if the template is missing or has no usable header.
- Walk errors are collected, not fatal.
- Per-file exceptions are caught individually, so one bad file cannot end the run.

## How to change things

| Task | Change |
|---|---|
| New template from the business | Update `TEMPLATE_PATH`. Nothing else. |
| A column still shows raw 1/2/3 | Add its exact template name to `YESNONA_COLUMNS`. |
| A column needs a flipped mapping | Add it to `CUSTOM_MAPPINGS` (override version) **and** make sure it is in `YESNONA_COLUMNS`. |
| The form misspells a field | Add `"misspelling": "RealColumn"` to `TAG_ALIASES`. |
| A template column must not be uploaded | Add it to `DROP_COLUMNS`, only after confirming it is blank in a full export. |
| Need date + time | `DATE_FMT = "%Y-%m-%d %H:%M:%S"`. |

Editing `YESNONA_COLUMNS`, `CUSTOM_MAPPINGS`, `DROP_COLUMNS` or `TAG_ALIASES` changes
every record. Confirm with the business owner first — these are business rules held as
plain lists precisely so they can be reviewed without reading code.

## Known issues and open items

- **`FinStmtReceived` override is inert.** See above.
- **`.0` values bypass `CUSTOM_MAPPINGS`.** See above.
- **`FSCertification`** is blank in every row of a full export but is deliberately *not*
  in `DROP_COLUMNS` — nobody has confirmed whether it should hold values. Check the
  `XML FIELD(S) HAVE NO COLUMN` warning and ask the business before deciding.
- **Leftover tags.** Fields present in the XML with no matching column are counted and
  reported, then dropped. Review that list against the mapping sheet periodically.
- **Duplicate checklists on the drive** are hidden from the upload, not resolved.

## File map

Both scripts follow the same four-part layout:

- **Part 1** — settings and business rules (the lists above).
- **Part 2** — helper functions (`long_path`, `show`, `key`, `clean`, `folder_parts`,
  `as_yesnona`, `as_boolean`).
- **Part 3** — reading the XML (`load_template`, `parse_xml`, `flatten`).
- **Part 4** — `run()`: the whole pipeline and the summary.

Comments marked `PYTHON NOTE` explain a language feature the first time it appears. Most
odd-looking code exists because of a real failure in this data, and the comment says
which — assume a line is load-bearing until you have checked it.

## Contacts

- Column mapping sheet and the Dataverse dataflow: **Rachelle**.
- Anything about which values a column should hold: the business owner, via the mapping
  sheet. The script only applies the rules it was given.
