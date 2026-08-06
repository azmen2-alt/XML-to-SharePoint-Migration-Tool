#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
Diagnose SharePoint "Some data couldn't be imported" row rejections.

SharePoint tells you the row numbers but not the column or the reason.
This finds them by profiling the rows that imported successfully and
reporting what is abnormal in the rows that failed.

It also writes a re-import-safe copy of the file: dates normalised,
over-length cells trimmed, control characters stripped, and person
columns optionally blanked.

Usage:
    python inspect_import_errors.py
    python inspect_import_errors.py "C:\path\sharepoint_upload_ready.csv"
"""

from __future__ import annotations

import os
import re
import sys
from collections import Counter

import pandas as pd

# =============================================================================
# CONFIGURATION
# =============================================================================
INPUT_FILE = r"C:\Users\BalfaqAh\Downloads\sharepoint_export\sharepoint_upload_ready.csv"
OUTPUT_DIR = ""          # "" = same folder as INPUT_FILE

# The row numbers SharePoint reported.
ERROR_ROWS = [2394, 2404, 2433, 2572, 2614, 2646, 2675, 2752, 2769, 2788, 2815]

SP_TEXT_LIMIT   = 255
BLANK_PEOPLE    = False  # True -> clear person columns in the safe copy
DATE_OUTPUT_FMT = "%Y-%m-%d"
# =============================================================================

# Columns SharePoint is likely to have typed as Person/Group.
PERSON_HINTS = re.compile(
    r"(assignedto|assignedby|createdby|clerk|officername|managername|"
    r"registrarname|sendto|username|currentuserfirst|currentuserlast|"
    r"tobeassigned|email)", re.I)

DATE_HINTS = re.compile(
    r"(date|expiry|due|receivedon|assignedon|createdon|modifiedon|"
    r"starttime|endtime)$|^(received|documentsreceived)$", re.I)

NUM_HINTS = re.compile(r"(id|number|days|fee|amount|count)$", re.I)

CTRL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
YESNO = {"yes", "no", "true", "false", "1", "0", ""}


def load(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        return pd.read_excel(path, dtype=str, engine="openpyxl").fillna("")
    return pd.read_csv(path, dtype=str, encoding="utf-8-sig",
                       keep_default_na=False, low_memory=False).fillna("")


def is_num(v: str) -> bool:
    return bool(re.fullmatch(r"[-+]?\d+(\.\d+)?", v.strip().replace(",", "")))


def as_date(v: str):
    v = v.strip()
    if not v or len(v) > 40:
        return None
    # A bare integer is not a date. pandas happily reads "3600" as the year
    # 3600, which makes numeric columns look like date columns.
    if v.isdigit() and len(v) != 8:
        return None
    for dayfirst in (False, True):
        try:
            d = pd.to_datetime(v, errors="raise", dayfirst=dayfirst)
            if pd.notna(d):
                return d
        except Exception:                                        # noqa: BLE001
            continue
    return None


def profile(df: pd.DataFrame, mask):
    """Per-column statistics computed from the rows that imported fine."""
    prof = {}
    good = df[mask]
    for c in df.columns:
        s = good[c].astype(str)
        nonempty = s[s.str.strip() != ""]
        n = len(nonempty)
        prof[c] = {
            "n": n,
            "fill_rate": n / max(len(s), 1),
            "num_rate": (sum(is_num(v) for v in nonempty) / n) if n else 0.0,
            "date_rate": (sum(as_date(v) is not None for v in nonempty.head(400))
                          / min(n, 400)) if n else 0.0,
            "max_len": int(nonempty.str.len().max()) if n else 0,
            "vals": Counter(nonempty.str.strip().str.lower()),
            "person": bool(PERSON_HINTS.search(c)),
            "datey": bool(DATE_HINTS.search(c)),
            "numy": bool(NUM_HINTS.search(c)),
        }
    return prof


def score_row(row, prof, columns):
    """Return a list of (severity, column, value, reason)."""
    out = []
    for c in columns:
        v = str(row[c]).strip()
        if not v:
            continue
        p = prof[c]

        # Only a cause if good rows do NOT also exceed the limit; if they do,
        # the column is already multi-line and length is not what failed.
        if len(v) > SP_TEXT_LIMIT and p["max_len"] <= SP_TEXT_LIMIT:
            out.append((3, c, v, f"{len(v)} chars, over the {SP_TEXT_LIMIT} "
                                 "limit; good rows max out at "
                                 f"{p['max_len']}"))
        if CTRL.search(v):
            out.append((3, c, v, "contains control characters"))

        # a person column holding a value seen nowhere else
        if p["person"]:
            seen = p["vals"].get(v.lower(), 0)
            if seen == 0:
                out.append((3, c, v, "PERSON-type column: this value appears in "
                                     "no successfully imported row - likely not "
                                     "resolvable in the address book"))
            elif seen <= 2:
                out.append((2, c, v, f"person column, rare value (seen {seen}x)"))

        # a date column holding something unparseable
        if as_date(v) is None and (p["date_rate"] >= 0.9
                                   or (p["datey"] and p["date_rate"] >= 0.5)):
            out.append((3, c, v, f"date column ({100*p['date_rate']:.0f}% of good "
                                 "rows parse as dates): this will not parse"))

        # a numeric column holding text
        if not is_num(v) and (p["num_rate"] >= 0.95
                              or (p["numy"] and p["num_rate"] >= 0.9)):
            out.append((3, c, v, f"numeric column ({100*p['num_rate']:.0f}% of good "
                                 "rows are numbers): this is not a number"))

        # a yes/no column holding something else
        low = {k for k in p["vals"] if k}
        if low and low <= YESNO and v.lower() not in YESNO:
            out.append((2, c, v, "column otherwise holds only Yes/No"))

        # a column that is almost always empty suddenly has content
        if 0 < p["fill_rate"] < 0.02 and p["n"] < 20:
            out.append((1, c, v, f"column is empty in {100*(1-p['fill_rate']):.0f}% "
                                 "of good rows; type may have been guessed wrong"))
    # One finding per column: keep the most severe reason.
    best: dict[str, tuple] = {}
    for sev, c, v, why in out:
        if c not in best or sev > best[c][0]:
            best[c] = (sev, c, v, why)
    return sorted(best.values(), key=lambda x: -x[0])


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else INPUT_FILE
    if not os.path.exists(path):
        sys.exit(f"ERROR: file not found:\n  {path}")
    out_dir = OUTPUT_DIR or os.path.dirname(os.path.abspath(path))
    os.makedirs(out_dir, exist_ok=True)

    df = load(path)
    n = len(df)
    print(f"Loaded {n} rows x {len(df.columns)} columns from\n  {path}\n")

    # SharePoint may or may not count the header row. Try both alignments
    # and keep whichever surfaces more hard problems.
    best = None
    for label, offset in (("data row N (1 = first record)", 1),
                          ("spreadsheet row N (1 = header)", 2)):
        idx = [r - offset for r in ERROR_ROWS if 0 <= r - offset < n]
        if len(idx) != len(ERROR_ROWS):
            continue
        mask = pd.Series(True, index=df.index)
        mask.iloc[idx] = False
        prof = profile(df, mask)
        findings = {r: score_row(df.iloc[r - offset], prof, df.columns)
                    for r in ERROR_ROWS}
        hard = sum(1 for f in findings.values() if any(s >= 3 for s, *_ in f))
        if best is None or hard > best[0]:
            best = (hard, label, offset, idx, prof, findings)

    if best is None:
        sys.exit("ERROR: the reported row numbers fall outside this file. "
                 "Are you pointing at the file you actually imported?")

    hard, label, offset, idx, prof, findings = best
    print(f"Row alignment: {label}  "
          f"({hard} of {len(ERROR_ROWS)} rows show a hard type conflict)\n")

    report = os.path.join(out_dir, "import_error_diagnosis.txt")
    with open(report, "w", encoding="utf-8") as f:
        w = f.write
        w("=" * 66 + "\n  SHAREPOINT IMPORT ERROR DIAGNOSIS\n" + "=" * 66 + "\n")
        w(f"File:      {path}\nRows:      {n}\n")
        w(f"Alignment: {label}\n")
        w(f"Rejected:  {', '.join(map(str, ERROR_ROWS))}\n\n")

        blame = Counter()
        for r in ERROR_ROWS:
            f_list = findings[r]
            key = df.iloc[r - offset].get("Missing_ID", "")
            w("-" * 66 + f"\nROW {r}   Missing_ID = {key}\n" + "-" * 66 + "\n")
            if not f_list:
                w("  No anomaly detected. Check this row by hand.\n\n")
                continue
            for sev, c, v, why in f_list[:8]:
                tag = {3: "LIKELY CAUSE", 2: "possible", 1: "note"}[sev]
                shown = v if len(v) <= 120 else v[:117] + "..."
                w(f"  [{tag}] {c}\n      value : {shown}\n      why   : {why}\n")
                if sev >= 3:
                    blame[c] += 1
            w("\n")

        w("=" * 66 + "\n  COLUMNS MOST OFTEN IMPLICATED\n" + "=" * 66 + "\n")
        for c, cnt in blame.most_common():
            p = prof[c]
            kind = ("person" if p["person"] else
                    "date" if p["date_rate"] >= 0.5 else
                    "numeric" if p["num_rate"] >= 0.9 else "text")
            w(f"  {c:<28}{cnt} rows   (looks like a {kind} column)\n")
        w("\nFIX OPTIONS\n")
        w("  1. In the SharePoint list settings, change the implicated columns\n")
        w("     to 'Single line of text' or 'Multiple lines of text'. Type\n")
        w("     validation then cannot reject anything. Convert later if needed.\n")
        w("  2. Or add the missing staff to the address book / correct the\n")
        w("     spelling so Person columns resolve.\n")
        w("  3. Or import sharepoint_upload_safe.csv (written alongside this\n")
        w("     report) and add the 11 rows in import_error_rows.csv by hand.\n")

    # ---- the rejected rows on their own, for manual entry
    rows_path = os.path.join(out_dir, "import_error_rows.csv")
    sub = df.iloc[idx].copy()
    sub.insert(0, "SharePoint_Row", ERROR_ROWS)
    sub.to_csv(rows_path, index=False, encoding="utf-8-sig")

    # ---- a safe re-import copy
    safe = df.copy()
    for c in safe.columns:
        p = prof[c]
        s = safe[c].astype(str).map(lambda v: CTRL.sub("", v))
        if p["datey"] or p["date_rate"] > 0.8:
            def fix_date(v):
                d = as_date(v)
                return d.strftime(DATE_OUTPUT_FMT) if d is not None else ""
            s = s.map(fix_date)
        if BLANK_PEOPLE and p["person"]:
            s = ""
        else:
            s = s.map(lambda v: v if len(v) <= SP_TEXT_LIMIT else v[:SP_TEXT_LIMIT])
        safe[c] = s
    safe_path = os.path.join(out_dir, "sharepoint_upload_safe.csv")
    safe.to_csv(safe_path, index=False, encoding="utf-8-sig", lineterminator="\r\n")

    print("Columns most often implicated:")
    for c, cnt in Counter(
            c for r in ERROR_ROWS for s, c, *_ in findings[r] if s >= 3
    ).most_common(10):
        print(f"  {c:<28}{cnt} rows")
    print(f"\nDiagnosis:   {report}")
    print(f"Bad rows:    {rows_path}")
    print(f"Safe re-import: {safe_path}")


if __name__ == "__main__":
    main()
