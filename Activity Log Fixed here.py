#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
LICENSING APPLICATIONS -> SHAREPOINT EXPORT

Reads the checklist XML in every application folder and writes one row per
application. Only reads the shared drive; never changes it.

RUN IT
    python xml_to_sharepoint_v19.py

    The template file must be defined in the SETTINGS section below. That is 
    where the column names come from. New template from the business? Update
    the path or drop the new file in and rerun.

IT WRITES
    sharepoint_upload_ready.csv      the deliverable, and the only file

    Everything else prints to the screen. Read the summary before you
    upload - the BALANCE line and any "!!" warnings are the checks that
    tell you whether the file is safe. To keep a copy of a run:
        python xml_to_sharepoint_v19.py > run_log.txt

FILE MAP
    Part 1  settings and business rules
    Part 2  helper functions
    Part 3  reading the XML
    Part 4  the main run

Comments marked "PYTHON NOTE" explain a language feature the first time it
appears. Most odd-looking code exists because of a real failure in this data;
the comment says which. Assume a line is load-bearing until you check it.

TWO KINDS OF NOTE IN THIS FILE
    A line starting with #      is a comment. Python ignores the whole line.
    Text wrapped in three       is a "docstring". Also not executed. It sits
    double-quotes               just under a def as its description, so
                                editors can show it as help text.

    The block you are reading now is a docstring - it opens with r and three
    double-quotes above, and closes the same way below. Nothing in it runs.
    Deleting a comment or docstring never changes what the script does; it
    only deletes the explanation.
"""
# ---------------------------------------------------------------------------
# IMPORTS - standard libraries, so this behaviour need not be rewritten.
#
# PYTHON NOTE: "import x" loads a toolbox. "import x as y" gives it a short
# nickname. "from x import y" takes one tool out of the box. These lines run
# first and must stay at the top.
#
#   csv          quoting rules for writing the output file
#   os           folders, file paths, file dates
#   re           text pattern matching
#   sys          command-line arguments, and stopping with an error
#   unicodedata  evens out accented characters before comparing names
#   ET           the XML reader (nicknamed, its real name is long)
#   Counter      counts how often something occurs
#   defaultdict  a dictionary that starts a new list on first use
#   datetime     dates and timestamps
#   pd           pandas: builds the table and writes the CSV (nicknamed)
#
# The "__future__" line lets newer type-hint syntax run on older Python
# versions. It does nothing else. Leave it alone.
# ---------------------------------------------------------------------------
from __future__ import annotations
import csv, os, re, sys, unicodedata
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from datetime import datetime
import pandas as pd

# ###########################################################################
#  PART 1 - SETTINGS AND BUSINESS RULES
# ###########################################################################
# The only three lines you normally change.
#
# PYTHON NOTE: the r before the quotes means "treat backslashes literally".
# Windows paths need it. Keep the r and both quotes; change only the text
# between them.
# ===========================================================================
INPUT_DIRECTORY = r"C:\Users\BalfaqAh\Government of Ontario\Licensing Unit - Licensing Applications"
OUTPUT_DIR      = r"C:\Users\BalfaqAh\Downloads\sharepoint_export"
TEMPLATE_PATH   = r"C:\Users\BalfaqAh\Downloads\template.xlsx"
# ===========================================================================

# ID_COLUMN      the extra column, e.g. "12151-Checklist"
# DATE_FMT       %Y year, %m month, %d day. Add " %H:%M:%S" for the time.
# SP_TEXT_LIMIT  SharePoint's single-line-text ceiling.
ID_COLUMN, DATE_FMT, SP_TEXT_LIMIT = "Missing_ID", "%Y-%m-%d", 255
LEGACY_COLUMNS = ["ApplicationFolderID", "ApplicationFolderName",
                  "XMLFileName"]                          # template dropped

# Columns present in the template that are never populated and must not be
# uploaded. Both were confirmed blank in every row of a full export.
#   group, test  leftovers from an earlier template revision
# FSCertification is ALSO blank in every row, but is deliberately NOT
# dropped: it is unconfirmed whether the field should hold values. Consult
# the business, and the "XML FIELD(S) HAVE NO COLUMN" warning in the run
# summary, before adding it above.
DROP_COLUMNS = ["group", "test"]
TS_COLUMNS = ["XML_Created", "XML_Modified"]              # the file's own dates

# ---------------------------------------------------------------------------
# BUSINESS RULES AS DATA
# From the business owner's mapping sheet. Kept as plain lists so they can be
# checked without reading code. Editing this section changes every record, so
# it is a business decision - confirm before changing.
# ---------------------------------------------------------------------------

# 55 columns where the form stores 1/2/3 and the new system wants Yes/No/N/A.
# Dataverse rejected 25 rows over a raw "2" here, so this list must be complete.
YESNONA_COLUMNS = [
    "ApplicationReceived", "ReferenceLetters2", "OneThousandBondProvided",
    "FiveThousandBondProvided", "EvidenceofExam", "SupervisionLetReceived",
    "ONBISComplete", "ONBISCorpPage", "ONBISDirector", "ONBISNameReg",
    "FederalRegistrationComplete", "LARSubmitted", "TermsConditions",
    "CriminalBckCheckComplete", "ConsumerReportConducted",
    "RiskProfileReport", "OustandingAP", "OperateInOntario",
    "StreetViewReview", "OfficeInOntario", "BusinessNameMatch",
    "DirectorOfficerConsistent", "ComplyWith24-3", "TrustAccountConsistent",
    "TrustAcctDesignation", "FinancialStatement", "TrustAccountConfirm",
    "FiveThousandBond", "BankNameCorrect", "OneThousandBond",
    "TwoYrsExperience", "AppOffDirOver18", "PassedExam", "ReferenceLetters",
    "AppOffDirInOntario", "DisclosureIssues", "AppCanadianCitizen",
    "SupervisionLetter", "BusinessPlan", "DebtorLetters", "MandatoryStmt",
    "SampleContract", "PaytoFund", "UpdatedCATS", "UpdatedICPS",
    "LicensesPrinted", "LicensedMailed", "RefusalLetter", "UpdateCATS",
    "UpdateICPS", "RefundApplication", "zzAppOffDirBankrupt",
    "AppOffDirBankrupt", "AppOffDirConvicted", "SubmittedFee",
]

# Form fields whose spelling matches no column. Never fires if the source is
# itself a template column.
# The form misspells some of its own fields. Maps the misspelling to the
# right column. Never fires if the misspelling is itself a real column.
TAG_ALIASES = {"FiancialStatement": "FinancialStatement",     # 543 files
               "MangerEndDate": "ManagerEndDate",             # 102 files
               "zzSubmittedFee": "SubmittedFee"}
BOOLEAN_COLUMNS = ["SendEmail", "ManagerDecision"]

YESNONA_MAP = {"1": "Yes", "2": "No", "3": "N/A"}
YESNONA_OK = {"yes": "Yes", "no": "No", "n/a": "N/A", "na": "N/A"}
RAW_CHOICE = set(YESNONA_MAP) | {"1.0", "2.0", "3.0"}

# Formatting tags. Used to tell rich text from a wrapper element (see flatten).
# PYTHON NOTE: { } with no colons is a SET - a bag for fast "is it in here?".
HTML_NOISE = {"div", "span", "b", "i", "u", "strong", "em", "p", "br", "a",
              "font", "ul", "ol", "li", "table", "tr", "td", "th", "tbody",
              "html", "head", "body", "meta", "title", "svg", "g", "path",
              "defs", "style", "rect", "circle"}

# ---------------------------------------------------------------------------
# FIVE TEXT PATTERNS ("regular expressions")
# ---------------------------------------------------------------------------
ILLEGAL_XML = re.compile("[^\u0009\u000A\u000D\u0020-\uD7FF\uE000-\uFFFD"
                         "\U00010000-\U0010FFFF]")
BARE_AMP = re.compile(r"&(?!(?:[A-Za-z][A-Za-z0-9]*|#[0-9]+|#[xX][0-9A-Fa-f]+);)")
FOLDER_NUM = re.compile(r"^\s*(\d+)\s+-\s")
FOLDER_TAIL = re.compile(r"^(.*?)\s*-\s*([^-]*?)\s*-\s*((?:19|20)\d{2})\s*$")
FILE_NUM = re.compile(r"^\s*(\d+)-")
ISO_DT = re.compile(r"^(\d{4}-\d{2}-\d{2})T\d{2}:\d{2}:\d{2}"
                    r"(?:\.\d+)?(?:Z|[+-]\d{2}:?\d{2})?$")
TRAILING_YEAR = re.compile(r"[\s\-]*(?:19|20)\d{2}\s*$")


# ###########################################################################
#  PART 2 - HELPER FUNCTIONS
# ###########################################################################

def long_path(p):
    if os.name != "nt": return p
    p = os.path.abspath(p)
    if p.startswith("\\\\?\\"): return p
    return ("\\\\?\\UNC\\" + p.lstrip("\\")) if p.startswith("\\\\") else "\\\\?\\" + p

def show(p):
    if p.startswith("\\\\?\\UNC\\"): return "\\\\" + p[8:]
    return p[4:] if p.startswith("\\\\?\\") else p

def key(s):
    return re.sub(r"[^0-9a-zA-Z]+", "", unicodedata.normalize("NFKD", str(s))).lower()

def clean(v):
    # Remove hidden zero-width spaces and non-breaking spaces that crash PowerApps
    v = v.replace('\u200b', '').replace('\xa0', ' ')
    
    # Keep standard newlines for plain text fields instead of HTML tags
    return re.sub(r"[\r\n]+", "\n", re.sub(r"[ \t]{2,}", " ", v)).strip()

def folder_parts(name):
    m = FOLDER_NUM.match(name)
    num = m.group(1) if m else ""
    rest = name[len(m.group(0)):].strip() if m else name.strip()

    parts = [x.strip() for x in re.split(r"\s+-\s+", rest) if x.strip()]
    if len(parts) >= 2 and re.fullmatch(r"(19|20)\d{2}", parts[-1]):
        year = parts.pop()
        typ = parts.pop() if len(parts) > 1 else ""
        return num, " - ".join(parts), typ, year

    t = FOLDER_TAIL.match(rest)
    if t:
        return num, t.group(1).strip(), t.group(2).strip(), t.group(3)

    return num, rest, "", ""

def as_yesnona(v):
    n = v[:-2] if v.endswith(".0") else v
    if n in YESNONA_MAP: return YESNONA_MAP[n]
    return YESNONA_OK.get(v.lower(), v)

def as_boolean(v):
    low = v.lower()
    if low in ("true", "yes", "1"): return "True"
    return "False" if low in ("false", "no", "0") else v


# ###########################################################################
#  PART 3 - READING THE XML
# ###########################################################################

def load_template():
    if not os.path.exists(TEMPLATE_PATH):
        sys.exit(f"ERROR: no template file found at {TEMPLATE_PATH}.")

    head = (pd.read_excel(TEMPLATE_PATH, nrows=0, engine="openpyxl")
            if TEMPLATE_PATH.lower().endswith((".xlsx", ".xlsm", ".xltx"))
            else pd.read_csv(TEMPLATE_PATH, nrows=0, encoding="utf-8-sig"))
            
    cols, seen = [], set()
    for c in head.columns:
        c = str(c).strip()
        if c and not c.lower().startswith("unnamed:") and c not in seen:
            seen.add(c); cols.append(c)
            
    if cols: 
        return cols, f"Template ({len(cols)} columns)"
        
    sys.exit(f"ERROR: {TEMPLATE_PATH} has no usable header row.")

def parse_xml(path):
    p = long_path(path)
    try:
        return ET.parse(p).getroot(), False
    except ET.ParseError as first:
        raw = open(p, "rb").read()
        if not raw.strip(): raise ET.ParseError("file is empty (0 bytes)")
        head = raw[:400].lstrip().lower()
        if head.startswith(b"<!doctype html") or head.startswith(b"<html"):
            raise ET.ParseError("file is HTML, not XML")
        text = ILLEGAL_XML.sub("", raw.decode("utf-8-sig", errors="replace"))
        cut = text.find("<")
        if cut > 0: text = text[cut:]
        for candidate in (text, BARE_AMP.sub("&amp;", text)):
            try: return ET.fromstring(candidate), True
            except ET.ParseError: continue
        raise first

def flatten(root, column_keys):
    vals, attrs, leaves = defaultdict(list), defaultdict(list), set()
    for el in root.iter():
        tag = el.tag.split("}", 1)[-1] if isinstance(el.tag, str) else None
        if not tag: continue
        structural = (key(tag) not in column_keys
                      and any(isinstance(c.tag, str)
                              and c.tag.split("}", 1)[-1].lower() not in HTML_NOISE
                              for c in el))
        text = ((el.text or "") if structural
                else " ".join(t for t in el.itertext() if t and t.strip()))
        text = re.sub(r"[ \t]+", " ", text).strip()
        if text: vals[tag].append(text)
        if len(el) == 0: leaves.add(tag)
        for k, v in el.attrib.items():
            v = str(v).strip()
            if v: attrs[k.split("}", 1)[-1]].append(v)
    lut = {}
    for src in (attrs, vals):
        for k, vs in src.items(): lut[key(k)] = "; ".join(dict.fromkeys(vs))
    for wrong, right in TAG_ALIASES.items():
        if (key(wrong) not in column_keys and lut.get(key(wrong))
                and not lut.get(key(right))):
            lut[key(right)] = lut[key(wrong)]
    return lut, vals, leaves


# ###########################################################################
#  PART 4 - THE MAIN RUN
# ###########################################################################

def run():
    started = datetime.now()
    src = sys.argv[1] if len(sys.argv) > 1 else INPUT_DIRECTORY
    out = sys.argv[2] if len(sys.argv) > 2 else OUTPUT_DIR
    if not os.path.isdir(src): sys.exit(f"ERROR: no such folder:\n  {src}")
    if os.path.abspath(out).startswith(os.path.abspath(src)):
        sys.exit("ERROR: OUTPUT_DIR sits inside the folder being scanned.")
    os.makedirs(out, exist_ok=True)

    template_columns, template_source = load_template()
    print(f"Columns from {template_source}")
    dropped = [c for c in template_columns if c in DROP_COLUMNS]
    if dropped:
        print(f"  Dropping unused column(s): {', '.join(dropped)}")
    columns = ([c for c in template_columns if c not in DROP_COLUMNS]
               + [c for c in LEGACY_COLUMNS if c not in template_columns])
    out_columns = columns + [ID_COLUMN] + TS_COLUMNS
    ckeys = {c: key(c) for c in columns}
    column_keys = set(ckeys.values())
    alias_keys = {key(k) for k in TAG_ALIASES}

    rule_columns = (set(YESNONA_COLUMNS) | set(BOOLEAN_COLUMNS)
                    | {"CPICRequired", "AppOffDirBankrupt",
                       "zzAppOffDirBankrupt", "ApplicationID",
                       "ApplicationType", "BusinessName"})
    rules_off = sorted(c for c in rule_columns if c not in set(columns))
    if rules_off:
        print(f"  !! {len(rules_off)} rule(s) will NOT run, column absent from "
              f"template: {', '.join(rules_off)}")

    root_abs = long_path(src)
    xmls, others, walk_errors, dirs_seen = [], [], [], []
    exts, per_folder, top_dirs, loose = Counter(), Counter(), set(), 0
    for r, ds, fs in os.walk(root_abs, onerror=walk_errors.append):
        dirs_seen.append(r)
        if os.path.abspath(r) == os.path.abspath(root_abs): top_dirs.update(ds)
        for f in fs:
            full = os.path.join(r, f)
            ext = os.path.splitext(f)[1].lower().strip()
            exts[ext or "(no extension)"] += 1
            rel = os.path.relpath(full, root_abs).split(os.sep)
            top = rel[0] if len(rel) > 1 else ""
            if not top: loose += 1
            if ext == ".xml": xmls.append(full); per_folder[top] += 1
            else: others.append(full)
    print(f"Scanning {src}\n  folders {len(dirs_seen)}   "
          f"files {len(xmls) + len(others)}   xml {len(xmls)}")

    rows, failures, leftover_tags = [], [], Counter()
    repaired = blank_rows = 0
    stats, odd_values = Counter(), defaultdict(set)

    for n, path in enumerate(xmls, 1):
        base = os.path.basename(path)
        stem = os.path.splitext(base)[0]
        rel = os.path.relpath(path, root_abs).split(os.sep)
        app_folder = rel[0] if len(rel) > 1 else ""
        f_num, f_biz, f_type, _ = folder_parts(app_folder)
        m = FILE_NUM.match(stem)
        stem_num = m.group(1) if m else ""
        try:
            xroot, was_repaired = parse_xml(path)
            lut, elements, leaves = flatten(xroot, column_keys)
            repaired += was_repaired
            
            row = {c: clean(lut.get(ckeys[c], "")) for c in columns}

            if row.get("ApplicationID", "").endswith(".0"):
                row["ApplicationID"] = row["ApplicationID"][:-2]

            for c in columns:
                if "T" in row[c]:
                    mm = ISO_DT.match(row[c])
                    if mm: row[c] = mm.group(1); stats["dates truncated"] += 1

            for c in YESNONA_COLUMNS:
                if row.get(c):
                    new = as_yesnona(row[c])
                    if new != row[c]: stats["1/2/3 converted"] += 1
                    elif new not in ("Yes", "No", "N/A"): odd_values[c].add(new)
                    row[c] = new

            if row.get("AppOffDirBankrupt"):
                if row.get("zzAppOffDirBankrupt") not in ("", row["AppOffDirBankrupt"]):
                    stats["zzAppOffDirBankrupt overwritten"] += 1
                row["zzAppOffDirBankrupt"] = row["AppOffDirBankrupt"]

            for c in BOOLEAN_COLUMNS:
                if row.get(c):
                    new = as_boolean(row[c])
                    if new not in ("True", "False"): odd_values[c].add(new)
                    row[c] = new

            if row.get("CPICRequired", "").lower() == "no nit":
                row["CPICRequired"] = "No Hit"; stats["CPICRequired typo fixed"] += 1

            for col, val in (("ApplicationID", stem_num or f_num),
                             ("ApplicationFolderID", f_num),
                             ("ApplicationFolderName", app_folder),
                             ("BusinessName", f_biz), ("ApplicationType", f_type),
                             ("XMLFileName", base)):
                if col in row and not row[col] and val: row[col] = val

            if row.get("ApplicationType"):
                new = TRAILING_YEAR.sub("", row["ApplicationType"]).strip(" -")
                if new != row["ApplicationType"]: stats["ApplicationType year stripped"] += 1
                row["ApplicationType"] = new

            row[ID_COLUMN] = f"{f_num}-{stem}" if not stem_num and f_num else stem

            blank_rows += not any(row[c] for c in columns)

            extra = {t: clean("; ".join(dict.fromkeys(v)))
                     for t, v in elements.items()
                     if t in leaves and t.lower() not in HTML_NOISE
                     and key(t) not in column_keys and key(t) not in alias_keys}
            for t in extra: leftover_tags[t] += 1

            try:
                st = os.stat(long_path(path))
                born = getattr(st, "st_birthtime", None) or st.st_ctime
                row["XML_Created"] = datetime.fromtimestamp(born).strftime(DATE_FMT)
                row["XML_Modified"] = datetime.fromtimestamp(st.st_mtime).strftime(DATE_FMT)
            except (OSError, OverflowError, ValueError):
                row["XML_Created"] = row["XML_Modified"] = ""
            rows.append(row)
        except Exception as e:                                   # noqa: BLE001
            failures.append((show(path), type(e).__name__, str(e)))

    df = pd.DataFrame(rows, columns=out_columns).fillna("").astype(str)
    if len(df):
        df = (df.assign(_n=pd.to_numeric(
                  df["ApplicationFolderID"].str.extract(r"^(\d+)", expand=False),
                  errors="coerce"))
                .sort_values(["_n", "ApplicationFolderID", ID_COLUMN],
                             na_position="last")
                .drop(columns="_n").reset_index(drop=True))

    DEDUPE_KEY = "ApplicationID"
    dupes, dropped_dupes = pd.Series(dtype=int), []
    if len(df) and DEDUPE_KEY in df.columns:
        real = df[df[DEDUPE_KEY] != ""]
        counts = real[DEDUPE_KEY].value_counts()
        dupes = counts[counts > 1]
        if len(dupes):
            keep = set(real.sort_values("XML_Modified", ascending=False,
                                        kind="stable")
                           .drop_duplicates(subset=[DEDUPE_KEY], keep="first")
                           .index) | set(df.index[df[DEDUPE_KEY] == ""])
            dropped_dupes = [(df.at[i, DEDUPE_KEY], df.at[i, "XMLFileName"],
                              df.at[i, "ApplicationFolderName"],
                              df.at[i, "XML_Modified"])
                             for i in df.index if i not in keep]
            df = df[df.index.isin(keep)].reset_index(drop=True)

    csv_path = os.path.join(out, "sharepoint_upload_ready.csv")
    df.to_csv(csv_path, index=False, encoding="utf-8-sig",
              quoting=csv.QUOTE_ALL, lineterminator="\r\n")

    cov = pd.DataFrame([
        dict(zip(("App_Folder", "App_Number", "BusinessName", "AppType", "Year"),
                 (d,) + folder_parts(d)),
             XML_Count=per_folder.get(d, 0),
             Status=("OK" if per_folder.get(d, 0) == 1 else
                     "NO XML FOUND" if per_folder.get(d, 0) == 0 else "MULTIPLE XML"))
        for d in sorted(top_dirs)]).sort_values(["XML_Count", "App_Folder"]) \
        if top_dirs else pd.DataFrame(columns=["App_Folder", "XML_Count", "Status"])

    still_choice = [(c, int((df[c] != "").sum())) for c in df.columns
                    if len(df) and (df[c] != "").any()
                    and set(df[c][df[c] != ""]) <= RAW_CHOICE]
    long_cols = [(c, int(df[c].str.len().max())) for c in df.columns
                 if len(df) and int(df[c].str.len().max() or 0) > SP_TEXT_LIMIT]
    no_xml = int((cov.XML_Count == 0).sum()) if len(cov) else 0
    multi = int((cov.XML_Count > 1).sum()) if len(cov) else 0

    L = "=" * 62
    def w(*ls):
        for l in ls: print(l)

    w("", L, "        EXPORT SUMMARY", L,
      f"Source   {src}", f"Columns  {template_source}")
    for label, val in [
            ("Explorer shows at top level", len(top_dirs) + loose),
            ("Application folders", len(top_dirs)),
            ("Loose files at top level", loose),
            ("Folders with NO checklist", no_xml),
            ("Rows exported", len(df)),
            ("Rows dropped as duplicates", len(dropped_dupes)),
            ("Failed files", len(failures)),
            ("Repaired while reading", repaired),
            ("Rows that mapped no fields", blank_rows),
            ("BALANCE", f"{len(df)}+{len(dropped_dupes)}+{len(failures)}="
                        f"{len(df) + len(dropped_dupes) + len(failures)}"
                        f" of {len(xmls)} xml  (these must be equal)")]:
        print(f"  {label:<32}{val}")

    if rules_off:
        w("", f"  !! {len(rules_off)} TRANSFORMATION RULE(S) DISABLED - column absent",
          "  !! from this template, so the conversion cannot run:",
          *[f"       {c}" for c in rules_off])
    if still_choice:
        w("", "  !! DO NOT UPLOAD. These columns still hold raw 1/2/3, and a",
          "  !! SharePoint Choice column will REJECT those rows. Add each one",
          "  !! to YESNONA_COLUMNS and run again:",
          *[f"       {c:<30}{n} cells" for c, n in still_choice])
    else:
        print(f"  {'Choice-field check':<32}PASSED")
    if failures:
        w("", f"  !! {len(failures)} FILE(S) COULD NOT BE READ - not in the CSV:",
          *[f"       [{k}] {pth}" for pth, k, _ in failures[:20]])

    if long_cols:
        w("", f"  Create these as 'Multiple lines of text' (over {SP_TEXT_LIMIT}):",
          *[f"       {c:<30}max {mx}"
            for c, mx in sorted(long_cols, key=lambda x: -x[1])])
    if dropped_dupes:
        w("", f"  {len(dropped_dupes)} duplicate row(s) dropped, newest file kept.",
          "  Two files for one application means a stray copy on the drive.",
          "  This hides it from the upload; it does not tidy the drive:",
          *[f"       {mid:<10}{fn:<32}in {fld}"
            for mid, fn, fld, _ in dropped_dupes[:20]])
    if leftover_tags:
        w("", f"  {len(leftover_tags)} XML FIELD(S) HAVE NO COLUMN - present in the",
          "  files, dropped from the CSV. Add a column if any are needed:",
          *[f"       {t:<38}{c} files"
            for t, c in leftover_tags.most_common(25)])
    if odd_values:
        w("", "  Values left as-is because they were not recognised:",
          *[f"       {c}: {sorted(v)[:8]}" for c, v in odd_values.items()])
    if no_xml:
        w("", f"  {no_xml} folder(s) contain no checklist - missing paperwork on",
          "  the drive, not a fault in the tool:",
          *[f"       {d}" for d in cov.loc[cov.XML_Count == 0, "App_Folder"][:20]])
    if walk_errors:
        w("", f"  {len(walk_errors)} UNREADABLE FOLDER(S):",
          *[f"       {show(getattr(e, 'filename', '?') or '?')}: {e}"
            for e in walk_errors])

    w("", "  Row counts rising between runs is normal - new applications arrive",
      "  continuously. Never force the total to match an older file; that",
      "  leaves real applications out of the system.",
      L, f"CSV   {csv_path}", L)

if __name__ == "__main__":
    run()