#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
XML -> SharePoint import builder  (v4)

Usage:
    python xml_to_sharepoint_v4.py
    python xml_to_sharepoint_v4.py "C:\source\folder" "C:\output\folder"

New in v4, driven by the v3 run of the real data set:
  * FOLDER COVERAGE. The tree has ~3,192 application folders but only
    2,875 XML files. v4 writes folder_coverage.csv listing every
    application folder and how many checklists it holds, so the folders
    with none can be chased down. This is the real source of the gap.
  * ID RECOVERY. The XML has no ApplicationID tag. The application
    number lives in the folder name ("13203 - My Money Tree Inc -
    Renewal - 2022") and the file name ("31021-Checklist.xml"). v4
    parses both and fills ApplicationID / ApplicationFolderID.
  * TAG ALIASES. The InfoPath form has misspelled fields. "FiancialStatement"
    (543 files) is really FinancialStatement; "AppOffDirBankrupt" is
    zzAppOffDirBankrupt; "Comments" is Comment. These are now mapped.
  * UNIQUE IDs. 176 files named "Master Checklist" and 87 named
    "checklist" collapsed into 4 duplicate IDs. Missing_ID is now
    prefixed with the folder's application number when the file name
    has no number of its own.
  * UNMAPPED FIELDS. The 30 XML tags with no template column are written
    to unmapped_fields.csv instead of being discarded.
  * FASTER. The v3 run took 22 minutes, mostly sniffing 20,447 non-XML
    files. Known binary/media extensions are now skipped.
"""

from __future__ import annotations

import csv
import os
import re
import sys
import unicodedata
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from datetime import datetime

import pandas as pd

# =============================================================================
# CONFIGURATION
# =============================================================================
INPUT_DIRECTORY = r"C:\Users\BalfaqAh\Government of Ontario\Licensing Unit - Licensing Applications"
OUTPUT_DIR      = r"C:\Users\BalfaqAh\Downloads\sharepoint_export"
TEMPLATE_FILE   = ""          # optional; "" uses the embedded 109 columns

ID_COLUMN            = "Missing_ID"
APPLY_VALUE_RULES    = True   # Rachelle's Aug-4 rules: booleans, Yes/No/N-A,
                              # person-name allowlist, RegistrarName restriction,
                              # AppOffDirBankrupt copy, CPIC typo, date truncation
WRITE_EXCEL          = True
FLATTEN_NEWLINES     = True
NEWLINE_TOKEN        = " | "
NORMALIZE_DUP_IDS    = True   # prefix folder app-number onto unnumbered names
FILL_IDS_FROM_NAMES  = True   # recover ApplicationID from folder / file name
APPLY_TAG_ALIASES    = True   # map the misspelled InfoPath tags
RETAIN_LEGACY_COLUMNS = True  # keep ApplicationFolderID + XMLFileName
WRITE_UNMAPPED       = True   # export XML fields with no template column
SNIFF_NON_XML        = True
INCLUDE_DIAG_COLUMNS = False  # adds folder/business/type/year columns
SP_TEXT_LIMIT        = 255
SORT_OUTPUT          = True   # deterministic row order by application number,
                              # so reruns diff cleanly and SharePoint's row
                              # numbers stay meaningful between imports
# =============================================================================

# Your new template dropped these two, but your SharePoint list and every
# previous export still have them. Kept so the columns keep filling.
# Set RETAIN_LEGACY_COLUMNS = False below to honour the template exactly.
LEGACY_COLUMNS = ["ApplicationFolderID", "XMLFileName"]

TEMPLATE_COLUMNS = [
    "ApplicationID", "ButtonSelection", "BusinessName", "LicenceType",
    "ApplicationType", "ReceivedDate", "ApplicationStage", "BusinessID",
    "CATSID", "ExpiryDate", "Priority", "ApplicationReceived",
    "DisclosureComplete", "FinStmtReceived", "FSCertification",
    "ResumeReceived", "BusinessPlanReceived", "LeaseReceived",
    "DebtorLetterReceived", "ContractReceived", "ReferenceLetters2",
    "OneThousandBondProvided", "FiveThousandBondProvided", "EvidenceofExam",
    "SupervisionLetReceived", "AdditionalInfoReq", "DocumentsReqDate",
    "AppContentDays", "DocumentsDue", "DocumentsReceived", "ReceiptNumber",
    "ONBISComplete", "ONBISCorpPage", "ONBISDirector", "ONBISNameReg",
    "FederalRegistrationComplete", "LARSubmitted", "TermsConditions",
    "CriminalBckCheckComplete", "ConsumerReportConducted",
    "RiskProfileReport", "OustandingAP", "ToBeAssigned",
    "ComplianceOfficerName", "RegistrationClerk", "CPICRequired",
    "OperateInOntario", "StreetViewReview", "OfficeInOntario",
    "AddressforServiceOntario", "BusinessNameMatch",
    "DirectorOfficerConsistent", "ComplyWith24-3", "TrustAccountConsistent",
    "TrustAcctDesignation", "FinancialStatement", "TrustAccountConfirm",
    "FiveThousandBond", "BankNameCorrect", "OneThousandBond",
    "BondIssuedDate", "TwoYrsExperience", "AppOffDirOver18", "PassedExam",
    "ConsumerReportReview", "zzAppOffDirBankrupt", "ReferenceLetters",
    "AppOffDirInOntario", "DisclosureIssues", "AppCanadianCitizen",
    "SupervisionLetter", "BusinessPlan", "DebtorLetters", "MandatoryStmt",
    "SampleContract", "PaytoFund", "AdditionalInfo", "ManagerEmail",
    "ComplianceOfficerEmail", "CollectionsExperience",
    "FinancialExperience", "TrustExperience", "DebtorCreditorLaw",
    "ManagementExperience", "ManagerDecision", "ManagerName",
    "ManagerSendTo", "RegistrarDecision", "RegistrarName",
    "RegistrarSendTo", "UpdatedCATS", "UpdatedICPS", "LicensesPrinted",
    "LicensedMailed", "ApprovalLetter", "RefusalLetter", "UpdateCATS",
    "UpdateICPS", "RefundApplication", "Comment", "ActivityLog",
    "AssignedTo", "CreatedBy", "AssignedOn", "CreatedOn", "AssignedBy",
    "ModifiedOn", "EmailMessage", "SendEmail", "UserName",
    "CurrentUserFirstName", "CurrentUserLastName", "group", "SubmittedFee",
    "ClerkComments", "Comments", "ClerkMessage", "ResearchStartDate",
    "ResearchEndDate", "ResearchCompleteDays", "AppOffDirBankrupt",
    "AppOffDirConvicted", "OfficerComments", "ComplianceOfficerMessage",
    "ComplianceStartDate", "ComplianceEndDate", "ComplianceReviewDays",
    "ManagerComments", "ManagerStartDate", "ManagerEndDate",
    "ManagerReviewDays", "RegistarComments", "RegistrarStartDate",
    "RegistrarEndDate", "RegistrarReviewDays", "CloseStartTime",
    "CloseEndTime", "CloseDays", "TotalApplicationTime", "BondEndDate",
]

# XML tag -> template column. Left side is what the form actually writes.
# XML tag -> template column, for form fields whose spelling does not match
# any column. Sources that ARE template columns are skipped automatically
# (see build_lookup), so adding one here can never duplicate a real field.
TAG_ALIASES = {
    "FiancialStatement": "FinancialStatement",   # form typo, 543 files
    "MangerEndDate":     "ManagerEndDate",       # form typo, 102 files
    "zzSubmittedFee":    "SubmittedFee",         # 1 file
    # "AssignedTo2":     "AssignedTo",           # 45 files - distinct field;
    #                                            # enable only if it really is
    #                                            # the same thing as AssignedTo
}

# ---------------------------------------------------------------------------
# Value-transformation rules requested 2026-08-04 (Rachelle's mapping sheet +
# Ahmed's follow-up answers). See VALUE_RULES_NOTES.txt written alongside the
# output for the full explanation of each rule.
# ---------------------------------------------------------------------------

# Column F in NotesRachelle: Yes/No -> True/False
BOOLEAN_COLUMNS = ["SendEmail", "ManagerDecision"]

# Column H in NotesRachelle: numeric 1/2/3 -> Yes/No/N/A. Order (1,2,3) is
# exactly as specified in the request. Idempotent: text already reading
# Yes/No/N/A is normalised in case, not altered in meaning.
YESNONA_COLUMNS = [
    "ApplicationReceived", "ReferenceLetters2", "OneThousandBondProvided",
    "FiveThousandBondProvided", "EvidenceofExam", "SupervisionLetReceived",
    "ONBISComplete", "ONBISCorpPage", "ONBISDirector", "ONBISNameReg",
    "FederalRegistrationComplete", "LARSubmitted", "TermsConditions",
    "CriminalBckCheckComplete", "ConsumerReportConducted", "RiskProfileReport",
    "OustandingAP", "OperateInOntario", "StreetViewReview", "OfficeInOntario",
    "BusinessNameMatch", "DirectorOfficerConsistent", "ComplyWith24-3",
    "TrustAccountConsistent", "TrustAcctDesignation", "FinancialStatement",
    "TrustAccountConfirm", "FiveThousandBond", "BankNameCorrect",
    "OneThousandBond", "TwoYrsExperience", "AppOffDirOver18", "PassedExam",
    "ReferenceLetters", "AppOffDirInOntario", "DisclosureIssues",
    "AppCanadianCitizen", "SupervisionLetter", "BusinessPlan", "DebtorLetters",
    "MandatoryStmt", "SampleContract", "PaytoFund", "UpdatedCATS",
    "UpdatedICPS", "LicensesPrinted", "LicensedMailed", "RefusalLetter",
    "UpdateCATS", "UpdateICPS", "RefundApplication", "zzAppOffDirBankrupt",
    "AppOffDirBankrupt",   # not in Rachelle's list, added for consistency with
                           # the AppOffDirBankrupt -> zzAppOffDirBankrupt copy
    "AppOffDirConvicted",  # confirmed by Rachelle 2026-08-04 after the
                           # Dataverse dataflow rejected value "2" on
                           # cree9_appoffdirconvicted (25 failed rows)
]

# NOT in Rachelle's column H. Found by scanning the real 2,880-row export:
# these columns contain ONLY 1/2/3 and no other values, so a Dataverse Choice
# column will reject them exactly the way AppOffDirConvicted did.
# SubmittedFee sounds monetary but holds no decimal or dollar values in any of
# its 222 filled cells - it is a choice field despite the name.
# Set this to [] to leave them as raw numbers instead.
YESNONA_UNCONFIRMED = ["SubmittedFee"]
YESNONA_MAP = {"1": "Yes", "2": "No", "3": "N/A"}
YESNONA_CANON = {"yes": "Yes", "no": "No", "n/a": "N/A", "na": "N/A"}

# Columns J-W in NotesRachelle: "Last, First (pronoun) (dept)" -> "First Last"
# ONLY for staff Ahmed confirmed are current. Everyone else -> blank. This
# dictionary was built automatically from every (source, target) pair in the
# sheet where a target was actually filled in.
PERSON_ALLOWLIST = {
    ("theodoulis", "joy"): "Joy Theodoulis",
    ("pittens", "chris"): "Chris Pittens",
    ("kotsopoulos", "kate"): "Kate Kotsopoulos",
    ("fimiani", "anthony"): "Anthony Fimiani",
    ("ross", "richard"): "Richard Ross",
    ("ajmani", "priya"): "Priya Ajmani",
    ("ngkamman", "jocelyne"): "Jocelyne Ng Kam Man",
}
# Raw values with no comma (already-short or truncated forms) seen in the
# sheet, mapped directly.
PERSON_EXACT_ALIASES = {"pittens": "Chris Pittens"}

GENERAL_PERSON_COLUMNS = [
    "ComplianceOfficerName", "RegistrationClerk", "ManagerEmail",
    "ComplianceOfficerEmail", "ManagerName", "ManagerSendTo", "RegistrarSendTo",
    "AssignedTo", "CreatedBy", "AssignedBy", "UserName",
]
REGISTRAR_NAME_COLUMN = "RegistrarName"
REGISTRAR_VALID = {("pittens", "chris"): "Chris Pittens",
                   ("ross", "richard"): "Richard Ross"}

# Confirmed from red-highlighted cells in NotesRachelle: in the
# RegistrarSendTo field specifically, Chris Pittens / Richard Ross (the
# registrars themselves) should be blanked even though they're valid staff
# for every other field. Elsewhere in the sheet, every red cell already had
# no target (i.e. was already covered by the "blank if unlisted" default),
# so this is the one place red changes the outcome.
FIELD_SPECIFIC_BLANK = {
    "RegistrarSendTo": {("pittens", "chris"), ("ross", "richard")},
}

# Corrupted ApplicationType values found in the real data: a stray trailing
# year fragment (e.g. "-2022", "Renewal-2022", "Renewal - 2022") left over
# from a folder-name parsing quirk. Strips the year, keeps whatever type text
# (if any) is left.
APPLICATIONTYPE_YEAR_RE = __import__("re").compile(
    r"[\s\-]*((?:19|20)\d{2})\s*$")

CPIC_TYPO_FIXES = {"no nit": "No Hit"}
DATETIME_T_RE = __import__("re").compile(
    r"^(\d{4}-\d{2}-\d{2})T\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:?\d{2})?$")

# ---------------------------------------------------------------------------

# Not worth opening to check whether they are secretly XML.
SKIP_SNIFF_EXT = {
    ".pdf", ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".msg",
    ".docx", ".doc", ".xlsx", ".xls", ".xlsm", ".pptx", ".ppt", ".zip", ".rar",
    ".7z", ".mp4", ".mov", ".avi", ".mp3", ".wav", ".exe", ".dll", ".odt",
    ".ods", ".oxps", ".rtf", ".csv", ".eml", ".url", ".lnk",
}

HTML_NOISE = {"div", "span", "b", "i", "u", "strong", "em", "p", "br", "a",
              "font", "ul", "ol", "li", "table", "tr", "td", "th", "tbody",
              "html", "head", "body", "meta", "title", "svg", "g", "path",
              "defs", "style", "rect", "circle"}

ILLEGAL_XML = re.compile(
    "[^\u0009\u000A\u000D\u0020-\uD7FF\uE000-\uFFFD\U00010000-\U0010FFFF]"
)
BARE_AMP = re.compile(r"&(?!(?:[A-Za-z][A-Za-z0-9]*|#[0-9]+|#[xX][0-9A-Fa-f]+);)")
LEADING_NUM = re.compile(r"^\s*(\d{2,7})\b")
MULTI_JOIN = "; "
ACTIVE_COLUMN_KEYS: set[str] = set()   # filled at runtime from the template


# --------------------------------------------------------------- path helpers
def long_path(path: str) -> str:
    if os.name != "nt":
        return path
    p = os.path.abspath(path)
    if p.startswith("\\\\?\\"):
        return p
    if p.startswith("\\\\"):
        return "\\\\?\\UNC\\" + p.lstrip("\\")
    return "\\\\?\\" + p


def display_path(path: str) -> str:
    r"""Strip the \\?\ prefix so report paths are copy-pasteable."""
    if path.startswith("\\\\?\\UNC\\"):
        return "\\\\" + path[8:]
    if path.startswith("\\\\?\\"):
        return path[4:]
    return path


def norm_key(s: str) -> str:
    s = unicodedata.normalize("NFKD", str(s))
    return re.sub(r"[^0-9a-zA-Z]+", "", s).lower()


def clean_cell(v):
    if not isinstance(v, str):
        return v
    if FLATTEN_NEWLINES:
        v = re.sub(r"[\r\n]+", NEWLINE_TOKEN, v)
    return re.sub(r"[ \t]{2,}", " ", v).strip()


def person_core(raw: str):
    """'Pittens, Chris (He/Him) (MPBSDP)' -> ('pittens','chris'). None if the
    value has no comma, i.e. is not in 'Last, First' person-picker form."""
    s = re.sub(r"\(.*?\)", "", raw).strip()
    if "," not in s:
        return None
    last, first = s.split(",", 1)
    norm = lambda x: re.sub(r"[^a-z]", "", x.lower())
    return (norm(last), norm(first))


def normalize_person(raw: str, field: str = "") -> str:
    """General person columns: known active staff -> 'First Last'; anyone
    else, including former staff and unrecognised formats, -> blank.
    A field can additionally blank specific people via FIELD_SPECIFIC_BLANK
    (e.g. RegistrarSendTo blanking the registrars themselves)."""
    raw = (raw or "").strip()
    if not raw:
        return ""
    exact = raw.strip().lower()
    key = person_core(raw)
    blocked = FIELD_SPECIFIC_BLANK.get(field, ())
    if key and key in blocked:
        return ""
    if exact in PERSON_EXACT_ALIASES:
        name = PERSON_EXACT_ALIASES[exact]
        return "" if key in blocked else name
    if key and key in PERSON_ALLOWLIST:
        return PERSON_ALLOWLIST[key]
    return ""


def clean_application_type(raw: str) -> str:
    """Strip a stray trailing year fragment ('-2022', 'Renewal-2022', ...)
    from ApplicationType. Keeps any real type text that remains."""
    v = (raw or "").strip()
    if not v:
        return v
    cleaned = APPLICATIONTYPE_YEAR_RE.sub("", v).strip(" -")
    return cleaned


def normalize_registrar(raw: str) -> str:
    """RegistrarName: only 'Chris Pittens' or 'Richard Ross' allowed. Any
    other value (former registrars, other staff, blank) -> blank."""
    raw = (raw or "").strip()
    if not raw:
        return ""
    exact = raw.lower()
    for (last, first), canon in REGISTRAR_VALID.items():
        if exact == canon.lower():
            return canon
    key = person_core(raw)
    if key and key in REGISTRAR_VALID:
        return REGISTRAR_VALID[key]
    if exact in PERSON_EXACT_ALIASES and PERSON_EXACT_ALIASES[exact] in REGISTRAR_VALID.values():
        return PERSON_EXACT_ALIASES[exact]
    return ""


def normalize_yesnona(raw: str) -> str:
    """'1'/'2'/'3' -> Yes/No/N/A. Existing Yes/No/N/A text is case-normalised
    and passed through unchanged in meaning. Anything else is left as-is."""
    v = (raw or "").strip()
    if not v:
        return v
    vnum = v[:-2] if v.endswith(".0") else v      # "1.0" from float coercion
    if vnum in YESNONA_MAP:
        return YESNONA_MAP[vnum]
    low = v.lower()
    if low in YESNONA_CANON:
        return YESNONA_CANON[low]
    return v   # unrecognised value: leave untouched, do not guess


def normalize_boolean(raw: str) -> str:
    """Yes/No/1/0/true/false -> 'True'/'False'. Blank stays blank. Anything
    else is left as-is (flagged in the report, not silently changed)."""
    v = (raw or "").strip()
    if not v:
        return v
    low = v.lower()
    if low in ("true", "yes", "1"):
        return "True"
    if low in ("false", "no", "0"):
        return "False"
    return v


def fix_cpic_typo(raw: str) -> str:
    v = (raw or "").strip()
    if v.lower() in CPIC_TYPO_FIXES:
        return CPIC_TYPO_FIXES[v.lower()]
    return v


def truncate_datetime(raw: str) -> str:
    """'2026-09-01T00:00:00' -> '2026-09-01'. Only fires on an exact ISO
    date+time match, so it cannot misfire on free-text fields."""
    v = (raw or "").strip()
    m = DATETIME_T_RE.match(v)
    return m.group(1) if m else v


def parse_folder_name(name: str):
    """
    '13203 - My Money Tree Inc - Renewal - 2022'
        -> ('13203', 'My Money Tree Inc', 'Renewal', '2022')
    Business names may contain hyphens, so split on ' - ' and work inwards.
    """
    num = biz = typ = year = ""
    m = LEADING_NUM.match(name)
    if m:
        num = m.group(1)
    parts = [p.strip() for p in re.split(r"\s+-\s+", name) if p.strip()]
    if parts and parts[0] == num:
        parts = parts[1:]
    if parts and re.fullmatch(r"(19|20)\d{2}", parts[-1]):
        year = parts.pop()
    if parts and len(parts) > 1:
        typ = parts.pop()
    biz = " - ".join(parts)
    return num, biz, typ, year


# ------------------------------------------------------------- template load
def resolve_columns():
    if TEMPLATE_FILE and os.path.exists(TEMPLATE_FILE):
        ext = os.path.splitext(TEMPLATE_FILE)[1].lower()
        try:
            if ext in (".xlsx", ".xlsm", ".xltx"):
                df = pd.read_excel(TEMPLATE_FILE, nrows=0, engine="openpyxl")
            elif ext == ".xls":
                df = pd.read_excel(TEMPLATE_FILE, nrows=0, engine="xlrd")
            else:
                df = pd.read_csv(TEMPLATE_FILE, nrows=0, encoding="utf-8-sig")
            cols, seen = [], set()
            for c in df.columns:
                c = str(c).strip()
                if c and not c.lower().startswith("unnamed:") and c not in seen:
                    seen.add(c)
                    cols.append(c)
            if cols:
                return cols, f"template file: {TEMPLATE_FILE}"
        except Exception as e:                                   # noqa: BLE001
            print(f"  ! could not read template ({e}); using embedded columns")
    elif TEMPLATE_FILE:
        print(f"  ! template not found at {TEMPLATE_FILE}; using embedded columns")
    return list(TEMPLATE_COLUMNS), f"embedded column list ({len(TEMPLATE_COLUMNS)} fields)"


# ------------------------------------------------------------------ xml parse
def strip_ns(tag):
    if isinstance(tag, str) and "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def looks_like_html(raw: bytes) -> bool:
    head = raw[:400].lstrip().lower()
    return head.startswith(b"<!doctype html") or head.startswith(b"<html")


def parse_xml(path: str):
    opath = long_path(path)
    try:
        return ET.parse(opath).getroot(), ""
    except ET.ParseError as first_err:
        with open(opath, "rb") as fh:
            raw = fh.read()
        if not raw.strip():
            raise ET.ParseError("file is empty (0 bytes)")
        if looks_like_html(raw):
            raise ET.ParseError("file is HTML, not XML")
        text = ILLEGAL_XML.sub("", raw.decode("utf-8-sig", errors="replace"))
        start = text.find("<")
        if start > 0:
            text = text[start:]
        for label, cand in (("cleaned control chars / BOM", text),
                            ("escaped bare ampersands", BARE_AMP.sub("&amp;", text))):
            try:
                return ET.fromstring(cand), f"recovered by {label}"
            except ET.ParseError:
                continue
        raise first_err


def flatten(root):
    """
    'group' is a column in the new template, but in InfoPath it is a wrapper
    holding other fields. Concatenating a wrapper would dump every child
    field into one cell, so a wrapper contributes only its own direct text.
    Rich-text fields (children are <div>/<b>/<span>) still get concatenated.
    """
    elements: dict[str, list[str]] = defaultdict(list)
    attrs: dict[str, list[str]] = defaultdict(list)
    leaf_tags: set[str] = set()
    container_tags: set[str] = set()
    for elem in root.iter():
        tag = strip_ns(elem.tag)
        if not isinstance(tag, str):
            continue
        structural = any(
            isinstance(strip_ns(ch.tag), str)
            and strip_ns(ch.tag).lower() not in HTML_NOISE for ch in elem)
        if structural:
            container_tags.add(tag)
            val = (elem.text or "").strip()
        else:
            val = " ".join(t.strip() for t in elem.itertext() if t and t.strip())
        val = re.sub(r"\s+", " ", val).strip()
        if val:
            elements[tag].append(val)
        if len(elem) == 0:
            leaf_tags.add(tag)
        for k, v in elem.attrib.items():
            k, v = strip_ns(k), str(v).strip()
            if v:
                attrs[f"{tag}@{k}"].append(v)
                attrs[k].append(v)
    return elements, attrs, leaf_tags, container_tags


def build_lookup(elements, attrs):
    lut = {}
    for src in (attrs, elements):
        for k, vals in src.items():
            lut[norm_key(k)] = MULTI_JOIN.join(dict.fromkeys(vals))
    if APPLY_TAG_ALIASES:
        for wrong, right in TAG_ALIASES.items():
            if norm_key(wrong) in ACTIVE_COLUMN_KEYS:
                continue        # it has its own column now; leave it alone
            wk, rk = norm_key(wrong), norm_key(right)
            if lut.get(wk) and not lut.get(rk):
                lut[rk] = lut[wk]
    return lut


# ------------------------------------------------------------------ tree walk
def scan_tree(src: str):
    root_abs = long_path(src)
    xml_files, other_files, walk_errors, all_dirs = [], [], [], []
    ext_counter = Counter()
    top_level_dirs: set[str] = set()
    top_xml_count: Counter = Counter()
    loose_files = 0

    def top_of(path: str) -> str:
        rel = os.path.relpath(path, root_abs)
        parts = rel.split(os.sep)
        return parts[0] if len(parts) > 1 else ""

    def on_error(err: OSError):
        walk_errors.append((getattr(err, "filename", "?") or "?", str(err)))

    for root, dirs, files in os.walk(root_abs, onerror=on_error):
        all_dirs.append(root)
        if os.path.abspath(root) == os.path.abspath(root_abs):
            top_level_dirs.update(dirs)
        for name in files:
            full = os.path.join(root, name)
            ext = os.path.splitext(name)[1].lower().strip()
            ext_counter[ext or "(no extension)"] += 1
            top = top_of(full)
            if not top:
                loose_files += 1
            if ext == ".xml":
                xml_files.append(full)
                top_xml_count[top] += 1
            else:
                other_files.append(full)

    return {"xml_files": xml_files, "other_files": other_files,
            "walk_errors": walk_errors, "all_dirs": all_dirs,
            "ext_counter": ext_counter, "top_level_dirs": sorted(top_level_dirs),
            "top_xml_count": top_xml_count, "loose_files": loose_files,
            "root_abs": root_abs}


def sniff_hidden_xml(paths):
    hits = []
    for p in paths:
        if os.path.splitext(p)[1].lower() in SKIP_SNIFF_EXT:
            continue
        try:
            with open(long_path(p), "rb") as fh:
                head = fh.read(2048).lstrip()
        except OSError:
            continue
        if looks_like_html(head) or head[:5].lower() == b"<svg " or b"<svg" in head[:200]:
            continue
        if head[:5].lower() == b"<?xml" or re.match(rb"<[A-Za-z_][\w:.-]*[\s>/]", head):
            hits.append(p)
    return hits


# ----------------------------------------------------------------------- main
def run():
    started = datetime.now()
    src = sys.argv[1] if len(sys.argv) > 1 else INPUT_DIRECTORY
    out = sys.argv[2] if len(sys.argv) > 2 else OUTPUT_DIR

    if not os.path.isdir(src):
        sys.exit(f"ERROR: input folder does not exist:\n  {src}")
    a, b = os.path.abspath(out), os.path.abspath(src)
    if a == b or a.startswith(b + os.sep):
        sys.exit("ERROR: OUTPUT_DIR sits inside the folder being scanned.")
    os.makedirs(out, exist_ok=True)

    columns, col_source = resolve_columns()
    if RETAIN_LEGACY_COLUMNS:
        for c in LEGACY_COLUMNS:
            if c not in columns:
                columns.append(c)
                print(f"  + retained legacy column: {c}")
    global ACTIVE_COLUMN_KEYS
    ACTIVE_COLUMN_KEYS = {norm_key(c) for c in columns}
    active_col_set = set(columns)
    containers_seen: set[str] = set()
    out_columns = list(columns)
    if ID_COLUMN not in out_columns:
        out_columns.append(ID_COLUMN)
    diag_cols = ["App_Folder", "App_Number", "Folder_BusinessName",
                 "Folder_AppType", "Folder_Year", "Source_Path", "Parse_Note"]
    if INCLUDE_DIAG_COLUMNS:
        out_columns += diag_cols

    col_keys = {c: norm_key(c) for c in columns}
    template_keys = set(col_keys.values())
    alias_keys = {norm_key(k) for k in TAG_ALIASES} if APPLY_TAG_ALIASES else set()

    print(f"Columns from: {col_source}  (+ '{ID_COLUMN}')")
    print(f"Scanning:     {src}")
    inv = scan_tree(src)
    xml_files = inv["xml_files"]
    total_files = len(xml_files) + len(inv["other_files"])
    print(f"  folders: {len(inv['all_dirs'])}   files: {total_files}   "
          f"xml: {len(xml_files)}")

    print("  sniffing non-xml files for stray XML content...")
    hidden_xml = sniff_hidden_xml(inv["other_files"]) if SNIFF_NON_XML else []

    records, failures, empty_records, recovered = [], [], [], []
    unmatched_tags, col_hits = Counter(), Counter()
    alias_hits = Counter()
    unmapped_rows = []
    multi_xml_folders = {k: v for k, v in inv["top_xml_count"].items() if v > 1 and k}

    dates_truncated = Counter()
    yesnona_converted = Counter()
    yesnona_unrecognised = defaultdict(set)
    booleans_unrecognised = defaultdict(set)
    names_kept, names_blanked = Counter(), Counter()
    registrar_kept, registrar_blanked = Counter(), Counter()
    bankrupt_overwritten = 0
    cpic_fixed = 0
    apptype_fixed = 0
    unconfirmed_converted = Counter()

    for n, path in enumerate(xml_files, 1):
        if n % 500 == 0:
            print(f"  ...{n}/{len(xml_files)}")
        base = os.path.basename(path)
        stem = os.path.splitext(base)[0]
        rel = os.path.relpath(path, inv["root_abs"])
        parts = rel.split(os.sep)
        app_folder = parts[0] if len(parts) > 1 else ""
        f_num, f_biz, f_type, f_year = parse_folder_name(app_folder)
        stem_num = LEADING_NUM.match(stem)
        stem_num = stem_num.group(1) if stem_num else ""

        try:
            root, note = parse_xml(path)
            elements, attrs, leaf_tags, container_tags = flatten(root)
            containers_seen.update(container_tags & active_col_set)
            lut = build_lookup(elements, attrs)

            row = {}
            for col in columns:
                row[col] = clean_cell(lut.get(col_keys[col], ""))

            # ----- ApplicationID must import as text, never as a number:
            # strip any ".0" left over from float coercion upstream.
            if APPLY_VALUE_RULES and "ApplicationID" in row and row["ApplicationID"].endswith(".0"):
                row["ApplicationID"] = row["ApplicationID"][:-2]

            # ----- date/time truncation, applied everywhere: safe because it
            # only matches an exact 'YYYY-MM-DDTHH:MM:SS[...]' string.
            if APPLY_VALUE_RULES:
                for col in columns:
                    if row[col] and "T" in row[col]:
                        new = truncate_datetime(row[col])
                        if new != row[col]:
                            dates_truncated[col] += 1
                            row[col] = new

            # ----- 1/2/3 -> Yes/No/N/A
            if APPLY_VALUE_RULES:
                for col in list(YESNONA_COLUMNS) + list(YESNONA_UNCONFIRMED):
                    if col in row and row[col]:
                        new = normalize_yesnona(row[col])
                        if new != row[col]:
                            yesnona_converted[col] += 1
                            if col in YESNONA_UNCONFIRMED:
                                unconfirmed_converted[col] += 1
                        elif new not in ("", "Yes", "No", "N/A"):
                            yesnona_unrecognised[col].add(new)
                        row[col] = new

            # ----- AppOffDirBankrupt is authoritative; copy into the legacy
            # zzAppOffDirBankrupt column (both already Yes/No/N-A by now)
            if APPLY_VALUE_RULES and "AppOffDirBankrupt" in row and "zzAppOffDirBankrupt" in row:
                if row["AppOffDirBankrupt"]:
                    if row["zzAppOffDirBankrupt"] and row["zzAppOffDirBankrupt"] != row["AppOffDirBankrupt"]:
                        bankrupt_overwritten += 1
                    row["zzAppOffDirBankrupt"] = row["AppOffDirBankrupt"]

            # ----- Yes/No booleans -> True/False
            if APPLY_VALUE_RULES:
                for col in BOOLEAN_COLUMNS:
                    if col in row and row[col]:
                        new = normalize_boolean(row[col])
                        if new not in ("True", "False"):
                            booleans_unrecognised[col].add(new)
                        row[col] = new

            # ----- person-name allowlist (blank if not a confirmed active name)
            if APPLY_VALUE_RULES:
                for col in GENERAL_PERSON_COLUMNS:
                    if col in row and row[col]:
                        new = normalize_person(row[col], field=col)
                        if new:
                            names_kept[new] += 1
                        else:
                            names_blanked[row[col]] += 1
                        row[col] = new

            # ----- RegistrarName restricted to two names
            if APPLY_VALUE_RULES and REGISTRAR_NAME_COLUMN in row and row[REGISTRAR_NAME_COLUMN]:
                original = row[REGISTRAR_NAME_COLUMN]
                new = normalize_registrar(original)
                if new:
                    registrar_kept[new] += 1
                else:
                    registrar_blanked[original] += 1
                row[REGISTRAR_NAME_COLUMN] = new

            # ----- CPIC typo
            if APPLY_VALUE_RULES and "CPICRequired" in row and row["CPICRequired"]:
                new = fix_cpic_typo(row["CPICRequired"])
                if new != row["CPICRequired"]:
                    cpic_fixed += 1
                row["CPICRequired"] = new

            # ----- identity columns the XML does not carry
            if FILL_IDS_FROM_NAMES:
                if "ApplicationID" in row and not row["ApplicationID"]:
                    row["ApplicationID"] = stem_num or f_num
                if "ApplicationFolderID" in row and not row["ApplicationFolderID"]:
                    row["ApplicationFolderID"] = f_num or app_folder
                if "BusinessName" in row and not row["BusinessName"] and f_biz:
                    row["BusinessName"] = f_biz
                if "ApplicationType" in row and not row["ApplicationType"] and f_type:
                    row["ApplicationType"] = f_type

            # ----- ApplicationType: strip corrupted trailing year fragments.
            # Runs after the folder-name fallback above, so it cleans the
            # FINAL value regardless of whether it came from the XML tag or
            # the folder name.
            if APPLY_VALUE_RULES and "ApplicationType" in row and row["ApplicationType"]:
                new = clean_application_type(row["ApplicationType"])
                if new != row["ApplicationType"]:
                    apptype_fixed += 1
                row["ApplicationType"] = new
            if "XMLFileName" in row and not row["XMLFileName"]:
                row["XMLFileName"] = base

            # ----- unique, human-readable ID
            mid = stem
            if NORMALIZE_DUP_IDS and not stem_num and f_num:
                mid = f"{f_num}-{stem}"
            row[ID_COLUMN] = mid

            populated = 0
            for col in columns:
                if row[col]:
                    populated += 1
                    col_hits[col] += 1

            for wrong in TAG_ALIASES:
                if APPLY_TAG_ALIASES and elements.get(wrong):
                    alias_hits[wrong] += 1

            leftovers = {}
            for t in elements:
                nk = norm_key(t)
                if (t in leaf_tags and t.lower() not in HTML_NOISE
                        and nk not in template_keys and nk not in alias_keys):
                    unmatched_tags[t] += 1
                    leftovers[t] = clean_cell(MULTI_JOIN.join(
                        dict.fromkeys(elements[t])))
            if WRITE_UNMAPPED and leftovers:
                leftovers[ID_COLUMN] = mid
                unmapped_rows.append(leftovers)

            if INCLUDE_DIAG_COLUMNS:
                row.update({"App_Folder": app_folder, "App_Number": f_num,
                            "Folder_BusinessName": f_biz, "Folder_AppType": f_type,
                            "Folder_Year": f_year,
                            "Source_Path": display_path(path), "Parse_Note": note})

            records.append(row)
            if note:
                recovered.append((path, note))
            if populated == 0:
                empty_records.append(path)

        except ET.ParseError as e:
            failures.append((path, "ParseError", str(e)))
        except Exception as e:                                   # noqa: BLE001
            failures.append((path, type(e).__name__, str(e)))

    # ------------------------------------------------------------- write files
    df = pd.DataFrame(records, columns=out_columns).fillna("").astype(str)
    df = df.replace({"nan": ""})
    if SORT_OUTPUT and len(df):
        key_col = "ApplicationFolderID" if "ApplicationFolderID" in df.columns else ID_COLUMN
        df = (df.assign(_n=pd.to_numeric(df[key_col].str.extract(r"^(\d+)", expand=False),
                                         errors="coerce"))
                .sort_values(["_n", key_col, ID_COLUMN], na_position="last")
                .drop(columns="_n").reset_index(drop=True))

    csv_path = os.path.join(out, "sharepoint_upload_ready.csv")
    df.to_csv(csv_path, index=False, encoding="utf-8-sig",
              quoting=csv.QUOTE_ALL, lineterminator="\r\n")

    xlsx_path = ""
    if WRITE_EXCEL:
        xlsx_path = os.path.join(out, "sharepoint_upload_ready.xlsx")
        try:
            df.to_excel(xlsx_path, index=False, engine="openpyxl")
        except Exception as e:                                   # noqa: BLE001
            print(f"  (Excel write skipped: {e})")
            xlsx_path = ""

    pd.DataFrame([(display_path(p), t, e) for p, t, e in failures],
                 columns=["Path", "ErrorType", "Detail"]).to_csv(
        os.path.join(out, "failed_files.csv"), index=False, encoding="utf-8-sig")

    pd.DataFrame(sorted(unmatched_tags.items(), key=lambda x: -x[1]),
                 columns=["XmlTag", "FileCount"]).to_csv(
        os.path.join(out, "xml_tags_not_in_template.csv"),
        index=False, encoding="utf-8-sig")

    # ---- folder coverage: the answer to "Explorer says more than this"
    cov = []
    for d in inv["top_level_dirs"]:
        num, biz, typ, year = parse_folder_name(d)
        cnt = inv["top_xml_count"].get(d, 0)
        cov.append({"App_Folder": d, "App_Number": num, "BusinessName": biz,
                    "AppType": typ, "Year": year, "XML_Count": cnt,
                    "Status": "OK" if cnt == 1 else
                              ("NO XML FOUND" if cnt == 0 else "MULTIPLE XML")})
    cov_df = pd.DataFrame(cov).sort_values(["XML_Count", "App_Folder"])
    cov_path = os.path.join(out, "folder_coverage.csv")
    cov_df.to_csv(cov_path, index=False, encoding="utf-8-sig")
    no_xml = cov_df[cov_df.XML_Count == 0]

    unmapped_path = ""
    if WRITE_UNMAPPED and unmapped_rows:
        keys = [ID_COLUMN] + [t for t, _ in unmatched_tags.most_common()]
        um = pd.DataFrame(unmapped_rows).reindex(columns=keys).fillna("")
        unmapped_path = os.path.join(out, "unmapped_fields.csv")
        um.to_csv(unmapped_path, index=False, encoding="utf-8-sig")

    # ---- pre-import safety net: any column still holding only 1/2/3 is a
    # Dataverse/SharePoint Choice-field rejection waiting to happen. Surface
    # it in the report BEFORE the import rather than after a failed dataflow.
    RAW_CHOICE = {"1", "2", "3", "1.0", "2.0", "3.0"}
    unconverted_choice = []
    for c in df.columns:
        vals = {v.strip() for v in df[c] if str(v).strip()}
        if vals and vals <= RAW_CHOICE:
            unconverted_choice.append(
                (c, sorted(vals), int(sum(1 for v in df[c] if str(v).strip()))))

    long_cols = []
    for c in df.columns:
        mx = int(df[c].str.len().max() or 0)
        if mx > SP_TEXT_LIMIT:
            long_cols.append((c, mx, int((df[c].str.len() > SP_TEXT_LIMIT).sum())))
    long_cols.sort(key=lambda x: -x[1])

    dup = df[ID_COLUMN].value_counts()
    dup = dup[dup > 1]
    never = [c for c in columns if col_hits[c] == 0]
    total_items = total_files + max(len(inv["all_dirs"]) - 1, 0)
    top_items = len(inv["top_level_dirs"]) + inv["loose_files"]

    audit_path = os.path.join(out, "processing_summary_report.txt")
    with open(audit_path, "w", encoding="utf-8") as f:
        w = f.write
        w("=" * 64 + "\n  XML -> SHAREPOINT EXTRACTION REPORT (v4)\n" + "=" * 64 + "\n")
        w(f"Started:   {started:%Y-%m-%d %H:%M:%S}\n")
        w(f"Finished:  {datetime.now():%Y-%m-%d %H:%M:%S}\n")
        w(f"Source:    {src}\nOutput:    {out}\nColumns:   {col_source}\n\n")

        w("--- WHAT WINDOWS EXPLORER IS SHOWING YOU ---\n")
        w("Explorer's status bar counts only what is DIRECTLY in the folder,\n")
        w("not the whole tree. That figure is:\n")
        w(f"  Application folders at top level:   {len(inv['top_level_dirs'])}\n")
        w(f"  Loose files at top level:           {inv['loose_files']}\n")
        w(f"  = items Explorer reports:           {top_items}\n\n")

        w("--- WHOLE TREE ---\n")
        w(f"Folders (all levels):            {max(len(inv['all_dirs']) - 1, 0)}\n")
        w(f"Files (all types):               {total_files}\n")
        w(f"TOTAL ITEMS:                     {total_items}\n")
        w(f"  .xml files:                    {len(xml_files)}\n")
        w(f"  other files:                   {len(inv['other_files'])}\n")
        w(f"Unreadable folders:              {len(inv['walk_errors'])}\n")
        w(f"XML content without .xml ext:    {len(hidden_xml)}\n\n")

        w("--- FOLDER COVERAGE (why the row count is below the folder count) ---\n")
        w(f"Folders with exactly one XML:    {int((cov_df.XML_Count == 1).sum())}\n")
        w(f"Folders with NO XML:             {len(no_xml)}   <<< THE GAP\n")
        w(f"Folders with MORE than one XML:  {int((cov_df.XML_Count > 1).sum())}\n")
        w("Full list: folder_coverage.csv (sorted with the empty ones first)\n\n")

        w("--- PROCESSING RESULT ---\n")
        w(f"Rows exported:                   {len(records)}\n")
        w(f"  repaired on read:              {len(recovered)}\n")
        w(f"Failed files:                    {len(failures)}\n")
        w(f"Rows with zero mapped fields:    {len(empty_records)}\n")
        w(f"Duplicate {ID_COLUMN}:             {len(dup)}\n")
        w(f"BALANCE CHECK: {len(records)} + {len(failures)} = "
          f"{len(records) + len(failures)}  (must equal {len(xml_files)})\n\n")

        if len(no_xml):
            w(f"--- APPLICATION FOLDERS WITH NO CHECKLIST XML ({len(no_xml)}) ---\n")
            for _, r in no_xml.iterrows():
                w(f"  {r['App_Folder']}\n")
            w("\n")

        if multi_xml_folders:
            w(f"--- FOLDERS WITH MORE THAN ONE XML ({len(multi_xml_folders)}) ---\n")
            w("Each XML becomes its own row, so these inflate the export.\n")
            for k, v in sorted(multi_xml_folders.items(), key=lambda x: -x[1]):
                w(f"  {v:>3} XMLs  {k}\n")
            w("\n")

        if containers_seen:
            w("--- COLUMNS THAT MATCH A WRAPPER ELEMENT, NOT A FIELD ---\n")
            w("These template columns name an XML container that holds other\n")
            w("fields. Only their own direct text is exported, so they will\n")
            w("usually be blank. The child fields land in their own columns.\n")
            for t in sorted(containers_seen):
                w(f"  {t}\n")
            w("\n")

        if alias_hits:
            w("--- MISSPELLED FORM FIELDS REMAPPED ---\n")
            for k, v in alias_hits.most_common():
                w(f"  {k:<26}-> {TAG_ALIASES[k]:<24}{v} files\n")
            w("\n")

        if APPLY_VALUE_RULES:
            w("=" * 64 + "\n  VALUE TRANSFORMATION RULES (Rachelle's sheet, 2026-08-04)\n"
              + "=" * 64 + "\n")

            w(f"ApplicationID forced to text (stripped '.0'): see report count "
              f"above if any.\n")
            w(f"CPICRequired 'No Nit' -> 'No Hit' fixed:  {cpic_fixed} rows\n")
            w(f"ApplicationType stray year fragment fixed: {apptype_fixed} rows "
              f"(e.g. '-2022' -> '', 'Renewal-2022' -> 'Renewal')\n")
            w(f"AppOffDirBankrupt copied into zzAppOffDirBankrupt on every row "
              f"where AppOffDirBankrupt had a value.\n")
            if bankrupt_overwritten:
                w(f"  NOTE: {bankrupt_overwritten} of those rows had a DIFFERENT "
                  f"value already in zzAppOffDirBankrupt from its own XML tag; "
                  f"that value was overwritten per your instruction that "
                  f"AppOffDirBankrupt is authoritative.\n")
            w("\n")

            w(f"--- DATE/TIME FIELDS TRUNCATED (removed T00:00:00 etc.) ---\n")
            if dates_truncated:
                for c, n in dates_truncated.most_common():
                    w(f"  {c:<28}{n} cells\n")
            else:
                w("  none found\n")
            w("\n")

            w(f"--- 1/2/3 -> YES/NO/N-A CONVERSIONS ---\n")
            for c in YESNONA_COLUMNS:
                if yesnona_converted.get(c):
                    w(f"  {c:<28}{yesnona_converted[c]} cells converted\n")
            if yesnona_unrecognised:
                w("\n  Values in these Yes/No/N-A columns that were NEITHER "
                  "1/2/3 NOR already\n  Yes/No/N/A, and so were left "
                  "unchanged (review these by hand):\n")
                for c, vals in yesnona_unrecognised.items():
                    w(f"    {c}: {sorted(vals)[:10]}\n")
            w("\n")

            if unconfirmed_converted:
                w("--- !! CONVERTED BUT NOT IN RACHELLE'S COLUMN H - CONFIRM !! ---\n")
                w("These columns held ONLY 1/2/3, so they were converted to\n")
                w("Yes/No/N-A to prevent the Dataverse Choice-field rejection\n")
                w("that AppOffDirConvicted caused. They were NOT on the sheet,\n")
                w("so please confirm the meaning before importing. To revert,\n")
                w("set YESNONA_UNCONFIRMED = [] at the top of the script.\n")
                for c, n in unconfirmed_converted.most_common():
                    w(f"  {c:<28}{n} cells converted\n")
                w("\n")

            w(f"--- BOOLEAN (Yes/No -> True/False) ---\n")
            w(f"  Columns affected: {', '.join(BOOLEAN_COLUMNS)}\n")
            if booleans_unrecognised:
                w("  Values that were neither Yes/No/True/False and were left "
                  "unchanged:\n")
                for c, vals in booleans_unrecognised.items():
                    w(f"    {c}: {sorted(vals)[:10]}\n")
            w("\n")

            w(f"--- PERSON NAME ALLOWLIST ---\n")
            w(f"Applied to: {', '.join(GENERAL_PERSON_COLUMNS)}\n")
            w(f"Recognised as current staff and normalised to 'First Last':\n")
            for name, cnt in names_kept.most_common():
                w(f"    {name:<26}{cnt} cells\n")
            w(f"\nNOT on the confirmed active list -> blanked per your "
              f"instruction ({sum(names_blanked.values())} cells, "
              f"{len(names_blanked)} distinct raw values):\n")
            for raw, cnt in names_blanked.most_common(60):
                w(f"    {cnt:>5}x  {raw}\n")
            if len(names_blanked) > 60:
                w(f"    ... and {len(names_blanked)-60} more distinct values\n")
            w("\nIf any of these are current staff who should be kept, add "
              "them to\nPERSON_ALLOWLIST at the top of this script and rerun "
              "- no other\nchange is needed.\n\n")

            w(f"--- REGISTRARNAME (restricted to Chris Pittens / Richard Ross) ---\n")
            for name, cnt in registrar_kept.most_common():
                w(f"    kept    {name:<20}{cnt} cells\n")
            for raw, cnt in registrar_blanked.most_common(30):
                w(f"    blanked {raw:<40}{cnt} cells\n")
            w("\n")

        if unconverted_choice:
            w("=" * 64 + "\n  !! ACTION REQUIRED BEFORE IMPORT !!\n" + "=" * 64 + "\n")
            w("These columns still contain ONLY raw 1/2/3 values. If the target\n")
            w("field is a Choice column, the import WILL reject those rows the\n")
            w("way AppOffDirConvicted did. Add the column name to\n")
            w("YESNONA_COLUMNS at the top of this script and rerun.\n")
            for c, vals, cnt in unconverted_choice:
                w(f"  {c:<30}values={vals}  {cnt} cells\n")
            w("\n")
        else:
            w("--- CHOICE-FIELD PRE-IMPORT CHECK: PASSED ---\n")
            w("No column contains raw 1/2/3 values.\n\n")

        w("--- FILE EXTENSION BREAKDOWN ---\n")
        for ext, cnt in inv["ext_counter"].most_common():
            w(f"  {ext:<26}{cnt}\n")
        w("\n")

        if failures:
            w("--- FAILURE REASONS ---\n")
            for t, c in Counter(t for _, t, _ in failures).most_common():
                w(f"  {t:<26}{c}\n")
            w("\n--- FAILED FILES ---\n")
            for p, t, e in failures:
                w(f"  [{t}] {display_path(p)}\n      {e}\n")
            w("\n")

        if inv["walk_errors"]:
            w("--- UNREADABLE FOLDERS ---\n")
            for p, e in inv["walk_errors"]:
                w(f"  {display_path(p)} | {e}\n")
            w("\n")

        if hidden_xml:
            w("--- XML CONTENT WITHOUT A .xml EXTENSION (not imported) ---\n")
            for p in hidden_xml:
                w(f"  {display_path(p)}\n")
            w("\n")

        if long_cols:
            w("--- SHAREPOINT COLUMN TYPES: THESE EXCEED 255 CHARS ---\n")
            w("Create these as 'Multiple lines of text'.\n")
            for c, mx, cnt in long_cols:
                w(f"  {c:<28}max={mx:<8}over limit={cnt}\n")
            w("\n")

        if never:
            w(f"--- TEMPLATE COLUMNS STILL EMPTY ({len(never)}) ---\n")
            for c in never:
                w(f"  {c}\n")
            w("\n")

        if unmatched_tags:
            w(f"--- XML FIELDS WITH NO TEMPLATE COLUMN ({len(unmatched_tags)}) ---\n")
            w("Values preserved in unmapped_fields.csv. Add SharePoint columns\n")
            w("for any you want to keep.\n")
            for t, c in unmatched_tags.most_common(60):
                w(f"  {t:<40}{c} files\n")
            w("\n")

        if len(dup):
            w(f"--- REMAINING DUPLICATE {ID_COLUMN} VALUES ---\n")
            for v, c in dup.items():
                w(f"  {v}: {c} rows\n")
            w("\n")

        if empty_records:
            w("--- PARSED BUT MAPPED NOTHING ---\n")
            for p in empty_records[:200]:
                w(f"  {display_path(p)}\n")

    # ---------------------------------------------------------------- console
    L = "=" * 56
    print("\n" + L + "\n        EXTRACTION SUMMARY\n" + L)
    print(f"Top-level application folders:   {len(inv['top_level_dirs'])}")
    print(f"Loose files at top level:        {inv['loose_files']}")
    print(f"  = what Explorer's bar shows:   {top_items}")
    print(L)
    print(f"Folders with one XML:            {int((cov_df.XML_Count == 1).sum())}")
    print(f"Folders with NO XML:             {len(no_xml)}   <-- THE GAP")
    print(f"Folders with >1 XML:             {int((cov_df.XML_Count > 1).sum())}")
    print(L)
    print(f"Files in tree (all types):       {total_files}")
    print(f"  .xml:                          {len(xml_files)}")
    print(f"Rows exported:                   {len(records)}")
    print(f"Failed files:                    {len(failures)}")
    print(f"BALANCE: {len(records)}+{len(failures)}={len(records)+len(failures)}"
          f" of {len(xml_files)} xml")
    print(f"Duplicate {ID_COLUMN}:             {len(dup)}")
    print(f"Template columns still empty:    {len(never)} of {len(columns)}")
    print(L)
    if APPLY_VALUE_RULES:
        print(L)
        print(f"Names normalised to active staff:  {sum(names_kept.values())}")
        print(f"Names blanked (not on allowlist):  {sum(names_blanked.values())} "
              f"({len(names_blanked)} distinct)")
        print(f"RegistrarName kept / blanked:      "
              f"{sum(registrar_kept.values())} / {sum(registrar_blanked.values())}")
        print(f"1/2/3 -> Yes/No/N-A conversions:   {sum(yesnona_converted.values())}")
        print(f"CPICRequired typo fixes:           {cpic_fixed}")
        print(f"ApplicationType year-fragment fixes:{apptype_fixed}")
        if unconfirmed_converted:
            print(f"  ! converted but NOT on Rachelle's sheet: "
                  f"{', '.join(unconfirmed_converted)} - confirm before import")
        if unconverted_choice:
            print(f"  !! {len(unconverted_choice)} column(s) STILL hold raw 1/2/3: "
                  f"{', '.join(c for c,_,_ in unconverted_choice)}")
            print(f"     These will be rejected on import. See the report.")
        else:
            print(f"Choice-field pre-import check:      PASSED")
        print(f"Date/time fields truncated:        {sum(dates_truncated.values())}")
    print(L)
    print(f"CSV:       {csv_path}")
    if xlsx_path:
        print(f"Excel:     {xlsx_path}")
    print(f"Coverage:  {cov_path}   <-- start here")
    if unmapped_path:
        print(f"Unmapped:  {unmapped_path}")
    print(f"Report:    {audit_path}")


if __name__ == "__main__":
    run()
