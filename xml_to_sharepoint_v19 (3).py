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
#
# A regular expression is a search pattern. Instead of looking for one exact
# word, it describes a SHAPE of text - "four digits", "a dash then a year".
# re.compile() prepares the pattern once here so it runs fast later.
#
# You do NOT need to be able to write these to maintain the script. The
# symbols below are the only ones used:
#
#       \d   any digit 0-9              ^    must be at the start
#       \s   a space or tab             $    must be at the end
#       +    one or more of these       ?    optional
#       *    zero or more               |    or
#       {2,7}  between 2 and 7 of them  ( )  remember this part
#
# What each pattern finds, and why it exists:
#
#   ILLEGAL_XML    characters that XML forbids outright. Some checklists
#                  contain them, which makes the file unreadable until they
#                  are stripped. The odd \u0009 style codes are just how you
#                  write "tab", "newline" and so on by number.
#
#   BARE_AMP       a lone "&" that is not part of a proper code like "&amp;".
#                  A lone & is fatal in XML, and business names are full of
#                  them - "Smith & Jones". This used to lose ~200 files.
#                  The (?! ... ) part means "an & NOT followed by this".
#
#   FOLDER_NUM     The application number at the start of a FOLDER name. It
#                  must be followed by " - ". That separator is the only
#                  thing distinguishing an application number from a
#                  corporation number. Compare two real folder names:
#                      "13203 - My Money Tree Inc - Renewal - 2022"
#                           13203 IS the application number
#                      "2611314 Ontario Limited-New-2019"
#                           2611314 is the CORPORATION number. This folder
#                           name holds no application number at all.
#                  DO NOT RELAX THIS PATTERN. A pattern that accepts digits
#                  without the separator writes corporation numbers into
#                  ApplicationFolderID, where they are indistinguishable
#                  from valid application numbers and cannot be detected
#                  downstream. Roughly 78 folders are affected.
#
#   FOLDER_TAIL    The other folder convention, "Business-Type-Year", used by
#                  697 folders. Separates the type and year so the business
#                  name is not left holding the whole folder name.
#
#   FILE_NUM       The number at the start of a FILE name, "13203-Checklist".
#                  A hyphen with no spaces, so it needs its own pattern.
#
#   ISO_DT         a full date-and-time such as 2026-09-01T00:00:00. The ( )
#                  captures just the date part so the time can be dropped.
#                  Deliberately strict, start to end, so ordinary text that
#                  happens to contain a "T" is never touched.
#
#   TRAILING_YEAR  a year at the very end of a value, with any spaces or
#                  dashes before it. Turns "Renewal-2022" into "Renewal".
#                  (?:19|20) means the year starts 19.. or 20..
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
# PYTHON NOTE: "def name(inputs):" defines a reusable piece of work.
# "return" hands a value back. Nothing here runs until called by name.
# ###########################################################################

def long_path(p):
    r"""
    Make a path openable on Windows however long it is.

    Windows fails past 260 characters and the failure is SILENT - the file
    just looks absent. The \\?\ prefix removes the limit.
    """
    # os.name is "nt" on Windows. Anywhere else, nothing to do.
    if os.name != "nt": return p
    # Turn a relative path ("..\\folder") into a full one from the drive down.
    p = os.path.abspath(p)
    # Already prefixed? Leave it, otherwise the prefix is applied twice.
    if p.startswith("\\\\?\\"): return p
    return ("\\\\?\\UNC\\" + p.lstrip("\\")) if p.startswith("\\\\") else "\\\\?\\" + p

def show(p):
    r"""Remove the \\?\ prefix so report paths are readable."""
    if p.startswith("\\\\?\\UNC\\"): return "\\\\" + p[8:]
    return p[4:] if p.startswith("\\\\?\\") else p

def key(s):
    """Reduce a name to lowercase letters and digits, for matching.

    An XML tag and a column name can mean the same field but differ in case
    or punctuation. Comparing the reduced form matches them anyway.
    """
    # Read this from the inside out:
    #   str(s)                     make sure it is text
    #   unicodedata.normalize      split accented letters into letter+accent
    #   re.sub("[^0-9a-zA-Z]+","")  delete anything not a letter or digit
    #   .lower()                   make it all lowercase
    # So "Comply With 24-3" and "complywith243" end up the same.
    return re.sub(r"[^0-9a-zA-Z]+", "", unicodedata.normalize("NFKD", str(s))).lower()

def clean(v):
    """Replace line breaks inside a value with " | ".

    SharePoint reads a newline inside a quoted cell as a NEW RECORD. One
    field here has newlines in ~2,600 of 2,750 rows; an early export came
    out as 35,708 lines for 2,750 records.
    """
    # Inner re.sub: every run of line breaks becomes " | ".
    # Outer re.sub: two or more spaces/tabs collapse to one space.
    # .strip(): remove leading and trailing spaces.
    return re.sub(r"[ \t]{2,}", " ", re.sub(r"[\r\n]+", " | ", v)).strip()

def folder_parts(name):
    """Split a folder name into number, business, type, year.

    TWO CONVENTIONS EXIST ON THE DRIVE. Both are handled here.

        A   "13203 - My Money Tree Inc - Renewal - 2022"     2,195 folders
            Spaces around the hyphens, and a real application number in
            front. -> ("13203", "My Money Tree Inc", "Renewal", "2022")

        B   "2611314 Ontario Limited-New-2019"                 697 folders
            No spaces, and NO application number. The leading digits are a
            corporation number and part of the business name.
            -> ("", "2611314 Ontario Limited", "New", "2019")

    Convention A is only accepted when the last piece is a four-digit year.
    Without that test, convention B folders containing a hyphenated business
    name were split in the wrong place: "Smith - Jones Holdings-Renewal-2021"
    came back as business "Smith", type "Jones Holdings-Renewal".
    """
    # 1. An application number is digits followed by " - ", and nothing else.
    #    Digits without that separator are a corporation number - see
    #    FOLDER_NUM. .match() returns a "match object" or None.
    m = FOLDER_NUM.match(name)
    num = m.group(1) if m else ""

    # 2. Everything after the application number, or the whole name if there
    #    was none. len(m.group(0)) is how many characters the number and its
    #    separator took up.
    rest = name[len(m.group(0)):].strip() if m else name.strip()

    # 3. TRY CONVENTION A. Split wherever " - " appears.
    #    PYTHON NOTE: [ ... for x in ... if ... ] is a "list comprehension" -
    #    it builds a list in one line. Here: take each piece, trim its
    #    spaces, and keep it only if something is left.
    parts = [x.strip() for x in re.split(r"\s+-\s+", rest) if x.strip()]

    # 4. Accept it ONLY if there are at least two pieces and the last one is
    #    a four-digit year. That is what tells convention A apart from a
    #    convention B name that happens to contain " - ".
    if len(parts) >= 2 and re.fullmatch(r"(19|20)\d{2}", parts[-1]):
        # .pop() removes the last item from the list and hands it back.
        year = parts.pop()
        # If more than one piece is left, the last is the application type.
        typ = parts.pop() if len(parts) > 1 else ""
        # Whatever remains is the business name. Rejoined with " - " because
        # a business name may itself contain one, e.g. "Smith - Jones".
        return num, " - ".join(parts), typ, year

    # 5. CONVENTION B. Take the last two hyphen-separated pieces as the type
    #    and the year. Anchored to the end of the name, so a business name
    #    with its own hyphen keeps it.
    t = FOLDER_TAIL.match(rest)
    if t:
        return num, t.group(1).strip(), t.group(2).strip(), t.group(3)

    # 6. Neither convention. Nothing to split, so it is all business name.
    #    PYTHON NOTE: returning several values separated by commas hands
    #    back a "tuple" - a fixed group. The caller unpacks it with
    #    "a, b, c, d = folder_parts(name)".
    return num, rest, "", ""

def as_yesnona(v):
    """1/2/3 -> Yes/No/N/A.

    Safe to run twice - "Yes" comes back unchanged. The real data mixes
    numbers and text in the same column. Anything unrecognised passes
    through and is flagged in the report rather than guessed at.
    """
    # v[:-2] means "everything except the last two characters", so "1.0"
    # becomes "1". Spreadsheets sometimes turn 1 into 1.0.
    n = v[:-2] if v.endswith(".0") else v
    # Is it one of our numbers? Then hand back the matching word.
    if n in YESNONA_MAP: return YESNONA_MAP[n]
    # Otherwise look it up in the already-correct spellings.
    # PYTHON NOTE: .get(key, fallback) returns the fallback instead of
    # crashing when the key is missing. Here the fallback is v itself, so an
    # unrecognised value is passed through untouched rather than guessed at.
    return YESNONA_OK.get(v.lower(), v)

def as_boolean(v):
    """Yes/No/1/0/true/false -> "True"/"False". Anything else untouched."""
    low = v.lower()
    if low in ("true", "yes", "1"): return "True"
    return "False" if low in ("false", "no", "0") else v


# ###########################################################################
#  PART 3 - READING THE XML
# ###########################################################################

def load_template():
    """Read column names from the template file specified in TEMPLATE_PATH.

    Required. If missing, the script stops and says what it looked for,
    rather than guessing and producing the wrong columns.
    """
    # Stop immediately if the file is not there. sys.exit() ends the script
    # and prints the message. Better than continuing with no columns.
    # PYTHON NOTE: the f before the quotes makes it an "f-string" - anything
    # in {curly braces} is replaced by that value. So {TEMPLATE_PATH} becomes
    # the actual path in the message.
    if not os.path.exists(TEMPLATE_PATH):
        sys.exit(f"ERROR: no template file found at {TEMPLATE_PATH}.")

    # Read ONLY the header row - nrows=0 means "no data rows, just the
    # column names". Excel and CSV need different readers, so pick by the
    # file extension. This is the "A if condition else B" form again, split
    # over three lines for readability.
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
    """Read one XML file, repairing it if needed.

    First line is the normal parse. The rest handles badly-formed files,
    least invasive fix first: strip forbidden characters and junk before the
    declaration, then escape bare ampersands and retry.

    The ampersand case matters most. A lone "&" is fatal XML, and business
    names contain them - "Smith & Jones Collections" used to kill the whole
    file. Empty files and HTML files are reported, not repaired, because a
    repair would be a guess.
    """
    p = long_path(path)
    try:
        # The normal case, and the only line that runs for a healthy file.
        # .getroot() hands back the outermost tag, from which everything
        # else can be reached. False means "no repair was needed".
        return ET.parse(p).getroot(), False
    except ET.ParseError as first:
        # It failed. Read the raw bytes so they can be repaired.
        # "rb" means read-binary: give us the exact bytes, do not interpret.
        raw = open(p, "rb").read()
        # Nothing in the file at all - report it, do not try to repair.
        if not raw.strip(): raise ET.ParseError("file is empty (0 bytes)")
        head = raw[:400].lstrip().lower()
        if head.startswith(b"<!doctype html") or head.startswith(b"<html"):
            raise ET.ParseError("file is HTML, not XML")
        # Turn bytes into text. "utf-8-sig" also strips the invisible marker
        # Windows puts at the start of some files. errors="replace" means a
        # byte that cannot be decoded becomes a placeholder instead of crashing.
        # .sub("") then deletes every character XML forbids.
        text = ILLEGAL_XML.sub("", raw.decode("utf-8-sig", errors="replace"))
        # Some files have junk before the first "<". Find it and cut it off.
        # text[cut:] means "from position cut to the end".
        cut = text.find("<")
        if cut > 0: text = text[cut:]
        # Try twice: first the cleaned text as-is, then the same text with
        # every bare & turned into the proper code &amp;. The first version
        # that parses wins. True means "this was repaired", which the report
        # counts.
        for candidate in (text, BARE_AMP.sub("&amp;", text)):
            try: return ET.fromstring(candidate), True
            except ET.ParseError: continue
        # Neither worked. Re-raise the ORIGINAL error, because it describes
        # the real problem better than anything our repairs produced.
        raise first


def flatten(root, column_keys):
    """Turn a parsed file into a flat "field name -> value" lookup.

    Reads each file once. (The original ran 109 searches per file and kept
    only the first line of any rich-text field.)

    THE SUBTLETY - wrappers versus rich text. Both contain other elements,
    but need opposite handling:

        <my:group>              WRAPPER holding real fields. Must NOT be
          <my:ResearchDate>     merged, or every child lands in one cell.
          <my:OfficerComments>

        <my:Comment>            RICH TEXT. Must be merged, or the value
          <div>Rich <b>text</b> comes out as just "Rich".

    Children that are HTML tags mean rich text; anything else is a wrapper.
    """
    # Three empty containers to collect into.
    # PYTHON NOTE: "a, b, c = 1, 2, 3" assigns three things on one line.
    #   vals    field name -> list of values found for it
    #   attrs   same, for attributes (extra info attached to a tag)
    #   leaves  names of tags with no tags inside them, i.e. real fields
    # defaultdict(list) is a dictionary that creates an empty list the first
    # time a new key is used, removing any need to check whether that key
    # already exists. set() is an unordered bag with no duplicates.
    vals, attrs, leaves = defaultdict(list), defaultdict(list), set()

    # root.iter() walks EVERY tag in the file, at any depth, one at a time.
    # "el" is the current tag on each pass through the loop.
    for el in root.iter():
        # Tag names arrive as "{long-namespace-url}ApplicationID". Split at
        # the "}" and take the last piece to get just "ApplicationID".
        # isinstance() checks the type first, because comments inside the XML
        # also appear here and their "tag" is not text.
        tag = el.tag.split("}", 1)[-1] if isinstance(el.tag, str) else None
        # "continue" skips the rest of this pass and moves to the next tag.
        if not tag: continue
        structural = any(isinstance(c.tag, str)
                         and c.tag.split("}", 1)[-1].lower() not in HTML_NOISE
                         for c in el)
        text = ((el.text or "") if structural
                else " ".join(t for t in el.itertext() if t and t.strip()))
        text = re.sub(r"\s+", " ", text).strip()
        # Only keep it if there is something there. .append() adds to the
        # list for this tag name - a name can legitimately appear twice.
        if text: vals[tag].append(text)
        # len(el) counts tags INSIDE this one. Zero means it is a real field
        # rather than a container.
        if len(el) == 0: leaves.add(tag)
        for k, v in el.attrib.items():
            v = str(v).strip()
            if v: attrs[k.split("}", 1)[-1]].append(v)
    # Now flatten everything into one lookup table: lut = "look-up table".
    # {} on its own is an empty dictionary, ready to be filled below.
    lut = {}
    # Go through attributes first, then tags. Because a later write to the
    # same key overwrites an earlier one, this means a real tag always wins
    # over an attribute of the same name.
    for src in (attrs, vals):
        # .items() gives the name and its list of values together.
        # dict.fromkeys(vs) removes duplicates while keeping the order.
        # "; ".join(...) glues the remaining values into one string.
        # key(k) is our matching helper, so lookups ignore case/punctuation.
        for k, vs in src.items(): lut[key(k)] = "; ".join(dict.fromkeys(vs))
    for wrong, right in TAG_ALIASES.items():
        if (key(wrong) not in column_keys and lut.get(key(wrong))
                and not lut.get(key(right))):
            lut[key(right)] = lut[key(wrong)]
    return lut, vals, leaves


# ###########################################################################
#  PART 4 - THE MAIN RUN
# ###########################################################################
# Everything above was definitions; nothing had run. Four stages:
#   A  set up and check settings
#   B  walk the folder tree
#   C  read every checklist into a row   <- the heart
#   D  write the two files and check them
# ###########################################################################

def run():
    started = datetime.now()
    src = sys.argv[1] if len(sys.argv) > 1 else INPUT_DIRECTORY
    out = sys.argv[2] if len(sys.argv) > 2 else OUTPUT_DIR
    if not os.path.isdir(src): sys.exit(f"ERROR: no such folder:\n  {src}")
    # Don't write output inside the scanned folder: earlier exports were left
    # in the source tree and got re-scanned as source data.
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

    # SAFEGUARD. Every rule targets a column by NAME, so a renamed or removed
    # column makes that rule stop firing with no error and no symptom. This
    # checks all 73 and warns. Silent on a correct template. Do not remove it.
    rule_columns = (set(YESNONA_COLUMNS) | set(BOOLEAN_COLUMNS)
                    | {"CPICRequired", "AppOffDirBankrupt",
                       "zzAppOffDirBankrupt", "ApplicationID",
                       "ApplicationType", "BusinessName"})
    rules_off = sorted(c for c in rule_columns if c not in set(columns))
    if rules_off:
        print(f"  !! {len(rules_off)} rule(s) will NOT run, column absent from "
              f"template: {', '.join(rules_off)}")

    # ---- STAGE B: walk the tree.
    # Counts every file and folder by type. This is what answered "why does
    # Explorer say 3,307 when the export has 2,875 rows?" - 3,307 was the
    # top-level item count, and the tree is mostly PDFs and emails.
    # onerror matters: without it, a folder Windows cannot read is skipped
    # silently, which looks like a clean run with fewer records.
    root_abs = long_path(src)
    xmls, others, walk_errors, dirs_seen = [], [], [], []
    exts, per_folder, top_dirs, loose = Counter(), Counter(), set(), 0
    # os.walk visits every folder in the tree. Each pass gives three things:
    #   r  = the folder it is currently in
    #   ds = the folder names inside it
    #   fs = the file names inside it
    # onerror=walk_errors.append records a folder it cannot open. Without
    # that, such a folder is skipped in total silence.
    for r, ds, fs in os.walk(root_abs, onerror=walk_errors.append):
        dirs_seen.append(r)
        if os.path.abspath(r) == os.path.abspath(root_abs): top_dirs.update(ds)
        for f in fs:
            full = os.path.join(r, f)
            ext = os.path.splitext(f)[1].lower().strip()
            exts[ext or "(no extension)"] += 1
            # relpath gives the path from the top folder down, then .split
            # breaks it at each "\". rel[0] is therefore the application
            # folder this file belongs to. os.sep is "\" on Windows and "/"
            # elsewhere, so this works on either.
            rel = os.path.relpath(full, root_abs).split(os.sep)
            top = rel[0] if len(rel) > 1 else ""
            if not top: loose += 1
            if ext == ".xml": xmls.append(full); per_folder[top] += 1
            else: others.append(full)
    print(f"Scanning {src}\n  folders {len(dirs_seen)}   "
          f"files {len(xmls) + len(others)}   xml {len(xmls)}")

    # ---- STAGE C: read every checklist. One row per file.
    # The steps below run IN ORDER; two of them depend on it, marked below.
    # PYTHON NOTE: "try:" attempts something that may fail, "except:" at the
    # bottom catches it. A bad file is recorded and the loop continues, so one
    # corrupt checklist cannot stop the other 2,899.
    rows, failures, leftover_tags = [], [], Counter()
    repaired = blank_rows = 0
    stats, odd_values = Counter(), defaultdict(set)

    # enumerate(xmls, 1) walks the list AND counts as it goes, starting at 1.
    # So n is the position and path is the file.
    for n, path in enumerate(xmls, 1):
        # basename = the file name without its folders.
        # splitext splits off the extension; [0] takes the part before the
        # dot. "13203-Checklist.xml" -> base, then stem "13203-Checklist".
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
            # STEP 1: pull all columns from the parsed file.
            # PYTHON NOTE: this is a "dict comprehension" - builds a whole
            # dictionary in one line. Read: for every column c, look up c.
            row = {c: clean(lut.get(ckeys[c], "")) for c in columns}

            # STEP 2: ApplicationID must import as text. Strip a trailing
            # ".0" left by earlier spreadsheet handling. The column TYPE must
            # still be set to text in the import mapping; a CSV cannot do it.
            if row.get("ApplicationID", "").endswith(".0"):
                row["ApplicationID"] = row["ApplicationID"][:-2]

            # STEP 3: dates. 2026-09-01T00:00:00 -> 2026-09-01. Only an
            # exact date-and-time is changed, so free text is safe.
            for c in columns:
                if "T" in row[c]:
                    mm = ISO_DT.match(row[c])
                    if mm: row[c] = mm.group(1); stats["dates truncated"] += 1

            # STEP 4: 1/2/3 -> Yes/No/N/A. THIS CAUSED A FAILED IMPORT -
            # Dataverse rejected 25 rows over a raw "2".
            for c in YESNONA_COLUMNS:
                if row.get(c):
                    new = as_yesnona(row[c])
                    if new != row[c]: stats["1/2/3 converted"] += 1
                    elif new not in ("Yes", "No", "N/A"): odd_values[c].add(new)
                    row[c] = new

            # STEP 5: copy into the older zz column the list still has.
            # AppOffDirBankrupt wins if they disagree; the report counts it.
            if row.get("AppOffDirBankrupt"):
                if row.get("zzAppOffDirBankrupt") not in ("", row["AppOffDirBankrupt"]):
                    stats["zzAppOffDirBankrupt overwritten"] += 1
                row["zzAppOffDirBankrupt"] = row["AppOffDirBankrupt"]

            # STEP 6: Yes/No -> True/False.
            for c in BOOLEAN_COLUMNS:
                if row.get(c):
                    new = as_boolean(row[c])
                    if new not in ("True", "False"): odd_values[c].add(new)
                    row[c] = new

            # STEP 7: Staff names are now pulled exactly as they appear in the 
            # XML file per management's instructions. No formatting or blanking is applied.

            # STEP 8: known typo in the source data.
            if row.get("CPICRequired", "").lower() == "no nit":
                row["CPICRequired"] = "No Hit"; stats["CPICRequired typo fixed"] += 1

            # STEP 9: the XML HAS NO ApplicationID AT ALL. The number
            # exists only in the folder name ("13203 - My Money Tree Inc -
            # Renewal - 2022") and the file name. That is the "missing ID"
            # the extra column is named for. Only blanks are filled, so a
            # real value in the XML always wins.
            # ApplicationFolderID holds A NUMBER OR NOTHING. Never fall back
            # to the folder name here. A name cannot be matched against a
            # number, so any comparison against SharePoint reports every
            # affected row as a missing record even though it is present.
            # Roughly 619 rows have no folder number. Their folder name is
            # preserved in ApplicationFolderName, which is its own column.
            for col, val in (("ApplicationID", stem_num or f_num),
                             ("ApplicationFolderID", f_num),
                             ("ApplicationFolderName", app_folder),
                             ("BusinessName", f_biz), ("ApplicationType", f_type),
                             ("XMLFileName", base)):
                if col in row and not row[col] and val: row[col] = val

            # STEP 10: strip a stray year, "Renewal-2022" -> "Renewal".
            # ORDER MATTERS: runs after step 9 so it cleans the final value,
            # whether that came from the XML or the folder name.
            if row.get("ApplicationType"):
                new = TRAILING_YEAR.sub("", row["ApplicationType"]).strip(" -")
                if new != row["ApplicationType"]: stats["ApplicationType year stripped"] += 1
                row["ApplicationType"] = new

            # STEP 11: Missing_ID. The file name, or the folder number
            # prefixed when the file name has no number - 176 files are named
            # "Master Checklist" and 87 just "checklist", which collided on
            # 263 rows before this.
            row[ID_COLUMN] = f"{f_num}-{stem}" if not stem_num and f_num else stem

            blank_rows += not any(row[c] for c in columns)

            # STEP 12: XML fields with no column. ~30 of them; their names
            # and counts go into the report so nothing vanishes unnoticed.
            extra = {t: clean("; ".join(dict.fromkeys(v)))
                     for t, v in elements.items()
                     if t in leaves and t.lower() not in HTML_NOISE
                     and key(t) not in column_keys and key(t) not in alias_keys}
            for t in extra: leftover_tags[t] += 1

            # STEP 13: the file's own dates. CAUTION: OneDrive resets
            # creation dates on sync, so XML_Created is indicative only.
            # XML_Modified is reliable; CreatedOn inside the XML is better.
            try:
                st = os.stat(long_path(path))
                born = getattr(st, "st_birthtime", None) or st.st_ctime
                row["XML_Created"] = datetime.fromtimestamp(born).strftime(DATE_FMT)
                row["XML_Modified"] = datetime.fromtimestamp(st.st_mtime).strftime(DATE_FMT)
            except (OSError, OverflowError, ValueError):
                row["XML_Created"] = row["XML_Modified"] = ""
            rows.append(row)
        # Record the failure, carry on. NOTHING IS DISCARDED SILENTLY -
        # this is what makes the balance check a proof and not a nice number.
        except Exception as e:                                   # noqa: BLE001
            failures.append((show(path), type(e).__name__, str(e)))

    # ---- STAGE D: write the files.
    # Sorted by application number so two runs can be compared directly.
    # Every field quoted, byte-order mark, Windows line endings - what the
    # SharePoint importer expects.
    # ---- assemble, sorted so reruns diff cleanly
    # A DataFrame is pandas' table: rows and named columns, like a sheet.
    # .fillna("") turns any missing cell into an empty string and .astype(str)
    # forces every value to text, so no number is reformatted on the way out.
    df = pd.DataFrame(rows, columns=out_columns).fillna("").astype(str)
    if len(df):
        df = (df.assign(_n=pd.to_numeric(
                  df["ApplicationFolderID"].str.extract(r"^(\d+)", expand=False),
                  errors="coerce"))
                .sort_values(["_n", "ApplicationFolderID", ID_COLUMN],
                             na_position="last")
                .drop(columns="_n").reset_index(drop=True))

    # ---- ONE ROW PER APPLICATION.
    # Some folders hold more than one checklist - a "Copy", or a second file
    # saved beside the first. Writing a row for each sends the destination
    # two records for one application, and raises no error in doing so.
    #
    # The key is ApplicationID, NOT Missing_ID. Missing_ID is built from the
    # file name, and "34931-Checklist Copy" is a different file name from
    # "34931-Checklist" - so keying on it misses the exact case this is here
    # to catch. Rows with no ApplicationID are never collapsed together.
    #
    # The most recently modified file wins. Every dropped row is reported by
    # name and folder, and counted in the balance check, so nothing goes
    # missing quietly. Review that list. Two folders holding the same
    # ApplicationID is a filing problem on the drive, not a fault in this
    # tool, and dropping the row hides it from the upload without fixing it.
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

    # os.path.join builds a path with the right separator for the system.
    csv_path = os.path.join(out, "sharepoint_upload_ready.csv")
    # index=False        do not add pandas' own row-number column
    # encoding utf-8-sig writes the marker Excel needs to show accents right
    # quoting QUOTE_ALL  wrap every value in quotes, so a comma inside a
    #                    value can never be mistaken for a column break
    # lineterminator     Windows-style line endings, which the SharePoint
    #                    importer expects
    df.to_csv(csv_path, index=False, encoding="utf-8-sig",
              quoting=csv.QUOTE_ALL, lineterminator="\r\n")
    # Folder coverage is still computed - the report needs the counts and the
    # list of folders with no checklist. It is just not written as its own CSV.
    cov = pd.DataFrame([
        dict(zip(("App_Folder", "App_Number", "BusinessName", "AppType", "Year"),
                 (d,) + folder_parts(d)),
             XML_Count=per_folder.get(d, 0),
             Status=("OK" if per_folder.get(d, 0) == 1 else
                     "NO XML FOUND" if per_folder.get(d, 0) == 0 else "MULTIPLE XML"))
        for d in sorted(top_dirs)]).sort_values(["XML_Count", "App_Folder"]) \
        if top_dirs else pd.DataFrame(columns=["App_Folder", "XML_Count", "Status"])

    # ---- the checks that decide whether this file is safe to upload
    still_choice = [(c, int((df[c] != "").sum())) for c in df.columns
                    if len(df) and (df[c] != "").any()
                    and set(df[c][df[c] != ""]) <= RAW_CHOICE]
    long_cols = [(c, int(df[c].str.len().max())) for c in df.columns
                 if len(df) and int(df[c].str.len().max() or 0) > SP_TEXT_LIMIT]
    no_xml = int((cov.XML_Count == 0).sum()) if len(cov) else 0
    multi = int((cov.XML_Count > 1).sum()) if len(cov) else 0

    # ---- SCREEN SUMMARY. The CSV above is the only file written.
    # Every check that determines whether the export is safe to upload is
    # printed here, so no failure passes unnoticed. To retain a copy of a
    # run, redirect the output:
    #     python xml_to_sharepoint_v19.py > run_log.txt
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

    # ---- the three things that must not pass unnoticed.
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

    # ---- the rest is advisory. Read it, then carry on.
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

# PYTHON NOTE: this last bit means "only start running if this file was
# launched directly". If another script ever imports this one to reuse a
# function, the export will not fire off unexpectedly. run() is the call
# that actually starts everything.
if __name__ == "__main__":
    run()