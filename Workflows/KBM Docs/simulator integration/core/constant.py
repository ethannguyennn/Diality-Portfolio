import os
import sys

if getattr(sys, 'frozen', False):
    _SCRIPT_DIR = os.path.dirname(sys.executable)
else:
    _SCRIPT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Folder paths — mutated at runtime when user picks a folder
RECORDS_FOLDER  = os.path.join(_SCRIPT_DIR, 'KBM Docs')
OUTDATED_FOLDER = os.path.join(RECORDS_FOLDER, 'Outdated')
REDLINE_FOLDER  = os.path.join(RECORDS_FOLDER, 'Redline')
PREV_REV_FOLDER = os.path.join(RECORDS_FOLDER, 'Previous Revisions')

_LEN_REVISION_STR    = 13   # len("Revision No: ")
_MAX_RETRIES         = 3    # Word COM retry limit
_DELAY               = 1    # base seconds between retries
_LEN_VERSION_STR     = 9    # len("Version: ")
_LEN_END_VERSION_STR = 14   # chars to read past "Version:" for value

REGEX_REVISION_PATTERN        = '[0-9]+\\.[0-9]+'
REGEX_DATE_PATTERN            = '[0-9]+/[0-9]+/[0-9]+'
REGEX_DOC_PATH_REV_WHITESPACE = r"[A-Za-z]+\s[0-9]+"
REGEX_DOC_PATH_REV            = r"[A-Za-z]+[0-9]+"

STR_REV_COMMENTARY = "External document updated."
STR_VERSION        = ""

KBM_MAP = {
    "DIA EF-732L.1": "101016",
    "DIA EF-732L.2": "101045",
    "DIA EF-732L.3": "101029",
    "DIA EF-733I.2": "103396",
    "DIA EF-736B.2": "101049",
    "DIA EF-736B.3": "101033",
    "DIA EF-736B.6": "102875",
    "DIA EF-737B.2": "101058",
    "DIA EF-737B.3": "101042",
    "DIA EF-737B.6": "102876",
    "DIA EF-738H": "103397",
    "DIA EF-753A.2": "103018",
    "DIA WF-7310": "101015",
    "DIA WF-732A": "100213",
    "DIA WF-732C": "100214",
    "DIA WF-732D": "100340",
    "DIA WF-732H": "101010",
    "DIA WF-733A.1": "100783",
    "DIA WF-733A.2": "101047",
    "DIA WF-733A.3": "101031",
    "DIA WF-733A.4": "101011",
    "DIA WF-733A.5": "103014",
    "DIA WF-733C": "100784",
    "DIA WF-733D": "100785",
    "DIA WF-733D.1": "100786",
    "DIA WF-734A.1": "101018",
    "DIA WF-734A.2": "101048",
    "DIA WF-734A.3": "101032",
    "DIA WF-735": "101012",
    "DIA WF-736B.1": "101019",
    "DIA WF-736B.4": "101050",
    "DIA WF-736B.5": "101034",
    "DIA WF-736E.1": "100864",
    "DIA WF-736E.2": "101051",
    "DIA WF-736E.3": "101035",
    "DIA WF-736G.1": "101024",
    "DIA WF-736G.2": "101052",
    "DIA WF-736G.3": "101036",
    "DIA WF-736H.1": "101025",
    "DIA WF-736H.2": "101053",
    "DIA WF-736H.3": "101037",
    "DIA WF-736J": "102786",
    "DIA WF-736L.1": "101026",
    "DIA WF-736L.2": "101054",
    "DIA WF-736L.3": "101038",
    "DIA WF-736N.2": "101055",
    "DIA WF-736N.3": "101039",
    "DIA WF-736N.4": "101056",
    "DIA WF-736N.5": "101040",
    "DIA WF-737A.1": "101027",
    "DIA WF-737A.2": "101057",
    "DIA WF-737A.3": "101041",
    "DIA WF-737B.1": "101028",
    "DIA WF-737B.4": "101059",
    "DIA WF-737B.5": "101043",
    "DIA WF-738A.1": "102429-001",
    "DIA WF-738A.2": "102427-001",
    "DIA WF-738A.3": "102430-000",
    "DIA WF-738G": "101013",
    "DIA WF-738H.1": "102545",
    "DIA WF-738H.2": "102546",
    "DIA WF-738H.3": "102547",
    "DIA WF-739A1": "101014",
    "DIA WF-753A": "102871",
    "DIA WF-753A.1": "102872",
    "DIA WF-753B.1": "102673",
    "DIA WF-754A": "102924",
    "DIA List of SOUP Items Project Whitney Web Application": "101046",
    "DIA List of SOUP Items Project Whitney Cloud Extensions": "101017",
    "DIA List of SOUP Items Project Whitney Mobile Application": "101030",
}
