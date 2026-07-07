"""Team roster and assignee-name normalization, shared by the Jira sync and
the sheet-update modules."""

SWVV_TEAM = [
    "Raghu Kallala", "Thomas Lippold", "Tejaskumar Patel", "Tiffany Mejia",
    "Zoltan Miskolci", "Sarina Cheung", "Tisha Patel", "Ethan Nguyen",
]
FW_TEAM = [
    "Arpita Srivastava", "Jashwant Gantyada", "Michael Garthwaite", "Praneeth Bunne",
    "Sameer Poyil", "Varshini Nagabooshanam", "Vijay Pamula", "Suresh Dharnala",
    "Santhos Kumar Reddy", "Vinayakam Mani", "Sean Nash",
]
SW_TEAM = ["Nicholas Ramirez", "Stephen Quong", "Dara Navaei", "Behrouz NematiPour"]
SYS_TEAM = [
    "Eliza Petersen", "Caitlynn Chang", "Christina Heine", "Abhijit Barman",
    "Chris Yu", "Vitas Buenaventura", "Emiline Hernandez",
]
TEAMS = [FW_TEAM, SWVV_TEAM, SW_TEAM, SYS_TEAM]
TEAM_NAMES = ["FW", "SWVV", "SW", "SYS"]
_SYS_TEAM_IDX = len(TEAMS) - 1
_UNMATCHED_TEAM_IDX = len(TEAMS)

NICKNAME_MAP = {
    "Tejaskumar Patel": "Tejas",
    "Nicholas Ramirez": "Nico",
    "Thomas Lippold": "Tom",
    "Jashwant Gantyada": "Jashwant",
    "Behrouz NematiPour": "Behrouz",
    "Sean Nash": "Sean",
    "Vinayakam Mani": "Vinay",
    "Raghu Kallala": "Raghu",
    "Tiffany Mejia": "Tiffany",
    "Zoltan Miskolci": "Zoltan",
    "Sarina Cheung": "Sarina",
    "Tisha Patel": "Tisha",
    "Ethan Nguyen": "Ethan",
    "Arpita Srivastava": "Arpita",
    "Michael Garthwaite": "Michael",
    "Praneeth Bunne": "Praneeth",
    "Sameer Poyil": "Sameer",
    "Varshini Nagabooshanam": "Varshini",
    "Vijay Pamula": "Vijay",
    "Suresh Dharnala": "Suresh",
    "Dara Navaei": "Dara",
    "Eliza Petersen": "Eliza",
    "Stephen Quong": "Stephen",
    "Christina Heine": "Christina",
    "Caitlynn Chang": "Caitlynn",
    "Santhos Kumar Reddy": "Santhos",
    "Abhijit Barman": "Abhijit",
    "Chris Yu": "Chris",
    "Vitas Buenaventura": "Vitas",
    "Emiline Hernandez": "Emiline",
}

# Build a fast lookup from full name or nickname to team index.
_ASSIGNEE_TO_TEAM_IDX: dict = {}
for _ti, _team in enumerate(TEAMS):
    for _full in _team:
        _ASSIGNEE_TO_TEAM_IDX[_full] = _ti
        _nick = NICKNAME_MAP.get(_full)
        if _nick:
            _ASSIGNEE_TO_TEAM_IDX[_nick] = _ti
del _ti, _team, _full, _nick


def _assignee_nickname(display_name: str) -> str:
    # Return the configured nickname for a display name, or the first name.
    if not display_name:
        return "Unassigned"
    return NICKNAME_MAP.get(display_name, display_name.split()[0])


def _team_for_nickname(nickname):
    # Look up the (team_name, team_index) pair for the given nickname.
    idx = _ASSIGNEE_TO_TEAM_IDX.get(nickname)
    if idx is None:
        return None, _UNMATCHED_TEAM_IDX
    return TEAM_NAMES[idx], idx
