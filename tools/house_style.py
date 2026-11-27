"""The VBA house style in one place: its vocabulary and the switch that turns it on.

``HOUSE_STYLE_ACTIVE`` is the single switch for the style described in
docs/VBA_HOUSE_STYLE.md. While it is False the repository's modules are held to
the previous normal form: the formatter's ``--check`` and the checks that name
error-handling labels read the previous spellings, and the checks that need
the new header fields and sections do not scan the tree. They still run their
own fixtures, which force the switch on, so every rule stays proved. Reformat
the modules with ``python3 tools/reformat.py --style house --write`` and set
the constant to True in the same change to move the repository to the style.
"""
import re
from typing import NamedTuple

HOUSE_STYLE_ACTIVE = False

# Every banner line is an apostrophe and 78 dashes.
RULE_WIDTH = 79
RULE_DASH = "'" + "-" * (RULE_WIDTH - 1)
RULE_EQUALS = "'" + "=" * (RULE_WIDTH - 1)
TITLE_COLUMNS = RULE_WIDTH - 1

# The nine documentation fields, in the order a header lists them.
FIELDS = ("PURPOSE", "WHY THIS EXISTS", "INPUTS", "RETURNS", "BEHAVIOR",
          "ERROR POLICY", "DEPENDENCIES", "NOTES", "UPDATED")
STANDARD_BEFORE_EXTRAS = FIELDS[:-1]
# A Private procedure with at most this many code lines may use a compact header.
COMPACT_MAX_CODE_LINES = 15
COMPACT_FIELDS = ("PURPOSE", "UPDATED")
# Extra header sections the repository defines. CALLED FROM follows NOTES; the
# dated history sits after every other section and before UPDATED.
CALLED_FROM = "CALLED FROM"
HISTORY = "HISTORY"
# Spellings of a standard field label that the formatter rewrites.
FIELD_ALIASES = {"WHY": "WHY THIS EXISTS"}

# Body sections of a procedure, in the order they appear.
DECLARE = "DECLARE"
INITIALIZE = "INITIALIZE"
EXIT_PROCEDURE = "EXIT PROCEDURE"
ERROR_HANDLER = "ERROR HANDLER"
SECTION_RANK = {DECLARE: 0, INITIALIZE: 1, EXIT_PROCEDURE: 3, ERROR_HANDLER: 4}
TASK_RANK = 2
# Titles the formatter replaces by the standard ones when they hold the label
# or the Exit statement the standard section is defined by.
EXIT_TITLES = frozenset({"RETURN SUCCESS", "RETURN", "SAFE EXIT", "CLEAN EXIT", "EXIT",
                         "SUCCESS", "RETURN VALUE AND EXIT"})
HANDLER_TITLES = frozenset({"FAIL", "CLEAN FAIL", "ERROR", "HANDLER", "ERROR HANDLING",
                            "ERR HANDLER"})

PROC_CONSTANT_OLD = "PROC"
PROC_CONSTANT = "PROC_NAME"

# Indentation: a statement sits at CODE_INDENT plus INDENT_STEP per open block,
# the one-line comment that introduces it at INTENT_INDENT plus the same steps,
# and a declaration under a section banner at DECLARATION_INDENT.
INTENT_INDENT = 4
CODE_INDENT = 8
INDENT_STEP = 4
DECLARATION_INDENT = 4


class LabelSet(NamedTuple):
    """The names of the labels a procedure uses to leave and to fail."""
    handler: str
    safe_exit: str
    clean_exit: str
    clean_fail: str
    fail_num: str

    def names(self) -> frozenset[str]:
        return frozenset(self)


PREVIOUS_LABELS = LabelSet("Err_Handler", "Safe_Exit", "Clean_Exit", "Clean_Fail", "Fail_Num")
HOUSE_LABELS = LabelSet("ErrorHandler", "SafeExit", "CleanExit", "CleanFail", "FailNum")

# Spellings that the house-style formatter rewrites, mapped to the house name.
# "Fail" is a label an older form of the formatter produced from.
LABEL_RENAMES = {
    "Err_Handler": HOUSE_LABELS.handler,
    "Fail": HOUSE_LABELS.handler,
    "Safe_Exit": HOUSE_LABELS.safe_exit,
    "Clean_Exit": HOUSE_LABELS.clean_exit,
    "Clean_Fail": HOUSE_LABELS.clean_fail,
    "Fail_Num": HOUSE_LABELS.fail_num,
}


def active(force: bool | None = None) -> bool:
    """Whether the style is on; ``force`` lets a fixture exercise either state."""
    return HOUSE_STYLE_ACTIVE if force is None else force


def labels(force: bool | None = None) -> LabelSet:
    """The label vocabulary the tree uses in the given (or current) state."""
    return HOUSE_LABELS if active(force) else PREVIOUS_LABELS


def centred_title(title: str) -> str:
    """A procedure or group title centred in the 78 columns after the apostrophe."""
    return "'" + " " * max(0, (TITLE_COLUMNS - len(title)) // 2) + title


WORD = re.compile(r"[A-Z]+(?![a-z])|[A-Z]?[a-z]+|[0-9]+")


def title_from_identifier(name: str) -> str:
    """Upper-case words of an identifier: UI_SetExcelUI becomes UI SET EXCEL UI."""
    return " ".join(word.upper() for part in name.split("_") for word in WORD.findall(part))
