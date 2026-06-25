# constants.py
"""
Project-wide constants: XML namespace map, Clark-notation prefixes, and the
ChangeType registry.

All modules should import from here rather than hard-coding namespace URIs or
change-type label strings.
"""

TO_PARSED_XML_PARTS = [
    "document",
    "styles",
    "footnotes",
    "endnotes",
    # Headers and footers are parsed separately by parse_headers_footers().
]

NS = {
    # Core WordprocessingML
    "w":   "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    # Drawing / positioning
    "wp":  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    # DrawingML content types
    "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "c":   "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    # Legacy VML
    "v":   "urn:schemas-microsoft-com:vml",
    "o":   "urn:schemas-microsoft-com:office:office",
    # Markup compatibility
    "mc":  "http://schemas.openxmlformats.org/markup-compatibility/2006",
    # Misc
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "relationships": "http://schemas.openxmlformats.org/package/2006/relationships",
}

# Pre-built Clark-notation prefixes for every namespace in NS.
# Use NSC["x"] wherever you need to construct an element or attribute name.
# This avoids scattering f"{NS['x']}" or hard-coded URI strings across the code.
NSC = {key: f"{{{uri}}}" for key, uri in NS.items()}


# Mapping from Word's named highlight colour values
# (w:highlight/@w:val) to 6-digit RGB hex strings
WORD_HIGHLIGHT_COLORS: dict[str, str] = {
    "yellow":       "#FFFF00",
    "green":        "#00FF00",
    "cyan":         "#00FFFF",
    "magenta":      "#FF00FF",
    "blue":         "#0000FF",
    "red":          "#FF0000",
    "darkBlue":     "#000080",
    "darkCyan":     "#008080",
    "darkGreen":    "#008000",
    "darkMagenta":  "#800080",
    "darkRed":      "#800000",
    "darkYellow":   "#808000",
    "darkGray":     "#808080",
    "lightGray":    "#C0C0C0",
    "black":        "#000000"
}

# ---- Size constants for conversion -----------------------------------------
EMU_PER_PX = 9525
EMU_PER_PT = 12700

# ---- Output / Excel rendering ----------------------------------------------

# Dimensions (in pixels) for embedded images in Excel cells.
IMAGE_WIDTH_PX  = 450
IMAGE_HEIGHT_PX = 450

# Colours used for inline diff highlighting in "Modify Content" rows.
DIFF_DEL_COLOR = "FF0000"
DIFF_INS_COLOR = "00B050"
DIFF_EQ_COLOR  = "000000"

# ---- Assets ----------------------------------------------------------------

# Directory for pre-rendered GUI assets (splash video, icons, etc.).
from pathlib import Path
import sys

# Resolves to the EXE's actual directory when frozen by PyInstaller (one-file or one-folder),
# or the source file's directory when running normally.
# NOTE: sys._MEIPASS is the temp extraction dir for one-file builds — we must NOT use it
# for external files (assets, data, output). We use sys.executable instead, which always
# points to the real EXE location.
BASE_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent

ASSETS_DIR        = BASE_DIR / "assets"

SPLASH_ASSETS_DIR = ASSETS_DIR / "splash_assets"
SPLASH_ASSETS_DIR.mkdir(exist_ok=True)

SPLASH_VIDEO_NAME = "splash.mp4"
SPLASH_VIDEO_PATH = SPLASH_ASSETS_DIR / SPLASH_VIDEO_NAME

APP_ICON_PATH = ASSETS_DIR / "DocGear.ico"

# Directory containing CPA Excel templates.
CPA_TEMPLATE_DIR = ASSETS_DIR / "CPA_template"
CPA_TEMPLATE_DIR.mkdir(exist_ok=True)


# ---- Template-based Excel output -------------------------------------------

# Preferred sheet name to write output into; falls back to the first sheet.
DEFAULT_TEMPLATE_SHEET = "CPA_JPN"

# Output mode identifiers.
OUTPUT_MODE_SCRATCH   = "scratch"    # generate a brand-new workbook
OUTPUT_MODE_TEMPLATE  = "template"   # copy a template then fill it in

# Default column letters used when no profile exists for a (template, sheet) pair.
DEFAULT_TEMPLATE_COLUMNS: list[str] = ["A", "B", "C", "D"]


# ---- GUI validation ranges -------------------------------------------------

# Valid pixel range for embedded image dimensions.
IMAGE_PX_MIN = 50
IMAGE_PX_MAX = 500

# Valid point range for image row height.
IMAGE_ROW_HEIGHT_MIN = 40
IMAGE_ROW_HEIGHT_MAX = 400

# ---- Image size mode -------------------------------------------------------

# Mode strings for image sizing behaviour in Excel output.
IMAGE_SIZE_MODE_FIXED    = "fixed"     # all images use the user-configured size
IMAGE_SIZE_MODE_ORIGINAL = "original"  # each image uses its wp:extent size; fallback to configured
IMAGE_SIZE_MODE          = IMAGE_SIZE_MODE_ORIGINAL   # default mode


# ---- rPr change detection --------------------------------------------------

# ISO-8601 date formats accepted when parsing w:rPrChange timestamps.
DATE_FORMATS = (
    "%Y-%m-%dT%H:%M:%SZ",
    "%Y-%m-%dT%H:%M:%S",
)

# Maximum gap (in seconds) between two rPrChange timestamps that are still
# considered the same edit for merging purposes.
MERGE_TOLERANCE_SECONDS = 3

# Specs used to detect highlight colour changes (w:highlight / w:shd fallback).
HIGHLIGHT_SPECS = (("highlight", "val", ("none", "auto")), ("shd", "fill", ("auto",)))

# Specs used to detect font colour changes (w:color).
COLOR_SPECS = (("color", "val", ("auto",)),)


# ---- Footnote / Endnote labels ---------------------------------------------

# Suffix labels appended to the heading for footnote/endnote changes.
FOOTNOTE_LABEL = "Footnote"
ENDNOTE_LABEL  = "Endnote"


# ---- Table delimiters ------------------------------------------------------

# Actual delimiter and output delimiter used to join cells in a row (stored: ¤).
CELL_DELIMITER = "¤"
CELL_OUTPUT_DELIMITER = " | "


# ---- New line delimiter --------------------------------------------------

# Actual delimiter and output delimiter used to join paragraphs (stored: ¶).
PARA_DELIMITER = "¶"
PARA_OUTPUT_DELIMITER = "\n"

# Actual delimiter and output delimiter used to represent <w:br> within a paragraph (stored: ↵).
LINE_BREAK_DELIMITER        = "↵"
LINE_BREAK_OUTPUT_DELIMITER = "\n"


# ---- Extraction defaults ---------------------------------------------------

# Default heading used when no heading has been encountered yet.
DEFAULT_HEADING = "Link to CP"

# Placeholder when a field has no value.
EMPTY_FIELD = "N/A"

# Prefix and zero-pad width for generated bookmark IDs (e.g. "docChg_00000001").
BOOKMARK_PREFIX    = "docChg_"
BOOKMARK_ID_WIDTH  = 8


def make_bookmark_id(n: int) -> str:
    """Return a bookmark ID string for the given integer index."""
    return f"{BOOKMARK_PREFIX}{n:0{BOOKMARK_ID_WIDTH}d}"


class ChangeType:
    """
    Central registry of all change-type label strings used throughout the project.

    Acts as a single source of truth for all valid change types.  Every
    ChangeRecord.change_type value must be one of these constants.
    """

    # ---- Text content -------------------------------------------------------
    MODIFY_CONTENT = "Modify Content"
    DELETE_CONTENT = "Delete Content"
    ADD_CONTENT    = "Add Content"

    # ---- Run-level format / specific rpr changes --------------
    STRIKETHROUGH_ADDED   = "Strike-through Added"
    STRIKETHROUGH_REMOVED = "Strike-through Removed"
    HIGHLIGHT_ADDED       = "Text Highlight Colour Added"
    HIGHLIGHT_REMOVED     = "Text Highlight Colour Removed"
    HIGHLIGHT_CHANGED     = "Text Highlight Colour Changed"
    FONT_COLOR_CHANGED    = "Font Colour Changed"
    FORMAT_CHANGE         = "Format Change"
    STYLE_CHANGE          = "Paragraph Style Change"
    UNKNOWN               = "**Unknown**\nReview manually!"

    # ---- Images / embedded objects ------------------------------------------
    ADD_OBJECT    = "Add non-text Object"
    DELETE_OBJECT = "Delete non-text Object"
    MODIFY_OBJECT = "Modify non-text Object"

    # ---- Footnote / Endnote structural changes ------------------------------
    ADD_NOTE    = "Add Note"
    DELETE_NOTE = "Delete Note"

    # ---- Caption changes ----------------------------------------------------
    CAPTION_CHANGE = "Caption Change"

    # ---- Table structure ----------------------------------------------------
    ADD_TABLE    = "Add Table"
    DELETE_TABLE = "Delete Table"
    ADD_ROW      = "Add Row"
    DELETE_ROW   = "Delete Row"
    ADD_CELL     = "Add Cell"
    DELETE_CELL  = "Delete Cell"
    MERGE_CELL   = "Merge Cell"


# ---- Change type category grouping (used by GUI) ---------------------------

# Maps display category names to their ChangeType members.
# Order determines the display order in the GUI tree.
# Add new categories here when new ChangeType entries are introduced.
CHANGE_TYPE_CATEGORIES: dict[str, list[str]] = {
    "Content Changes": [
        ChangeType.ADD_CONTENT,
        ChangeType.DELETE_CONTENT,
        ChangeType.MODIFY_CONTENT,
    ],
    "Strikethrough Changes": [
        ChangeType.STRIKETHROUGH_ADDED,
        ChangeType.STRIKETHROUGH_REMOVED,
    ],
    "Highlight Changes": [
        ChangeType.HIGHLIGHT_ADDED,
        ChangeType.HIGHLIGHT_REMOVED,
        ChangeType.HIGHLIGHT_CHANGED,
    ],
    "Font Color Changes": [
        ChangeType.FONT_COLOR_CHANGED,
    ],
    "Format Changes": [
        ChangeType.FORMAT_CHANGE,
    ],
    "Style Changes": [
        ChangeType.STYLE_CHANGE,
    ],
    "Object Changes": [
        ChangeType.ADD_OBJECT,
        ChangeType.DELETE_OBJECT,
        ChangeType.MODIFY_OBJECT,
    ],
    "Note Changes": [
        ChangeType.ADD_NOTE,
        ChangeType.DELETE_NOTE,
    ],
    "Caption Changes": [
        ChangeType.CAPTION_CHANGE,
    ],
    "Table Changes": [
        ChangeType.ADD_TABLE,
        ChangeType.DELETE_TABLE,
        ChangeType.ADD_ROW,
        ChangeType.DELETE_ROW,
        ChangeType.ADD_CELL,
        ChangeType.DELETE_CELL,
        ChangeType.MERGE_CELL,
    ],
}

# Flat set of all selectable change types (excludes UNKNOWN, which is always included).
ALL_CHANGE_TYPES: frozenset[str] = frozenset(
    ct for types in CHANGE_TYPE_CATEGORIES.values() for ct in types
)
