# run_utils.py
"""
Low-level helpers for extracting text and run-property states from lxml run elements.

All functions operate on individual ``<w:r>`` (run) or ``<w:rPr>``
(run-properties) elements and have no side effects.
"""

import re

from lxml import etree

from constants import NS, NSC, WORD_HIGHLIGHT_COLORS, LINE_BREAK_DELIMITER

# Validates a resolved hex colour string: exactly 6 or 8 hex characters, no prefix.
_ARGB_RE = re.compile(r"^[A-Fa-f0-9]{6}$|^[A-Fa-f0-9]{8}$")


def is_line_break(br_el) -> bool:
    """Return True if *br_el* is a manual line break (textWrapping), not a page or column break."""
    br_type = br_el.get(NSC["w"] + "type", "textWrapping")
    return br_type == "textWrapping"


def run_text(run) -> str:
    """Collect text from a run, emitting LINE_BREAK_DELIMITER for manual line breaks."""
    parts = []
    for child in run:
        local = etree.QName(child.tag).localname
        if local == "t":
            parts.append(child.text or "")
        elif local == "br" and is_line_break(child):
            parts.append(LINE_BREAK_DELIMITER)
    return "".join(parts)


def del_text(del_run) -> str:
    """Collect text from a deleted run, emitting LINE_BREAK_DELIMITER for manual line breaks."""
    parts = []
    for child in del_run:
        local = etree.QName(child.tag).localname
        if local == "delText":
            parts.append(child.text or "")
        elif local == "br" and is_line_break(child):
            parts.append(LINE_BREAK_DELIMITER)
    return "".join(parts)


def has_rpr_property(rpr, *tag_names: str) -> bool:
    """
    Return True if the run-properties element contains any of the given ``w:`` child tags.

    Examples::

        has_rpr_property(rpr, "highlight", "shd")   # highlight detection
        has_rpr_property(rpr, "strike", "dstrike")  # strike-through detection
    """
    if rpr is None:
        return False
    return any(rpr.find(f"w:{tag}", NS) is not None for tag in tag_names)


def get_rpr_property_states(run, *tag_names: str) -> tuple[bool, bool]:
    """
    Return ``(old_state, new_state)`` for any rPr property identified by *tag_names*.
    """
    rpr, old_rpr = _get_old_rpr(run)
    return has_rpr_property(old_rpr, *tag_names), has_rpr_property(rpr, *tag_names)


def get_rpr_attr_val(
    rpr, 
    tag: str, 
    attr: str, 
    ignore_vals: tuple[str, ...] = ()
) -> str | None:
    """
    Extract an attribute value from a named child element of a ``<w:rPr>``.
    """
    if rpr is None:
        return None
    el = rpr.find(f"w:{tag}", NS)
    if el is None:
        return None
    val = el.get(NSC["w"] + attr, "")
    return None if val.lower() in ignore_vals or not val else val


def _get_old_rpr(run):
    """Return the old ``<w:rPr>`` from inside ``<w:rPrChange>``, or ``None``."""
    rpr = run.find("w:rPr", NS)
    if rpr is None:
        return None, None
    rpr_change = rpr.find("w:rPrChange", NS)
    old_rpr = rpr_change.find("w:rPr", NS) if rpr_change is not None else None
    return rpr, old_rpr


def resolve_color(val: str | None) -> str | None:
    """
    Resolve a colour value to a 6- or 8-digit RGB/ARGB hex string.
    """
    if not val:
        return None
    resolved = WORD_HIGHLIGHT_COLORS.get(val)
    hex_val = (resolved or val).lstrip("#").upper()
    return hex_val if _ARGB_RE.match(hex_val) else None


def get_rpr_val_states(
    run,
    *specs: tuple[str, str, tuple[str, ...]],
) -> tuple[str | None, str | None]:
    """
    Return ``(old_val, new_val)`` for a run property defined by one or more specs.

    Each spec is ``(tag, attr, ignore_vals)`` passed directly to
    ``get_rpr_attr_val``.  Specs are tried in order; the first non-None
    result is used.  This covers both simple properties (single spec) and
    fallback chains (e.g. highlight → shd).

    Examples::

        get_rpr_val_states(run, ("color", "val", ("auto",)))
        get_rpr_val_states(run, ("highlight", "val", ("none", "auto")), ("shd", "fill", ("auto",)))
    """
    rpr, old_rpr = _get_old_rpr(run)

    def _val(r):
        for tag, attr, ignore_vals in specs:
            v = get_rpr_attr_val(r, tag, attr, ignore_vals=ignore_vals)
            if v is not None:
                return v
        return None

    return _val(old_rpr), _val(rpr)
