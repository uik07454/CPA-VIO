# rpr_change_utils.py
"""
Helpers for detecting, classifying, and merging w:rPrChange (run property
change) tracked-change elements.

Public API used by para_utils.py / extractor.py:
    - get_rpr_change_meta   — extract author/date from a run's w:rPrChange
    - can_merge_rpr_runs    — decide whether two consecutive rPrChange runs can be merged
    - flush_rpr_buffer      — emit one ChangeRecord for all runs in the buffer
"""

from datetime import datetime, timezone
from lxml import etree

from constants import (
    NS, NSC, ChangeType,
    DATE_FORMATS,
    MERGE_TOLERANCE_SECONDS,
    HIGHLIGHT_SPECS,
    COLOR_SPECS,
)
from models import ChangeRecord
from run_utils import run_text, get_rpr_property_states, get_rpr_val_states, resolve_color
from classifier import classify_change_type, RprDetection, ValDetection


def _parse_date(date_str: str):
    """Parse an ISO-8601 date string into a datetime object, or return None."""
    if not date_str:
        return None
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(date_str, fmt).replace(tzinfo=timezone.utc)
        except ValueError:
            pass
    return None


def get_rpr_change_meta(run) -> tuple[str, str] | None:
    """Extract author and date metadata from a **w:rPrChange** element in *run*."""
    rpr = run.find("w:rPr", NS)
    if rpr is None:
        return None
    rpr_change = rpr.find("w:rPrChange", NS)
    if rpr_change is None:
        return None
    author = rpr_change.get(NSC["w"] + "author", "")
    date_str = rpr_change.get(NSC["w"] + "date", "")
    return author, date_str


def _dates_within_tolerance(date_str_a: str, date_str_b: str) -> bool:
    """
    Return True if the two date strings represent datetimes 
    within *_MERGE_TOLERANCE_SECONDS* of each other.
    """
    dt_a = _parse_date(date_str_a)
    dt_b = _parse_date(date_str_b)
    if dt_a is None or dt_b is None:
        return date_str_a == date_str_b
    return abs((dt_a - dt_b).total_seconds()) <= MERGE_TOLERANCE_SECONDS


def _classify_run_change_type(run) -> str:
    """
    Classify the change type of a single w:r carrying w:rPrChange.

    Detection priority:
      1. Strike-through Added/Removed  — w:strike or w:dstrike presence changed
      2. Text Highlight Added/Removed  — w:highlight or w:shd presence changed
      3. Text Highlight Colour Changed — highlight/shd value changed between old and new
      4. Font Colour Changed           — w:color/@w:val added, removed, or changed
      5. Format Change                 — any other rPrChange
      6. Unknown                       — rPrChange present but nothing classifiable
    """
    old_st, new_st = get_rpr_property_states(run, "strike", "dstrike")
    old_hl, new_hl = get_rpr_property_states(run, "highlight", "shd")
    old_hlv, new_hlv = (resolve_color(v) for v in get_rpr_val_states(run, *HIGHLIGHT_SPECS))
    old_fc,  new_fc  = get_rpr_val_states(run, *COLOR_SPECS)
    is_fmt = get_rpr_change_meta(run) is not None
    return classify_change_type(
        rpr_detections=[
            RprDetection(old_st, new_st, ChangeType.STRIKETHROUGH_ADDED, ChangeType.STRIKETHROUGH_REMOVED),
            RprDetection(old_hl, new_hl, ChangeType.HIGHLIGHT_ADDED,     ChangeType.HIGHLIGHT_REMOVED),
        ],
        val_detections=[
            ValDetection(old_hlv, new_hlv, ChangeType.HIGHLIGHT_CHANGED),
            ValDetection(old_fc,  new_fc,  ChangeType.FONT_COLOR_CHANGED),
        ],
        is_format_only=is_fmt,
    )


def _get_rpr_merge_key(run) -> tuple:
    """
    Return a value fingerprint for the property change on *run*.
    """
    change_type = _classify_run_change_type(run)
    if change_type in (ChangeType.HIGHLIGHT_ADDED, ChangeType.HIGHLIGHT_REMOVED, ChangeType.HIGHLIGHT_CHANGED):
        return tuple(resolve_color(v) for v in get_rpr_val_states(run, *HIGHLIGHT_SPECS))
    if change_type == ChangeType.FONT_COLOR_CHANGED:
        return get_rpr_val_states(run, *COLOR_SPECS)
    return ()


def can_merge_rpr_runs(run_a, run_b) -> bool:
    """
    Return True if two consecutive w:r elements carrying w:rPrChange can be
    merged into a single ChangeRecord.

    Merge criteria (all must hold):
      1. Same author on their w:rPrChange
      2. Dates within _MERGE_TOLERANCE_SECONDS of each other
      3. Same classified change type (e.g. both "Font Colour Changed")
         — prevents mixing different change types in one record
      4. Same property value fingerprint (e.g. both green→auto, not green→auto vs. red→blue)
         — prevents merging runs whose values differ within the same change type
    """
    meta_a = get_rpr_change_meta(run_a)
    meta_b = get_rpr_change_meta(run_b)
    if meta_a is None or meta_b is None:
        return False
    author_a, date_a = meta_a
    author_b, date_b = meta_b
    if author_a != author_b or not _dates_within_tolerance(date_a, date_b):
        return False
    if _classify_run_change_type(run_a) != _classify_run_change_type(run_b):
        return False
    return _get_rpr_merge_key(run_a) == _get_rpr_merge_key(run_b)


def flush_rpr_buffer(
    buffer: list[etree._Element],
    current_heading: str,
    records: list[ChangeRecord]
) -> None:
    """
    Emit one ChangeRecord for all runs in the buffer.

    All runs in the buffer are guaranteed to share the same change type
    (enforced by _can_merge_rpr_runs), so the type is derived from the
    first run only.  The text is the concatenation of all runs.
    """
    if not buffer:
        return
    text = "".join(run_text(r) for r in buffer)
    change_type = _classify_run_change_type(buffer[0])

    style_meta = {}
    merge_key = _get_rpr_merge_key(buffer[0])
    if len(merge_key) == 2:
        role = "font" if change_type == ChangeType.FONT_COLOR_CHANGED else "highlight"
        style_meta = {"color_role": role, "old_color": merge_key[0], "new_color": merge_key[1]}

    records.append(
        ChangeRecord(
            heading=current_heading,
            old_text=text,
            new_text=text,
            change_type=change_type,
            style_meta=style_meta,
        )
    )
    buffer.clear()
