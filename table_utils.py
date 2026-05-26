# table_utils.py
"""
Helpers for reading text out of table cells/rows, detecting vertical-merge
span geometry, and extracting tracked changes from table structures.

Public API used by extractor.py:
    - CELL_PARA_DELIMITER         — delimiter used to join paragraphs inside a cell (stored: ¶)
    - CELL_DELIMITER              — delimiter used to join cells in a row (stored: ¤)

    - first_para_in_tr           — first w:p inside the first w:tc of a w:tr
    - check_whole_table_change    — collapse all-ins/all-del table into one record
    - process_row                 — handle row-level ins/del
    - process_cell                — handle cell-level changes
    - extract_table_changes       — walk a <w:tbl> and emit all change records
"""

import re
from typing import TYPE_CHECKING

from constants import (
    NS, NSC, ChangeType,
    PARA_DELIMITER, CELL_DELIMITER,
)
from models import ChangeRecord
from classifier import classify_change_type
from para_utils import get_paragraph_text
from para_utils import unwrap_sdts

if TYPE_CHECKING:
    from extractor import _Counter

# Regex patterns to normalise w:cellMerge vMerge attribute values.
# Word may use short ("rest", "cont") or long ("restart", "continue") forms.
_VMERGE_RESTART_RE = re.compile(r"^rest(art)?$", re.IGNORECASE)
_VMERGE_CONT_RE    = re.compile(r"^cont(inue)?$", re.IGNORECASE)


def _get_cell_text(
    tc, 
    include_del: bool = False, 
    include_ins: bool = True
) -> str:
    """
    Collect all paragraph text inside a <w:tc>, joining 
    paragraphs with the CELL_PARA_DELIMITER.
    """
    parts = []
    for para in tc.findall("w:p", NS):
        para_text = get_paragraph_text(para, include_del=include_del, include_ins=include_ins)
        if para_text:
            parts.append(para_text)
    return PARA_DELIMITER.join(parts)


def _get_row_text(
    tr, 
    include_del: bool = False, 
    include_ins: bool = True
) -> str:
    """
    Collect all cell texts in a <w:tr> and format them as aligned paragraph lines.
    """
    # 1d list, [[cell1_para1¶cell1_para2], [cell2_para1¶cell2_para2], ...]
    cell_texts = [_get_cell_text(tc, include_del, include_ins) for tc in tr.findall("w:tc", NS)]
    # 2d list, [[cell1_para1, cell1_para2], [cell2_para1, cell2_para2], ...]
    para_lists = [ct.split(PARA_DELIMITER) for ct in cell_texts]
    max_lines = max((len(p) for p in para_lists), default=1)
    # Width of the widest paragraph in each cell column (for alignment)
    col_widths = [
        max((len(p[i]) for i in range(len(p))), default=0)
        for p in para_lists
    ]
    last_col = len(para_lists) - 1
    lines = []
    # transpose the 2d list of paragraphs and join each line with CELL_DELIMITER
    # cell1_para1 | cell2_para1 | cell3_para1
    # cell1_para2 | cell2_para2 | cell3_para2
    for line_idx in range(max_lines):
        parts = []
        for col_idx, p in enumerate(para_lists):
            entry = p[line_idx] if line_idx < len(p) else ""
            # Pad all columns except the last so the delimiter aligns vertically
            if col_idx < last_col:
                entry = entry.ljust(col_widths[col_idx])
            parts.append(entry)
        lines.append(CELL_DELIMITER.join(parts))
    return PARA_DELIMITER.join(lines)


def _first_para_in_tc(tc):
    """Return the first <w:p> inside a <w:tc>, or None."""
    return tc.find("w:p", NS)


def first_para_in_tr(tr):
    """Return the first <w:p> inside the first <w:tc> of a <w:tr>, or None."""
    tc = tr.find("w:tc", NS)
    if tc is not None:
        return _first_para_in_tc(tc)
    return None


def _get_cell_merge_vmerge(tc) -> str | None:
    """
    Return vMerge role: "restart" or "continue".
    If the cell carries a w:cellMerge tracked-change element
    """
    tc_pr = tc.find("w:tcPr", NS)
    if tc_pr is None:
        return None
    cell_merge = tc_pr.find("w:cellMerge", NS)
    if cell_merge is None:
        return None
    val = cell_merge.get(NSC["w"] + "vMerge", "")
    if _VMERGE_RESTART_RE.match(val):
        return "restart"
    if _VMERGE_CONT_RE.match(val):
        return "continue"
    return None


def _is_vmerge_continuation(tc) -> bool:
    """
    Return True if this cell is a vertical-merge *continuation* cell.
    """
    tc_pr = tc.find("w:tcPr", NS)
    if tc_pr is None:
        return False
    vmerge = tc_pr.find("w:vMerge", NS)
    if vmerge is None:
        return False
    val = vmerge.get(NSC["w"] + "val", "")

    return val != "restart"


def _is_row_vmerge_continuation(tr) -> bool:
    """
    Return True if *every* cell in this row is a vertical-merge continuation.
    """
    cells = tr.findall("w:tc", NS)
    if not cells:
        return False
    return all(_is_vmerge_continuation(tc) for tc in cells)


# ---------------------------------------------------------------------------
# Table-level tracked-change extraction
# ---------------------------------------------------------------------------

_CONTENT_TYPES = frozenset({
    ChangeType.ADD_CONTENT,
    ChangeType.DELETE_CONTENT,
    ChangeType.MODIFY_CONTENT,
})


def check_whole_table_change(
    real_rows: list,
    table_heading: str,
    counter: "_Counter",
) -> list[ChangeRecord] | None:
    """
    If every non-continuation row is inserted or every one is deleted,
    collapse the whole table into a single ADD_TABLE / DELETE_TABLE record.
    Returns None if the table has mixed changes.
    """
    all_ins = all(tr.find("w:trPr/w:ins", NS) is not None for tr in real_rows)
    all_del = all(tr.find("w:trPr/w:del", NS) is not None for tr in real_rows)
    if not (all_ins or all_del):
        return None

    is_ins = all_ins
    change_type = classify_change_type(
        structural_type=ChangeType.ADD_TABLE if is_ins else ChangeType.DELETE_TABLE
    )
    return [ChangeRecord(
        heading=table_heading,
        old_text="" if is_ins else "[Table]",
        new_text="[Table]" if is_ins else "",
        change_type=change_type,
        bookmark_id=counter.next_id(),
        target_element=first_para_in_tr(real_rows[0]),
    )]


def process_row(
    tr,
    table_heading: str,
    counter: "_Counter",
    records: list[ChangeRecord],
) -> bool:
    """
    If the row is inserted or deleted as a whole, append one ADD_ROW /
    DELETE_ROW record and return True (caller should skip cell processing).
    Continuation rows of a vertical-merge span are silently skipped (returns True).
    """
    tr_pr = tr.find("w:trPr", NS)
    row_inserted = tr_pr is not None and tr_pr.find("w:ins", NS) is not None
    row_deleted  = tr_pr is not None and tr_pr.find("w:del", NS) is not None

    if not (row_inserted or row_deleted):
        return False

    if _is_row_vmerge_continuation(tr):
        counter.next_id()
        return True

    row_text = _get_row_text(tr, include_del=row_deleted)
    if row_inserted:
        change_type = classify_change_type(structural_type=ChangeType.ADD_ROW)
        old_text, new_text = "", row_text
    else:
        change_type = classify_change_type(structural_type=ChangeType.DELETE_ROW)
        old_text, new_text = row_text, ""

    records.append(ChangeRecord(
        heading=table_heading,
        old_text=old_text,
        new_text=new_text,
        change_type=change_type,
        bookmark_id=counter.next_id(),
        target_element=first_para_in_tr(tr),
    ))
    return True


def process_cell(
    tc,
    table_heading: str,
    counter: "_Counter",
    include_sdt: bool,
    records: list[ChangeRecord],
    image_map: dict[str, bytes] = {},
) -> None:
    """
    Handle one w:tc.  Dispatches to the first matching case:

    1. Cell inserted / deleted (w:cellIns / w:cellDel) — one ADD_CELL /
       DELETE_CELL record.  Continuation vMerge cells are silently skipped.
    2. Tracked vertical merge (w:cellMerge) — one MERGE_CELL record for the
       restart cell; continuation cells are silently skipped.
    3. Mixed changes — content changes are aggregated into one record;
       non-content changes (objects, highlights, format, style) are always
       emitted individually alongside it so nothing is silently discarded.
    """
    # Lazy import to avoid circular dependency (extractor imports table_utils).
    from extractor import extract_changes_from_paragraph

    tc_pr = tc.find("w:tcPr", NS)
    cell_inserted = tc_pr is not None and tc_pr.find("w:cellIns", NS) is not None
    cell_deleted  = tc_pr is not None and tc_pr.find("w:cellDel", NS) is not None

    # ---- Case 1: whole cell insertion / deletion --------------------------
    if cell_inserted or cell_deleted:
        if _is_vmerge_continuation(tc):
            counter.next_id()
            return
        cell_text = _get_cell_text(tc, include_del=cell_deleted)
        if cell_inserted:
            change_type = classify_change_type(structural_type=ChangeType.ADD_CELL)
            old_text, new_text = "", cell_text
        else:
            change_type = classify_change_type(structural_type=ChangeType.DELETE_CELL)
            old_text, new_text = cell_text, ""
        records.append(ChangeRecord(
            heading=table_heading,
            old_text=old_text,
            new_text=new_text,
            change_type=change_type,
            bookmark_id=counter.next_id(),
            target_element=_first_para_in_tc(tc),
        ))
        return

    # ---- Case 2: tracked vertical merge (w:cellMerge) ----------------------
    cell_merge_val = _get_cell_merge_vmerge(tc)
    if cell_merge_val == "continue":
        counter.next_id()
        return
    if cell_merge_val == "restart":
        records.append(ChangeRecord(
            heading=table_heading,
            old_text=_get_cell_text(tc, include_del=True,  include_ins=False),
            new_text=_get_cell_text(tc, include_del=False, include_ins=True),
            change_type=classify_change_type(structural_type=ChangeType.MERGE_CELL),
            bookmark_id=counter.next_id(),
            target_element=_first_para_in_tc(tc),
        ))
        return

    # ---- Case 3: content-only changes inside the cell ----------------------
    cell_para_records: list[ChangeRecord] = []
    for para in tc.findall("w:p", NS):
        cell_para_records.extend(
            extract_changes_from_paragraph(para, table_heading, include_sdt=include_sdt, image_map=image_map)
        )

    if not cell_para_records:
        counter.next_id()
        return

    bm_id  = counter.next_id()
    target = _first_para_in_tc(tc)

    content_records     = [r for r in cell_para_records if r.change_type in _CONTENT_TYPES]
    non_content_records = [r for r in cell_para_records if r.change_type not in _CONTENT_TYPES]

    if content_records:
        records.append(ChangeRecord(
            heading=table_heading,
            old_text=_get_cell_text(tc, include_del=True,  include_ins=False),
            new_text=_get_cell_text(tc, include_del=False, include_ins=True),
            change_type=classify_change_type(
                _get_cell_text(tc, include_del=True,  include_ins=False),
                _get_cell_text(tc, include_del=False, include_ins=True),
            ),
            bookmark_id=bm_id,
            target_element=target,
        ))

    for rec in non_content_records:
        rec.bookmark_id    = bm_id
        rec.target_element = target
    records.extend(non_content_records)


def extract_table_changes(
    tbl,
    current_heading: str,
    counter: "_Counter",
    include_sdt: bool = True,
    image_map: dict[str, bytes] = {},
) -> list[ChangeRecord]:
    """
    Walk a <w:tbl> element and return ChangeRecord objects for:
      - Whole-table insertion / deletion  (all rows inserted or all deleted)
      - Inserted / deleted rows           (w:trPr/w:ins or w:trPr/w:del)
      - Inserted / deleted cells          (w:tcPr/w:cellIns or w:tcPr/w:cellDel)
      - Tracked vertical cell merges      (w:tcPr/w:cellMerge)
      - Content-only changes inside cells (delegates to extract_changes_from_paragraph)

    Control flow
    ------------
    1. check_whole_table_change — early return if every row is ins/del.
    2. For each row  → process_row  — handles row-level ins/del; returns True to skip cells.
    3. For each cell → process_cell — handles cell ins/del, merge, or content changes.
    """
    from lxml import etree

    records: list[ChangeRecord] = []
    table_heading = current_heading + " - Table"

    tbl_rows = [
        el for el in (unwrap_sdts(list(tbl)) if include_sdt else list(tbl))
        if etree.QName(el.tag).localname == "tr"
    ]
    real_rows = [tr for tr in tbl_rows if not _is_row_vmerge_continuation(tr)]

    if real_rows:
        whole_table = check_whole_table_change(real_rows, table_heading, counter)
        if whole_table:
            return whole_table

    for tr in tbl_rows:
        row_handled = process_row(tr, table_heading, counter, records)
        if row_handled:
            continue
        for tc in tr.findall("w:tc", NS):
            process_cell(tc, table_heading, counter, include_sdt, records, image_map)

    return records
