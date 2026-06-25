# para_utils.py
"""
Paragraph-level helpers for walking and dispatching tracked changes.

Provides the building blocks used by extractor.py to process a single
paragraph's child elements into ChangeRecord objects.

Public API used by extractor.py:
    - TRANSPARENT_TAGS          -- frozenset of tag localnames to skip silently
    - unwrap_sdts               -- replace <w:sdt> with its content children
    - next_non_transparent      -- find next non-bookkeeping sibling
    - get_ppr_style_change      -- detect paragraph style changes
    - is_caption_paragraph      -- True if paragraph uses the "Caption" style
    - should_skip_paragraph     -- True if paragraph is inside a tracked textbox
    - collect_consecutive_del   -- collect text + objects from consecutive <w:del>
    - collect_consecutive_ins   -- collect text + objects from consecutive <w:ins>
    - handle_del                -- emit records for a del group (pairs with ins)
    - handle_ins                -- emit records for an ins group
    - handle_run                -- accumulate or flush rPrChange runs
"""

from lxml import etree

from constants import NS, NSC, ChangeType
from models import ChangeRecord
from run_utils import run_text, del_text, is_line_break
from classifier import classify_change_type
from rpr_change_utils import get_rpr_change_meta, can_merge_rpr_runs, flush_rpr_buffer
from embedded_content import (
    CollectedContent,
    has_embedded_content, classify_embedded_content,
    extract_image_data, extract_image_size_px,
    emit_object_records,
    has_note_reference,
    emit_note_records, emit_plain_note_records,
)


# Tags that are meaningless bookkeeping noise — skip without flushing any buffer.
TRANSPARENT_TAGS: frozenset[str] = frozenset({
    "commentRangeStart", "commentRangeEnd",
    "bookmarkStart", "bookmarkEnd",
    "permStart", "permEnd",
    "proofErr",
})


def unwrap_sdts(elements: list[etree._Element]) -> list[etree._Element]:
    """Replace every <w:sdt> with its <w:sdtContent> children, recursively."""
    result = []
    for el in elements:
        if etree.QName(el.tag).localname == "sdt":
            sdt_content = el.find("w:sdtContent", NS)
            if sdt_content is not None:
                result.extend(unwrap_sdts(list(sdt_content)))
        else:
            result.append(el)
    return result


def next_non_transparent(
    children: list[etree._Element],
    start: int,
) -> tuple[int | None, etree._Element | None]:
    """Return (index, element) of the first child at or after *start* not in TRANSPARENT_TAGS."""
    i = start
    while i < len(children):
        if etree.QName(children[i].tag).localname not in TRANSPARENT_TAGS:
            return i, children[i]
        i += 1
    return None, None


def get_ppr_style_change(para) -> tuple[str, str] | None:
    """Return (old_style_id, new_style_id) if the paragraph style changed, else None."""
    pPr = para.find("w:pPr", NS)
    if pPr is None:
        return None
    pPrChange = pPr.find("w:pPrChange", NS)
    if pPrChange is None:
        return None

    new_style_el = pPr.find("w:pStyle", NS)
    new_style = new_style_el.get(NSC["w"] + "val", "Normal") if new_style_el is not None else "Normal"

    old_pPr = pPrChange.find("w:pPr", NS)
    old_style_el = old_pPr.find("w:pStyle", NS) if old_pPr is not None else None
    old_style = old_style_el.get(NSC["w"] + "val", "Normal") if old_style_el is not None else "Normal"

    if old_style == new_style:
        return None  # style didn't actually change — ignore
    return old_style, new_style


# ---- Caption paragraph detection -------------------------------------------

# Style IDs that Word uses for caption paragraphs.  The primary built-in value
# is "Caption"; some documents use localised variants (e.g. "Bildunterschrift").
# Extend this set if additional caption style IDs need to be recognised.
_CAPTION_STYLE_IDS: frozenset[str] = frozenset({"Caption"})


def is_caption_paragraph(para: etree._Element) -> bool:
    """Return True if *para* is styled as a caption paragraph.

    Checks the current paragraph style (``<w:pStyle w:val="...">`` inside
    ``<w:pPr>``).  Does **not** look inside ``<w:pPrChange>`` — we care about
    what the paragraph *is now*, not what it used to be.

    Args:
        para: A ``<w:p>`` lxml element.

    Returns:
        ``True`` when the paragraph style ID matches a known caption style.
    """
    pPr = para.find("w:pPr", NS)
    if pPr is None:
        return False
    pStyle = pPr.find("w:pStyle", NS)
    if pStyle is None:
        return False
    style_val = pStyle.get(NSC["w"] + "val", "")
    return style_val in _CAPTION_STYLE_IDS


# ---- Textbox-inside-tracked-change detection --------------------------------

def should_skip_paragraph(para: etree._Element) -> bool:
    """Return True when *para* is inside a textbox that is itself under a
    tracked insertion or deletion.

    When a textbox is inserted (``<w:ins>…<w:drawing>…<wps:txbx>…</w:ins>``)
    or deleted (``<w:del>…``), the object-level record is already emitted by
    ``handle_ins`` / ``handle_del`` via ``emit_object_records``.  Walking the
    textbox paragraphs again through ``_extract_embedded_container_changes``
    would produce duplicate content-level records.  This guard prevents that.

    Detection strategy (no XPath ancestor axis needed):
      1. Walk up the element tree from *para*.
      2. If we reach a ``<wps:txbx>`` or ``<w:txbxContent>`` ancestor, we are
         inside a textbox.
      3. Continue walking up from that textbox anchor.  If we reach a ``<w:ins>``
         or ``<w:del>`` before hitting the document body, the textbox itself is
         under a tracked change → skip.

    Args:
        para: A ``<w:p>`` lxml element.

    Returns:
        ``True`` when the paragraph should be skipped to avoid duplicate records.
    """
    TEXTBOX_TAGS  = frozenset({"txbx", "txbxContent"})
    TRACKED_TAGS  = frozenset({"ins", "del"})
    BODY_TAGS     = frozenset({"body", "hdr", "ftr"})

    # Walk up to find a textbox ancestor
    node = para.getparent()
    textbox_anchor = None
    while node is not None:
        local = etree.QName(node.tag).localname
        if local in TEXTBOX_TAGS:
            textbox_anchor = node
            break
        if local in BODY_TAGS:
            # Reached document body without finding a textbox — not inside one
            return False
        node = node.getparent()

    if textbox_anchor is None:
        return False  # not inside any textbox

    # Now walk up from the textbox anchor to see if it sits inside w:ins/w:del
    node = textbox_anchor.getparent()
    while node is not None:
        local = etree.QName(node.tag).localname
        if local in TRACKED_TAGS:
            return True   # textbox is inside a tracked change → skip paragraph
        if local in BODY_TAGS:
            break         # reached body without finding a tracked change
        node = node.getparent()

    return False


def collect_consecutive_del(
    children: list[etree._Element], start: int
) -> tuple[CollectedContent, int]:
    """
    Starting at *start*, collect all consecutive <w:del> elements.
    Separates text runs and object runs into a single CollectedContent.
    """
    texts: list[str] = []
    object_runs: list[etree._Element] = []
    note_runs: list[etree._Element] = []
    i = start
    while i < len(children):
        tag = etree.QName(children[i].tag).localname
        if tag in TRANSPARENT_TAGS:
            i += 1
            continue
        if tag != "del":
            break
        for r in children[i].findall("w:r", NS):
            if has_embedded_content(r):
                object_runs.append(r)
            elif has_note_reference(r):
                note_runs.append(r)
            else:
                texts.append(del_text(r))
        i += 1
    return CollectedContent(text="".join(texts), object_runs=object_runs, note_runs=note_runs), i


def collect_consecutive_ins(
    children: list[etree._Element], start: int
) -> tuple[CollectedContent, int]:
    """
    Starting at *start*, collect all consecutive <w:ins> elements.
    Separates text runs and object runs into a single CollectedContent.
    """
    texts: list[str] = []
    object_runs: list[etree._Element] = []
    note_runs: list[etree._Element] = []
    i = start
    while i < len(children):
        tag = etree.QName(children[i].tag).localname
        if tag in TRANSPARENT_TAGS:
            i += 1
            continue
        if tag != "ins":
            break
        for r in children[i].findall("w:r", NS):
            if has_embedded_content(r):
                object_runs.append(r)
            elif has_note_reference(r):
                note_runs.append(r)
            else:
                texts.append(run_text(r))
        i += 1
    return CollectedContent(text="".join(texts), object_runs=object_runs, note_runs=note_runs), i


def handle_del(
    children: list[etree._Element],
    i: int,
    current_heading: str,
    records: list[ChangeRecord],
    image_map: dict[str, bytes] = {},
    note_roots: dict[str, "etree._Element"] = {},
) -> int:
    """
    Handle one or more consecutive w:del elements, emitting records for both
    text content, embedded objects, and note references.
    Pairs with a following w:ins group when present.
    """
    del_content, i = collect_consecutive_del(children, i)

    next_idx, next_el = next_non_transparent(children, i)
    has_following_ins = next_el is not None and etree.QName(next_el.tag).localname == "ins"

    if has_following_ins:
        ins_content, i = collect_consecutive_ins(children, next_idx)
        old_str, new_str = del_content.text.strip(), ins_content.text.strip()
        if old_str or new_str:
            records.append(ChangeRecord(
                heading=current_heading,
                old_text=old_str,
                new_text=new_str,
                change_type=classify_change_type(old_str, new_str),
            ))
        # Use pairwise handling to preserve the connection between related del/ins runs & make MODIFY
        emit_object_records(del_content.object_runs, ins_content.object_runs, current_heading, records, image_map)
        # Notes only have add/delete semantics, no MODIFY
        emit_note_records(del_content.note_runs, ins_content.note_runs, note_roots, current_heading, records)
    else:
        old_str = del_content.text.strip()
        if old_str:
            records.append(ChangeRecord(
                heading=current_heading,
                old_text=old_str,
                new_text="",
                change_type=classify_change_type(old_str, ""),
            ))
        for r in del_content.object_runs:
            records.append(ChangeRecord(
                heading=current_heading,
                old_text=classify_embedded_content(r),
                new_text="",
                change_type=classify_change_type(structural_type=ChangeType.DELETE_OBJECT),
                old_image_data=extract_image_data(r, image_map),
                old_image_size_px=extract_image_size_px(r),
            ))
        emit_note_records(del_content.note_runs, [], note_roots, current_heading, records)
    return i


def handle_ins(
    children: list[etree._Element],
    i: int,
    current_heading: str,
    records: list[ChangeRecord],
    image_map: dict[str, bytes] = {},
    note_roots: dict[str, "etree._Element"] = {},
) -> int:
    """
    Handle one or more consecutive w:ins elements, emitting records for both
    text content, embedded objects, and note references.
    """
    ins_content, i = collect_consecutive_ins(children, i)

    new_str = ins_content.text.strip()
    if new_str:
        records.append(ChangeRecord(
            heading=current_heading,
            old_text="",
            new_text=new_str,
            change_type=classify_change_type("", new_str),
        ))
    for r in ins_content.object_runs:
        records.append(ChangeRecord(
            heading=current_heading,
            old_text="",
            new_text=classify_embedded_content(r),
            change_type=classify_change_type(structural_type=ChangeType.ADD_OBJECT),
            new_image_data=extract_image_data(r, image_map),
            new_image_size_px=extract_image_size_px(r),
        ))
    emit_note_records([], ins_content.note_runs, note_roots, current_heading, records)
    return i


def handle_run(
    child: etree._Element,
    rpr_buffer: list[etree._Element],
    current_heading: str,
    records: list[ChangeRecord],
    note_roots: dict[str, "etree._Element"] = {},
) -> None:
    """
    Runs carrying w:rPrChange are accumulated in rpr_buffer for merging.
    Plain runs flush the buffer unless they contain only whitespace.
    Plain runs with note references are dispatched to emit_plain_note_records.
    """
    br = child.find("w:br", NS)
    has_no_text = child.find("w:t", NS) is None

    if br is not None and has_no_text:
        # Manual line break: keep within active rPrChange spans
        if is_line_break(br) and rpr_buffer:
            rpr_buffer.append(child)
        # Other break types (e.g. page & column) are discarded
        return

    if has_note_reference(child):
        emit_plain_note_records([child], note_roots, current_heading, records)
        return

    meta = get_rpr_change_meta(child)
    if meta is not None:
        # Compare against buffer[0] (always a valid rPrChange run) rather than
        # buffer[-1], which may be a break-only run with no rPrChange metadata.
        if rpr_buffer and can_merge_rpr_runs(rpr_buffer[0], child):
            rpr_buffer.append(child)
        else:
            flush_rpr_buffer(rpr_buffer, current_heading, records)
            rpr_buffer.append(child)
    else:
        # Plain run: whitespace-only runs do not break an active rPr span.
        if rpr_buffer and not run_text(child).strip():
            pass
        else:
            flush_rpr_buffer(rpr_buffer, current_heading, records)


def get_paragraph_text(
    para,
    include_del: bool = False,
    include_ins: bool = True,
) -> str:
    """Collect text from a paragraph with control over tracked-change runs.

    Rules:
        - Plain ``<w:t>`` nodes (run NOT inside ``<w:ins>`` or ``<w:del>``)
          are always included.
        - ``<w:t>`` nodes inside ``<w:ins>`` are included only when
          ``include_ins=True``.
        - ``<w:delText>`` nodes inside ``<w:del>`` are included only when
          ``include_del=True``.

    Args:
        para: A ``<w:p>`` lxml element.
        include_del: When ``True``, deleted text (``<w:delText>``) is included.
        include_ins: When ``True``, inserted text (``<w:t>`` inside ``<w:ins>``)
            is included.

    Returns:
        The stripped, concatenated text content of the paragraph.
    """
    parts = []
    for child in para:
        tag = etree.QName(child.tag).localname

        if tag == "r":
            parts.append(run_text(child))

        elif tag == "ins":
            if include_ins:
                for r in child.findall("w:r", NS):
                    parts.append(run_text(r))

        elif tag == "del":
            if include_del:
                for r in child.findall("w:r", NS):
                    parts.append(del_text(r))

    return "".join(parts).strip()
