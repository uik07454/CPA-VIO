# extractor.py
"""
Tracked-change extraction pipeline.

Provides two public extraction functions:
    - ``extract_tracked_changes`` — table-aware body extraction
    - ``extract_hf_changes``      — header/footer extraction

All functions return a list of ``ChangeRecord`` objects in document order.
"""

from dataclasses import dataclass
from lxml import etree

from constants import NS, ChangeType, DEFAULT_HEADING
from models import ChangeRecord
from heading_resolver import get_heading_text
from classifier import classify_change_type
from para_utils import (
    TRANSPARENT_TAGS,
    unwrap_sdts,
    get_ppr_style_change,
    handle_del,
    handle_ins,
    handle_run,
)
from rpr_change_utils import flush_rpr_buffer
from table_utils import (
    first_para_in_tr,
    extract_table_changes,
)


class _Counter:
    """Monotonically increasing counter passed by reference; assigns unique bookmark IDs."""

    def __init__(self) -> None:
        self._value = 0

    def next_id(self) -> int:
        val = self._value
        self._value += 1
        return val


# To add a new container type: append a ContainerDescriptor to _BODY_CONTAINER_REGISTRY.
@dataclass
class ContainerDescriptor:
    """
    xpath : findall() XPath locating container instances within a body element (w:p or w:tbl).
    label : suffix appended to the heading, e.g. "Textbox" -> "<heading> - Textbox".
    """
    xpath: str
    label: str


_BODY_CONTAINER_REGISTRY: list[ContainerDescriptor] = [
    # Modern DrawingML textboxes  (wps:txbx > w:txbxContent > w:p)
    ContainerDescriptor(xpath=".//wps:txbx", label="Textbox"),
]


def _extract_embedded_container_changes(
    element: etree._Element,
    current_heading: str,
    registry: list[ContainerDescriptor] = _BODY_CONTAINER_REGISTRY,
    image_map: dict[str, bytes] = {},
) -> list[ChangeRecord]:
    """Scan element for embedded text containers and extract tracked changes from their paragraphs."""
    records = []
    for descriptor in registry:
        for container in element.findall(descriptor.xpath, NS):
            heading = f"{current_heading} - {descriptor.label}"
            for para in container.findall(".//w:p", NS):
                para_records = extract_changes_from_paragraph(para, current_heading=heading, image_map=image_map)
                for rec in para_records:
                    rec.target_element = para  # pinpoint bookmark to the textbox paragraph
                records.extend(para_records)
    return records


def extract_changes_from_paragraph(
    para: etree._Element,
    current_heading: str,
    include_sdt: bool = True,
    image_map: dict[str, bytes] = {},
    note_roots: dict[str, etree._Element] = {},
) -> list[ChangeRecord]:
    """
    Walk the direct children of a paragraph and emit ChangeRecord objects
    for every tracked change found.

    Dispatch table for each child tag
    ----------------------------------
    "r"             -> handle_run (rPrChange buffering + note reference detection)
    transparent tag -> skip (must not flush rPr buffer)
    "del"           -> handle_del (text + objects + notes, pairs with following ins)
    "ins"           -> handle_ins (text + objects + notes)
    anything else   -> skip
    """
    records = []

    # ---- Paragraph-level style change (w:pPrChange) ------------------------
    style_change = get_ppr_style_change(para)
    if style_change:
        old_style, new_style = style_change
        records.append(ChangeRecord(
            heading=current_heading,
            old_text=f"Style: {old_style}",
            new_text=f"Style: {new_style}",
            change_type=classify_change_type(structural_type=ChangeType.STYLE_CHANGE),
        ))

    children = unwrap_sdts(list(para)) if include_sdt else list(para)
    rpr_buffer: list[etree._Element] = []
    i = 0

    while i < len(children):
        child = children[i]
        tag = etree.QName(child.tag).localname

        # ---- w:r: run-level format change (rPrChange) or plain run ----------
        if tag == "r":
            handle_run(child, rpr_buffer, current_heading, records, note_roots)
            i += 1
            continue

        # ---- Transparent bookkeeping tags: skip without flushing rPr buffer -
        if tag in TRANSPARENT_TAGS:
            i += 1
            continue

        # ---- Any content tag flushes the pending rPr buffer -----------------
        flush_rpr_buffer(rpr_buffer, current_heading, records)

        # ---- w:del: deletion or modification --------------------------------
        if tag == "del":
            i = handle_del(children, i, current_heading, records, image_map, note_roots)
            continue

        # ---- w:ins: insertion -----------------------------------------------
        if tag == "ins":
            i = handle_ins(children, i, current_heading, records, image_map, note_roots)
            continue

        i += 1  # unrecognised tag — skip

    flush_rpr_buffer(rpr_buffer, current_heading, records)
    return records


def extract_tracked_changes(
    doc_root: etree._Element,
    outline_map: dict[str, int],
    include_sdt: bool = True,
    image_map: dict[str, bytes] = {},
    note_roots: dict[str, etree._Element] = {},
    default_heading: str = DEFAULT_HEADING,
) -> list[ChangeRecord]:
    """
    Detects both w:p and w:tbl tracked changes in a single pass by walking the document body.
    All records are emitted in document flow order.
    """
    all_records = []
    current_heading = default_heading

    body = doc_root.find(".//w:body", NS)
    if body is None:
        return all_records

    counter = _Counter()

    for child in (unwrap_sdts(list(body)) if include_sdt else list(body)):
        tag = etree.QName(child.tag).localname

        if tag == "p":
            # Update current heading if this paragraph is a heading with non-empty text
            current_heading = get_heading_text(child, outline_map) or current_heading

            para_records = extract_changes_from_paragraph(child, current_heading, include_sdt=include_sdt, image_map=image_map, note_roots=note_roots)
            if para_records:
                bm_id = counter.next_id()
                for rec in para_records:
                    # If record still doesn't have a target, assign shared bookmark
                    if rec.target_element is None:
                        # Body-paragraph record — assign the shared body bookmark
                        rec.bookmark_id = bm_id
                        rec.target_element = child
                    else:
                        # Note content record — already has its own target_element (note <w:p>)
                        rec.bookmark_id = counter.next_id()
                all_records.extend(para_records)

            embedded = _extract_embedded_container_changes(child, current_heading, image_map=image_map)
            if embedded:
                target_to_bm: dict[int, str] = {}
                for rec in embedded:
                    if rec.target_element is None:
                        rec.target_element = child
                    key = id(rec.target_element)
                    if key not in target_to_bm:
                        target_to_bm[key] = counter.next_id()
                    rec.bookmark_id = target_to_bm[key]
                all_records.extend(embedded)
            else:
                counter.next_id()
            continue

        if tag == "tbl":
            tbl_records = extract_table_changes(child, current_heading, counter, include_sdt=include_sdt, image_map=image_map)
            all_records.extend(tbl_records)

            first_tr = child.find("w:tr", NS)
            tbl_anchor = first_para_in_tr(first_tr) if first_tr is not None else None
            embedded = _extract_embedded_container_changes(child, current_heading, image_map=image_map)
            if embedded:
                target_to_bm: dict[int, str] = {}
                for rec in embedded:
                    if rec.target_element is None:
                        rec.target_element = tbl_anchor
                    key = id(rec.target_element)
                    if key not in target_to_bm:
                        target_to_bm[key] = counter.next_id()
                    rec.bookmark_id = target_to_bm[key]
                all_records.extend(embedded)
            else:
                counter.next_id()
            continue

    return all_records


def extract_hf_changes(hf_roots: list[tuple[str, etree._Element]]) -> list[ChangeRecord]:
    """Extract tracked changes from all header and footer XML roots."""
    all_records = []

    for label, root in hf_roots:
        for para in root.findall(".//w:p", NS):
            records = extract_changes_from_paragraph(para, current_heading=label)
            all_records.extend(records)

    return all_records
