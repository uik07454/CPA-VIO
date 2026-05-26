# embedded_content.py
"""
Helpers for detecting and labelling embedded non-text content (drawings,
images, textboxes, charts, OLE objects) inside w:ins and w:del elements.

Public API used by extractor.py / para_utils.py:
    - has_embedded_content      -- True if element contains any drawing/pict/object
    - classify_embedded_content -- human-readable label for the embedded content
    - extract_image_data        -- raw bytes for the first image found, or None
    - has_note_reference        -- True if run contains a footnote/endnote reference
    - get_note_ref_info         -- returns (note_type, note_id) for a note-reference run
    - CollectedContent          -- dataclass holding text + object + note runs from del/ins
    - emit_object_records       -- pair del/ins object runs into MODIFY/DELETE/ADD_OBJECT records
    - emit_note_records         -- emit ADD_NOTE/DELETE_NOTE records for ins/del note runs
    - emit_plain_note_records   -- emit content-change records for note runs in plain runs
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from constants import (
    NS, NSC, ChangeType, 
    PARA_DELIMITER, FOOTNOTE_LABEL, ENDNOTE_LABEL,
    EMU_PER_PT, EMU_PER_PX,
)
from models import ChangeRecord
from classifier import classify_change_type

if TYPE_CHECKING:
    from lxml import etree


def has_embedded_content(element) -> bool:
    """
    Return True if *element* (a w:ins or w:del) contains 
    any embedded content: (w:drawing, w:pict, w:object).
    """
    return (
        element.find(".//w:drawing", NS) is not None or
        element.find(".//w:pict",    NS) is not None or
        element.find(".//w:object",  NS) is not None
    )


def classify_embedded_content(element) -> str:
    """
    Inspect the subtree of *element* (a w:ins or w:del) and return a
    human-readable label that identifies both the type and name of the
    embedded content.

    Detection priority
    ------------------
    w:drawing (modern DrawingML):
      pic:pic           -> [Image: <name>]
      wps:txbx          -> [Textbox: <name>]
      c:chart           -> [Chart: <name>]
      wpg:wgp           -> [Shape: <name>]
      anything else     -> [Drawing: <name>]

    w:pict (legacy VML):
      v:imagedata       -> [Image: <name>]
      v:textbox         -> [Textbox: <name>]
      other v:shape     -> [Shape: <name>]

    w:object (OLE embedded object):
      o:OLEObject       -> [OLE: <ProgID>]
      anything else     -> [Object]
    """
    def _docpr_name(el) -> str:
        docPr = el.find(".//wp:docPr", NS)
        if docPr is not None:
            return docPr.get("name") or docPr.get("descr") or ""
        return ""

    def _vml_name(shape) -> str:
        return shape.get("title") or shape.get("alt") or ""

    # ---- w:drawing (modern DrawingML) ------------------------------------
    drawing = element.find(".//w:drawing", NS)
    if drawing is not None:
        name = _docpr_name(drawing)
        suffix = f": {name}" if name else ""
        if drawing.find(".//pic:pic", NS) is not None:
            return f"[Image{suffix}]"
        if drawing.find(".//wps:txbx", NS) is not None:
            text = _extract_textbox_text(drawing.find(".//wps:txbx", NS))
            return text if text else f"[Textbox{suffix}]"
        if drawing.find(".//c:chart", NS) is not None:
            return f"[Chart{suffix}]"
        if drawing.find(".//wpg:wgp", NS) is not None:
            return f"[Shape{suffix}]"
        return f"[Drawing{suffix}]"

    # ---- w:pict (legacy VML) --------------------------------------------
    pict = element.find(".//w:pict", NS)
    if pict is not None:
        shape = pict.find(".//v:shape", NS)
        name = _vml_name(shape) if shape is not None else ""
        suffix = f": {name}" if name else ""
        if pict.find(".//v:imagedata", NS) is not None:
            return f"[Image{suffix}]"
        if pict.find(".//v:textbox", NS) is not None:
            text = _extract_textbox_text(pict.find(".//v:textbox", NS))
            return text if text else f"[Textbox{suffix}]"
        if shape is not None:
            return f"[Shape{suffix}]"
        return "[Object]"

    # ---- w:object (OLE embedded object) ---------------------------------
    obj = element.find(".//w:object", NS)
    if obj is not None:
        ole = obj.find(".//o:OLEObject", NS)
        if ole is not None:
            prog_id = ole.get(NSC["o"] + "ProgID") or ole.get("ProgID") or ""
            if prog_id:
                return f"[OLE: {prog_id}]"
    
    return "[Object]"


# Maps note-reference tag localnames to their canonical type string.
# Extend this dict to support additional note-like reference types.
_NOTE_REF_TAGS: dict[str, str] = {
    "footnoteReference": "footnote",
    "endnoteReference":  "endnote",
}


def has_note_reference(run) -> bool:
    """Return True if *run* contains a footnote or endnote reference element."""
    return any(run.find(f"w:{tag}", NS) is not None for tag in _NOTE_REF_TAGS)


def get_note_ref_info(run) -> tuple[str, str] | None:
    """Return ``(note_type, note_id)`` for the first note reference in *run*, or None."""
    for tag, note_type in _NOTE_REF_TAGS.items():
        ref_el = run.find(f"w:{tag}", NS)
        if ref_el is not None:
            return note_type, ref_el.get(NSC["w"] + "id", "")
    return None


def extract_image_data(element, image_map: dict[str, bytes]) -> bytes | None:
    """
    Return raw image bytes for the first image found in *element*, or None.

    Resolves the relationship ID (r:embed for DrawingML, r:id for VML) against
    *image_map* (a ``{rId: bytes}`` dict built from the document relationships).
    Returns None for non-image embedded content (charts, textboxes, OLE, etc.).
    """
    # ---- Modern DrawingML: a:blip @r:embed under pic:pic ------------------
    drawing = element.find(".//w:drawing", NS)
    if drawing is not None and drawing.find(".//pic:pic", NS) is not None:
        blip = drawing.find(".//a:blip", NS)
        if blip is not None:
            r_embed = blip.get(NSC["r"] + "embed")
            if r_embed:
                return image_map.get(r_embed)

    # ---- Legacy VML: v:imagedata @r:id ------------------------------------
    pict = element.find(".//w:pict", NS)
    if pict is not None:
        imagedata = pict.find(".//v:imagedata", NS)
        if imagedata is not None:
            r_id = imagedata.get(NSC["r"] + "id")
            if r_id:
                return image_map.get(r_id)

    return None

def extract_image_size_px(element) -> tuple[int, int] | None:
    """
    Return (width, height) in pixels for the first image found in *element*, or None.

    For DrawingML images, looks for wp:extent @cx and @cy under the first wp:inline or wp:anchor.
    For VML images, looks for v:shape/@style and parses out width and height.
    Returns None for non-image embedded content or if size attributes are missing.
    """
    # ---- Modern DrawingML: wp:extent @cx and @cy under wp:inline/wp:anchor ----
    drawing = element.find(".//w:drawing", NS)
    if drawing is not None and drawing.find(".//pic:pic", NS) is not None:
        extent = drawing.find(".//wp:extent", NS)
        if extent is not None:
            try:
                cx = int(extent.get("cx", "0"))
                cy = int(extent.get("cy", "0"))
                if cx > 0 and cy > 0:
                    return (cx // EMU_PER_PX, cy // EMU_PER_PX)  # convert EMUs to pixels
            except ValueError:
                pass  # ignore invalid integer values
    
    # ---- Legacy VML: v:shape/@style --------------------------------------
    pict = element.find(".//w:pict", NS)
    if pict is not None:
        shape = pict.find(".//v:shape", NS)
        if shape is not None:
            style = shape.get("style", "")
            width, height = 0, 0
            for part in style.split(";"):
                if part.strip().startswith("width:"):
                    try:
                        width = int(float(part.split(":")[1].strip().rstrip("pt")) * EMU_PER_PT)  # convert points to EMUs
                    except ValueError:
                        pass
                elif part.strip().startswith("height:"):
                    try:
                        height = int(float(part.split(":")[1].strip().rstrip("pt")) * EMU_PER_PT)  # convert points to EMUs
                    except ValueError:
                        pass
            if width > 0 and height > 0:
                return (width // EMU_PER_PX, height // EMU_PER_PX)  # convert EMUs to pixels

    return None
@dataclass
class CollectedContent:
    """Aggregated text, object runs, and note-reference runs from consecutive del/ins elements."""
    text: str = ""
    object_runs: list["etree._Element"] = field(default_factory=list)
    note_runs:   list["etree._Element"] = field(default_factory=list)


# Maps note_type string to its heading suffix label.
_NOTE_LABELS: dict[str, str] = {
    "footnote": FOOTNOTE_LABEL,
    "endnote":  ENDNOTE_LABEL,
}


def _extract_textbox_text(element) -> str:
    """Return joined paragraph text from all w:p inside a textbox element, or empty string."""
    from para_utils import get_paragraph_text
    parts = [
        get_paragraph_text(p, include_del=True, include_ins=True)
        for p in element.findall(".//w:p", NS)
    ]
    return PARA_DELIMITER.join(p for p in parts if p)


def _get_note_element(
    note_roots: dict[str, "etree._Element"],
    note_type: str,
    note_id: str,
) -> "etree._Element | None":
    """Return the <w:footnote>/<w:endnote> element matching *note_id*, or None."""
    root = note_roots.get(note_type)
    if root is None:
        return None
    tag = f"w:{note_type}"
    for el in root.findall(tag, NS):
        if el.get(NSC["w"] + "id") == note_id:
            return el
    return None


def _note_plain_text(note_el: "etree._Element", include_del: bool, include_ins: bool) -> str:
    """Collect plain text from all paragraphs inside a note element."""
    from para_utils import get_paragraph_text
    parts = [get_paragraph_text(p, include_del, include_ins) for p in note_el.findall(".//w:p", NS)]
    return PARA_DELIMITER.join(p for p in parts if p)


def emit_note_records(
    del_runs: list["etree._Element"],
    ins_runs: list["etree._Element"],
    note_roots: dict[str, "etree._Element"],
    current_heading: str,
    records: list[ChangeRecord],
) -> None:
    """
    Emit ADD_NOTE / DELETE_NOTE records for note-reference runs found inside
    w:ins (ins_runs) or w:del (del_runs).
    Content of the note is included as new_text / old_text respectively.
    """
    for run in ins_runs:
        info = get_note_ref_info(run)
        if info is None:
            continue
        note_type, note_id = info
        note_el = _get_note_element(note_roots, note_type, note_id)
        note_text = _note_plain_text(note_el, include_del=False, include_ins=True) if note_el is not None else ""
        label = _NOTE_LABELS.get(note_type, note_type.capitalize())
        records.append(ChangeRecord(
            heading=f"{current_heading} - {label}",
            old_text="",
            new_text=note_text,
            change_type=classify_change_type(structural_type=ChangeType.ADD_NOTE),
        ))

    for run in del_runs:
        info = get_note_ref_info(run)
        if info is None:
            continue
        note_type, note_id = info
        note_el = _get_note_element(note_roots, note_type, note_id)
        note_text = _note_plain_text(note_el, include_del=True, include_ins=False) if note_el is not None else ""
        label = _NOTE_LABELS.get(note_type, note_type.capitalize())
        records.append(ChangeRecord(
            heading=f"{current_heading} - {label}",
            old_text=note_text,
            new_text="",
            change_type=classify_change_type(structural_type=ChangeType.DELETE_NOTE),
        ))


def emit_plain_note_records(
    note_runs: list["etree._Element"],
    note_roots: dict[str, "etree._Element"],
    current_heading: str,
    records: list[ChangeRecord],
) -> None:
    """
    For note-reference runs in plain (non-ins/del) runs, look up the note content
    and emit content-change records from extract_changes_from_paragraph.
    """
    from extractor import extract_changes_from_paragraph

    for run in note_runs:
        info = get_note_ref_info(run)
        if info is None:
            continue
        note_type, note_id = info
        note_el = _get_note_element(note_roots, note_type, note_id)
        if note_el is None:
            continue
        label = _NOTE_LABELS.get(note_type, note_type.capitalize())
        note_heading = f"{current_heading} - {label}"
        for para in note_el.findall(".//w:p", NS):
            para_records = extract_changes_from_paragraph(
                para, current_heading=note_heading, note_roots=note_roots
            )
            for rec in para_records:
                rec.target_element = para  # anchor to the note <w:p> for bookmark injection
            records.extend(para_records)


def emit_object_records(
    del_runs: list["etree._Element"],
    ins_runs: list["etree._Element"],
    current_heading: str,
    records: list[ChangeRecord],
    image_map: dict[str, bytes],
) -> None:
    """
    Pair deleted and inserted object runs positionally into MODIFY_OBJECT records.
    Unpaired extras become DELETE_OBJECT or ADD_OBJECT.
    """
    paired = min(len(del_runs), len(ins_runs))
    for j in range(paired):
        records.append(ChangeRecord(
            heading=current_heading,
            old_text=classify_embedded_content(del_runs[j]),
            new_text=classify_embedded_content(ins_runs[j]),
            change_type=classify_change_type(structural_type=ChangeType.MODIFY_OBJECT),
            old_image_data=extract_image_data(del_runs[j], image_map),
            new_image_data=extract_image_data(ins_runs[j], image_map),
            old_image_size_px=extract_image_size_px(del_runs[j]),
            new_image_size_px=extract_image_size_px(ins_runs[j]),
        ))
    for r in del_runs[paired:]:
        records.append(ChangeRecord(
            heading=current_heading,
            old_text=classify_embedded_content(r),
            new_text="",
            change_type=classify_change_type(structural_type=ChangeType.DELETE_OBJECT),
            old_image_data=extract_image_data(r, image_map),
            old_image_size_px=extract_image_size_px(r),
        ))
    for r in ins_runs[paired:]:
        records.append(ChangeRecord(
            heading=current_heading,
            old_text="",
            new_text=classify_embedded_content(r),
            change_type=classify_change_type(structural_type=ChangeType.ADD_OBJECT),
            new_image_data=extract_image_data(r, image_map),
            new_image_size_px=extract_image_size_px(r),
        ))
