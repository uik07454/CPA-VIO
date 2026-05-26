# book_marker.py
"""
Bookmark injection utility.

Injects ``<w:bookmarkStart>`` / ``<w:bookmarkEnd>`` pairs into the annotated
.docx so that Excel HYPERLINK formulas can navigate directly to each change.
"""

import zipfile
from collections import defaultdict
from pathlib import Path
from lxml import etree
from constants import NS, NSC, make_bookmark_id


# ---------------------------------------------------------------------------
# Path helpers — used to locate a target_element in a freshly-parsed tree
# ---------------------------------------------------------------------------

def _element_path(element: etree._Element) -> list[int]:
    """Return the child-index path from the tree root down to *element*.

    Example: ``root -> child[2] -> child[0] -> element`` → ``[2, 0, <idx>]``

    This path is stable across two independently-parsed copies of the same
    XML document, so it can be used to locate the matching node in a
    freshly-parsed tree.

    Args:
        element: Any lxml element that belongs to a parsed tree.

    Returns:
        A list of integer child indices from the root to *element*.
    """
    path = []
    el = element
    while True:
        parent = el.getparent()
        if parent is None:
            break
        path.append(list(parent).index(el))
        el = parent
    path.reverse()
    return path


def _resolve_path(root: etree._Element, path: list[int]) -> etree._Element | None:
    """Walk *root* following the child indices in *path* and return the target element.

    Args:
        root: The root element of a freshly-parsed XML tree.
        path: A list of integer child indices as returned by ``_element_path``.

    Returns:
        The element at the end of *path*, or ``None`` if any index is out of range.
    """
    node = root
    for idx in path:
        children = list(node)
        if idx >= len(children):
            return None
        node = children[idx]
    return node


def _ensure_para(target_el: etree._Element) -> etree._Element | None:
    """Return a ``<w:p>`` suitable for bookmark injection from *target_el*.

    Resolution strategy:
        1. If *target_el* is itself a ``<w:p>``, return it directly.
        2. Otherwise return the first descendant ``<w:p>``.
        3. If *target_el* is a ``<w:tc>`` with no ``<w:p>`` (empty cell —
           rare but possible), create a minimal ``<w:p>`` and append it so
           the bookmark has a valid anchor.  Word requires every ``<w:tc>``
           to contain at least one ``<w:p>``, so this also repairs the
           document structure.
        4. For any other element type with no ``<w:p>`` descendant, return
           ``None`` (bookmark will be skipped by the caller).

    Args:
        target_el: A resolved lxml element from the freshly-parsed tree.

    Returns:
        A ``<w:p>`` element to inject the bookmark into, or ``None`` if no
        suitable anchor can be found or created.
    """
    tag = etree.QName(target_el.tag).localname

    if tag == "p":
        return target_el

    # Try to find an existing paragraph descendant
    para = target_el.find(".//w:p", NS)
    if para is not None:
        return para

    # Empty cell — create a minimal <w:p> to anchor the bookmark
    if tag == "tc":
        new_para = etree.SubElement(target_el, f"{NSC['w']}p")
        return new_para

    return None  # Can't anchor — caller will skip this bookmark


def _get_part_name(element: etree._Element) -> str:
    """Return the word/ part name for the XML tree that contains *element*.

    Walks to the tree root and reads its tag localname, which matches the
    ZIP entry name (e.g. ``"document"`` → ``word/document.xml``).
    """
    return etree.QName(element.getroottree().getroot().tag).localname


def _inject_into_xml(
    bm_map: dict[str, list[int]],
    xml_root: etree._Element,
    bm_id_start: int,
) -> tuple[bytes, int]:
    """Inject bookmarks into *xml_root* and return ``(serialised_bytes, next_bm_id)``."""
    sorted_items = sorted(bm_map.items(), key=lambda kv: len(kv[1]), reverse=True)
    bm_id = bm_id_start
    for bm_name, elem_path in sorted_items:
        target_el = _resolve_path(xml_root, elem_path)
        if target_el is None:
            continue
        para = _ensure_para(target_el)
        if para is None:
            continue
        bm_start = etree.Element(f"{NSC['w']}bookmarkStart")
        bm_start.set(NSC["w"] + "id",   str(bm_id))
        bm_start.set(NSC["w"] + "name", bm_name)
        bm_end = etree.Element(f"{NSC['w']}bookmarkEnd")
        bm_end.set(NSC["w"] + "id", str(bm_id))
        bm_id += 1
        para.insert(0, bm_start)
        para.insert(1, bm_end)
    return etree.tostring(xml_root, xml_declaration=True, encoding="UTF-8", standalone=True), bm_id


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def inject_bookmarks(
    docx_path: str, 
    records: list, 
    out_path: str
) -> etree._Element:
    """Inject bookmarks into each change's target paragraph and save a new .docx.

    Each ``ChangeRecord`` that carries a ``bookmark_id`` also carries a
    ``target_element`` — the lxml element pointing to the injection site in
    the original parsed tree.  Because this function parses a *fresh* copy of
    ``document.xml``, ``target_element`` cannot be used directly.  Instead,
    the child-index path is computed with ``_element_path`` and replayed on
    the fresh tree with ``_resolve_path``.

    Records without a ``target_element`` (e.g. header/footer records) are
    silently skipped.

    Args:
        docx_path: Path to the source ``.docx`` file.
        records: List of ``ChangeRecord`` objects produced by the extractor.
        out_path: Destination path for the annotated ``.docx`` file.

    Returns:
        The modified ``doc_root`` lxml element so the caller can save the
        annotated XML for inspection without re-parsing the written file.
    """
    # ------------------------------------------------------------------
    # 1. Build per-part bookmark maps:  part_name -> {bm_name -> elem_path}
    #
    #    Part name is derived from the XML root tag of each target_element
    #    (e.g. "document", "footnotes", "endnotes").
    # ------------------------------------------------------------------
    part_bm_maps: dict[str, dict[str, list[int]]] = defaultdict(dict)
    for r in records:
        if r.bookmark_id is None or r.target_element is None:
            continue
        bm_name = make_bookmark_id(r.bookmark_id)
        part = _get_part_name(r.target_element)
        if bm_name not in part_bm_maps[part]:
            part_bm_maps[part][bm_name] = _element_path(r.target_element)

    if not part_bm_maps:
        # Nothing to inject — just copy the file unchanged
        Path(out_path).write_bytes(Path(docx_path).read_bytes())
        with zipfile.ZipFile(docx_path, "r") as z:
            return etree.fromstring(z.read("word/document.xml"))

    # ------------------------------------------------------------------
    # 2. Load all ZIP contents into memory
    # ------------------------------------------------------------------
    with zipfile.ZipFile(docx_path, "r") as zin:
        file_contents: dict[str, bytes] = {n: zin.read(n) for n in zin.namelist()}

    # ------------------------------------------------------------------
    # 3. Inject bookmarks into each XML part — deepest paths first.
    #    bm_id_counter carries over across parts for globally unique IDs.
    # ------------------------------------------------------------------
    bm_id_counter = 0
    doc_root = None

    for part_name, bm_map in part_bm_maps.items():
        xml_key = f"word/{part_name}.xml"
        xml_root = etree.fromstring(file_contents[xml_key])
        patched_bytes, bm_id_counter = _inject_into_xml(bm_map, xml_root, bm_id_counter)
        file_contents[xml_key] = patched_bytes
        if part_name == "document":
            doc_root = xml_root

    # Ensure doc_root is always set (in case no document-part records existed)
    if doc_root is None:
        doc_root = etree.fromstring(file_contents["word/document.xml"])

    # ------------------------------------------------------------------
    # 4. Serialise and write the annotated .docx
    # ------------------------------------------------------------------
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in file_contents.items():
            zout.writestr(name, data)

    print(f"[INFO] Annotated docx saved to: {out_path}")
    return doc_root
