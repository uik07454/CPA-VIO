# parser.py
"""
.docx XML parsing utilities.

Handles opening the .docx ZIP archive, parsing the core document and styles
XML trees, discovering header/footer files via the relationship file, and
serialising lxml element trees to pretty-printed XML files for inspection.
"""

import zipfile
from pathlib import Path
from xml.etree import ElementTree as StdET

from lxml import etree

from constants import BASE_DIR, NS, NSC

# Default folder where serialised XML files are written.
OUTPUT_DIR = BASE_DIR / "output"

# Relationship types that identify header and footer parts.
_HF_REL_TYPES: dict[str, str] = {
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header": "Header",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer": "Footer",
}


def parse_any_xml(
    docx_path: str,
    parse_parts: list[str],
    save_xml: bool = False,
) -> dict[str, etree._Element]:
    """Open a .docx archive and return the parsed parts.

    Args:
        docx_path: Path to the ``.docx`` file.
        parse_parts: List of parts to parse.
        save_xml: When ``True``, XML trees are written to ``output/``
            for inspection.

    Returns:
        A dictionary mapping part names to their parsed lxml element trees.
    """
    results: dict[str, etree._Element] = {}

    with zipfile.ZipFile(docx_path, "r") as z:
        for part in parse_parts:
            word_part = f"word/{part}.xml"
            with z.open(word_part) as f:
                root = etree.parse(f).getroot()
                results[part] = root

                if not save_xml:
                    continue
                suffix = part.replace("word/", "").replace(".xml", "")
                save_xml_to_output(root, docx_path, suffix=suffix)

    return results


def save_xml_to_output(
    root: etree._Element,
    docx_path: str,
    suffix: str = "document",
    output_dir: Path | None = None,
) -> Path:
    """Serialise an lxml element tree to a pretty-printed XML file in ``output/``.

    Args:
        root: The lxml element to serialise.
        docx_path: Path to the source ``.docx`` file; its stem is used as the
            output filename prefix.
        suffix: String appended to the stem to distinguish multiple XML files
            from the same document (e.g. ``"document"``, ``"styles"``).
        output_dir: Override the default output directory.  Defaults to
            ``OUTPUT_DIR`` (``<project_root>/output``).

    Returns:
        The ``Path`` of the written XML file.
    """
    target_dir = output_dir if output_dir is not None else OUTPUT_DIR
    target_dir.mkdir(parents=True, exist_ok=True)

    stem = Path(docx_path).stem
    out_path = target_dir / f"{stem}_{suffix}.xml"

    etree.ElementTree(root).write(
        str(out_path),
        pretty_print=True,
        xml_declaration=True,
        encoding="UTF-8",
    )
    return out_path


def build_outline_level_map(styles_root: etree._Element) -> dict[str, int]:
    """Return a mapping of ``styleId -> outline level`` for paragraph styles.

    Only styles that explicitly carry a ``<w:outlineLvl>`` with a value of
    0-8 are included.  Level 9 is the OOXML sentinel for body text and is
    excluded.

    Args:
        styles_root: The root element of ``word/styles.xml``.

    Returns:
        A ``{styleId: level}`` dict where level is an integer 0-8.
    """
    outline_map: dict[str, int] = {}
    for style in styles_root.findall(".//w:style[@w:type='paragraph']", NS):
        style_id = style.get(NSC["w"] + "styleId")
        outline_el = style.find(".//w:pPr/w:outlineLvl", NS)
        if style_id and outline_el is not None:
            val = outline_el.get(NSC["w"] + "val")
            if val is not None and val.isdigit():
                lvl = int(val)
                if lvl <= 8:
                    outline_map[style_id] = lvl
    return outline_map


def build_image_map(docx_path: str) -> dict[str, bytes]:
    """Return ``{rId: raw_bytes}`` for every image relationship in the document.

    Reads ``word/_rels/document.xml.rels`` and extracts all relationships whose
    type ends with ``/image``, then reads the corresponding bytes from the ZIP.
    """
    _IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    image_map: dict[str, bytes] = {}

    with zipfile.ZipFile(docx_path, "r") as z:
        rels_path = "word/_rels/document.xml.rels"
        if rels_path not in z.namelist():
            return image_map

        rels_root = StdET.fromstring(z.read(rels_path))
        for rel in rels_root:
            if rel.get("Type", "") != _IMAGE_REL_TYPE:
                continue
            r_id = rel.get("Id", "")
            target = rel.get("Target", "")          # e.g. "media/image1.png"
            full_path = f"word/{target}"             # e.g. "word/media/image1.png"
            if r_id and full_path in z.namelist():
                image_map[r_id] = z.read(full_path)

    return image_map


def parse_headers_footers(
    docx_path: str,
    save_xml: bool = False,
) -> list[tuple[str, etree._Element]]:
    """Discover and parse all header/footer XML files in the docx archive.

    Reads ``word/_rels/document.xml.rels`` to find header and footer parts,
    then parses each one with lxml.

    Args:
        docx_path: Path to the ``.docx`` file.
        save_xml: When ``True``, each parsed XML tree is written to ``output/``
            for inspection.

    Returns:
        A list of ``(label, xml_root)`` tuples, e.g.::

            [("Header", <Element>), ("Footer", <Element>), ...]
    """
    results: list[tuple[str, etree._Element]] = []

    with zipfile.ZipFile(docx_path, "r") as z:
        rels_path = "word/_rels/document.xml.rels"
        if rels_path not in z.namelist():
            return results

        rels_root = StdET.fromstring(z.read(rels_path))
        for rel in rels_root:
            rel_type = rel.get("Type", "")
            if rel_type not in _HF_REL_TYPES:
                continue

            target = rel.get("Target", "")       # e.g. "header1.xml"
            full_path = f"word/{target}"         # e.g. "word/header1.xml"
            label = _HF_REL_TYPES[rel_type]      # "Header" or "Footer"

            if full_path not in z.namelist():
                continue

            with z.open(full_path) as f:
                root = etree.parse(f).getroot()

            if save_xml:
                save_xml_to_output(root, docx_path, suffix=target.replace(".xml", ""))

            results.append((label, root))

    return results
