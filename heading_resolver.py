# heading_resolver.py
"""
Paragraph heading and text extraction utilities.

Provides helpers to determine whether a paragraph is a navigation heading
and to collect its plain text content with fine-grained control over which
tracked-change runs are included.
"""


from constants import NS, NSC
from para_utils import get_paragraph_text


def get_paragraph_outline_level(para, outline_map: dict) -> int | None:
    """Return the outline level (0-8) for a paragraph, or None if body text.

    Resolution order:
        1. ``w:outlineLvl`` directly on the paragraph's ``w:pPr`` (inline
           override from document.xml — takes priority).
        2. Inherited from the paragraph style via *outline_map* (built from
           styles.xml).

    Outline levels 0-8 correspond to navigation headings (Heading 1 through
    Heading 9).  Level 9 or absent means the paragraph is body text.

    Args:
        para: A ``<w:p>`` lxml element from document.xml.
        outline_map: Mapping of ``styleId -> outline_level`` built from
            styles.xml by ``build_outline_level_map``.

    Returns:
        An integer 0-8 if the paragraph is a heading, or ``None`` for body text.
    """
    # 1. Direct outline level on the paragraph (from document.xml)
    outline_el = para.find(".//w:pPr/w:outlineLvl", NS)
    if outline_el is not None:
        val = outline_el.get(NSC["w"] + "val")
        if val is not None and val.isdigit():
            lvl = int(val)
            if lvl <= 8:
                return lvl

    # 2. Inherited from the paragraph style (from styles.xml via outline_map)
    p_style = para.find(".//w:pPr/w:pStyle", NS)
    if p_style is not None:
        style_id = p_style.get(NSC["w"] + "val")
        if style_id and style_id in outline_map:
            return outline_map[style_id]

    return None


def is_heading(outline_level: int | None) -> bool:
    """Return True if the outline level represents a navigation heading (0-8).

    Args:
        outline_level: An integer outline level, or ``None`` for body text.

    Returns:
        ``True`` for levels 0-8, ``False`` for ``None`` or any other value.
    """
    return outline_level is not None and 0 <= outline_level <= 8


def get_heading_text(para, outline_map: dict) -> str:
    """Return heading text if para is a heading with non-empty text, else None."""
    outline_level = get_paragraph_outline_level(para, outline_map)
    if not is_heading(outline_level):
        return None
    return get_paragraph_text(para) or None