# models.py
"""
Data model for a single tracked change extracted from a .docx document.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Optional

if TYPE_CHECKING:
    from lxml import etree


@dataclass
class ChangeRecord:
    """
    Represents a single tracked change found in a Word document.

    Attributes:
        heading: str
        old_text: str
        new_text: str
        change_type: str
        bookmark_id: Optional[int] — None means no bookmark assigned; make_bookmark_id() should be called on it before use
        target_element: Optional["etree._Element"]
        style_meta: dict
        old_image_data: Optional[bytes] — raw image bytes for the old/deleted image
        new_image_data: Optional[bytes] — raw image bytes for the new/inserted image
        old_image_size_px: Optional[tuple[int, int]] — (width, height) in pixels for the old/deleted image
        new_image_size_px: Optional[tuple[int, int]] — (width, height) in pixels for the new/inserted image
    """

    heading:        str
    old_text:       str
    new_text:       str
    change_type:    str
    bookmark_id:    Optional[int] = None
    target_element: Optional["etree._Element"] = field(default=None, repr=False)
    style_meta:     dict = field(default_factory=dict)
    old_image_data: Optional[bytes] = field(default=None, repr=False)
    new_image_data: Optional[bytes] = field(default=None, repr=False)
    old_image_size_px: Optional[tuple[int, int]] = field(default=None)
    new_image_size_px: Optional[tuple[int, int]] = field(default=None)
    # True when the record originates from a paragraph styled as a caption
    # (e.g. <w:pStyle w:val="Caption">). Used by AppConfig.filter_records()
    # to optionally suppress caption changes from the output.
    is_caption:     bool = False
    is_caption:     bool = False
