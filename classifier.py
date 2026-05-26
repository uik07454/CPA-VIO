# classifier.py
"""
Change-type classification logic.

Provides a single entry point, ``classify_change_type``, used by every part
of the pipeline that needs to assign a semantic label to a tracked change.
"""

from typing import NamedTuple

from constants import ChangeType


class RprDetection(NamedTuple):
    """
    Encapsulates the old/new state of a single detectable rPr property and the
    ChangeType constants that should be returned when the property is added or removed.

    Fields:
        old_state:    Whether the property was present before the change.
        new_state:    Whether the property is present after the change.
        added_type:   ChangeType to return when new_state=True and old_state=False.
        removed_type: ChangeType to return when old_state=True and new_state=False.
    """
    old_state: bool
    new_state: bool
    added_type: str
    removed_type: str


class ValDetection(NamedTuple):
    """
    Encapsulates the old/new value of a run property for value-based change detection.

    Fields:
        old_val:      Property value before the change, or None if absent.
        new_val:      Property value after the change, or None if absent.
        changed_type: ChangeType to return when old_val != new_val.
    """
    old_val:      str | None
    new_val:      str | None
    changed_type: str


def classify_change_type(
    old_text: str = "",
    new_text: str = "",
    rpr_detections: list[RprDetection] | None = None,
    val_detections: list[ValDetection] | None = None,
    is_format_only: bool = False,
    structural_type: str | None = None,
) -> str:
    """
    Determine the semantic category of a tracked change.

    This is the single, unified entry point for all change-type classification.
    Every ChangeRecord should have its change_type set via this function.

    Priority order (when structural_type is None):
        1. Modify Content  — both sides have text and they differ
        2. Delete Content  — only old text present
        3. Add Content     — only new text present
        4. rPr detections  — evaluated in the order provided via rpr_detections
        5. Val detections  — value-based changes (highlight colour, font colour, etc.)
        6. Format Change   — any other rPrChange (bold, font, size, etc.)
        7. Unknown
    """
    # Structural changes take unconditional priority.
    if structural_type is not None:
        return structural_type

    if old_text and new_text and old_text != new_text:
        return ChangeType.MODIFY_CONTENT
    if old_text and not new_text:
        return ChangeType.DELETE_CONTENT
    if new_text and not old_text:
        return ChangeType.ADD_CONTENT

    for det in (rpr_detections or []):
        if det.new_state and not det.old_state:
            return det.added_type
        if det.old_state and not det.new_state:
            return det.removed_type

    for det in (val_detections or []):
        if det.old_val != det.new_val:
            return det.changed_type

    if is_format_only:
        return ChangeType.FORMAT_CHANGE
    return ChangeType.UNKNOWN
