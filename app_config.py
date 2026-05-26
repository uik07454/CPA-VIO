# app_config.py
"""
AppConfig: user-configurable settings with JSON persistence.

Mirrors the configurable constants from constants.py and adds
change-type selection state. All GUI settings flow through this module.
"""

import json
import re
from dataclasses import dataclass, field
from pathlib import Path

from constants import (
    ALL_CHANGE_TYPES,
    BASE_DIR,
    CPA_TEMPLATE_DIR,
    DEFAULT_HEADING,
    DEFAULT_TEMPLATE_COLUMNS,
    DIFF_DEL_COLOR,
    DIFF_EQ_COLOR,
    DIFF_INS_COLOR,
    EMPTY_FIELD,
    IMAGE_HEIGHT_PX,
    IMAGE_PX_MAX,
    IMAGE_PX_MIN,
    IMAGE_SIZE_MODE,
    IMAGE_SIZE_MODE_FIXED,
    IMAGE_SIZE_MODE_ORIGINAL,
    IMAGE_WIDTH_PX,
    OUTPUT_MODE_SCRATCH,
    ChangeType,
)

_HEX_COLOR_RE = re.compile(r"^[0-9A-Fa-f]{6}$")

CONFIG_PATH = BASE_DIR / "data" / "config.json"


def _validate_hex_color(value: str, field_name: str) -> None:
    if not _HEX_COLOR_RE.match(value):
        raise ValueError(f"{field_name} must be a 6-character hex RGB string (e.g. 'FF0000'), got: '{value}'")


def _validate_px(value: int, field_name: str) -> None:
    if not (IMAGE_PX_MIN <= value <= IMAGE_PX_MAX):
        raise ValueError(f"{field_name} must be between {IMAGE_PX_MIN} and {IMAGE_PX_MAX} px, got: {value}")


@dataclass
class AppConfig:
    """
    User-configurable runtime settings.

    All fields default to the project-wide constants. Mutate fields directly,
    then call validate() before use and save() to persist.
    """

    # ---- Text Settings ------------------------------------------------------
    empty_field:     str = EMPTY_FIELD
    default_heading: str = DEFAULT_HEADING

    # ---- Image Settings -----------------------------------------------------
    image_width_px:  int = IMAGE_WIDTH_PX
    image_height_px: int = IMAGE_HEIGHT_PX
    image_size_mode: str = IMAGE_SIZE_MODE

    # ---- Diff Color Settings ------------------------------------------------
    diff_del_color: str = DIFF_DEL_COLOR
    diff_ins_color: str = DIFF_INS_COLOR
    diff_eq_color:  str = DIFF_EQ_COLOR

    # ---- Change Type Selection ----------------------------------------------
    # Stores the set of change type strings to include in Excel output.
    # ChangeType.UNKNOWN is always included regardless of this set.
    selected_change_types: set[str] = field(default_factory=lambda: set(ALL_CHANGE_TYPES))

    # ---- Header/Footer & SDT Extraction -------------------------------------
    include_hf_changes: bool = True
    include_sdt:        bool = True

    # ---- Template Output Settings -------------------------------------------
    # "scratch" → generate a new workbook; "template" → copy a template file.
    output_mode: str = OUTPUT_MODE_SCRATCH

    # Filename (basename only) of the selected template inside CPA_TEMPLATE_DIR.
    selected_template: str = ""

    # Per-template, per-sheet column mapping:
    #   {filename: {sheet_name: [col_heading, col_change_type, col_old, col_new]}}
    # Column values are Excel column letters (e.g. "A", "B", "C", "D").
    template_profiles: dict[str, dict[str, list[str]]] = field(default_factory=dict)

    # Selected sheet name per template: {filename: sheet_name}
    template_selected_sheets: dict[str, str] = field(default_factory=dict)

    def validate(self) -> list[str]:
        """
        Validate all fields. Returns a list of error messages (empty = valid).

        Does not raise; callers decide how to surface errors.
        """
        errors: list[str] = []

        if not self.empty_field:
            errors.append("empty_field must not be empty.")
        if not self.default_heading:
            errors.append("default_heading must not be empty.")

        for name, val in (
            ("image_width_px", self.image_width_px),
            ("image_height_px", self.image_height_px),
        ):
            try:
                _validate_px(val, name)
            except ValueError as e:
                errors.append(str(e))

        if self.image_size_mode not in (IMAGE_SIZE_MODE_FIXED, IMAGE_SIZE_MODE_ORIGINAL):
            errors.append(f"image_size_mode must be '{IMAGE_SIZE_MODE_FIXED}' or '{IMAGE_SIZE_MODE_ORIGINAL}'.")

        for name, val in (
            ("diff_del_color", self.diff_del_color),
            ("diff_ins_color", self.diff_ins_color),
            ("diff_eq_color",  self.diff_eq_color),
        ):
            try:
                _validate_hex_color(val, name)
            except ValueError as e:
                errors.append(str(e))

        if not self.selected_change_types:
            errors.append("At least one change type must be selected.")

        return errors

    def is_valid(self) -> bool:
        return len(self.validate()) == 0

    def save(self, path: Path = CONFIG_PATH) -> None:
        """Persist current config to a JSON file."""
        data = {
            "empty_field":              self.empty_field,
            "default_heading":          self.default_heading,
            "image_width_px":           self.image_width_px,
            "image_height_px":          self.image_height_px,
            "image_size_mode":          self.image_size_mode,
            "diff_del_color":           self.diff_del_color,
            "diff_ins_color":           self.diff_ins_color,
            "diff_eq_color":            self.diff_eq_color,
            "selected_change_types":    sorted(self.selected_change_types),
            "include_hf_changes":       self.include_hf_changes,
            "include_sdt":              self.include_sdt,
            "output_mode":              self.output_mode,
            "selected_template":        self.selected_template,
            "template_profiles":        self.template_profiles,
            "template_selected_sheets": self.template_selected_sheets,
        }
        path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")

    @classmethod
    def load(cls, path: Path = CONFIG_PATH) -> "AppConfig":
        """Load config from JSON, falling back to defaults for missing keys."""
        if not path.exists():
            return cls()
        raw = json.loads(path.read_text(encoding="utf-8"))
        cfg = cls()
        cfg.empty_field            = raw.get("empty_field",            cfg.empty_field)
        cfg.default_heading        = raw.get("default_heading",        cfg.default_heading)
        cfg.image_width_px         = raw.get("image_width_px",         cfg.image_width_px)
        cfg.image_height_px        = raw.get("image_height_px",        cfg.image_height_px)
        cfg.image_size_mode        = raw.get("image_size_mode",        cfg.image_size_mode)
        cfg.diff_del_color         = raw.get("diff_del_color",         cfg.diff_del_color)
        cfg.diff_ins_color         = raw.get("diff_ins_color",         cfg.diff_ins_color)
        cfg.diff_eq_color          = raw.get("diff_eq_color",          cfg.diff_eq_color)
        saved_types = raw.get("selected_change_types")
        if saved_types is not None:
            # Only keep types that still exist in the current registry.
            cfg.selected_change_types = set(saved_types) & ALL_CHANGE_TYPES
        cfg.include_hf_changes = raw.get("include_hf_changes", cfg.include_hf_changes)
        cfg.include_sdt        = raw.get("include_sdt",        cfg.include_sdt)
        cfg.output_mode        = raw.get("output_mode",        cfg.output_mode)
        cfg.selected_template  = raw.get("selected_template",  cfg.selected_template)
        cfg.template_profiles         = raw.get("template_profiles",         cfg.template_profiles)
        cfg.template_selected_sheets  = raw.get("template_selected_sheets",  cfg.template_selected_sheets)
        return cfg

    def reset_to_default(self) -> None:
        """Reset all fields to the project-wide defaults in-place."""
        default = AppConfig()
        self.empty_field            = default.empty_field
        self.default_heading        = default.default_heading
        self.image_width_px         = default.image_width_px
        self.image_height_px        = default.image_height_px
        self.image_size_mode        = default.image_size_mode
        self.diff_del_color         = default.diff_del_color
        self.diff_ins_color         = default.diff_ins_color
        self.diff_eq_color          = default.diff_eq_color
        self.selected_change_types  = set(ALL_CHANGE_TYPES)
        self.include_hf_changes     = default.include_hf_changes
        self.include_sdt            = default.include_sdt
        # output_mode, selected_template, and template_profiles are intentionally
        # NOT reset — user-defined column mappings should survive a settings reset.

    def available_templates(self) -> list[str]:
        """Return sorted list of template filenames found in CPA_TEMPLATE_DIR."""
        if not CPA_TEMPLATE_DIR.is_dir():
            return []
        return sorted(p.name for p in CPA_TEMPLATE_DIR.glob("*.xlsx"))

    def get_selected_sheet(self, template_name: str = "") -> str:
        """Return the selected sheet name for *template_name*."""
        name = template_name or self.selected_template
        return self.template_selected_sheets.get(name, "")

    def get_column_mapping(self, template_name: str = "", sheet_name: str = "") -> list[str]:
        """Return the column mapping for a (template, sheet) pair, falling back to the default."""
        name  = template_name or self.selected_template
        sheet = sheet_name or self.get_selected_sheet(name)
        sheet_map = self.template_profiles.get(name, {})
        return list(sheet_map.get(sheet, DEFAULT_TEMPLATE_COLUMNS))

    def filter_records(self, records: list) -> list:
        """
        Return only the records whose change_type is selected or is UNKNOWN.

        UNKNOWN is always passed through so reviewers never miss unclassified changes.
        """
        allowed = self.selected_change_types | {ChangeType.UNKNOWN}
        return [r for r in records if r.change_type in allowed]
