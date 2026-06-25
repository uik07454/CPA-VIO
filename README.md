# CPAuto — Word Compare Difference Extractor

## Overview

**CPAuto** is a Python-based tool that parses Microsoft Word comparison documents (`.docx` files with tracked changes) and extracts every detected change into a structured, navigable Excel report. It reads OOXML (Office Open XML) directly — no Microsoft Word installation is required.

The system operates in two modes:

- **GUI Mode** (`gui.py`) — a PyQt6 desktop application with a full settings panel, change-type filter tree, and one-click pipeline execution.
- **CLI Mode** (`main.py`) — a lightweight command-line entry point intended for development and scripting.

Both modes run the same underlying pipeline: parse → extract → classify → annotate → output.

## Features

- **PyQt6 GUI** with a scrollable settings panel, hierarchical change-type filter tree, and real-time status feedback.
- **Excel output** (`.xlsx`) generated either from scratch or by filling in a pre-existing CPA template workbook.
- **Delete Toggle Feature** — Optional toggle column in Excel output with TRUE/FALSE dropdowns for marking rows for deletion. Includes auto-delete watcher, batch deletion with automatic backup, and GUI integration. See [DELETE_TOGGLE_GUIDE.md](DELETE_TOGGLE_GUIDE.md) for complete guide.
- **Bookmark injection** into an annotated copy of the source `.docx`, enabling Excel `HYPERLINK()` formulas to jump directly to each change location.
- **Inline word-level diff highlighting** for *Modify Content* rows — deleted words in red, inserted words in green, unchanged words in black.
- **Image extraction and embedding** — images inside `<w:ins>` / `<w:del>` elements are extracted and embedded directly into Excel cells, with configurable dimensions and size modes.
- **Header and footer change extraction** — optional, toggleable per run.
- **SDT (Structured Document Tag) unwrapping** — content inside Word content controls is optionally included in extraction.
- **Footnote and endnote change detection** — add/delete note records with full note content.
- **Table structure change detection** — whole-table, row, cell, and vertical-merge tracked changes.
- **Embedded object classification** — images, textboxes, charts, shapes, and OLE objects are identified and labelled.
- **Paragraph style change detection** — `<w:pPrChange>` elements are captured as *Paragraph Style Change* records.
- **Run-property change detection** — strikethrough, highlight colour, font colour, and generic format changes from `<w:rPrChange>` elements.
- **JSON-persisted user configuration** — all GUI settings are saved to `data/config.json` and restored on next launch.
- **Animated splash screen** on GUI launch (rendered with Manim, played via Qt Multimedia).
- **PyInstaller-ready** — includes a `build.bat` and `.spec` file for packaging into a standalone executable.

## Architecture Overview

The system is organized into four logical layers:

```
┌─────────────────────────────────────────────────────────────┐
│                        GUI Layer                            │
│   gui.py  ·  app_config.py  ·  splash_scene.py             │
└────────────────────────┬────────────────────────────────────┘
                         │ orchestrates
┌────────────────────────▼────────────────────────────────────┐
│                     Parsing Layer                           │
│   parser.py  ·  constants.py                                │
│   (unzip .docx, parse XML parts, build maps)                │
└────────────────────────┬────────────────────────────────────┘
                         │ XML element trees
┌────────────────────────▼────────────────────────────────────┐
│               Extraction & Classification Layer             │
│   extractor.py  ·  para_utils.py  ·  table_utils.py        │
│   heading_resolver.py  ·  run_utils.py                      │
│   rpr_change_utils.py  ·  embedded_content.py               │
│   classifier.py  ·  models.py                               │
│   (walk paragraphs/tables, emit ChangeRecord objects)       │
└────────────────────────┬────────────────────────────────────┘
                         │ list[ChangeRecord]
┌────────────────────────▼────────────────────────────────────┐
│                      Output Layer                           │
│   book_marker.py  ·  output.py                              │
│   (inject bookmarks → annotated .docx, write .xlsx)         │
└─────────────────────────────────────────────────────────────┘
```

**Pipeline steps (in order):**

1. `parse_any_xml()` — opens the `.docx` ZIP and parses `document.xml`, `styles.xml`, `footnotes.xml`, and `endnotes.xml` into lxml element trees.
2. `build_outline_level_map()` — builds a `styleId → outline level` map from `styles.xml` for heading detection.
3. `build_image_map()` — reads `word/_rels/document.xml.rels` and extracts raw image bytes keyed by relationship ID.
4. `extract_tracked_changes()` — walks every `<w:p>` and `<w:tbl>` in the document body, emitting `ChangeRecord` objects in document order.
5. `extract_hf_changes()` — optionally parses all header/footer XML parts and extracts their tracked changes.
6. `AppConfig.filter_records()` — filters the combined record list to only the change types selected by the user.
7. `inject_bookmarks()` — writes an annotated copy of the `.docx` with `<w:bookmarkStart>` / `<w:bookmarkEnd>` pairs injected at each change location.
8. `generate_excel_output()` or `generate_excel_from_template()` — renders the filtered records into an `.xlsx` workbook with hyperlinks, diff highlighting, and embedded images.

## Project Structure

```
compare_crs_docx/
│
├── constants.py          # XML namespace map, ChangeType registry, project-wide constants
├── models.py             # ChangeRecord dataclass
├── parser.py             # .docx ZIP parsing, XML tree construction, image/HF discovery
├── heading_resolver.py   # Paragraph outline-level resolution and heading text extraction
├── run_utils.py          # Low-level run/rPr text and property helpers
├── classifier.py         # classify_change_type() — unified change classification
├── para_utils.py         # Paragraph-level dispatch: del/ins/run handlers, SDT unwrapping
├── rpr_change_utils.py   # rPrChange detection, merging, and flushing
├── embedded_content.py   # Embedded object detection, image extraction, note helpers
├── table_utils.py        # Table/row/cell change extraction
├── extractor.py          # Top-level extraction engine (body + header/footer)
├── book_marker.py        # Bookmark injection into annotated .docx
├── output.py             # Excel output generation (scratch and template modes)
├── app_config.py         # AppConfig dataclass with JSON persistence
├── gui.py                # PyQt6 GUI entry point
├── splash_scene.py       # Manim splash animation scene (run once to regenerate)
├── main.py               # CLI entry point (development/testing)
│
└── build_exe_tools/
    ├── build.bat                          # PyInstaller build script
    └── CPAuto.spec                        # PyInstaller spec file

```

## File Descriptions

| File | Responsibility |
|---|---|
| `constants.py` | Defines the OOXML XML namespace map (`NS`), Clark-notation prefixes (`NSC`), the `ChangeType` class (central registry of all change-type label strings), `CHANGE_TYPE_CATEGORIES` (GUI grouping), `ALL_CHANGE_TYPES` (flat frozenset), and all project-wide tunable constants (image dimensions, diff colours, delimiters, bookmark prefix, etc.) |
| `models.py` | Defines the `ChangeRecord` dataclass: `heading`, `old_text`, `new_text`, `change_type`, `bookmark_id`, `target_element`, `style_meta`, `old_image_data`, `new_image_data`, `old_image_size_px`, `new_image_size_px` |
| `parser.py` | Opens the `.docx` ZIP archive; parses `document.xml`, `styles.xml`, `footnotes.xml`, and `endnotes.xml` via `parse_any_xml()`; builds the `styleId → outline level` map via `build_outline_level_map()`; builds the `rId → image bytes` map via `build_image_map()`; discovers and parses header/footer XML parts via `parse_headers_footers()` |
| `heading_resolver.py` | Resolves the outline level (0–8) of a paragraph via `get_paragraph_outline_level()`, checks whether a paragraph is a navigation heading via `is_heading()`, and extracts its plain text via `get_heading_text()` |
| `run_utils.py` | Low-level helpers for extracting text from `<w:r>` runs (`run_text()`) and `<w:del>` runs (`del_text()`), detecting run-property states (`has_rpr_property()`, `get_rpr_property_states()`), reading attribute values (`get_rpr_attr_val()`), and resolving colour strings (`resolve_color()`) |
| `classifier.py` | Provides `classify_change_type()` — the single, unified entry point for assigning a semantic `ChangeType` label to any tracked change, using a priority-ordered rule set over text content, `RprDetection` flags, `ValDetection` values, and structural type overrides |
| `para_utils.py` | Paragraph-level dispatch engine: `unwrap_sdts()` (flatten `<w:sdt>` elements), `get_ppr_style_change()` (detect paragraph style changes), `collect_consecutive_del/ins()`, `handle_del()`, `handle_ins()`, `handle_run()` (rPrChange buffering), and `get_paragraph_text()` |
| `rpr_change_utils.py` | Handles `<w:rPrChange>` run-property changes: `get_rpr_change_meta()` (extract author/date), `can_merge_rpr_runs()` (merge consecutive same-type rPr runs), `flush_rpr_buffer()` (emit one `ChangeRecord` for a merged run group) |
| `embedded_content.py` | Detects and classifies embedded non-text content (`has_embedded_content()`, `classify_embedded_content()`), extracts image bytes and size (`extract_image_data()`, `extract_image_size_px()`), handles footnote/endnote references (`has_note_reference()`, `get_note_ref_info()`), and emits object/note records (`emit_object_records()`, `emit_note_records()`, `emit_plain_note_records()`) |
| `table_utils.py` | Walks `<w:tbl>` elements: `check_whole_table_change()` (collapse all-ins/all-del tables), `process_row()` (row-level ins/del), `process_cell()` (cell ins/del, vertical merge, content changes), `extract_table_changes()` (top-level table dispatcher) |
| `extractor.py` | Top-level extraction engine: `extract_changes_from_paragraph()` (processes a single `<w:p>`), `extract_tracked_changes()` (walks the full document body), `extract_hf_changes()` (processes header/footer roots), `_extract_embedded_container_changes()` (textbox extraction) |
| `book_marker.py` | Injects `<w:bookmarkStart>` / `<w:bookmarkEnd>` pairs into a fresh copy of the `.docx` via `inject_bookmarks()`, using child-index paths (`_element_path()`, `_resolve_path()`) to locate target elements across independently-parsed trees |
| `output.py` | Renders `ChangeRecord` lists into Excel workbooks: `generate_excel_output()` (from-scratch mode, 4-column layout), `generate_excel_from_template()` (template mode, configurable column mapping); includes `_make_diff_rich_text()` (word-level diff), `_embed_image()` (TwoCellAnchor image embedding), and `_apply_row_style()` (colour/strikethrough styling) |
| `app_config.py` | `AppConfig` dataclass: all user-configurable settings (text, image, diff colours, change-type selection, HF/SDT toggles, output mode, template selection, column mapping); `save()` / `load()` for JSON persistence; `validate()` for GUI-side validation; `filter_records()` for change-type filtering |
| `gui.py` | PyQt6 main window (`MainWindow`), settings panel (`_SettingsPanel`), change-type tree (`_ChangeTypeTree`), output settings panel (`_OutputSettingsPanel`), template profile manager dialog (`_TemplateProfileDialog`), background pipeline worker thread (`_PipelineWorker`), splash screen (`_SplashScreen`), and `launch()` entry point |
| `splash_scene.py` | Manim `CPAutoSplash` scene that renders the animated splash video to `assets/splash_assets/splash.mp4`. Run once manually to regenerate the video after modifying the animation. |
| `main.py` | CLI entry point for development and scripting. Runs the full pipeline (parse → extract → annotate → Excel) without the GUI. Accepts an optional `.docx` path as a command-line argument. |

## Prerequisites

### Python Version
Python **3.9 or newer** is required.

### Dependencies

Install all runtime dependencies with:

```bash
pip install lxml openpyxl PyQt6 PyQt6-Qt6 PyQt6-Multimedia
```

| Package | Purpose | Required for |
|---|---|---|
| `lxml` | XPath-based OOXML XML parsing | Core pipeline (always required) |
| `openpyxl` | Excel workbook generation and template manipulation | Core pipeline (always required) |
| `PyQt6` | GUI framework (widgets, layouts, dialogs) | GUI mode only |
| `PyQt6-Qt6` | Qt6 runtime libraries | GUI mode only |
| `PyQt6-Multimedia` | Video playback for the splash screen | GUI mode only |
| `manim` | Rendering the splash animation video | Development only — run `splash_scene.py` once to regenerate `splash.mp4` |

> All other modules (`zipfile`, `dataclasses`, `difflib`, `json`, `shutil`, `re`, `sys`, `pathlib`, `typing`) are part of the Python standard library and require no installation.

> **Microsoft Word is not required.** The tool reads `.docx` files directly as ZIP archives and parses their OOXML content offline.

## Input Requirements

The input must be a `.docx` file generated via **Microsoft Word's built-in Compare feature**:

1. Open Microsoft Word.
2. Go to **Review → Compare → Compare Documents**.
3. Select the original document and the revised document.
4. Word generates a new document containing all differences as tracked changes.
5. **Save that resulting document** as a `.docx` file (e.g. `compare.docx`).

> ⚠️ **CPAuto does not perform its own document comparison.** It only parses and extracts the tracked changes that Word has already embedded in the comparison document. The input must be a Word-generated comparison file, not an arbitrary `.docx` with manual tracked changes.

## Usage

### GUI Mode (Recommended)

Launch the desktop application:

```bash
python gui.py
```

**Workflow:**

1. A splash screen plays on launch (click anywhere to skip).
2. The main window opens. Click **Browse…** or click the document path field to select a `.docx` comparison file.
3. Adjust settings in the left panel (text, image, diff colours, output mode).
4. Select or deselect change types in the right panel tree.
5. Click **▶ Run** to execute the pipeline. Progress is shown in the status bar.
6. When complete, a dialog confirms the output path. Both output files are saved to the `output/` directory.

### CLI Mode (Development / Scripting)

```bash
python main.py [docx_path]
```

| Argument | Required | Default | Description |
|---|---|---|---|
| `docx_path` | No | `compare.docx` | Path to the Word comparison document |

**Example:**

```bash
python main.py my_compare.docx
```

Output files are written to the `output/` directory (created automatically):

- `output/<stem>_changes.xlsx` — Excel change table
- `output/<stem>_annotated.docx` — Bookmark-annotated copy of the source document

> The CLI mode uses module-level constant defaults and does not apply GUI settings or change-type filtering. It is intended for development and testing only.

## GUI Overview

The main window is divided into three zones:

```
┌──────────────────────────────────────────────────────────────────┐
│  Document: [_______________________________]  [ Browse… ]        │  ← File Bar
├────────────────────────┬─────────────────────────────────────────┤
│  Settings Panel        │  Change Type Tree                       │
│  (scrollable)          │                                         │
│                        │  ☑ Include Header/Footer Changes        │
│  ┌ Text Settings ────┐ │  ☑ Include SDT Changes                  │
│  │ Empty Field       │ │                                         │
│  │ Default Heading   │ │  [ Select All ]  [ Deselect All ]       │
│  └───────────────────┘ │                                         │
│                        │  ▼ Content Changes                      │
│  ┌ Image Settings ───┐ │    ☑ Add Content                        │
│  │ ○ Original Size   │ │    ☑ Delete Content                     │
│  │ ○ Fixed Size      │ │    ☑ Modify Content                     │
│  │ Width  [slider]   │ │  ▼ Highlight Changes                    │
│  │ Height [slider]   │ │    ☑ Text Highlight Colour Added        │
│  └───────────────────┘ │    ☑ Text Highlight Colour Removed      │
│                        │    ☑ Text Highlight Colour Changed      │
│  ┌ Diff Colors ──────┐ │  ▼ Table Changes                        │
│  │ Delete  [■] [hex] │ │    ☑ Add Table  ☑ Delete Table …        │
│  │ Insert  [■] [hex] │ │                                         │
│  │ Equal   [■] [hex] │ │                                         │
│  └───────────────────┘ │                                         │
│                        │                                         │
│  ┌ Output Settings ──┐ │                                         │
│  │ ○ From Scratch    │ │                                         │
│  │ ○ From Template   │ │                                         │
│  │ Template: [combo] │ │                                         │
│  │ Sheet:    [combo] │ │                                         │
│  │ Col mapping …     │ │                                         │
│  └───────────────────┘ │                                         │
│                        │                                         │
│  ┌ Delete Toggle ────┐ │                                         │
│  │ ☑ Enable Toggle   │ │                                         │
│  │ Column: [E]       │ │                                         │
│  │ [Delete Toggled…] │ │                                         │
│  └───────────────────┘ │                                         │
├────────────────────────┴─────────────────────────────────────────┤
│  [ Reset to Default ]          [ Save Config ]  [ ▶  Run ]       │  ← Action Bar
├──────────────────────────────────────────────────────────────────┤
│  Status bar: pipeline log messages                               │
└──────────────────────────────────────────────────────────────────┘
```

### Settings Panel (Left)

**Text Settings**
- **Empty Field Placeholder** — text displayed in Excel when a field has no content (default: `N/A`).
- **Default Heading** — heading label used for changes that appear before any document heading (default: `Link to CP`).

**Image Settings**
- **Image Size Mode** — `Original Size` uses each image's own `wp:extent` dimensions from the document (falling back to configured size if unavailable); `Fixed Size` applies the configured width/height to all images uniformly.
- **Default Image Width / Height** — pixel dimensions for embedded images in Excel (range: 50–500 px). The auto-calculated row height (in points) is displayed below the sliders.

**Diff Colors**
- **Delete Color** — hex RGB colour for deleted words in *Modify Content* diff highlighting (default: `FF0000` red).
- **Insert Color** — hex RGB colour for inserted words (default: `00B050` green).
- **Equal Color** — hex RGB colour for unchanged words (default: `000000` black).
- Click the colour swatch button to open a colour picker, or read the hex value from the adjacent read-only field.

**Output Settings**
- **Generate from Scratch** — creates a new `.xlsx` workbook with a fixed 4-column layout.
- **Generate from Template** — copies a CPA template file from `assets/CPA_template/` and writes records into a user-selected sheet with a configurable column mapping.
- **Template / Sheet selectors** — populated from the available `.xlsx` files in `assets/CPA_template/` and the sheets within the selected template.
- **Column mapping** — four letter fields (Heading, Change Type, Old Content, New Content) specifying the Excel column for each output field.
- **Manage Template Profiles…** — opens a dialog to add, edit, or delete per-sheet column-mapping profiles that are persisted across sessions.

**Delete Toggle Feature**
- **Enable Delete Toggle Column** — when checked, adds a toggle column with TRUE/FALSE dropdowns to the Excel output for marking rows for deletion.
- **Toggle Column** — Excel column letter where the toggle column will be added (default: "E").
- **Delete Toggled Rows from Excel…** — opens a file dialog to select an Excel file with toggle column, shows a summary of rows marked for deletion, and batch-deletes them with automatic backup creation.
- **Auto-Delete Watcher** — Run `python auto_delete_watcher.py <excel_file>` to automatically delete rows when you save Excel (recommended for best user experience).
- See [DELETE_TOGGLE_GUIDE.md](DELETE_TOGGLE_GUIDE.md) for complete guide with all methods, workflows, and troubleshooting.

### Change Type Tree (Right)

A hierarchical checkbox tree grouped by category (Content Changes, Strikethrough Changes, Highlight Changes, Font Color Changes, Format Changes, Style Changes, Object Changes, Note Changes, Table Changes). Each leaf node corresponds to one `ChangeType` constant.

- **Include Header/Footer Changes** — when checked, header and footer XML parts are parsed and their tracked changes are included in the output.
- **Include SDT (Content Control) Changes** — when checked, `<w:sdt>` elements are unwrapped and their content is included in extraction.
- **Select All / Deselect All** — convenience buttons to toggle all change types at once.
- `**Unknown**` changes are always included regardless of the tree selection, so unclassified changes are never silently discarded.

### Action Bar

- **Reset to Default** — resets all settings to their `constants.py` defaults in-place (template profiles and column mappings are intentionally preserved).
- **Save Config** — validates and persists the current settings to `data/config.json`.
- **▶ Run** — validates settings, then runs the full pipeline on a background thread (`_PipelineWorker`) to keep the UI responsive.

## Configuration and Customization

All user-configurable settings are managed by the `AppConfig` dataclass in `app_config.py` and persisted to `data/config.json`. The GUI reads and writes this file automatically; manual edits to the JSON are also supported.

### Configurable Constants

The following constants in `constants.py` serve as the system-wide defaults. `AppConfig` mirrors each of these and overrides them at runtime:

| Constant | Default | Description |
|---|---|---|
| `EMPTY_FIELD` | `"N/A"` | Placeholder text for empty old/new content fields in Excel |
| `DEFAULT_HEADING` | `"Link to CP"` | Heading label for changes before the first document heading |
| `IMAGE_WIDTH_PX` | `380` | Default image width in pixels for Excel embedding |
| `IMAGE_HEIGHT_PX` | `380` | Default image height in pixels for Excel embedding |
| `IMAGE_SIZE_MODE` | `"original"` | Image sizing mode: `"original"` or `"fixed"` |
| `IMAGE_PX_MIN` | `50` | Minimum allowed image dimension (px) |
| `IMAGE_PX_MAX` | `500` | Maximum allowed image dimension (px) |
| `DIFF_DEL_COLOR` | `"FF0000"` | Hex RGB for deleted words in Modify Content diff |
| `DIFF_INS_COLOR` | `"00B050"` | Hex RGB for inserted words in Modify Content diff |
| `DIFF_EQ_COLOR` | `"000000"` | Hex RGB for unchanged words in Modify Content diff |
| `OUTPUT_MODE_SCRATCH` | `"scratch"` | Output mode: generate a new workbook |
| `OUTPUT_MODE_TEMPLATE` | `"template"` | Output mode: fill in a CPA template |
| `DEFAULT_TEMPLATE_COLUMNS` | `["A","B","C","D"]` | Default column letters for template output |
| `BOOKMARK_PREFIX` | `"docChg_"` | Prefix for injected bookmark names |
| `MERGE_TOLERANCE_SECONDS` | `3` | Max timestamp gap (seconds) for merging consecutive rPrChange runs |

### Configuration File (`data/config.json`)

The configuration file is auto-generated on first save. A typical file looks like:

```json
{
  "empty_field": "N/A",
  "default_heading": "Link to CP",
  "image_width_px": 380,
  "image_height_px": 380,
  "image_size_mode": "original",
  "diff_del_color": "FF0000",
  "diff_ins_color": "00B050",
  "diff_eq_color": "000000",
  "selected_change_types": ["Add Content", "Delete Content", "Modify Content", "..."],
  "include_hf_changes": true,
  "include_sdt": true,
  "include_caption_changes": false,
  "output_mode": "scratch",
  "selected_template": "",
  "template_profiles": {
    "CPA_Template_noAutoCount.xlsx": {
      "CPA_JPN": ["A", "B", "C", "D"]
    }
  },
  "template_selected_sheets": {
    "CPA_Template_noAutoCount.xlsx": "CPA_JPN"
  },
  "enable_delete_toggle": false,
  "delete_toggle_column": "E"
}
```

### Change Type Filtering

The `selected_change_types` field stores the set of `ChangeType` string values that should appear in the output. Any change type not in this set is filtered out by `AppConfig.filter_records()` before output generation. The `ChangeType.UNKNOWN` type (`"**Unknown**\nReview manually!"`) is always passed through regardless of the filter, ensuring no unclassified changes are silently discarded.

### Template Profile Management

Per-template, per-sheet column mappings are stored in `template_profiles`. The key is the template filename (basename only); the value is a dict mapping sheet names to a list of four Excel column letters `[heading, change_type, old, new]`. These profiles survive a **Reset to Default** operation.

## Excel Output System

All Excel rendering is handled by `output.py` using the `openpyxl` library.

### Output Modes

**From Scratch (`generate_excel_output`)**

Creates a new `.xlsx` workbook with a single sheet named `Changes` and a fixed 4-column layout:

| Column A | Column B | Column C | Column D |
|---|---|---|---|
| Heading / Sub-heading | Change Type | Old Content | New Content |

**From Template (`generate_excel_from_template`)**

Copies a CPA template file from `assets/CPA_template/` to the output directory, then writes records into the target sheet starting at row 2 (row 1 is the template header and is never modified). Column positions are determined by the user-configured column mapping.

### Heading Cell Hyperlinks

Each heading cell is written as an Excel `HYPERLINK()` formula that links to the corresponding bookmark in the annotated `.docx`:

```
=HYPERLINK("<stem>_annotated.docx#docChg_00000001", "<heading text>")
```

The heading cell is styled with blue underlined font to visually indicate the hyperlink. Clicking the cell in Excel opens the annotated document and jumps directly to the change location.

### Inline Diff Highlighting (Modify Content)

For *Modify Content* rows, the Old Content and New Content cells are rendered as `CellRichText` objects with word-level diff colouring:

- **Deleted words** — coloured with `diff_del_color` (default: red `FF0000`)
- **Inserted words** — coloured with `diff_ins_color` (default: green `00B050`)
- **Unchanged words** — coloured with `diff_eq_color` (default: black `000000`)

The diff is computed by `difflib.SequenceMatcher` at the word level. Newlines are preserved as explicit line-break tokens.

### Cell Styling by Change Type

| Change Type | Styling Applied |
|---|---|
| Modify Content | Word-level diff rich text in Old/New cells |
| Strikethrough Added | New Content cell font has strikethrough; Old Content does not |
| Strikethrough Removed | Old Content cell font has strikethrough; New Content does not |
| Text Highlight Colour Added/Removed/Changed | Cell background fill set to the old/new highlight colour |
| Font Colour Changed | Cell font colour set to the old/new font colour |
| All others | Monospace (`Consolas`) font applied to content cells |

### Image Embedding

When a `ChangeRecord` carries `old_image_data` or `new_image_data`, the image bytes are embedded into the corresponding Excel cell using a `TwoCellAnchor` with `fLocksWithSheet=True`. This causes the image to hide with its row when rows are filtered, preventing images from floating over unrelated rows.

Image dimensions are calculated as follows:
- **Original Size mode** — uses the `wp:extent` dimensions from the document (in EMUs, converted to pixels), clamped to the configured min/max range. Falls back to the configured default if the original size is unavailable.
- **Fixed Size mode** — always uses the configured `image_width_px` × `image_height_px`.

Row height is automatically set to accommodate the image: `height_pt = int(image_height_px × 0.75) + 20`.

### Table Cell Text Formatting

Multi-paragraph cell content uses `¶` as an internal paragraph delimiter (stored) and `\n` for display. Multi-column table rows use `¤` as an internal cell delimiter (stored) and ` | ` for display. These conversions are applied by `_format_content()` before writing to Excel.

## Supported Change Types

All change types are defined as string constants on the `ChangeType` class in `constants.py` and grouped into categories for the GUI tree. The table below lists every supported type, its category, and the OOXML source element that triggers it.

### Content Changes

| Change Type | `ChangeType` Constant | Description | OOXML Source |
|---|---|---|---|
| Add Content | `ADD_CONTENT` | Text present only in the new version | `<w:ins>` containing `<w:r>` with `<w:t>` |
| Delete Content | `DELETE_CONTENT` | Text present only in the original version | `<w:del>` containing `<w:r>` with `<w:delText>` |
| Modify Content | `MODIFY_CONTENT` | Text changed between old and new versions | `<w:del>` immediately followed by `<w:ins>` |

### Strikethrough Changes

| Change Type | `ChangeType` Constant | Description |
|---|---|---|
| Strike-through Added | `STRIKETHROUGH_ADDED` | `<w:strike>` or `<w:dstrike>` added via `<w:rPrChange>` |
| Strike-through Removed | `STRIKETHROUGH_REMOVED` | `<w:strike>` or `<w:dstrike>` removed via `<w:rPrChange>` |

### Highlight Changes

| Change Type | `ChangeType` Constant | Description |
|---|---|---|
| Text Highlight Colour Added | `HIGHLIGHT_ADDED` | `<w:highlight>` or `<w:shd>` added via `<w:rPrChange>` |
| Text Highlight Colour Removed | `HIGHLIGHT_REMOVED` | `<w:highlight>` or `<w:shd>` removed via `<w:rPrChange>` |
| Text Highlight Colour Changed | `HIGHLIGHT_CHANGED` | Highlight colour value changed via `<w:rPrChange>` |

### Font Color Changes

| Change Type | `ChangeType` Constant | Description |
|---|---|---|
| Font Colour Changed | `FONT_COLOR_CHANGED` | `<w:color>/@w:val` changed via `<w:rPrChange>` |

### Format Changes

| Change Type | `ChangeType` Constant | Description |
|---|---|---|
| Format Change | `FORMAT_CHANGE` | Any other run-property change (bold, italic, font size, etc.) via `<w:rPrChange>` |

### Style Changes

| Change Type | `ChangeType` Constant | Description |
|---|---|---|
| Paragraph Style Change | `STYLE_CHANGE` | Paragraph style changed via `<w:pPrChange>` |

### Object Changes

| Change Type | `ChangeType` Constant | Description |
|---|---|---|
| Add non-text Object | `ADD_OBJECT` | Drawing, image, chart, textbox, or OLE object inserted (`<w:ins>`) |
| Delete non-text Object | `DELETE_OBJECT` | Drawing, image, chart, textbox, or OLE object deleted (`<w:del>`) |
| Modify non-text Object | `MODIFY_OBJECT` | A deleted object immediately followed by an inserted object |

### Note Changes

| Change Type | `ChangeType` Constant | Description |
|---|---|---|
| Add Note | `ADD_NOTE` | Footnote or endnote reference inserted (`<w:ins>` containing `<w:footnoteReference>` / `<w:endnoteReference>`) |
| Delete Note | `DELETE_NOTE` | Footnote or endnote reference deleted (`<w:del>` containing a note reference) |

### Table Changes

| Change Type | `ChangeType` Constant | Description |
|---|---|---|
| Add Table | `ADD_TABLE` | All rows in a `<w:tbl>` are inserted (`<w:trPr/w:ins>` on every row) |
| Delete Table | `DELETE_TABLE` | All rows in a `<w:tbl>` are deleted (`<w:trPr/w:del>` on every row) |
| Add Row | `ADD_ROW` | A single table row is inserted (`<w:trPr/w:ins>`) |
| Delete Row | `DELETE_ROW` | A single table row is deleted (`<w:trPr/w:del>`) |
| Add Cell | `ADD_CELL` | A table cell is inserted (`<w:tcPr/w:cellIns>`) |
| Delete Cell | `DELETE_CELL` | A table cell is deleted (`<w:tcPr/w:cellDel>`) |
| Merge Cell | `MERGE_CELL` | A vertical cell merge is tracked (`<w:tcPr/w:cellMerge @w:vMerge="restart">`) |

### Special

| Change Type | `ChangeType` Constant | Description |
|---|---|---|
| Unknown | `UNKNOWN` | `<w:rPrChange>` present but no classifiable property change detected. Always included in output regardless of filter settings. |

## Delete Toggle Feature

The Delete Toggle feature provides a workflow for filtering Excel output by marking unwanted rows for deletion. This is particularly useful for large change reports where manual review is needed.

### Quick Start

1. **Enable in GUI**: Check "Enable Delete Toggle Column" in the Delete Toggle Feature section
2. **Run Pipeline**: Generate Excel output as normal - it will include a toggle column
3. **Mark Rows**: Open the Excel file and set toggle to TRUE for rows you want to delete
4. **Delete**: Click "Delete Toggled Rows from Excel..." in the GUI, select the file, and confirm

### Features

- **Dropdown Validation**: Toggle column uses Excel data validation with TRUE/FALSE dropdown
- **Batch Deletion**: Delete all marked rows in one operation
- **Automatic Backup**: Creates `.backup.xlsx` before modifying files
- **Summary Preview**: Shows count of rows to be deleted before confirmation
- **Column Removal**: Option to remove toggle column after deletion
- **Command-Line Tools**: Standalone CLI for adding toggle columns, showing summaries, and deleting rows
- **Python API**: Programmatic access for custom workflows

### Detailed Documentation

For complete usage instructions, workflow examples, troubleshooting, and API reference, see:
- **[DELETE_TOGGLE_GUIDE.md](DELETE_TOGGLE_GUIDE.md)** — Complete delete toggle feature guide (all methods consolidated)

### Testing

The delete toggle feature includes comprehensive test coverage:

```bash
# Run all delete toggle tests
python -m pytest tests/test_delete_toggle.py -v

# Expected: 16 tests covering all functionality
```

## Output Format

Each pipeline run produces two files in the `output/` directory:

### `<stem>_changes.xlsx` — Excel Change Table

A structured workbook where each row represents one `ChangeRecord`. The default (from-scratch) layout uses four columns:

| Heading / Sub-heading | Change Type | Old Content | New Content |
|---|---|---|---|
| `=HYPERLINK(...)` (blue, underlined) | e.g. `Modify Content` | Original text or image | Revised text or image |

- **Heading column** — contains a `HYPERLINK()` formula linking to the exact paragraph in the annotated `.docx`. Clicking navigates directly to the change.
- **Change Type column** — contains the `ChangeType` string label.
- **Old Content / New Content columns** — contain plain text, rich-text diff (for Modify Content), or embedded images.
- Footnote/endnote changes append a label suffix to the heading (e.g. `Section 3 - Footnote`).
- Table changes append ` - Table` to the heading (e.g. `Section 2 - Table`).
- Textbox changes append ` - Textbox` to the heading.

### `<stem>_annotated.docx` — Bookmark-Annotated Document

A copy of the source `.docx` with `<w:bookmarkStart>` / `<w:bookmarkEnd>` pairs injected at the start of each change's target paragraph. Bookmark names follow the pattern `docChg_XXXXXXXX` (zero-padded 8-digit integer). These bookmarks are the targets of the `HYPERLINK()` formulas in the Excel output.

## Dependency Flow

```
gui.py / main.py
 ├── app_config.py        →  constants.py
 ├── parser.py            →  constants.py
 ├── extractor.py         →  constants.py, models.py
 │    ├── heading_resolver.py  →  constants.py, para_utils.py
 │    ├── para_utils.py        →  constants.py, models.py, run_utils.py,
 │    │                            classifier.py, rpr_change_utils.py,
 │    │                            embedded_content.py
 │    ├── rpr_change_utils.py  →  constants.py, models.py, run_utils.py,
 │    │                            classifier.py
 │    ├── embedded_content.py  →  constants.py, models.py, classifier.py
 │    ├── table_utils.py       →  constants.py, models.py, classifier.py,
 │    │                            para_utils.py
 │    └── classifier.py        →  constants.py
 ├── book_marker.py       →  constants.py
 └── output.py            →  constants.py, models.py, run_utils.py
```

## Known Limitations & Notes

- **No self-comparison capability** — CPAuto only reads tracked changes already embedded by Word's Compare feature. It does not compare two arbitrary documents itself.
- **Header/footer section context** — changes inside headers and footers are labelled `"Header"` or `"Footer"` without a more specific section heading, since headers/footers exist outside the main document body flow.
- **Theme-based highlight colours** — highlights expressed via Word theme references (rather than explicit `<w:highlight>` or `<w:shd>` values) may not be detected or resolved to a colour.
- **Adjacent del/ins pairs** — the tool classifies a `<w:del>` immediately followed by `<w:ins>` as *Modify Content*. In rare cases, Word may produce adjacent del/ins pairs that represent two independent changes rather than a single modification.
- **Partial run splitting** — Word sometimes splits text across multiple `<w:r>` runs within the same tracked-change block. The tool concatenates them correctly, but very fragmented runs may result in slightly unexpected diff boundaries.
- **rPrChange merging** — consecutive runs with `<w:rPrChange>` are merged into a single record when they share the same author, timestamp (within `MERGE_TOLERANCE_SECONDS`), change type, and property value fingerprint. Runs that differ on any of these axes produce separate records.
- **Vertical-merge continuation rows** — table rows that are pure vertical-merge continuations (`<w:vMerge>` without `val="restart"`) are silently skipped to avoid duplicate records for the same logical cell span.
- **SDT auto-content** — `<w:sdt>` elements such as Table of Contents fields are typically auto-calculated by Word and are not genuine user edits. The *Include SDT Changes* toggle allows users to exclude these from the output.
- **Splash video dependency** — the animated splash screen requires `assets/splash_assets/splash.mp4` to be present. If the file is missing, the GUI launches directly without a splash screen. Run `splash_scene.py` once to regenerate the video.
- **PyInstaller one-file mode** — `BASE_DIR` is resolved to `sys.executable`'s parent (not `sys._MEIPASS`) to ensure external files (`assets/`, `data/`, `output/`) are always located relative to the real executable, not the temporary extraction directory.

## Building an Executable

A standalone Windows executable can be built using PyInstaller via the provided build tools in `build_exe_tools/`.

### Prerequisites

Install PyInstaller into a dedicated virtual environment:

```bash
python -m venv .build_venv
.build_venv\Scripts\activate
pip install pyinstaller lxml openpyxl PyQt6 PyQt6-Qt6 PyQt6-Multimedia
```

### Build

```bash
cd build_exe_tools
build.bat
```

The `build.bat` script invokes PyInstaller with `CPAuto.spec`. The resulting executable is placed in `dist/CPAuto/` (one-folder mode) or `dist/CPAuto.exe` (one-file mode), depending on the spec configuration.

### Notes

- The `assets/`, `data/`, and `output/` directories must reside alongside the executable (not inside it) so that templates, config, and output files are accessible at runtime.
- `BASE_DIR` in `constants.py` is resolved via `sys.executable` (not `sys._MEIPASS`) to ensure correct path resolution in both one-file and one-folder PyInstaller builds.
- The `output/` and `data/` directories are created automatically at runtime if they do not exist.

---

## Extensibility and Future Improvements

The system is designed for incremental extension. The following areas are straightforward to extend:

### Adding a New Change Type

1. Add a new string constant to the `ChangeType` class in `constants.py`.
2. Add it to the appropriate category in `CHANGE_TYPE_CATEGORIES` (or create a new category).
3. Update `classify_change_type()` in `classifier.py` to return the new type when appropriate, or pass it as a `structural_type` override from the extractor.
4. Add any new Excel styling logic to `_apply_row_style()` in `output.py` if needed.

### Adding a New Embedded Container Type

To extract tracked changes from a new type of embedded container (e.g. a new DrawingML shape type):

1. Append a new `ContainerDescriptor` to `_BODY_CONTAINER_REGISTRY` in `extractor.py`:
   ```python
   ContainerDescriptor(xpath=".//new:element", label="MyContainer")
   ```
2. Ensure the corresponding namespace is registered in `NS` in `constants.py`.

### Adding a New Run-Property Detection

To detect a new `<w:rPrChange>` property (e.g. underline added/removed):

1. Add new `ChangeType` constants (e.g. `UNDERLINE_ADDED`, `UNDERLINE_REMOVED`).
2. In `rpr_change_utils.py`, add an `RprDetection` entry to `_classify_run_change_type()`.
3. Add the new types to `CHANGE_TYPE_CATEGORIES`.

### Adding a New Output Format

To add a new output format (e.g. CSV, HTML):

1. Create a new function in `output.py` (or a new module) that accepts `list[ChangeRecord]` and writes the desired format.
2. Wire it into the pipeline in `gui.py` (`_PipelineWorker.run()`) and `main.py` as needed.
3. Add any new GUI controls for the format selection to `_OutputSettingsPanel` in `gui.py`.

### Adding a New Template

Place any `.xlsx` file in `assets/CPA_template/`. It will automatically appear in the Template selector dropdown on the next GUI launch. Use **Manage Template Profiles…** to configure the column mapping for each sheet.
