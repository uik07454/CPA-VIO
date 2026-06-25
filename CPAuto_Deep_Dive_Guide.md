# CPAuto — Complete Deep-Dive Guide

> **Reading order:** Sections 1 → 10. Each section builds on the previous one.
> After reading this guide you will be able to trace any change from raw XML all the way to an Excel cell.

---

## 1. High-Level Overview

### What is CPAuto?

CPAuto is a **Word Compare Difference Extractor**. Its job is simple to state:

> Take a `.docx` file that was produced by Microsoft Word's **Compare Documents** feature, find every tracked change inside it, and write a clean, navigable Excel report.

### Why does this problem exist?

When Word compares two documents it does **not** give you a spreadsheet — it gives you a `.docx` full of coloured markup (`w:ins` / `w:del` XML tags). Reading that markup manually is slow and error-prone. CPAuto automates the extraction.

### The Four Layers

```
┌─────────────────────────────────────────────────────────┐
│  Layer 1 — GUI  (gui.py)                                │
│  User picks a file, sets options, clicks Run.           │
├─────────────────────────────────────────────────────────┤
│  Layer 2 — Parsing  (parser.py)                         │
│  Opens the .docx ZIP, parses XML trees with lxml.       │
├─────────────────────────────────────────────────────────┤
│  Layer 3 — Extraction & Classification                  │
│  (extractor.py, para_utils.py, table_utils.py,          │
│   embedded_content.py, classifier.py,                   │
│   rpr_change_utils.py, run_utils.py,                    │
│   heading_resolver.py)                                  │
│  Walks the XML tree, finds every change, labels it.     │
├─────────────────────────────────────────────────────────┤
│  Layer 4 — Output  (book_marker.py, output.py)          │
│  Injects bookmarks into a copy of the .docx,            │
│  then writes an Excel file with hyperlinks.             │
└─────────────────────────────────────────────────────────┘
```

### Why is it split into layers?

| Layer | Reason for separation |
|---|---|
| GUI | Keeps visual code away from logic; the pipeline can also run headless (CLI) |
| Parsing | Isolates all ZIP/XML I/O; nothing else touches the file system for reading |
| Extraction | The most complex layer — split further into helpers so each file has one job |
| Output | Isolates openpyxl (Excel library) from the rest; easy to swap output format |

### The Pipeline in One Sentence

```
.docx file
  → unzip & parse XML          (parser.py)
  → walk every paragraph/table (extractor.py)
  → classify each change        (classifier.py)
  → filter by user settings     (app_config.py)
  → inject bookmarks            (book_marker.py)
  → write Excel with hyperlinks (output.py)
```

---

## 2. End-to-End Execution Flow

### Case 1 — GUI (`gui.py`)

```
User launches the app
  │
  ▼
launch()                         # gui.py — creates QApplication
  │
  ├─► _SplashScreen              # plays splash.mp4 fullscreen
  │
  └─► MainWindow.__init__()      # loads AppConfig from data/config.json
        │
        ▼
      User selects .docx file via Browse button
        │
        ▼
      User adjusts settings (change types, image size, colors, output mode)
        │
        ▼
      User clicks ▶ Run
        │
        ▼
      MainWindow._on_run()
        ├─ _collect_config_from_ui()   # reads all widget values into AppConfig
        └─ _PipelineWorker.start()     # runs pipeline on a background QThread
              │
              ▼
           _PipelineWorker.run()       # ← THE REAL WORK HAPPENS HERE
              │
              ├─ parse_any_xml()           → doc_root, styles_root, footnotes, endnotes
              ├─ build_outline_level_map() → outline_map  {styleId: level}
              ├─ build_image_map()         → image_map    {rId: bytes}
              ├─ extract_tracked_changes() → records      [ChangeRecord, ...]
              ├─ extract_hf_changes()      → hf_records   [ChangeRecord, ...]
              ├─ AppConfig.filter_records()→ all_records  (filtered)
              ├─ inject_bookmarks()        → annotated .docx saved to output/
              └─ generate_excel_output()   → .xlsx saved to output/
                    │
                    ▼
              _PipelineWorker emits finished signal
                    │
                    ▼
              MainWindow._on_pipeline_finished()
                    └─ shows QMessageBox "Done!"
```

### Case 2 — CLI (`main.py`)

```
python main.py compare.docx
  │
  ▼
main(docx_path)
  │
  ├─ parse_any_xml(docx_path, TO_PARSED_XML_PARTS)
  │     └─ returns: {document, styles, footnotes, endnotes} as lxml roots
  │
  ├─ build_outline_level_map(styles_root)
  │     └─ returns: {styleId: outline_level}  e.g. {"Heading1": 0, "Heading2": 1}
  │
  ├─ build_image_map(docx_path)
  │     └─ returns: {rId: raw_bytes}  e.g. {"rId5": b'\x89PNG...'}
  │
  ├─ extract_tracked_changes(doc_root, outline_map, image_map, note_roots)
  │     └─ returns: [ChangeRecord, ChangeRecord, ...]
  │
  ├─ parse_headers_footers(docx_path)
  │     └─ returns: [("Header", root), ("Footer", root), ...]
  │
  ├─ extract_hf_changes(hf_roots)
  │     └─ returns: [ChangeRecord, ...]  (header/footer changes)
  │
  ├─ inject_bookmarks(docx_path, all_records, annotated_path)
  │     └─ saves annotated .docx, returns modified doc_root
  │
  └─ generate_excel_output(all_records, annotated_path, excel_out)
        └─ saves changes.xlsx
```

### Step-by-Step: What Data is Passed at Each Stage

| Step | Function | Input | Output |
|---|---|---|---|
| 1 | `parse_any_xml` | `.docx` path (string) | `dict[str, etree._Element]` — XML roots |
| 2 | `build_outline_level_map` | `styles_root` (XML element) | `dict[str, int]` — style→level |
| 3 | `build_image_map` | `.docx` path | `dict[str, bytes]` — rId→image bytes |
| 4 | `extract_tracked_changes` | doc_root, outline_map, image_map, note_roots | `list[ChangeRecord]` |
| 5 | `extract_hf_changes` | list of (label, XML root) | `list[ChangeRecord]` |
| 6 | `filter_records` | `list[ChangeRecord]` | `list[ChangeRecord]` (filtered) |
| 7 | `inject_bookmarks` | `.docx` path + records | annotated `.docx` file on disk |
| 8 | `generate_excel_output` | records + annotated path | `.xlsx` file on disk |

---

## 3. File-by-File Explanation

### 3.1 constants.py
- **Purpose:** Single source of truth for XML namespaces, change-type labels, delimiters, paths.
- **When used:** Imported by every other module.
- **Key contents:**
  - `NS` dict: maps prefix to URI (e.g. `w` to WordprocessingML URI)
  - `NSC` dict: Clark-notation prefixes for building element names
  - `ChangeType` class: registry of all change-type label strings
  - `CHANGE_TYPE_CATEGORIES`: groups change types for GUI display
  - Size/color/delimiter constants for output rendering
  - `make_bookmark_id(n)`: returns bookmark name like `docChg_00000001`

### 3.2 models.py
- **Purpose:** Defines the `ChangeRecord` dataclass - the single data object that flows through the entire pipeline.
- **When used:** Created during extraction, consumed during output.
- **Fields:**
  - `heading` (str): section heading where the change was found
  - `old_text` / `new_text` (str): before/after content
  - `change_type` (str): one of the ChangeType constants
  - `bookmark_id` (Optional[int]): assigned during extraction
  - `target_element`: lxml element for bookmark injection
  - `style_meta` (dict): color info for highlight/font changes
  - `old_image_data` / `new_image_data`: raw image bytes
  - `old_image_size_px` / `new_image_size_px`: pixel dimensions
  - `is_caption` (bool): True if from a caption paragraph

### 3.3 parser.py
- **Purpose:** Opens the .docx ZIP, parses XML parts into lxml trees.
- **When used:** First step of the pipeline.
- **Key functions:**
  - `parse_any_xml(docx_path, parts, save_xml)` -> dict of XML roots
  - `build_outline_level_map(styles_root)` -> {styleId: outline_level}
  - `build_image_map(docx_path)` -> {rId: raw_bytes}
  - `parse_headers_footers(docx_path)` -> [(label, xml_root)]
  - `save_xml_to_output(root, path, suffix)` -> saves XML for debugging

### 3.4 extractor.py
- **Purpose:** Orchestrates the extraction - walks the document body, dispatches paragraphs and tables.
- **When used:** Core of the extraction layer.
- **Key functions:**
  - `extract_tracked_changes(doc_root, outline_map, ...)` -> list[ChangeRecord]
    - Walks `<w:body>` children: `<w:p>` (paragraphs) and `<w:tbl>` (tables)
    - Updates current_heading when a heading paragraph is found
    - Assigns bookmark_id via `_Counter`
  - `extract_changes_from_paragraph(para, heading, ...)` -> list[ChangeRecord]
    - Walks paragraph children: `<w:r>`, `<w:del>`, `<w:ins>`
    - Dispatches to `handle_run`, `handle_del`, `handle_ins` from para_utils
    - Detects paragraph style changes via `get_ppr_style_change`
    - Stamps caption flag if paragraph is a caption
  - `extract_hf_changes(hf_roots)` -> list[ChangeRecord]
    - Extracts changes from header/footer XML roots
  - `_extract_embedded_container_changes(element, heading)`
    - Finds textboxes inside paragraphs/tables and extracts their changes

### 3.5 para_utils.py
- **Purpose:** Paragraph-level helpers for walking and dispatching tracked changes.
- **When used:** Called by extractor.py for each paragraph.
- **Key functions:**
  - `handle_del(children, i, heading, records, image_map, note_roots)` -> next index
    - Collects consecutive `<w:del>` elements, pairs with following `<w:ins>` if present
    - Emits DELETE_CONTENT, MODIFY_CONTENT, or object/note records
  - `handle_ins(children, i, heading, records, image_map, note_roots)` -> next index
    - Collects consecutive `<w:ins>` elements, emits ADD_CONTENT or object/note records
  - `handle_run(child, rpr_buffer, heading, records, note_roots)`
    - Accumulates runs with rPrChange for merging, flushes on non-matching runs
  - `collect_consecutive_del/ins` -> (CollectedContent, next_index)
    - Separates text, object runs, and note runs
  - `unwrap_sdts(elements)` -> flattened list without `<w:sdt>` wrappers
  - `get_ppr_style_change(para)` -> (old_style, new_style) or None
  - `is_caption_paragraph(para)` -> bool
  - `should_skip_paragraph(para)` -> bool (skips textbox paras inside tracked changes)
  - `get_paragraph_text(para, include_del, include_ins)` -> str

### 3.6 classifier.py
- **Purpose:** Single entry point for change-type classification.
- **When used:** Called whenever a ChangeRecord needs a change_type label.
- **Key function:**
  - `classify_change_type(old_text, new_text, rpr_detections, val_detections, is_format_only, structural_type)` -> str
  - Priority: structural_type > text content > rpr detections > val detections > format > unknown
- **Helper types:**
  - `RprDetection(old_state, new_state, added_type, removed_type)` - for boolean property changes
  - `ValDetection(old_val, new_val, changed_type)` - for value-based changes

### 3.7 embedded_content.py
- **Purpose:** Detects and labels non-text content (images, textboxes, charts, OLE objects).
- **When used:** Called by para_utils when a run contains drawings/pictures.
- **Key functions:**
  - `has_embedded_content(element)` -> bool: checks for w:drawing, w:pict, w:object
  - `classify_embedded_content(element)` -> str: returns label like `[Image: pic1]`
  - `extract_image_data(element, image_map)` -> bytes or None
  - `extract_image_size_px(element)` -> (width, height) or None
  - `has_note_reference(run)` -> bool: checks for footnote/endnote references
  - `get_note_ref_info(run)` -> (note_type, note_id) or None
  - `emit_object_records(del_runs, ins_runs, heading, records, image_map)`
    - Pairs del/ins objects positionally into MODIFY/DELETE/ADD_OBJECT
  - `emit_note_records(del_runs, ins_runs, note_roots, heading, records)`
    - Emits ADD_NOTE/DELETE_NOTE records
  - `emit_plain_note_records(note_runs, note_roots, heading, records)`
    - For note refs in plain runs, extracts changes from note content
  - `CollectedContent` dataclass: holds text + object_runs + note_runs

### 3.8 table_utils.py
- **Purpose:** Handles table-level tracked changes (row/cell/merge/content).
- **When used:** Called by extractor.py when a `<w:tbl>` is encountered.
- **Key functions:**
  - `extract_table_changes(tbl, heading, counter, ...)` -> list[ChangeRecord]
    - Checks whole-table ins/del first, then walks rows and cells
  - `check_whole_table_change(rows, heading, counter)` -> records or None
  - `process_row(tr, heading, counter, records)` -> bool (True = skip cells)
  - `process_cell(tc, heading, counter, include_sdt, records, image_map)`
    - Case 1: cell inserted/deleted -> ADD_CELL/DELETE_CELL
    - Case 2: tracked vertical merge -> MERGE_CELL
    - Case 3: content changes -> aggregated into one record

### 3.9 run_utils.py
- **Purpose:** Low-level text extraction and run-property state reading.
- **When used:** Called by para_utils and rpr_change_utils.
- **Key functions:**
  - `run_text(run)` -> str: collects `<w:t>` text, emits line-break delimiters
  - `del_text(del_run)` -> str: collects `<w:delText>` text
  - `has_rpr_property(rpr, *tags)` -> bool
  - `get_rpr_property_states(run, *tags)` -> (old_state, new_state)
  - `get_rpr_val_states(run, *specs)` -> (old_val, new_val)
  - `resolve_color(val)` -> hex string or None

### 3.10 rpr_change_utils.py
- **Purpose:** Detects, classifies, and merges w:rPrChange (run property change) elements.
- **When used:** Called by para_utils.handle_run and extractor flush logic.
- **Key functions:**
  - `get_rpr_change_meta(run)` -> (author, date) or None
  - `can_merge_rpr_runs(run_a, run_b)` -> bool (same author, date, type, values)
  - `flush_rpr_buffer(buffer, heading, records)` -> emits one ChangeRecord for all buffered runs
  - `_classify_run_change_type(run)` -> str (strikethrough/highlight/font color/format)

### 3.11 heading_resolver.py
- **Purpose:** Determines if a paragraph is a heading and extracts its text.
- **When used:** Called by extractor.py to update current_heading.
- **Key functions:**
  - `get_heading_text(para, outline_map)` -> str or None
  - `get_paragraph_outline_level(para, outline_map)` -> int (0-8) or None
  - `is_heading(outline_level)` -> bool

### 3.12 book_marker.py
- **Purpose:** Injects bookmark pairs into the annotated .docx for Excel hyperlinks.
- **When used:** After extraction, before Excel output.
- **Key function:**
  - `inject_bookmarks(docx_path, records, out_path)` -> modified doc_root
    - Computes element paths from original tree
    - Re-parses fresh XML, replays paths to find injection targets
    - Inserts `<w:bookmarkStart>` / `<w:bookmarkEnd>` pairs
    - Writes new .docx ZIP

### 3.13 output.py
- **Purpose:** Renders ChangeRecords into Excel (.xlsx) with hyperlinks and formatting.
- **When used:** Final step of the pipeline.
- **Key functions:**
  - `generate_excel_output(records, annotated_path, out_path, config)` -> .xlsx
  - `generate_excel_from_template(records, annotated_path, out_path, template_path, columns, config, sheet)` -> .xlsx
  - `_make_diff_rich_text(old, new, del_color, ins_color, eq_color)` -> word-level diff
  - `_embed_image(ws, image_bytes, cell, width, height)` -> embeds image in cell

### 3.14 app_config.py / gui.py / main.py
- **app_config.py:** User-configurable settings with JSON persistence. Provides `filter_records()` to remove unwanted change types.
- **gui.py:** PyQt6 GUI with settings panel, change-type tree, template management. Runs pipeline on background QThread.
- **main.py:** CLI entry point for headless execution. Same pipeline, no filtering.

## 4. Deep Dive into Important Files

### 4.1 extractor.py — The Orchestrator

**`extract_tracked_changes(doc_root, outline_map, ...)`**

Step-by-step logic:
1. Find `<w:body>` in the document root
2. Create a `_Counter` for unique bookmark IDs
3. For each direct child of body:
   - If `<w:p>` (paragraph):
     a. Check if it's a heading -> update `current_heading`
     b. Call `extract_changes_from_paragraph()` -> get records
     c. Assign bookmark_id to records that don't have a target_element
     d. Call `_extract_embedded_container_changes()` for textboxes
   - If `<w:tbl>` (table):
     a. Call `extract_table_changes()` from table_utils
     b. Also check for embedded textboxes

**`extract_changes_from_paragraph(para, heading, ...)`**

Step-by-step logic:
1. Check if paragraph is a caption (`is_caption_paragraph`)
2. Check for paragraph style change (`get_ppr_style_change`)
3. Unwrap SDT elements if enabled
4. Initialize empty `rpr_buffer` for format change merging
5. Walk children with index `i`:
   - `w:r` (run) -> `handle_run()` (may buffer rPrChange runs)
   - transparent tag -> skip (don't flush buffer)
   - `w:del` -> flush buffer, then `handle_del()`
   - `w:ins` -> flush buffer, then `handle_ins()`
   - other -> skip
6. Final flush of rpr_buffer
7. If caption paragraph: stamp all records with `is_caption=True`, override change_type

### 4.2 para_utils.py — The Dispatcher

**`handle_del()` — The Most Important Function**

This is where "Modify Content" detection happens:

```
Found <w:del>
  |
  +-> collect_consecutive_del() -> CollectedContent(text, objects, notes)
  |
  +-> Look ahead: is next non-transparent element <w:ins>?
       |
       YES -> collect_consecutive_ins() -> CollectedContent
       |      This is a MODIFICATION (del + ins = old + new)
       |      - Text: classify(old, new) -> MODIFY_CONTENT
       |      - Objects: emit_object_records() -> pairs into MODIFY_OBJECT
       |      - Notes: emit_note_records()
       |
       NO -> This is a pure DELETION
              - Text: classify(old, "") -> DELETE_CONTENT
              - Objects: each -> DELETE_OBJECT
              - Notes: emit_note_records()
```

**`handle_run()` — rPrChange Buffering**

```
Run has w:rPrChange?
  |
  YES -> Can merge with buffer[0]? (same author, date, type, values)
  |      |
  |      YES -> append to buffer
  |      NO  -> flush buffer, start new buffer with this run
  |
  NO -> Is it whitespace-only and buffer is active?
        |
        YES -> keep buffer alive (don't break span)
        NO  -> flush buffer
```

### 4.3 classifier.py — The Decision Tree

```
classify_change_type() priority:

1. structural_type provided?  -> return it directly
   (ADD_TABLE, DELETE_ROW, ADD_OBJECT, etc.)

2. old_text AND new_text AND different?  -> MODIFY_CONTENT
3. old_text only?                        -> DELETE_CONTENT
4. new_text only?                        -> ADD_CONTENT

5. rpr_detections (boolean property changes):
   - strikethrough added/removed?
   - highlight added/removed?

6. val_detections (value changes):
   - highlight color changed?
   - font color changed?

7. is_format_only? -> FORMAT_CHANGE

8. fallback -> UNKNOWN
```

### 4.4 embedded_content.py — Object Detection

**Detection priority for `classify_embedded_content()`:**

```
w:drawing (modern DrawingML):
  pic:pic     -> [Image: <name>]
  wps:txbx    -> [Textbox: <text>]  (extracts actual text!)
  c:chart     -> [Chart: <name>]
  wpg:wgp     -> [Shape: <name>]
  fallback    -> [Drawing: <name>]

w:pict (legacy VML):
  v:imagedata -> [Image: <name>]
  v:textbox   -> [Textbox: <text>]
  v:shape     -> [Shape: <name>]

w:object (OLE):
  o:OLEObject -> [OLE: <ProgID>]
  fallback    -> [Object]
```

**Image data extraction:**
- DrawingML: find `a:blip` -> get `r:embed` attribute -> look up in image_map
- VML: find `v:imagedata` -> get `r:id` attribute -> look up in image_map

### 4.5 table_utils.py — Table Processing

**`extract_table_changes()` control flow:**

```
1. Filter rows (unwrap SDTs, skip non-tr elements)
2. Identify "real rows" (exclude vMerge continuation rows)
3. check_whole_table_change():
   - ALL rows inserted? -> one ADD_TABLE record, return
   - ALL rows deleted?  -> one DELETE_TABLE record, return
4. For each row:
   a. process_row():
      - Row has w:trPr/w:ins? -> ADD_ROW record, skip cells
      - Row has w:trPr/w:del? -> DELETE_ROW record, skip cells
      - vMerge continuation row? -> skip silently
   b. If not handled, for each cell:
      process_cell():
      - w:cellIns? -> ADD_CELL
      - w:cellDel? -> DELETE_CELL
      - w:cellMerge restart? -> MERGE_CELL
      - w:cellMerge continue? -> skip
      - Otherwise: extract paragraph changes from cell content
        - Content changes (add/del/modify) -> aggregated into ONE record
        - Non-content changes (format, highlight) -> emitted individually
```

## 5. Data Flow (DOCX to Excel)

### The Complete Journey of Data

```
DOCX ZIP file (compare.docx)
  |
  v
[parser.py] unzip -> word/document.xml, word/styles.xml, etc.
  |
  v
lxml.etree._Element trees (in-memory XML)
  |
  v
[extractor.py] walks <w:body> children
  |
  +-- <w:p> paragraphs --> extract_changes_from_paragraph()
  |     |
  |     +-- <w:del> + <w:ins> --> ChangeRecord(old_text, new_text, MODIFY_CONTENT)
  |     +-- <w:del> alone    --> ChangeRecord(old_text, "",       DELETE_CONTENT)
  |     +-- <w:ins> alone    --> ChangeRecord("",       new_text, ADD_CONTENT)
  |     +-- <w:r> with rPrChange --> ChangeRecord(text, text, HIGHLIGHT_CHANGED)
  |
  +-- <w:tbl> tables --> extract_table_changes()
        |
        +-- whole table ins/del --> ChangeRecord(ADD_TABLE / DELETE_TABLE)
        +-- row ins/del         --> ChangeRecord(ADD_ROW / DELETE_ROW)
        +-- cell ins/del        --> ChangeRecord(ADD_CELL / DELETE_CELL)
        +-- cell content        --> ChangeRecord(MODIFY_CONTENT / etc.)
  |
  v
list[ChangeRecord]  (raw, unfiltered)
  |
  v
[app_config.py] filter_records() -- removes unwanted change types
  |
  v
list[ChangeRecord]  (filtered)
  |
  v
[book_marker.py] inject_bookmarks()
  |  - reads target_element from each record
  |  - computes child-index path in original tree
  |  - re-parses fresh XML copy
  |  - replays path to find same element in fresh tree
  |  - inserts <w:bookmarkStart>/<w:bookmarkEnd> pairs
  |  - writes annotated .docx to output/
  |
  v
annotated .docx (with bookmarks)
  |
  v
[output.py] generate_excel_output()
  |  - for each ChangeRecord:
  |    - heading cell: =HYPERLINK("file.docx#docChg_00000001", "heading")
  |    - change_type cell: plain text
  |    - old_text cell: formatted (diff colors for MODIFY, strikethrough, etc.)
  |    - new_text cell: formatted
  |    - if image data: embed image in cell
  |
  v
changes.xlsx (final output)
```

### Where ChangeRecord is Created

| Location | What creates it |
|---|---|
| `para_utils.handle_del()` | Text deletions and modifications |
| `para_utils.handle_ins()` | Text insertions |
| `rpr_change_utils.flush_rpr_buffer()` | Format/highlight/font changes |
| `embedded_content.emit_object_records()` | Image/textbox/chart changes |
| `embedded_content.emit_note_records()` | Footnote/endnote add/delete |
| `embedded_content.emit_plain_note_records()` | Note content changes |
| `table_utils.process_row()` | Row insertions/deletions |
| `table_utils.process_cell()` | Cell ins/del/merge/content |
| `table_utils.check_whole_table_change()` | Whole table ins/del |
| `extractor.extract_changes_from_paragraph()` | Paragraph style changes |

### Where ChangeRecord is Modified

| Location | What changes |
|---|---|
| `extractor.extract_tracked_changes()` | Sets `bookmark_id` and `target_element` |
| `extractor._extract_embedded_container_changes()` | Sets `target_element` for textbox records |
| `extractor.extract_changes_from_paragraph()` | Stamps `is_caption` and overrides `change_type` |
| `app_config.filter_records()` | Removes records (does not modify them) |

### Where ChangeRecord is Consumed

| Location | What reads it |
|---|---|
| `book_marker.inject_bookmarks()` | Reads `bookmark_id` and `target_element` |
| `output.generate_excel_output()` | Reads all fields to write Excel rows |

## 6. Relationships Between Modules

### Dependency Graph (who imports whom)

```
constants.py  <-- imported by ALL modules
     |
models.py     <-- imported by extractor, para_utils, table_utils,
     |             embedded_content, rpr_change_utils, output, book_marker
     |
classifier.py <-- imported by para_utils, rpr_change_utils,
     |             embedded_content, table_utils
     |
run_utils.py  <-- imported by para_utils, rpr_change_utils, output
     |
rpr_change_utils.py <-- imported by para_utils, extractor
     |
embedded_content.py <-- imported by para_utils
     |
heading_resolver.py <-- imported by extractor
     |
para_utils.py <-- imported by extractor, table_utils, heading_resolver,
     |             embedded_content (lazy)
     |
table_utils.py <-- imported by extractor
     |
extractor.py  <-- imported by main, gui, table_utils (lazy),
     |             embedded_content (lazy)
     |
parser.py     <-- imported by main, gui
     |
book_marker.py <-- imported by main, gui
     |
output.py     <-- imported by main, gui
     |
app_config.py <-- imported by gui
     |
gui.py / main.py  <-- entry points (import everything above)
```

### Circular Dependency Management

Two circular dependencies exist and are handled with **lazy imports**:
1. `table_utils.process_cell()` imports `extractor.extract_changes_from_paragraph` inside the function
2. `embedded_content.emit_plain_note_records()` imports `extractor.extract_changes_from_paragraph` inside the function

This is necessary because extractor imports table_utils and embedded_content, but they need to call back into extractor for paragraph-level extraction.

### Why This Structure?

| Design choice | Reason |
|---|---|
| `constants.py` separate | Single source of truth; change a namespace URI in one place |
| `models.py` separate | Data model independent of logic; easy to extend fields |
| `classifier.py` separate | Classification logic reusable; easy to add new change types |
| `run_utils.py` separate | Low-level XML property reading; no side effects; testable |
| `rpr_change_utils.py` separate | Complex merging logic isolated from paragraph walking |
| `embedded_content.py` separate | Object detection is complex; keeps para_utils focused on text |
| `table_utils.py` separate | Table processing has its own hierarchy (table > row > cell) |
| `heading_resolver.py` separate | Heading detection is independent of change extraction |
| `para_utils.py` separate | Paragraph dispatching is the most complex helper; deserves its own file |
| `book_marker.py` separate | Bookmark injection is a post-processing step, not extraction |
| `output.py` separate | Excel rendering is independent; could be swapped for CSV/HTML |

## 7. Key Concepts You Must Understand

### 7.1 OOXML Structure (What's Inside a .docx?)

A `.docx` file is just a **ZIP archive** containing XML files:

```
compare.docx (ZIP)
  |
  +-- word/
  |     +-- document.xml      <-- the main document body
  |     +-- styles.xml        <-- style definitions (Heading1, Normal, etc.)
  |     +-- footnotes.xml     <-- footnote content
  |     +-- endnotes.xml      <-- endnote content
  |     +-- header1.xml       <-- header content
  |     +-- footer1.xml       <-- footer content
  |     +-- media/            <-- embedded images (image1.png, etc.)
  |     +-- _rels/
  |           +-- document.xml.rels  <-- relationship IDs (rId -> file path)
  |
  +-- [Content_Types].xml
  +-- _rels/.rels
```

### 7.2 Key XML Elements

```xml
<w:body>                          <!-- document body -->
  <w:p>                           <!-- paragraph -->
    <w:pPr>                       <!-- paragraph properties -->
      <w:pStyle w:val="Heading1"/>  <!-- paragraph style -->
      <w:pPrChange>               <!-- paragraph property change (tracked) -->
    </w:pPr>
    <w:r>                         <!-- run (a chunk of text with same formatting) -->
      <w:rPr>                     <!-- run properties (bold, color, etc.) -->
        <w:rPrChange>             <!-- run property change (tracked) -->
      </w:rPr>
      <w:t>Hello world</w:t>     <!-- actual text content -->
    </w:r>
    <w:ins>                       <!-- tracked insertion -->
      <w:r><w:t>new text</w:t></w:r>
    </w:ins>
    <w:del>                       <!-- tracked deletion -->
      <w:r><w:delText>old text</w:delText></w:r>
    </w:del>
  </w:p>
  <w:tbl>                        <!-- table -->
    <w:tr>                        <!-- table row -->
      <w:tc>                      <!-- table cell -->
        <w:p>...</w:p>            <!-- cell contains paragraphs -->
      </w:tc>
    </w:tr>
  </w:tbl>
</w:body>
```

### 7.3 How Word Stores Tracked Changes

Word uses three mechanisms:

**1. Content changes: `<w:ins>` and `<w:del>`**
- Inserted text is wrapped in `<w:ins><w:r><w:t>new</w:t></w:r></w:ins>`
- Deleted text is wrapped in `<w:del><w:r><w:delText>old</w:delText></w:r></w:del>`
- Note: deleted text uses `<w:delText>`, not `<w:t>`!

**2. Format changes: `<w:rPrChange>` inside `<w:rPr>`**
- The current `<w:rPr>` shows the NEW formatting
- `<w:rPrChange>` contains a nested `<w:rPr>` showing the OLD formatting
- Example: text was highlighted yellow, now unhighlighted:
  ```xml
  <w:r>
    <w:rPr>
      <!-- no highlight = current state -->
      <w:rPrChange w:author="John" w:date="2024-01-01T00:00:00Z">
        <w:rPr>
          <w:highlight w:val="yellow"/>  <!-- old state: was highlighted -->
        </w:rPr>
      </w:rPrChange>
    </w:rPr>
    <w:t>some text</w:t>
  </w:r>
  ```

**3. Structural changes: attributes on table/row/cell elements**
- `<w:trPr><w:ins/></w:trPr>` = row was inserted
- `<w:trPr><w:del/></w:trPr>` = row was deleted
- `<w:tcPr><w:cellIns/></w:tcPr>` = cell was inserted
- `<w:tcPr><w:cellDel/></w:tcPr>` = cell was deleted
- `<w:tcPr><w:cellMerge w:vMerge="restart"/></w:tcPr>` = cell merge tracked

### 7.4 How "Modify Content" (del + ins) is Detected

When Word modifies text, it stores it as a deletion immediately followed by an insertion:

```xml
<w:p>
  <w:del>
    <w:r><w:delText>old version</w:delText></w:r>
  </w:del>
  <w:ins>
    <w:r><w:t>new version</w:t></w:r>
  </w:ins>
</w:p>
```

CPAuto's `handle_del()` detects this pattern:
1. Collect all consecutive `<w:del>` elements -> old_text
2. Look ahead: is the next element `<w:ins>`?
3. If YES: collect `<w:ins>` -> new_text, classify as MODIFY_CONTENT
4. If NO: classify as DELETE_CONTENT

### 7.5 How Tables are Processed

Tables have a hierarchy: `<w:tbl>` > `<w:tr>` (rows) > `<w:tc>` (cells) > `<w:p>` (paragraphs)

Processing order:
1. Check if ALL rows are inserted/deleted -> whole table change
2. For each row: check if row itself is inserted/deleted
3. For each cell: check cell-level changes (ins/del/merge)
4. For remaining cells: extract paragraph-level changes from cell content
5. Content changes within a cell are aggregated into ONE record

### 7.6 How Objects (Images/Textboxes) are Handled

Objects are detected by `has_embedded_content()` which checks for:
- `<w:drawing>` (modern DrawingML - images, textboxes, charts)
- `<w:pict>` (legacy VML - older format images/shapes)
- `<w:object>` (OLE embedded objects)

When found inside `<w:ins>` or `<w:del>`:
- Image data is extracted from the ZIP via relationship IDs
- Objects are paired positionally (del[0] with ins[0]) for MODIFY detection
- Textbox text content is extracted and included in the record

### 7.7 How Footnotes/Endnotes are Handled

Footnotes and endnotes have two parts:
1. A **reference** in the document body: `<w:footnoteReference w:id="1"/>`
2. The **content** in footnotes.xml/endnotes.xml: `<w:footnote w:id="1">...</w:footnote>`

CPAuto handles three cases:
- Reference inside `<w:ins>` -> ADD_NOTE (content from footnotes.xml)
- Reference inside `<w:del>` -> DELETE_NOTE (content from footnotes.xml)
- Reference in plain run -> extract changes FROM the note content itself

## 8. Visual Flow Diagrams

### Main Pipeline Flow
```
main.py / gui.py
    |
    v
parser.py -----> extractor.py -----> app_config.py
  |                  |                    |
  |  parse XML       |  extract changes  |  filter
  |                  |                    |
  v                  v                    v
lxml trees      ChangeRecord[]      ChangeRecord[] (filtered)
                     |                    |
                     v                    v
              book_marker.py -------> output.py
                  |                      |
                  v                      v
            annotated.docx          changes.xlsx
```

### Extraction Call Graph
```
extract_tracked_changes()
  |
  +-- get_heading_text()           [heading_resolver.py]
  |
  +-- extract_changes_from_paragraph()  [extractor.py]
  |     |
  |     +-- get_ppr_style_change()      [para_utils.py]
  |     +-- handle_run()                [para_utils.py]
  |     |     +-- get_rpr_change_meta() [rpr_change_utils.py]
  |     |     +-- can_merge_rpr_runs()  [rpr_change_utils.py]
  |     |     +-- flush_rpr_buffer()    [rpr_change_utils.py]
  |     |
  |     +-- handle_del()                [para_utils.py]
  |     |     +-- collect_consecutive_del()
  |     |     +-- collect_consecutive_ins()
  |     |     +-- classify_change_type()  [classifier.py]
  |     |     +-- emit_object_records()   [embedded_content.py]
  |     |     +-- emit_note_records()     [embedded_content.py]
  |     |
  |     +-- handle_ins()                [para_utils.py]
  |
  +-- extract_table_changes()           [table_utils.py]
  |     +-- check_whole_table_change()
  |     +-- process_row()
  |     +-- process_cell()
  |
  +-- _extract_embedded_container_changes()  [extractor.py]
```

## 9. How to Read This Code Effectively

### Recommended Study Order
1. **constants.py** - understand namespaces and change types first
2. **models.py** - understand the ChangeRecord data structure
3. **classifier.py** - understand how changes are labelled
4. **run_utils.py** - understand low-level text extraction
5. **parser.py** - understand how XML is loaded
6. **heading_resolver.py** - simple, builds confidence
7. **extractor.py** - the orchestrator (read top-level functions first)
8. **para_utils.py** - the core dispatcher (handle_del is key)
9. **rpr_change_utils.py** - format change detection
10. **embedded_content.py** - object/note handling
11. **table_utils.py** - table hierarchy processing
12. **book_marker.py** - bookmark injection
13. **output.py** - Excel rendering
14. **app_config.py** - user settings
15. **gui.py** - UI (read last, it just calls the pipeline)

### How to Trace One Change from Start to End

Example: tracing a text modification "hello" -> "world"

1. **parser.py**: `parse_any_xml()` opens ZIP, parses `document.xml`
2. **extractor.py**: `extract_tracked_changes()` finds `<w:p>` containing the change
3. **heading_resolver.py**: `get_heading_text()` determines current heading
4. **extractor.py**: `extract_changes_from_paragraph()` walks paragraph children
5. **para_utils.py**: finds `<w:del>`, calls `handle_del()`
6. **para_utils.py**: `collect_consecutive_del()` extracts "hello"
7. **para_utils.py**: looks ahead, finds `<w:ins>`, calls `collect_consecutive_ins()`
8. **para_utils.py**: extracts "world"
9. **classifier.py**: `classify_change_type("hello", "world")` -> MODIFY_CONTENT
10. **models.py**: `ChangeRecord(heading="Section 1", old="hello", new="world", type=MODIFY_CONTENT)`
11. **extractor.py**: assigns `bookmark_id=0`, `target_element=<w:p>`
12. **app_config.py**: `filter_records()` keeps it (MODIFY_CONTENT is selected)
13. **book_marker.py**: injects `<w:bookmarkStart name="docChg_00000000"/>` into the paragraph
14. **output.py**: writes Excel row with HYPERLINK formula, red/green diff coloring

### How to Debug

Add print statements at these key points:
- `extractor.py:extract_changes_from_paragraph()` - print paragraph text
- `para_utils.py:handle_del()` - print del_content and ins_content
- `classifier.py:classify_change_type()` - print inputs and result
- `rpr_change_utils.py:flush_rpr_buffer()` - print buffer contents

## 10. Beginner Cheat-Sheet

### Quick Reference: XML Tag -> What CPAuto Does

| XML Tag | CPAuto Action |
|---|---|
| `<w:p>` | Walk children for changes |
| `<w:r>` | Extract text or check for rPrChange |
| `<w:t>` | Plain text content |
| `<w:delText>` | Deleted text content |
| `<w:ins>` | Tracked insertion -> ADD or MODIFY |
| `<w:del>` | Tracked deletion -> DELETE or MODIFY |
| `<w:rPrChange>` | Format change -> HIGHLIGHT/FONT/STRIKE/FORMAT |
| `<w:pPrChange>` | Style change -> STYLE_CHANGE |
| `<w:tbl>` | Table -> delegate to table_utils |
| `<w:tr>` | Table row -> check row ins/del |
| `<w:tc>` | Table cell -> check cell ins/del/merge |
| `<w:drawing>` | Embedded object -> classify and extract |
| `<w:pict>` | Legacy VML object |
| `<w:sdt>` | Content control -> unwrap |
| `<w:footnoteReference>` | Note reference -> look up content |

### Quick Reference: ChangeType -> What Triggers It

| ChangeType | Trigger |
|---|---|
| MODIFY_CONTENT | `<w:del>` followed by `<w:ins>` |
| DELETE_CONTENT | `<w:del>` alone |
| ADD_CONTENT | `<w:ins>` alone |
| STRIKETHROUGH_ADDED | rPrChange: strike absent -> present |
| HIGHLIGHT_ADDED | rPrChange: highlight absent -> present |
| HIGHLIGHT_CHANGED | rPrChange: highlight value changed |
| FONT_COLOR_CHANGED | rPrChange: color value changed |
| FORMAT_CHANGE | rPrChange: any other property |
| STYLE_CHANGE | pPrChange: pStyle value changed |
| ADD/DELETE_OBJECT | `<w:drawing>` inside ins/del |
| ADD/DELETE_NOTE | footnoteReference inside ins/del |
| ADD/DELETE_TABLE | all rows inserted/deleted |
| ADD/DELETE_ROW | trPr contains ins/del |
| ADD/DELETE_CELL | tcPr contains cellIns/cellDel |
| MERGE_CELL | tcPr contains cellMerge |

### The Golden Rule

> Every tracked change in Word XML becomes exactly one `ChangeRecord`.
> Every `ChangeRecord` becomes exactly one row in Excel.
> The `heading` field tells you WHERE. The `change_type` tells you WHAT.
> The `old_text`/`new_text` tell you the BEFORE/AFTER.

---
*End of CPAuto Deep-Dive Guide*
