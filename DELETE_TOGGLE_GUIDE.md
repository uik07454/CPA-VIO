# Delete Toggle Feature - Complete Guide

## 📋 Table of Contents

1. [Quick Start (Recommended)](#quick-start-recommended)
2. [How to Use the Toggle Column](#how-to-use-the-toggle-column)
3. [Three Methods Explained](#three-methods-explained)
4. [Complete Workflows](#complete-workflows)
5. [Troubleshooting](#troubleshooting)
6. [API Reference](#api-reference)

---

## Quick Start (Recommended)

### ✨ Method 1: Auto-Delete Watcher (EASIEST!)

**Best for: Most users - simplest workflow**

```bash
# 1. Generate Excel with toggle column (CPAuto GUI → Enable Delete Toggle)
# 2. Start the watcher
python auto_delete_watcher.py output/changes.xlsx

# 3. Open Excel, mark rows with TRUE, save (Ctrl+S)
# 4. Rows auto-delete instantly! ✨
# 5. Keep working... repeat steps 3-4 as needed
```

**Advantages:**
- ✅ Just mark and save - automatic deletion
- ✅ No need to close Excel
- ✅ No button clicking
- ✅ Batch delete on each save
- ✅ Auto backup created

---

## How to Use the Toggle Column

### Understanding the Checkbox Column

The toggle column shows **☐** (unchecked box) by default. **This looks like a real checkbox!**

**To toggle the checkbox:**
1. **Click** on any cell in the "Delete Toggle" column (yellow header)
2. A **dropdown arrow** appears on the right side of the cell
3. **Click the arrow** and select:
   - **☐** = Keep row (unchecked)
   - **☑** = Delete row (checked)

### Marking Rows for Deletion

**Method 1: Using Dropdown (Recommended)**
- Click cell → Click dropdown arrow → Select ☑ (checked box)

**Method 2: Copy-Paste**
- Set one cell to ☑ → Copy (Ctrl+C) → Select other cells → Paste (Ctrl+V)

**Method 3: Direct Typing** (if needed)
- Click cell → Type the checkbox symbol → Press Enter

### Visual Indicators

- **Yellow header** = Delete Toggle column
- **☐** (unchecked) = Row will be kept  
- **☑** (checked) = Row will be deleted
- **Larger font (14pt)** = Checkboxes are easy to see

---

## Three Methods Explained

### Method 1: Auto-Delete Watcher ⭐ RECOMMENDED

**What it does:** Watches your Excel file and auto-deletes rows when you save

**Workflow:**
```bash
python auto_delete_watcher.py output/changes.xlsx
```

Then in Excel:
1. Mark rows with TRUE
2. Save (Ctrl+S)
3. Rows auto-delete!
4. Continue working...

**Pros:**
- ✅ Easiest workflow (2 steps: mark + save)
- ✅ No Excel closing needed
- ✅ Batch delete
- ✅ Auto backup
- ✅ No macros

**Cons:**
- ⚠️ Need to keep terminal/command prompt open

---

### Method 2: Manual Delete Button

**What it does:** Use CPAuto GUI button to delete marked rows

**Workflow:**
1. Mark rows with TRUE in Excel
2. Save and close Excel
3. CPAuto GUI → "Delete Toggled Rows from Excel..."
4. Select file
5. Confirm deletion

**Pros:**
- ✅ Built into CPAuto GUI
- ✅ Batch delete
- ✅ Auto backup
- ✅ No macros

**Cons:**
- ❌ Must close Excel
- ❌ Must reopen file
- ❌ More steps

---

### Method 3: VBA Macro (Advanced)

**What it does:** Instant deletion when you select TRUE (requires Excel macros)

**Setup:**
```bash
python delete_toggle_feature_macro.py output/changes.xlsx
# Follow instructions to add VBA code to Excel
```

**Pros:**
- ✅ Instant deletion
- ✅ No external tools

**Cons:**
- ❌ Requires VBA macros (security risk)
- ❌ Complex setup
- ❌ One-by-one deletion
- ❌ No auto backup

---

## Complete Workflows

### Workflow A: Auto-Watcher (Recommended)

```
┌─────────────────────────────────────────────┐
│ 1. CPAuto GUI                               │
│    ☑ Enable Delete Toggle Column           │
│    ▶ Run → generates output/changes.xlsx   │
└─────────────────────────────────────────────┘
                    ↓
┌─────────────────────────────────────────────┐
│ 2. Terminal/Command Prompt                  │
│    python auto_delete_watcher.py \          │
│           output/changes.xlsx               │
│    [Watcher running - keep this open]       │
└─────────────────────────────────────────────┘
                    ↓
┌─────────────────────────────────────────────┐
│ 3. Excel                                    │
│    Open changes.xlsx                        │
│    Mark rows: TRUE = delete, FALSE = keep   │
│    Save (Ctrl+S)                            │
└─────────────────────────────────────────────┘
                    ↓
┌─────────────────────────────────────────────┐
│ 4. Auto-Delete Happens! ✨                  │
│    Terminal shows:                          │
│    ✓ Deleted 5 row(s)                       │
│    Backup created: .auto_backup.xlsx        │
└─────────────────────────────────────────────┘
                    ↓
┌─────────────────────────────────────────────┐
│ 5. Continue Working                         │
│    Excel still open                         │
│    Mark more rows → Save → Auto-delete!     │
│    Repeat as needed                         │
└─────────────────────────────────────────────┘
```

### Workflow B: Manual Button

```
1. CPAuto GUI → Enable Delete Toggle → Run
2. Open Excel → Mark rows TRUE → Save → Close Excel
3. CPAuto GUI → "Delete Toggled Rows from Excel..."
4. Select file → Confirm
5. Reopen Excel to continue
```

---

## Troubleshooting

### Toggle Column Issues

**Q: I don't see a dropdown arrow**
- **A:** Click the cell first - the arrow appears when selected

**Q: Cells just show "FALSE" text**
- **A:** That's correct! Click the cell to access the dropdown

**Q: Can I just type TRUE/FALSE?**
- **A:** Yes! You don't need to use the dropdown

---

### Auto-Watcher Issues

**Q: Watcher says "File might be locked"**
- **A:** Excel is still writing. Wait a moment and save again

**Q: Nothing happens when I save**
- **A:** 
  - Make sure at least one row is marked TRUE
  - Check watcher is still running (not stopped)
  - Verify you saved the correct file

**Q: Can I close the terminal?**
- **A:** No, keep it open (minimize it instead)

**Q: Where are backups?**
- **A:** Same folder as Excel file: `filename.auto_backup.xlsx`

---

### General Issues

**Q: Toggle column not in Excel output**
- **A:** 
  - Check "Enable Delete Toggle Column" in CPAuto GUI
  - Click "Save Config"
  - Re-run the pipeline

**Q: Can I undo deletion?**
- **A:** Yes! Open the `.backup.xlsx` or `.auto_backup.xlsx` file

**Q: Deletion removes wrong rows**
- **A:** 
  - Verify TRUE/FALSE values before saving
  - Check the backup file

---

## API Reference

### Command-Line Tools

#### Auto-Delete Watcher
```bash
python auto_delete_watcher.py <excel_file> [toggle_column]

# Examples:
python auto_delete_watcher.py output/changes.xlsx
python auto_delete_watcher.py output/changes.xlsx F
```

#### Manual Operations
```bash
# Add toggle column to existing file
python delete_toggle_feature.py add <excel_file> [toggle_col]

# Show summary
python delete_toggle_feature.py summary <excel_file> [toggle_col]

# Delete toggled rows
python delete_toggle_feature.py delete <excel_file> [output_file] [toggle_col]
```

#### VBA Macro Setup
```bash
python delete_toggle_feature_macro.py <excel_file> [toggle_col]
```

---

### Python API

```python
from delete_toggle_feature import (
    add_toggle_column_to_workbook,
    read_toggle_status,
    delete_toggled_rows,
    create_delete_toggle_summary,
)
import openpyxl

# Add toggle column
wb = openpyxl.load_workbook("file.xlsx")
ws = wb.active
add_toggle_column_to_workbook(ws, toggle_col="E")
wb.save("file.xlsx")
wb.close()

# Read status
status = read_toggle_status("file.xlsx", toggle_col="E")
print(f"Rows to delete: {sum(1 for v in status.values() if v)}")

# Create summary
summary = create_delete_toggle_summary("file.xlsx", toggle_col="E")
print(f"Total: {summary['total']}, Delete: {summary['marked_for_deletion']}")

# Delete rows
deleted_count, output_path = delete_toggled_rows(
    "file.xlsx",
    output_path="filtered.xlsx",
    toggle_col="E",
    remove_toggle_column=True
)
print(f"Deleted {deleted_count} rows")
```

---

## Configuration

### GUI Settings

In CPAuto GUI → Delete Toggle Feature section:

| Setting | Description | Default |
|---------|-------------|---------|
| Enable Delete Toggle Column | Add toggle column to output | Disabled |
| Toggle Column | Excel column letter | E |

Settings are saved in `data/config.json`

---

## Comparison Table

| Feature | Auto-Watcher | Manual Button | VBA Macro |
|---------|--------------|---------------|-----------|
| **Ease of Use** | ⭐⭐⭐⭐⭐ | ⭐⭐⭐ | ⭐⭐ |
| **Setup** | One command | Built-in | Complex |
| **Excel Closing** | Not needed | Required | Not needed |
| **Batch Delete** | ✅ Yes | ✅ Yes | ❌ One-by-one |
| **Auto Backup** | ✅ Yes | ✅ Yes | ❌ No |
| **Security** | ✅ Safe | ✅ Safe | ⚠️ Macros |
| **Speed** | ✅ Fast | ⚠️ Slow | ⚠️ Slow |

**Recommendation: Use Auto-Watcher for best experience!**

---

## Testing

Run comprehensive tests:

```bash
python -m pytest tests/test_delete_toggle.py -v
# Expected: 16 tests pass
```

---

## Support

For issues or questions:
1. Check this guide first
2. Review test cases in `tests/test_delete_toggle.py`
3. Examine source code in `delete_toggle_feature.py`
4. Report bugs using `/reportbug` in VIO Coder

---

**Last Updated:** June 23, 2026  
**Version:** 2.0 (Consolidated)
