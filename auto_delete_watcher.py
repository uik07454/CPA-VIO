"""
Auto-Delete File Watcher

This script watches an Excel file and automatically deletes rows marked with TRUE
whenever you save the file. Much more user-friendly than the manual workflow!

Usage:
    python auto_delete_watcher.py <excel_file>

Workflow:
    1. Run this script (it stays running in background)
    2. Open Excel file normally
    3. Mark rows with TRUE
    4. Save Excel (Ctrl+S)
    5. Rows are automatically deleted!
    6. Continue working...
"""

import time
import os
from pathlib import Path
from datetime import datetime
import shutil


def watch_and_auto_delete(excel_path: str, toggle_col: str = "E", check_interval: float = 1.0):
    """
    Watch an Excel file and auto-delete rows when file is saved.
    
    Args:
        excel_path: Path to Excel file to watch
        toggle_col: Toggle column letter
        check_interval: How often to check file (seconds)
    """
    from delete_toggle_feature import delete_toggled_rows, create_delete_toggle_summary
    
    excel_path = Path(excel_path).resolve()
    
    if not excel_path.exists():
        print(f"[ERROR] File not found: {excel_path}")
        return
    
    print("=" * 70)
    print("AUTO-DELETE WATCHER - Running")
    print("=" * 70)
    print(f"Watching: {excel_path}")
    print(f"Toggle Column: {toggle_col}")
    print(f"\nInstructions:")
    print("  1. Open the Excel file in Excel")
    print("  2. Mark rows with TRUE in the toggle column")
    print("  3. Save the file (Ctrl+S)")
    print("  4. Rows will be AUTOMATICALLY deleted!")
    print(f"\nPress Ctrl+C to stop watching")
    print("=" * 70)
    print()
    
    last_modified = os.path.getmtime(excel_path)
    last_check_time = time.time()
    
    try:
        while True:
            time.sleep(check_interval)
            
            # Check if file still exists
            if not excel_path.exists():
                print(f"[WARNING] File no longer exists: {excel_path}")
                print("[INFO] Waiting for file to reappear...")
                time.sleep(2)
                continue
            
            # Check if file was modified
            current_modified = os.path.getmtime(excel_path)
            
            if current_modified > last_modified:
                # File was saved!
                current_time = datetime.now().strftime("%H:%M:%S")
                print(f"[{current_time}] File saved detected! Checking for rows to delete...")
                
                # Wait a bit to ensure file is fully written
                time.sleep(0.5)
                
                try:
                    # Check summary
                    summary = create_delete_toggle_summary(str(excel_path), toggle_col=toggle_col)
                    
                    if summary["marked_for_deletion"] > 0:
                        print(f"[INFO] Found {summary['marked_for_deletion']} row(s) marked for deletion")
                        
                        # Create backup
                        backup_path = excel_path.with_suffix('.auto_backup.xlsx')
                        shutil.copy2(excel_path, backup_path)
                        print(f"[INFO] Backup created: {backup_path.name}")
                        
                        # Delete rows
                        deleted_count, _ = delete_toggled_rows(
                            str(excel_path),
                            output_path=str(excel_path),
                            toggle_col=toggle_col,
                            remove_toggle_column=False  # Keep toggle column for future use
                        )
                        
                        print(f"[SUCCESS] ✓ Deleted {deleted_count} row(s)")
                        print(f"[INFO] Remaining rows: {summary['kept']}")
                        print(f"[INFO] You can continue editing the file...")
                        print()
                    else:
                        print(f"[INFO] No rows marked for deletion (all FALSE)")
                        print()
                
                except Exception as e:
                    print(f"[ERROR] Failed to process file: {e}")
                    print(f"[INFO] The file might be locked by Excel. Try again after saving.")
                    print()
                
                last_modified = current_modified
            
            # Show "still watching" indicator every 30 seconds
            if time.time() - last_check_time > 30:
                current_time = datetime.now().strftime("%H:%M:%S")
                print(f"[{current_time}] Still watching... (file is open and ready)")
                last_check_time = time.time()
    
    except KeyboardInterrupt:
        print("\n" + "=" * 70)
        print("[INFO] Watcher stopped by user")
        print("=" * 70)


def main():
    import sys
    
    if len(sys.argv) < 2:
        print("=" * 70)
        print("AUTO-DELETE WATCHER")
        print("=" * 70)
        print("\nUsage:")
        print("  python auto_delete_watcher.py <excel_file> [toggle_column]")
        print("\nExample:")
        print("  python auto_delete_watcher.py output/changes.xlsx E")
        print("\nWhat this does:")
        print("  1. Watches your Excel file")
        print("  2. When you save (Ctrl+S), automatically deletes rows marked TRUE")
        print("  3. Creates backup before each deletion")
        print("  4. You can keep editing and saving - it keeps watching!")
        print("\nThis is MUCH easier than the manual workflow!")
        print("=" * 70)
        sys.exit(1)
    
    excel_file = sys.argv[1]
    toggle_col = sys.argv[2] if len(sys.argv) > 2 else "E"
    
    watch_and_auto_delete(excel_file, toggle_col)


if __name__ == "__main__":
    main()
