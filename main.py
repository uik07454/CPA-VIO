# main.py
# main() entry point and __main__ guard
# used for terminal testing without rendering gui (for dev)

import sys
from pathlib import Path

from parser import (
    parse_any_xml, 
    build_outline_level_map, 
    build_image_map, 
    parse_headers_footers, 
    save_xml_to_output
)
from extractor import extract_tracked_changes, extract_hf_changes
from book_marker import inject_bookmarks
from output import generate_excel_output
from constants import BASE_DIR, TO_PARSED_XML_PARTS

# DEBUG FLAG: Set to True to save all XML files for debugging purposes
# This should be False for production/end-user usage
_SAVE_DEBUG_XML = False
_INCLUDE_SDT = False


def main(docx_path: str = "compare.docx", save_debug_xml: bool = _SAVE_DEBUG_XML) -> None:
    """
    Orchestrate the full pipeline:
      1. Parse the .docx XML.
      2. Build a style-ID → outline level map.
      3. Extract all tracked changes.
      4. Render and save Excel output.
    """
    print(f"[INFO] Reading: {docx_path}")
    parse_results = parse_any_xml(docx_path, TO_PARSED_XML_PARTS, save_debug_xml)
    doc_root = parse_results["document"]
    styles_root = parse_results["styles"]
    footnotes_root = parse_results["footnotes"]
    endnotes_root = parse_results["endnotes"]

    print("[INFO] Building outline level map ...")
    outline_map = build_outline_level_map(styles_root)

    print("[INFO] Building image map ...")
    image_map = build_image_map(docx_path)

    note_roots = {
        "footnote": footnotes_root,
        "endnote":  endnotes_root,
    }

    print("[INFO] Extracting tracked changes ...")
    records = extract_tracked_changes(doc_root, outline_map, include_sdt=_INCLUDE_SDT, image_map=image_map, note_roots=note_roots)
    print(f"[INFO] Found {len(records)} document body change(s).")

    hf_roots = parse_headers_footers(docx_path, save_xml=save_debug_xml)
    hf_records = extract_hf_changes(hf_roots)
    print(f"[INFO] Found {len(hf_records)} header/footer change(s).")

    all_records = hf_records + records

    output_dir = BASE_DIR / "output"
    output_dir.mkdir(exist_ok=True)
    annotated_path = str(output_dir / (Path(docx_path).stem + "_annotated.docx"))
    annotated_doc_root = inject_bookmarks(docx_path, all_records, annotated_path)

    # Save the annotated document XML (with bookmarks injected) for inspection
    if save_debug_xml:
        save_xml_to_output(annotated_doc_root, annotated_path, suffix="document")
        print(f"[INFO] Annotated document XML saved to output/")

    excel_out = str(output_dir / "changes.xlsx")
    generate_excel_output(all_records, annotated_path, excel_out)


if __name__ == "__main__":
    docx = sys.argv[1] if len(sys.argv) > 1 else "compare.docx"
    main(docx)
