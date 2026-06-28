"""
SolidWorks Drawing Automation Pipeline
=======================================
Automatically extracts dimensions from 3D parts (.SLDPRT)
and inserts them into 2D drawings (.SLDDRW).

Requirements:
    - SolidWorks must be open and running
    - pip install pywin32

Usage:
    py solidworks_dimension_pipeline.py                  # Process all parts in current folder
    py solidworks_dimension_pipeline.py --dry-run        # Test connection only
    py solidworks_dimension_pipeline.py --part Part1.SLDPRT  # Process single part
    py solidworks_dimension_pipeline.py --dir "C:\\Models"   # Process specific folder
"""

import os
import sys
import time
import argparse
import logging
from pathlib import Path
import pythoncom
import win32com.client

# ─── Logging Setup ────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("solidworks_pipeline.log", encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)

# ─── SolidWorks Constants ─────────────────────────────────────────────────────
SW_DOC_PART        = 1
SW_DOC_DRAWING     = 3
SW_SAVE_AS_CURRENT = 0
SW_AUTO_ARRANGE    = 32768   # swAlignDimensionType_AutoArrange

# User Preference String index for Default Drawing Template path
SW_DEFAULT_TEMPLATE_DRAWING = 10

# Annotation Type: Dimensions marked for drawing
SW_INSERT_DIM_MARKED_FOR_DWG = 32768


def make_byref_int(val=0):
    """Wrap integer value in a COM VARIANT to pass it by reference (VT_BYREF)."""
    return win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, val)


# ══════════════════════════════════════════════════════════════════════════════
# 1.  COM CONNECTION
# ══════════════════════════════════════════════════════════════════════════════

def connect_to_solidworks():
    """Connect to the already-running SolidWorks instance via COM."""
    try:
        import win32com.client
    except ImportError:
        log.error("pywin32 is not installed. Run: py -m pip install pywin32")
        sys.exit(1)

    log.info("Connecting to SolidWorks ...")
    try:
        # Get active object or dispatch a new one
        sw = win32com.client.Dispatch("SldWorks.Application")
        sw.Visible = True
        version = sw.RevisionNumber
        log.info(f"Connected to SolidWorks (revision {version})")
        return sw
    except Exception as exc:
        log.error(
            "Cannot connect to SolidWorks. "
            "Make sure SolidWorks is open and running.\n"
            f"  Detail: {exc}"
        )
        sys.exit(1)


# ══════════════════════════════════════════════════════════════════════════════
# 2.  DRAWING TEMPLATE
# ══════════════════════════════════════════════════════════════════════════════

def get_drawing_template(sw) -> str:
    """Return the default drawing template path configured in SolidWorks."""
    template = sw.GetUserPreferenceStringValue(SW_DEFAULT_TEMPLATE_DRAWING)
    if template and os.path.isfile(template):
        log.info(f"Drawing template found in preferences: {template}")
        return template

    # Fallback - search common installation paths
    fallback_dirs = [
        r"C:\ProgramData\SolidWorks\SolidWorks 2026\templates",
        r"C:\ProgramData\SolidWorks\SolidWorks 2025\templates",
        r"C:\ProgramData\SolidWorks\templates",
    ]
    for d in fallback_dirs:
        candidate = os.path.join(d, "Drawing.drwdot")
        if os.path.isfile(candidate):
            log.info(f"Using fallback template: {candidate}")
            return candidate

    log.warning(
        "No drawing template found. "
        "Set one in SolidWorks: Tools > Options > System Options > Default Templates"
    )
    return ""


# ══════════════════════════════════════════════════════════════════════════════
# 3.  PROCESS A SINGLE PART
# ══════════════════════════════════════════════════════════════════════════════

def close_document(sw, doc):
    """Safely close a SolidWorks document using multiple identifier heuristics."""
    if doc is None:
        return
    try:
        title = doc.GetTitle
        if title:
            sw.CloseDoc(title)
    except Exception:
        pass
    try:
        path = doc.GetPathName
        if path:
            filename = os.path.basename(path)
            sw.CloseDoc(filename)
            stem = Path(path).stem
            sw.CloseDoc(stem)
    except Exception:
        pass


def process_part(sw, part_path: str, template: str, export_pdf: bool = True) -> bool:
    """
    Full pipeline for one .SLDPRT file:
      open part → create drawing → insert views →
      pull dimensions → auto-arrange → save → (optional PDF) → close
    Returns True on success, False on failure.
    """
    part_path = str(Path(part_path).resolve())
    drawing_path = str(Path(part_path).with_suffix(".SLDDRW"))
    pdf_path     = str(Path(part_path).with_suffix(".pdf"))

    log.info(f"{'-'*60}")
    log.info(f"Processing part: {part_path}")

    part_doc = None
    draw_doc = None

    # ── 3a. Open part ──────────────────────────────────────────────────────
    try:
        errors   = make_byref_int(0)
        warnings = make_byref_int(0)
        part_doc = sw.OpenDoc6(
            part_path,
            SW_DOC_PART,
            0,          # swOpenDocOptions_Silent
            "",
            errors,
            warnings,
        )
        if part_doc is None:
            log.error(f"  Failed to open part (errors={errors})")
            return False
        log.info("  Part opened successfully")
    except Exception as exc:
        log.error(f"  Error opening part: {exc}")
        return False

    try:
        # ── 3b. Create drawing ─────────────────────────────────────────────
        draw_doc = sw.NewDocument(template, 0, 0, 0)
        if draw_doc is None:
            log.error("  Failed to create drawing document")
            return False
        log.info("  Drawing document created")

        # Get sheet dimensions to position Isometric view dynamically
        iso_x, iso_y = 0.30, 0.20 # defaults
        try:
            sheet = draw_doc.GetCurrentSheet
            props = sheet.GetProperties2
            width = props[5]
            height = props[6]
            log.info(f"  Sheet dimensions: {width:.3f}m x {height:.3f}m")
            # Position Isometric view in the top-right quadrant
            iso_x = width * 0.75
            iso_y = height * 0.70
        except Exception as exc:
            log.warning(f"  Could not determine sheet size: {exc}. Using standard defaults.")

        # ── 3c. Insert standard orthographic views (3rd-angle) ─────────────
        # Place Front/Top/Right views automatically
        try:
            success = draw_doc.Create3rdAngleViews2(part_path)
            if success:
                log.info("  Inserted standard orthographic views (Front, Top, Right)")
            else:
                log.warning("  Create3rdAngleViews2 returned False")
        except Exception as exc:
            log.warning(f"  3rd-angle views creation error: {exc}")

        # Insert Isometric view
        try:
            iso_view = draw_doc.CreateDrawViewFromModelView3(
                part_path,
                "*Isometric",
                iso_x,
                iso_y,
                0.0,
            )
            if iso_view:
                log.info(f"  Inserted Isometric view at coordinates ({iso_x:.3f}, {iso_y:.3f})")
            else:
                log.warning("  Isometric view creation returned None")
        except Exception as exc:
            log.warning(f"  Isometric view creation error: {exc}")

        # Small pause so SolidWorks finishes rendering views
        time.sleep(1.5)

        # ── 3d. Auto-insert model dimensions ──────────────────────────────
        try:
            # InsertModelAnnotations3 has 6 parameters:
            # (Option, Types, AllViews, DuplicateDims, HiddenFeatureDims, UsePlacementInSketch)
            # Source = 0 (entire model)
            # Types = SW_INSERT_DIM_MARKED_FOR_DWG
            inserted = draw_doc.InsertModelAnnotations3(
                0,                             # Option: 0 = entire model
                SW_INSERT_DIM_MARKED_FOR_DWG,  # Types: dimensions marked for drawing
                True,                          # AllViews: True
                True,                          # DuplicateDims: True (eliminates duplicates)
                False,                         # HiddenFeatureDims: False
                False,                         # UsePlacementInSketch: False
            )
            log.info("  Model dimensions imported successfully")
        except Exception as exc:
            log.warning(f"  InsertModelAnnotations error: {exc}")

        # ── 3e. Auto-arrange dimensions in every view ──────────────────────
        try:
            view = draw_doc.GetFirstView
            if view:
                while view:
                    if view.Type != 1:  # Skip the sheet view itself (swDrawingSheet = 1)
                        _arrange_view_dimensions(draw_doc, view)
                    view = view.GetNextView
                log.info("  Dimension auto-arrange completed")
            else:
                log.warning("  No views found in drawing for auto-arrange")
        except Exception as exc:
            log.warning(f"  Auto-arrange error: {exc}")

        # ── 3f. Save drawing ───────────────────────────────────────────────
        try:
            save_errors   = make_byref_int(0)
            save_warnings = make_byref_int(0)
            result = draw_doc.SaveAs4(
                drawing_path,
                SW_SAVE_AS_CURRENT,
                0,
                save_errors,
                save_warnings,
            )
            if result:
                log.info(f"  Drawing saved: {drawing_path}")
            else:
                log.error(f"  Drawing save failed (errors={save_errors.value}, warnings={save_warnings.value})")
        except Exception as exc:
            log.error(f"  Error saving drawing: {exc}")

        # ── 3g. Optional PDF export ────────────────────────────────────────
        if export_pdf:
            try:
                pdf_errors = make_byref_int(0)
                pdf_warnings = make_byref_int(0)
                result = draw_doc.SaveAs4(
                    pdf_path,
                    SW_SAVE_AS_CURRENT,
                    1,   # swSaveAsOptions_Silent
                    pdf_errors,
                    pdf_warnings,
                )
                if result:
                    log.info(f"  PDF exported: {pdf_path}")
                else:
                    log.warning(f"  PDF export failed (errors={pdf_errors.value}, warnings={pdf_warnings.value})")
            except Exception as exc:
                log.warning(f"  PDF export error: {exc}")

        return True

    finally:
        # ── 3h. Close documents safely ─────────────────────────────────────
        close_document(sw, draw_doc)
        close_document(sw, part_doc)
        log.info("  Closed part and drawing documents")


def _arrange_view_dimensions(draw_doc, view):
    """Collect all DisplayDimension objects in a view and auto-arrange them."""
    try:
        view_name = view.Name
        draw_doc.ActivateView(view_name)
        
        # Traverse display dimensions
        dim_list = []
        disp_dim = view.GetFirstDisplayDimension5
        while disp_dim:
            dim_list.append(disp_dim)
            disp_dim = disp_dim.GetNext5
            
        if not dim_list:
            return

        # Select all collected display dimensions individually to avoid COM array mapping mismatch
        draw_doc.ClearSelection2(True)
        select_count = 0
        for disp_dim in dim_list:
            try:
                ann = disp_dim.GetAnnotation
                if ann:
                    ann.Select4(True, None)
                    select_count += 1
            except Exception:
                pass
                
        if select_count > 0:
            # Auto arrange the selected dimensions with 1.5mm gap
            align_success = draw_doc.Extension.AlignDimensions(SW_AUTO_ARRANGE, 0.0015)
            log.info(f"    View '{view_name}': Selected {select_count} and auto-arranged dimensions (result={align_success})")
        else:
            log.warning(f"    View '{view_name}': No dimensions were successfully selected")
            
        draw_doc.ClearSelection2(True)

    except Exception as exc:
        log.warning(f"    Error arranging dimensions in view '{view.Name}': {exc}")


# ══════════════════════════════════════════════════════════════════════════════
# 4.  BATCH PROCESSING
# ══════════════════════════════════════════════════════════════════════════════

def process_directory(sw, directory: str, template: str, export_pdf: bool = True):
    """Scan directory for .SLDPRT files and process each one."""
    parts = list(Path(directory).glob("*.SLDPRT"))
    # Filter out lock files/temporary files
    parts = [p for p in parts if not p.name.startswith("~$")]
    
    if not parts:
        log.warning(f"No .SLDPRT files found in: {directory}")
        return

    log.info(f"Found {len(parts)} part(s) to process in {directory}")
    success = 0
    failed  = 0

    for part in parts:
        ok = process_part(sw, str(part), template, export_pdf)
        if ok:
            success += 1
        else:
            failed += 1

    log.info(f"{'='*60}")
    log.info(f"Done! [OK] {success} succeeded, [FAIL] {failed} failed")


# ══════════════════════════════════════════════════════════════════════════════
# 5.  DRY-RUN / CONNECTION TEST
# ══════════════════════════════════════════════════════════════════════════════

def dry_run(sw):
    """Verify COM connectivity and print active options."""
    log.info("--- DRY RUN ---")
    log.info(f"  SolidWorks revision : {sw.RevisionNumber}")
    log.info(f"  Visible             : {sw.Visible}")
    log.info(f"  Active doc          : {sw.ActiveDoc}")
    log.info("  pywin32 COM connection is working correctly [OK]")
    log.info("  Run without --dry-run to process actual parts.")


# ══════════════════════════════════════════════════════════════════════════════
# 6.  ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════

def parse_args():
    p = argparse.ArgumentParser(
        description="SolidWorks Drawing Automation Pipeline"
    )
    p.add_argument(
        "--dry-run", action="store_true",
        help="Test COM connection only; do not process any files"
    )
    p.add_argument(
        "--part", metavar="FILE",
        help="Process a single .SLDPRT file"
    )
    p.add_argument(
        "--dir", metavar="DIR", default=".",
        help="Directory to scan for .SLDPRT files (default: current folder)"
    )
    p.add_argument(
        "--no-pdf", action="store_true",
        help="Skip PDF export"
    )
    return p.parse_args()


def main():
    args = parse_args()

    log.info("==========================================================")
    log.info("   SolidWorks Drawing Automation Pipeline")
    log.info("==========================================================")

    # Step 1: Connect
    sw = connect_to_solidworks()

    # Step 2: Dry run?
    if args.dry_run:
        dry_run(sw)
        return

    # Step 3: Get template
    template = get_drawing_template(sw)
    if not template:
        log.error("Unable to proceed: No valid drawing template (.drwdot) found.")
        sys.exit(1)

    export_pdf = not args.no_pdf

    # Step 4: Single part or batch?
    if args.part:
        part = str(Path(args.part).resolve())
        if not os.path.isfile(part):
            log.error(f"Part file not found: {part}")
            sys.exit(1)
        process_part(sw, part, template, export_pdf)
    else:
        process_directory(sw, args.dir, template, export_pdf)


if __name__ == "__main__":
    main()
