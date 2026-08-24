"""One-command CATIA annotation-tree capture and Excel export.

Run this file from the project directory.  It never saves or closes the CATIA
document: it only activates the visible window, reads its tree, and writes a
new isolated capture run plus its Excel export.
"""

from __future__ import annotations

import argparse
from datetime import datetime
import json
from pathlib import Path
import sys
import time
import unicodedata
import re

import cv2
import numpy as np
import pytesseract

from annotation_visual_scope import VisualSubtreeError, normalise_label, select_visual_annotation_subtree
from annotation_text_recovery import recover_annotation_text
from catia_tree_pipeline import extract_all_captures, save_review_crops, write_review_queue
from excel_export import Exporter
from ocr_config import (
    ANNOTATION_TARGET_LABEL,
    FULL_IMAGE_CONFIG,
    OCR_LANGUAGE,
    RESULTS_DIR,
)
from screen_capture import activate_catia_window, get_catia_window, grab_catia_tree, save_tree_capture, tree_monitor
from tree_builder import TreeBuilder
from tree_scroller import MAX_SCROLL_CLICKS, MIN_SCROLL_CLICKS, SCROLL_CLICKS, focus_tree, go_to_top, scroll_down


PROJECT_ROOT = Path(__file__).resolve().parent
CAPTURE_RUNS_DIR = PROJECT_ROOT / "captures" / "runs"
RUN_RESULTS_DIR = RESULTS_DIR / "runs"
TARGET_KEY = normalise_label(ANNOTATION_TARGET_LABEL)


class WorkflowError(RuntimeError):
    """A clear, user-actionable stop condition for the CATIA workflow."""


def _plain_tokens(value: str) -> set[str]:
    plain = "".join(
        character
        for character in unicodedata.normalize("NFKD", value)
        if not unicodedata.combining(character)
    ).casefold()
    return set(re.findall(r"[a-z0-9]+", plain))


def _target_is_visible(image: np.ndarray) -> bool:
    """Fast, target-specific OCR used only while navigating the live tree."""
    if image is None or image.size == 0:
        return False
    height, width = image.shape[:2]
    crop = image[:height, min(55, max(0, width - 1)):min(width, 680)]
    if crop.size == 0:
        return False
    # A modest enlargement improves CATIA's compact anti-aliased labels and
    # costs far less than the geometry-aware full OCR pass used after capture.
    enlarged = cv2.resize(crop, None, fx=1.35, fy=1.35, interpolation=cv2.INTER_CUBIC)
    try:
        text = pytesseract.image_to_string(enlarged, lang=OCR_LANGUAGE, config=FULL_IMAGE_CONFIG)
    except Exception as error:
        raise WorkflowError(f"Tesseract cannot search the CATIA tree: {error}") from error

    key = normalise_label(text)
    if TARGET_KEY and TARGET_KEY in key:
        return True

    # Fallback for a selected orange row where one word may be partially
    # masked.  These three anchors together identify this CATIA UI group.
    tokens = _plain_tokens(text)
    has_result = any(token.startswith("result") for token in tokens)
    has_ensemble = any(token.startswith("ensembl") for token in tokens)
    has_annotation = any(token.startswith("annotat") for token in tokens)
    return has_annotation and (has_result or has_ensemble)


def _frame_change_ratio(previous: np.ndarray, current: np.ndarray) -> float:
    """Measure whether the tree content changed, ignoring most empty viewport."""
    if previous is None or current is None or previous.shape != current.shape:
        return 1.0
    height, width = previous.shape[:2]
    # The right CAD viewport is excluded.  The watermark is stable, but the
    # crop also makes the metric independent of its large empty background.
    right = min(width, 520)
    left = min(30, max(0, right - 1))
    before = previous[:, left:right]
    after = current[:, left:right]
    difference = cv2.absdiff(before, after)
    gray = cv2.cvtColor(difference, cv2.COLOR_BGR2GRAY)
    changed = np.count_nonzero(gray > 18)
    return float(changed) / float(gray.size or 1)


def _wait_for_catia_window(timeout_seconds: float = 35.0):
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        window = get_catia_window()
        if window is not None:
            return activate_catia_window()
        time.sleep(0.5)
    return None


def _ensure_catia(document: Path | None) -> object:
    """Attach to CATIA, launching its registered COM server only if necessary."""
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as error:
        raise WorkflowError("pywin32 is required to start or connect to CATIA.") from error

    try:
        catia = win32com.client.GetActiveObject("CATIA.Application")
        print("CATIA instance found.")
    except Exception:
        try:
            catia = win32com.client.Dispatch("CATIA.Application")
            print("CATIA started through its COM registration.")
        except Exception as error:
            raise WorkflowError(
                "CATIA could not be started. Open CATIA manually or check its COM installation."
            ) from error

    try:
        catia.Visible = True
    except Exception:
        pass

    if document is not None:
        document = document.expanduser().resolve()
        if not document.is_file():
            raise WorkflowError(f"CATIA document not found: {document}")
        active_name = ""
        try:
            active_name = str(catia.ActiveDocument.FullName)
        except Exception:
            pass
        if Path(active_name).resolve() != document:
            try:
                catia.Documents.Open(str(document))
                print(f"CATIA document opened: {document.name}")
            except Exception as error:
                raise WorkflowError(f"CATIA could not open '{document}'.") from error

    try:
        document_count = int(catia.Documents.Count)
    except Exception:
        document_count = 0
    if document_count < 1:
        raise WorkflowError(
            "CATIA is open but no document is loaded. Open the CATPart/CATProduct first, "
            "or run with --document <path>."
        )

    window = _wait_for_catia_window()
    if window is None:
        raise WorkflowError("CATIA started but its window was not found within 35 seconds.")
    return catia


def _new_run_directory() -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    destination = CAPTURE_RUNS_DIR / timestamp
    suffix = 2
    while destination.exists():
        destination = CAPTURE_RUNS_DIR / f"{timestamp}_{suffix}"
        suffix += 1
    destination.mkdir(parents=True, exist_ok=False)
    return destination


def _locate_target(maximum_steps: int) -> np.ndarray:
    """Return the first live tree frame that contains the annotation root."""
    if not go_to_top():
        raise WorkflowError("CATIA tree could not receive focus.")

    previous: np.ndarray | None = None
    stable_frames = 0
    for step in range(maximum_steps):
        frame = grab_catia_tree(activate=False)
        if frame is None:
            raise WorkflowError("Unable to capture CATIA's left tree panel.")
        if _target_is_visible(frame):
            print(f"Annotation root found after {step} navigation step(s).")
            return frame

        if previous is not None and _frame_change_ratio(previous, frame) < 0.0008:
            stable_frames += 1
            if stable_frames >= 2:
                break
        else:
            stable_frames = 0
        previous = frame
        # Searching uses a safer half-size step: the target label must never
        # slip entirely between two live frames.
        if not scroll_down(focus=False, clicks=max(MIN_SCROLL_CLICKS, SCROLL_CLICKS // 2)):
            break

    raise WorkflowError(
        "The annotation root was not found. In CATIA, make the left tree visible and "
        "expand 'Résultat d’un ensemble d’annotations', then retry."
    )


def _wait_for_user_at_target() -> np.ndarray:
    """Let the user position the visible CATIA tree before capture begins."""
    print("\nCATIA is ready.")
    print("1. In CATIA, manually scroll to 'Résultat d’un ensemble d’annotations'.")
    print("2. Expand the branches you want to export.")
    print("3. Return here and press Enter. No automatic scrolling occurs before then.")
    while True:
        input("Press Enter when the annotation root is visible in CATIA... ")
        frame = grab_catia_tree(activate=False)
        if frame is not None and _target_is_visible(frame):
            print("Annotation root confirmed. Starting captures from the current position.")
            return frame
        print("[WARN] The annotation root is not readable in the current CATIA view. Adjust the tree and press Enter again.")


def _capture_target_to_bottom(run_dir: Path, first_frame: np.ndarray, maximum_captures: int) -> list[Path]:
    """Save overlapping screenshots only from the annotation root to the end."""
    paths: list[Path] = []
    first_path = run_dir / "tree_000.png"
    if not save_tree_capture(first_frame, first_path):
        raise WorkflowError("The first annotation capture could not be saved.")
    paths.append(first_path)
    # Pressing Enter in PowerShell makes the terminal the foreground window.
    # Restore CATIA/tree focus *without clicking a node* before the first
    # wheel action; otherwise all scroll messages go to PowerShell and two
    # identical images look like an end-of-tree condition.
    if not focus_tree():
        raise WorkflowError("CATIA tree could not regain focus before scrolling.")
    focused_frame = grab_catia_tree(activate=False)
    previous = focused_frame if focused_frame is not None else first_frame
    print("CATIA tree focused. Capturing from the confirmed annotation root to the end...")
    stable_frames = 0
    scroll_clicks = SCROLL_CLICKS

    for index in range(1, maximum_captures):
        if not scroll_down(focus=False, clicks=scroll_clicks):
            raise WorkflowError("CATIA tree lost focus while scrolling.")
        frame = grab_catia_tree(activate=False)
        if frame is None:
            raise WorkflowError("A CATIA tree capture failed during scrolling.")
        ratio = _frame_change_ratio(previous, frame)
        if ratio < 0.0008:
            stable_frames += 1
            # Two unchanged frames are required: CATIA sometimes paints one
            # delayed frame immediately after a wheel action.
            if stable_frames >= 2:
                print("End of tree detected.")
                break
            # A wheel event occasionally lands between two redraws.  Increase
            # slightly before retrying, but never beyond the safe cap.
            scroll_clicks = min(MAX_SCROLL_CLICKS, max(MIN_SCROLL_CLICKS, int(round(scroll_clicks * 1.20))))
            continue

        stable_frames = 0
        # Adapt to a workstation's wheel settings.  The thresholds are based
        # on the tree-only pixel crop: below 1.8% means too much duplicate
        # content; above 20% risks losing an intermediate CATIA row.
        if ratio < 0.018:
            scroll_clicks = min(MAX_SCROLL_CLICKS, max(MIN_SCROLL_CLICKS, int(round(scroll_clicks * 1.25))))
        elif ratio > 0.20:
            scroll_clicks = max(MIN_SCROLL_CLICKS, int(round(scroll_clicks * 0.75)))
        path = run_dir / f"tree_{index:03d}.png"
        if not save_tree_capture(frame, path):
            raise WorkflowError(f"Capture {index:03d} could not be saved.")
        paths.append(path)
        previous = frame
        print(f"Capture {index:03d}: change={ratio:.3%}, next scroll={scroll_clicks}")
    else:
        raise WorkflowError(
            f"Capture stopped at the safety limit ({maximum_captures}). "
            "Increase --max-captures only after checking the CATIA tree."
        )
    return paths


def _export_annotation_tree(capture_dir: Path, run_results: Path) -> Path:
    """Run OCR once and export the complete visual annotation subtree."""
    print("\n[OCR] Reading and merging overlapping CATIA captures...")
    data = extract_all_captures(capture_dir)
    if data.empty:
        raise WorkflowError("No readable CATIA labels were found in the new capture run.")

    builder = TreeBuilder()
    builder.load_dataframe(data)
    full_tree = builder.build()
    try:
        selection = select_visual_annotation_subtree(full_tree, ANNOTATION_TARGET_LABEL)
    except VisualSubtreeError as error:
        raise WorkflowError(str(error)) from error
    result = selection.dataframe
    result = recover_annotation_text(result)

    review_dir = run_results / "review" / "annotations"
    result = save_review_crops(result, list(capture_dir.glob("*.png")), review_dir=review_dir)
    write_review_queue(result, review_dir=review_dir)

    exporter = Exporter(
        result,
        excel_file=run_results / "excel" / "annotation_tree.xlsx",
        csv_file=run_results / "csv" / "annotation_tree.csv",
        json_file=run_results / "json" / "annotation_tree.json",
    )
    outputs = exporter.export_all()
    print(
        f"[OK] {len(result)} annotation node(s), end={selection.end_reason}, "
        f"visual nodes absent from old parent graph={selection.graph_missing_nodes}."
    )
    return outputs["excel"]


def _write_manifest(run_dir: Path, *, capture_count: int | None = None, error: str | None = None) -> None:
    window = get_catia_window()
    payload = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "target": ANNOTATION_TARGET_LABEL,
        "capture_count": capture_count,
        "error": error,
        "window": (
            {
                "title": str(window.title),
                "left": int(window.left),
                "top": int(window.top),
                "width": int(window.width),
                "height": int(window.height),
                "tree_monitor": tree_monitor(window),
            }
            if window is not None
            else None
        ),
    }
    (run_dir / "manifest.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Capture CATIA annotations and export one Excel tree.")
    parser.add_argument("--document", type=Path, help="Optional CATPart/CATProduct path to open in CATIA.")
    parser.add_argument("--capture-only", action="store_true", help="Save screenshots but skip OCR/export.")
    parser.add_argument(
        "--auto-find-target",
        action="store_true",
        help="Automatically scroll from the top to find the annotation root. By default the user positions CATIA manually.",
    )
    parser.add_argument("--max-search-steps", type=int, default=180, help="Safety limit while locating the annotation root.")
    parser.add_argument("--max-captures", type=int, default=500, help="Safety limit for overlapping screenshots.")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    if args.max_search_steps < 1 or args.max_captures < 2:
        print("[ERROR] --max-search-steps must be >= 1 and --max-captures must be >= 2.")
        return 2

    started = time.time()
    run_dir: Path | None = None
    try:
        print("=" * 64)
        print("CATIA ANNOTATION TREE - ONE COMMAND EXPORT")
        print("=" * 64)
        _ensure_catia(args.document)
        run_dir = _new_run_directory()
        first_frame = _locate_target(args.max_search_steps) if args.auto_find_target else _wait_for_user_at_target()
        paths = _capture_target_to_bottom(run_dir, first_frame, args.max_captures)
        _write_manifest(run_dir, capture_count=len(paths))

        if args.capture_only:
            print(f"[OK] {len(paths)} capture(s) saved in: {run_dir}")
            return 0

        run_results = RUN_RESULTS_DIR / run_dir.name
        excel = _export_annotation_tree(run_dir, run_results)
        print("=" * 64)
        print(f"Excel: {excel}")
        print(f"Captures: {run_dir}")
        print(f"Elapsed: {time.time() - started:.1f} s")
        print("=" * 64)
        return 0
    except WorkflowError as error:
        if run_dir is not None:
            _write_manifest(run_dir, error=str(error))
        print(f"[ERROR] {error}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
