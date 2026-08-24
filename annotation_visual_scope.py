"""Geometry-first selection of the CATIA annotation-results subtree.

The OCR pipeline already registers overlapping screenshots into a stable visual
order.  A normal parent-id graph is useful, but it can be incomplete when a
capture begins in the middle of an expanded branch.  This module therefore
uses the on-screen indentation lane as the authority for the boundaries of
``Résultat d'un ensemble d'annotations`` and rebuilds its local hierarchy.
"""

from __future__ import annotations

from collections import defaultdict, deque
from dataclasses import dataclass
from difflib import SequenceMatcher
import math
import re
import unicodedata

import pandas as pd


SECTION_PARENT_KEYS = {
    "captures",
    "vues",
    "references",
    "cadresdetolerances",
    "tolerancegeometrique",
    "notes",
}

# In CATIA's annotation-results tree these two UI folders expose a flat list
# of direct entries.  A connector/icon OCR box can move one entry to the next
# indentation lane; treating that one-pixel artefact as a nested node produced
# errors such as ``Z2`` incorrectly becoming a child of ``X1Z1``.
FLAT_SECTION_KEYS = {"tolerancegeometrique", "notes"}


class VisualSubtreeError(ValueError):
    """Raised when a requested visual CATIA subtree cannot be selected safely."""


@dataclass(frozen=True)
class VisualSubtreeSelection:
    dataframe: pd.DataFrame
    target_node_id: str
    target_text: str
    match_method: str
    match_score: float
    end_reason: str
    graph_missing_nodes: int


def _string(value: object) -> str:
    return "" if value is None or pd.isna(value) else str(value).strip()


def _drop_instance_suffix(value: str) -> str:
    return re.sub(r"\s*\.?\s*\d+\s*$", "", value.strip())


def normalise_label(value: object) -> str:
    """Return an accent/separator insensitive key while preserving export text."""
    text = _drop_instance_suffix(_string(value)).replace("’", "'").replace("‘", "'")
    plain = "".join(
        character
        for character in unicodedata.normalize("NFKD", text)
        if not unicodedata.combining(character)
    )
    return re.sub(r"[^a-z0-9]+", "", plain.casefold())


def _word_anchors(value: object) -> set[str]:
    text = _drop_instance_suffix(_string(value)).replace("’", "'").replace("‘", "'")
    plain = "".join(
        character
        for character in unicodedata.normalize("NFKD", text)
        if not unicodedata.combining(character)
    )
    return set(re.findall(r"[a-z0-9]+", plain.casefold()))


def _node_options(row: pd.Series) -> list[tuple[str, str]]:
    options = [("text", _string(row.get("text", "")))]
    for value in _string(row.get("ocr_alternatives", "")).split("|"):
        value = _string(value)
        if value and value not in {item for _, item in options}:
            options.append(("alternative", value))
    return options


def _find_target(dataframe: pd.DataFrame, target_label: str) -> tuple[int, str, float]:
    """Find the unique annotation root, including an OCR-truncated fallback."""
    target_key = normalise_label(target_label)
    if not target_key:
        raise VisualSubtreeError("The annotation target label is empty.")

    exact: list[tuple[int, str, float]] = []
    fuzzy: list[tuple[int, str, float]] = []
    fragment: list[tuple[int, str, float]] = []
    required = {"resultat", "ensemble", "annotations"}

    for index, row in dataframe.iterrows():
        for method, candidate in _node_options(row):
            key = normalise_label(candidate)
            if not key:
                continue
            if key == target_key:
                exact.append((index, method, 1.0))
                break

            anchors = _word_anchors(candidate)
            has_required = all(
                any(anchor.startswith(word[:-1]) for anchor in anchors)
                for word in required
            )
            score = SequenceMatcher(None, key, target_key).ratio()
            if has_required and score >= 0.78:
                fuzzy.append((index, method, score))
                break

            # The selected orange CATIA row can occasionally hide the first
            # words from a line OCR pass, leaving only ``annotations``.  It is
            # accepted only when it is the unique annotation-looking row.
            if "annotat" in key and len(key) >= 8:
                fragment.append((index, method, max(0.50, score)))
                break

    candidates = exact or fuzzy
    if not candidates and len({index for index, _, _ in fragment}) == 1:
        candidates = fragment
        method_suffix = "_annotation_fragment"
    else:
        method_suffix = ""
    if not candidates:
        preview = ", ".join(_string(value) for value in dataframe.get("text", pd.Series(dtype=str)).head(15))
        raise VisualSubtreeError(
            f"Annotation root '{target_label}' was not found. OCR preview: {preview}"
        )

    best_score = max(score for _, _, score in candidates)
    best = [item for item in candidates if abs(item[2] - best_score) < 0.001]
    indexes = sorted({index for index, _, _ in best})
    if len(indexes) != 1:
        raise VisualSubtreeError(
            f"Annotation root '{target_label}' is ambiguous in {len(indexes)} OCR rows."
        )
    index = indexes[0]
    method = next(method for candidate_index, method, _ in best if candidate_index == index)
    return index, f"{method}{method_suffix}", float(best_score)


def _ordering_columns(dataframe: pd.DataFrame) -> list[str]:
    columns = [
        column
        for column in ("segment", "global_y", "capture_index", "top", "line", "node_id")
        if column in dataframe.columns
    ]
    return columns or ["node_id"]


def _number(value: object) -> float | None:
    try:
        number = float(value)
        return number if math.isfinite(number) else None
    except (TypeError, ValueError):
        return None


def _indent_step(dataframe: pd.DataFrame) -> float:
    """Infer the CATIA indentation step from connector/branch x positions."""
    column = "branch_x" if "branch_x" in dataframe.columns else "left"
    values = sorted({round(value) for value in (_number(item) for item in dataframe[column]) if value is not None})
    gaps = [later - earlier for earlier, later in zip(values, values[1:]) if 16 <= later - earlier <= 42]
    return float(pd.Series(gaps).median()) if gaps else 28.0


def _lane_value(row: pd.Series) -> float | None:
    for column in ("branch_x", "left"):
        if column in row:
            value = _number(row.get(column))
            if value is not None:
                return value
    return None


def _graph_descendants(dataframe: pd.DataFrame, root_id: str) -> set[str]:
    children: dict[str, list[str]] = defaultdict(list)
    for _, row in dataframe.iterrows():
        child = _string(row.get("node_id"))
        parent = _string(row.get("parent_id"))
        if child and parent:
            children[parent].append(child)
    selected = {root_id}
    queue: deque[str] = deque([root_id])
    while queue:
        parent = queue.popleft()
        for child in children.get(parent, []):
            if child not in selected:
                selected.add(child)
                queue.append(child)
    return selected


def _visual_end(
    ordered: pd.DataFrame,
    start: int,
    target_label: str,
) -> tuple[int, str]:
    """Return the first sibling/ancestor after the selected visual branch."""
    root = ordered.iloc[start]
    root_lane = _lane_value(root)
    root_level = int(_number(root.get("level")) or 0)
    step = _indent_step(ordered)
    lane_tolerance = max(5.0, round(step * 0.35))
    target_seen_child = False
    publication_key = normalise_label("Publication")

    for position in range(start + 1, len(ordered)):
        row = ordered.iloc[position]
        row_lane = _lane_value(row)
        row_level = int(_number(row.get("level")) or 0)
        text_key = normalise_label(row.get("text", ""))

        if root_lane is not None and row_lane is not None:
            below_root = row_lane > root_lane + lane_tolerance
            at_or_above_root = row_lane <= root_lane + lane_tolerance
        else:
            below_root = row_level > root_level
            at_or_above_root = row_level <= root_level
        target_seen_child = target_seen_child or below_root

        if not target_seen_child or not at_or_above_root:
            continue
        if text_key == publication_key:
            return position, "publication_boundary"

        # Any later row on a stable root lane, or clearly to its left, starts
        # a sibling/ancestor branch.  The graph is deliberately not used here:
        # it is the component that failed in the reported captures.
        clearly_left = root_lane is not None and row_lane is not None and row_lane < root_lane - lane_tolerance
        stable_root_lane = root_lane is not None and row_lane is not None and abs(row_lane - root_lane) <= lane_tolerance
        if clearly_left or row_level < root_level or (stable_root_lane and row_level <= root_level):
            return position, "visual_sibling_boundary"

    return len(ordered), "end_of_capture"


def _rebuild_hierarchy(selected: pd.DataFrame, target_label: str, match_method: str, match_score: float, end_reason: str) -> pd.DataFrame:
    """Rebase one visual subtree and rebuild its parent/child paths locally."""
    result = selected.copy().reset_index(drop=True)
    root = result.iloc[0]
    root_id = _string(root.get("node_id"))
    root_lane = _lane_value(root)
    root_level = int(_number(root.get("level")) or 0)
    step = _indent_step(result)

    result["source_level"] = pd.to_numeric(result.get("level", 0), errors="coerce").fillna(0).astype(int)
    result["source_parent_id"] = result.get("parent_id", "")
    result["source_parent"] = result.get("parent", "")
    result["source_full_path"] = result.get("full_path", "")

    raw_levels: list[int] = []
    for index, row in result.iterrows():
        if index == 0:
            raw_levels.append(0)
            continue
        lane = _lane_value(row)
        if root_lane is not None and lane is not None:
            inferred = int(round((lane - root_lane) / step))
        else:
            inferred = int(row["source_level"]) - root_level
        # Every remaining row belongs to the root visually.  A weak OCR box
        # can land on the root lane, but must not become a second root.
        raw_levels.append(max(1, inferred))

    # Parent sections are more reliable than a single text-box x coordinate.
    # Once a known CATIA section parent is found, it opens a visual interval
    # that ends at the following section parent.  For Notes and geometric
    # tolerance, every visible item in that interval is a direct child.
    semantic_levels = list(raw_levels)
    active_section = ""
    section_rows: set[int] = set()
    for index, row in result.iterrows():
        if index == 0:
            continue
        key = normalise_label(row.get("text", ""))
        if key in SECTION_PARENT_KEYS:
            semantic_levels[index] = 1
            active_section = key
            section_rows.add(index)
        elif active_section in FLAT_SECTION_KEYS:
            semantic_levels[index] = 2

    # A genuine parent remains visible for a range of rows.  Conversely an
    # isolated deeper reading surrounded by two same-level siblings is almost
    # always icon/connector geometry, so flatten only that isolated outlier.
    for index in range(1, len(semantic_levels) - 1):
        if index in section_rows:
            continue
        previous_level = semantic_levels[index - 1]
        current_level = semantic_levels[index]
        next_level = semantic_levels[index + 1]
        if current_level > previous_level and previous_level == next_level:
            semantic_levels[index] = previous_level

    # A missing intermediate label must not create an impossible one-row
    # depth jump.  Use the immediately preceding row, not the maximum depth
    # ever seen in the document, so returning to later siblings is correct.
    normalised_levels: list[int] = []
    previous_level = 0
    for index, level in enumerate(semantic_levels):
        if index == 0:
            normalised_levels.append(0)
            continue
        level = min(max(1, int(level)), previous_level + 1)
        normalised_levels.append(level)
        previous_level = level
    result["level"] = normalised_levels

    parent_ids: list[str] = []
    parent_names: list[str] = []
    statuses: list[str] = []
    paths: list[str] = []
    stack: dict[int, tuple[str, str, str]] = {}
    for index, row in result.iterrows():
        node_id = _string(row.get("node_id"))
        text = _string(row.get("text"))
        level = int(row["level"])
        if index == 0:
            parent_ids.append("")
            parent_names.append("ROOT")
            statuses.append("selected_annotation_root")
            path = text
        else:
            candidate_levels = [depth for depth in stack if depth < level]
            parent_level = max(candidate_levels) if candidate_levels else 0
            parent_id, parent_text, parent_path = stack[parent_level]
            parent_ids.append(parent_id)
            parent_names.append(parent_text)
            statuses.append(
                "semantic_section_child"
                if normalise_label(parent_text) in FLAT_SECTION_KEYS
                else "visual_geometry"
            )
            path = f"{parent_path} > {text}" if parent_path else text
        paths.append(path)
        stack = {depth: item for depth, item in stack.items() if depth < level}
        stack[level] = (node_id, text, path)

    # When the selected orange root was OCR-truncated, export the known CATIA
    # UI label rather than a meaningless fragment.
    if "annotation_fragment" in match_method:
        result.at[0, "text"] = target_label
        paths[0] = target_label
        for index in range(1, len(paths)):
            if paths[index].startswith(_string(root.get("text"))):
                paths[index] = f"{target_label}{paths[index][len(_string(root.get('text'))):]}"

    result["parent_id"] = parent_ids
    result["parent"] = parent_names
    result["hierarchy_status"] = statuses
    result["full_path"] = paths
    result["selection_target_node_id"] = root_id
    result["selection_match_method"] = match_method
    result["selection_match_score"] = round(float(match_score), 4)
    result["selection_end_reason"] = end_reason
    result["line"] = range(1, len(result) + 1)
    return result


def select_visual_annotation_subtree(dataframe: pd.DataFrame, target_label: str) -> VisualSubtreeSelection:
    """Keep every visually nested row from the annotation root to its sibling.

    This is intentionally a visual preorder selection.  It is the correct
    authority for CATIA scrolling screenshots when a generic parent graph has
    lost a whole visible branch such as ``Vues`` or ``Références``.
    """
    required = {"node_id", "text", "level"}
    missing = required.difference(dataframe.columns)
    if missing:
        raise VisualSubtreeError(f"Tree is missing required columns: {sorted(missing)}")
    if dataframe.empty:
        raise VisualSubtreeError("Cannot select an annotation subtree from an empty OCR result.")

    ordered = dataframe.sort_values(_ordering_columns(dataframe), kind="stable").reset_index(drop=True)
    target_index, method, score = _find_target(ordered, target_label)
    target = ordered.iloc[target_index]
    end, end_reason = _visual_end(ordered, target_index, target_label)
    selected = ordered.iloc[target_index:end].copy()
    if selected.empty:
        raise VisualSubtreeError("The annotation root was found but its visual range is empty.")

    root_id = _string(target.get("node_id"))
    graph_ids = _graph_descendants(ordered, root_id)
    graph_missing = int(sum(_string(value) not in graph_ids for value in selected["node_id"]))
    result = _rebuild_hierarchy(selected, target_label, method, score, end_reason)
    result["graph_missing_from_selection"] = result["node_id"].map(lambda value: _string(value) not in graph_ids)

    return VisualSubtreeSelection(
        dataframe=result,
        target_node_id=root_id,
        target_text=_string(result.iloc[0].get("text")),
        match_method=method,
        match_score=score,
        end_reason=end_reason,
        graph_missing_nodes=graph_missing,
    )


__all__ = [
    "VisualSubtreeError",
    "VisualSubtreeSelection",
    "normalise_label",
    "select_visual_annotation_subtree",
]
