"""Geometry-aware OCR for CATIA Specification Tree screenshots.

The previous pipeline OCRed the complete screenshot then tried to remove icon
letters afterwards.  That cannot be reliable: an icon can receive a higher
OCR confidence than the real label.  This module first finds the indentation
grid and crop rows after their icons, then uses other overlapping screenshots
as independent evidence for the same tree node.
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from pathlib import Path
import json
import math
import re
import unicodedata

import cv2
import numpy as np
import pandas as pd
import pytesseract

from batch_line_ocr import batch_line_ocr
from ocr_config import (
    DEFAULT_INDENT_STEP,
    FULL_IMAGE_CONFIG,
    LINE_IMAGE_CONFIG,
    OCR_LANGUAGE,
    REGISTRATION_MAX_SHIFT_MARGIN,
    REGISTRATION_MIN_SHIFT,
    REVIEW_DIR,
    ROW_GROUP_TOLERANCE,
    TREE_PANEL_LEFT,
    TREE_PANEL_RIGHT_MARGIN,
)


MIN_WORD_HEIGHT = 7
MAX_WORD_HEIGHT = 45
MIN_WORD_CONFIDENCE = 0.0
# Registration is deliberately compressed only horizontally.  The vertical
# coordinate must remain in pixels because it becomes the global tree order.
# 80 columns preserves the layout of labels/branches while making a complete
# 40-capture registration take seconds instead of minutes.
REGISTRATION_FEATURE_WIDTH = 80

# CATIA exposes a small set of standard, localised tree labels.  They are not
# global spelling substitutions: a correction is made only after a close
# fuzzy match to one of these UI-owned labels.  Arbitrary part names are left
# untouched and remain subject to normal OCR evidence/review.
CATIA_UI_LABELS = (
    "Références",
    "Cadres de tolérances",
    "Tolérance géométrique",
    "Résultat d'un ensemble d'annotations",
    "Captures",
    "Vues",
    "Notes",
    "Publication",
    "Corps principal",
)


def _natural_key(path: Path) -> tuple:
    """Sort capture files by their numeric sequence, not lexical accidents."""
    chunks = re.split(r"(\d+)", path.name.casefold())
    return tuple(int(chunk) if chunk.isdigit() else chunk for chunk in chunks)


def _normalise_for_match(text: str) -> str:
    # Diacritics must be ignored for matching but retained in exported text.
    # ``NFKD`` makes `é` comparable with `e` without turning arbitrary French
    # labels into ASCII in the actual result.
    decomposed = unicodedata.normalize("NFKD", str(text))
    without_marks = "".join(char for char in decomposed if not unicodedata.combining(char))
    return re.sub(r"[^a-z0-9]+", "", without_marks.casefold())


def _canonical_catia_ui_label(text: str) -> str:
    """Restore a CATIA built-in French label only when the match is strong."""
    value = str(text).strip()
    if not value:
        return value
    suffix_match = re.search(r"(\s*\.?\s*\d+)$", value)
    suffix = suffix_match.group(1).replace(" ", "") if suffix_match else ""
    base = value[:suffix_match.start()].strip() if suffix_match else value
    normalised = _normalise_for_match(base)
    if len(normalised) < 4:
        return value
    best_label = None
    best_similarity = 0.0
    for label in CATIA_UI_LABELS:
        reference = _normalise_for_match(label)
        similarity = SequenceMatcher(None, normalised, reference).ratio()
        if similarity > best_similarity:
            best_similarity = similarity
            best_label = label
    # The length check rejects accidental matches such as a short part code
    # that happens to share a few French characters with a UI label.
    if (
        best_label is not None
        and best_similarity >= 0.82
        and abs(len(normalised) - len(_normalise_for_match(best_label))) <= max(3, len(normalised) // 4)
    ):
        return f"{best_label}{suffix}"
    return value


def _is_textual(text: str) -> bool:
    return any(character.isalnum() for character in str(text))


def _text_plausibility(text: str) -> float:
    """Small, explainable score used with OCR confidence, never alone."""
    value = str(text).strip()
    alnum = "".join(character for character in value if character.isalnum())
    if not alnum:
        return -70.0
    score = min(len(alnum), 24) * 0.8
    if len("".join(character for character in value if character.isalpha())) >= 3:
        score += 7.0
    if re.fullmatch(r"[A-Za-z]{1,3}\d+[A-Za-z0-9.-]*", value):
        score += 12.0
    if re.fullmatch(r"[0-9.]+", value):
        # A cropped first character such as "1.0" is usually an incomplete
        # identifier, not a standalone CATIA node.  It remains available for
        # review when no better candidate exists.
        score -= 30.0
    if len(alnum) == 1 and not any(char.isdigit() for char in alnum):
        score -= 28.0
    if "\ufffd" in value or "?" in value:
        score -= 18.0
    if re.match(r"^[^A-Za-z0-9]", value):
        score -= 12.0
    return score


def _is_structured_short_identifier(text: str) -> bool:
    """Return True for common CATIA short identifiers such as Y2 or X1Z1."""
    compact = re.sub(r"[^A-Za-z0-9_.-]", "", str(text))
    return bool(re.fullmatch(r"[A-Za-z]{1,3}\d+[A-Za-z0-9_.-]*", compact))


def _clean_text(text: str) -> str:
    """Keep CATIA identifiers intact; only remove OCR control noise."""
    text = (
        str(text)
        .replace("\x0c", " ")
        .replace("\ufffd", "")
        .replace("\u2018", "'")
        .replace("\u2019", "'")
    )
    # Python's Unicode-aware ``\w`` preserves CATIA's French accents while
    # still removing OCR control noise.  Earlier ASCII-only matching turned
    # words such as `Découpe` into `Dcoupe`.
    text = re.sub(r"[^\w .|/()'\-]+", "", text, flags=re.UNICODE)
    # A crop can start inside CATIA's branch/icon and therefore OCR a leading
    # closing bracket or a trailing separator (for example ")Z2" or "Z2|").
    # Remove only unbalanced crop-edge noise; keep punctuation inside names.
    text = re.sub(r"^[\]\)}|]+", "", text)
    text = re.sub(r"[\[\({|]+$", "", text)
    text = re.sub(r"\s+", " ", text).strip()

    # A lone lower-case letter in an otherwise all-capital acronym is an OCR
    # case error (cPC -> CPC).  Names such as Startup remain unchanged.
    words = []
    for word in text.split():
        letters = "".join(character for character in word if character.isalpha())
        if len(letters) >= 2 and sum(char.isupper() for char in letters) >= len(letters) - 1:
            prefix = "".join(char.upper() if char.isalpha() else char for char in word)
            words.append(prefix)
        else:
            words.append(word)
    return " ".join(words)


def _line_words(data: pd.DataFrame) -> pd.DataFrame:
    required = {"text", "conf", "left", "top", "width", "height"}
    if data.empty or not required.issubset(data.columns):
        return pd.DataFrame(columns=["text", "conf", "left", "top", "width", "height", "center_y"])
    result = data.copy()
    result["text"] = result["text"].astype(str).str.strip()
    result["conf"] = pd.to_numeric(result["conf"], errors="coerce").fillna(-1.0)
    result = result[
        (result["text"] != "")
        & (result["conf"] >= MIN_WORD_CONFIDENCE)
        & result["height"].between(MIN_WORD_HEIGHT, MAX_WORD_HEIGHT)
    ].copy()
    result = result[result["text"].apply(_is_textual)]
    if result.empty or not required.issubset(result.columns):
        return pd.DataFrame(columns=["text", "conf", "left", "top", "width", "height", "center_y"])
    try:
        result["center_y"] = result["top"] + result["height"] / 2.0
    except KeyError:
        return pd.DataFrame(columns=["text", "conf", "left", "top", "width", "height", "center_y"])
    return result.reset_index(drop=True)


def _ocr_data(image: np.ndarray, config: str) -> pd.DataFrame:
    data = pytesseract.image_to_data(
        image,
        lang=OCR_LANGUAGE,
        config=config,
        output_type=pytesseract.Output.DATAFRAME,
    )
    return _line_words(data.dropna(subset=["text"]))


def _bright_tree_mask(image: np.ndarray) -> np.ndarray:
    """Mask light, low-saturation CATIA text and connector lines."""
    hsv = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)
    return ((hsv[:, :, 1] < 95) & (hsv[:, :, 2] > 135)).astype(np.uint8)


def _cluster_positions(positions: list[int], maximum_gap: int = 3) -> list[float]:
    if not positions:
        return []
    clusters: list[list[int]] = [[positions[0]]]
    for value in positions[1:]:
        if value - clusters[-1][-1] <= maximum_gap:
            clusters[-1].append(value)
        else:
            clusters.append([value])
    return [float(np.median(cluster)) for cluster in clusters]


def _infer_text_lanes(image: np.ndarray) -> tuple[list[int], float]:
    """Infer text starts from the long vertical CATIA connector lines."""
    mask = _bright_tree_mask(image)
    h, w = mask.shape
    search_right = min(260, w - TREE_PANEL_RIGHT_MARGIN)

    # Long vertical connectors survive this opening while glyph strokes do not.
    vertical = cv2.morphologyEx(mask[:, :search_right], cv2.MORPH_OPEN, np.ones((45, 1), np.uint8))
    column_score = vertical.sum(axis=0)
    eligible = column_score[TREE_PANEL_LEFT:]
    if len(eligible) == 0 or int(eligible.max(initial=0)) == 0:
        return [75 + DEFAULT_INDENT_STEP * depth for depth in range(9)], float(DEFAULT_INDENT_STEP)

    threshold = max(35, int(float(eligible.max()) * 0.42))
    columns = [
        column
        for column in range(TREE_PANEL_LEFT, search_right)
        if column_score[column] >= threshold
    ]
    axes = _cluster_positions(columns)
    axes = [axis for axis in axes if axis >= TREE_PANEL_LEFT]

    differences = [
        later - earlier
        for earlier, later in zip(axes, axes[1:])
        if 20.0 <= later - earlier <= 40.0
    ]
    step = float(np.median(differences)) if differences else float(DEFAULT_INDENT_STEP)

    if axes:
        # Choose the axis that has the most neighbours one indentation apart.
        def support(axis: float) -> int:
            return sum(
                1
                for other in axes
                if other > axis and abs(((other - axis) / step) - round((other - axis) / step)) < 0.16
            )

        base_axis = min(axes, key=lambda axis: (-support(axis), axis))
    else:
        base_axis = 47.0

    lanes = [int(round(base_axis + step * (depth + 1))) for depth in range(10)]
    return lanes, step


def _group_word_rows(words: pd.DataFrame) -> list[pd.DataFrame]:
    """Group full-image OCR words by their physical CATIA row."""
    if words.empty:
        return []
    rows: list[list[pd.Series]] = []
    centers: list[float] = []
    for _, word in words.sort_values(["center_y", "left"]).iterrows():
        center = float(word["center_y"])
        best_index = None
        best_distance = math.inf
        for index, previous_center in enumerate(centers):
            distance = abs(center - previous_center)
            if distance <= ROW_GROUP_TOLERANCE and distance < best_distance:
                best_index = index
                best_distance = distance
        if best_index is None:
            rows.append([word])
            centers.append(center)
        else:
            rows[best_index].append(word)
            centers[best_index] = float(np.median([item["center_y"] for item in rows[best_index]]))
    return [pd.DataFrame(group).sort_values("left").reset_index(drop=True) for group in rows]


def _nearest_lane(value: float, lanes: list[int]) -> tuple[int, int]:
    index = min(range(len(lanes)), key=lambda item: abs(lanes[item] - value))
    return lanes[index], index


def _full_image_candidate(group: pd.DataFrame, lanes: list[int], indent_step: float) -> dict:
    max_confidence = float(group["conf"].max())
    strong_minimum = max(22.0, max_confidence - 25.0)
    strong = group[group["conf"] >= strong_minimum]
    signal = strong if not strong.empty else group
    signal = signal[signal["text"].apply(_is_textual)]
    left_signal = float(signal["left"].min()) if not signal.empty else float(group["left"].min())
    lane, level = _nearest_lane(left_signal, lanes)

    # Anything clearly left of the inferred text lane is an icon, connector,
    # or expand/collapse marker, not a part of the label.
    label_words = group[(group["left"] >= lane - 4) & (group["conf"] >= 12.0)].copy()
    # Keep this provenance separately from the fallback below.  A group made
    # only of an expand icon/watermark has no reliable row centre for crop
    # OCR: a 38px crop can otherwise read parts of the two neighbouring
    # labels and manufacture a convincing but non-existent node.
    has_label_evidence = not label_words.empty
    if label_words.empty:
        label_words = signal.copy()
    label_words = label_words.sort_values("left")
    # A lowercase icon artefact immediately before a real short code ("os
    # Y3") must not set the node's indentation.  This is geometry-aware: a
    # genuine long label such as "Plan xy" is never removed here.
    label_words = _drop_crop_icon_prefixes(label_words)

    # Full-image OCR occasionally gives the first code of a multiword label a
    # very low confidence while reading the second word correctly.  Example:
    # "Z2 Line" becomes a weak "22" at level 3 plus strong "Line" at level
    # 4.  Keep a nearby structured/numbered prefix rather than turning it
    # into a false child called "Line".  Icon garbage like "SQ? EXTRACT" or
    # "SY CPC TRACE" has no digit and is still excluded.
    if not label_words.empty:
        first_left = int(label_words["left"].min())
        prefixes = group[group["left"] < first_left].sort_values("left", ascending=False)
        for prefix_index, prefix in prefixes.iterrows():
            prefix_text = _clean_text(prefix["text"])
            compact_prefix = re.sub(r"[^A-Za-z0-9]", "", prefix_text)
            prefix_right = int(prefix["left"] + prefix["width"])
            one_lane_before = abs(float(prefix["left"]) - (lane - indent_step)) <= max(9.0, indent_step * 0.36)
            adjacent = prefix_right >= first_left - max(8, int(round(indent_step * 0.40)))
            structured = bool(re.search(r"\d", compact_prefix)) and len(compact_prefix) <= 10
            if one_lane_before and adjacent and structured:
                label_words = pd.concat([group.loc[[prefix_index]], label_words]).sort_values("left")
                break

    text = _clean_text(" ".join(label_words["text"].astype(str)))
    confidence = float(label_words["conf"].mean()) if not label_words.empty else 0.0
    left = int(label_words["left"].min()) if not label_words.empty else int(round(lane))
    top = int(round(float(np.median(group["top"]))))
    bottom = int((group["top"] + group["height"]).max())
    actual_lane, actual_level = _nearest_lane(left, lanes)
    low_prefixes = group[(group["left"] < left - 10) & (group["conf"] < 25.0)]
    prefix_lane = None
    if not low_prefixes.empty:
        prefix_lane, _ = _nearest_lane(float(low_prefixes["left"].min()), lanes)
    score = confidence + _text_plausibility(text)
    return {
        "text": text,
        "left": left,
        "top": top,
        "height": max(1, bottom - top),
        "width": int((label_words["left"] + label_words["width"]).max() - left) if not label_words.empty else 0,
        "conf": confidence,
        "score": score,
        "lane": actual_lane,
        "level": actual_level,
        "branch_x": int(round(actual_lane - indent_step)),
        "method": "full_image",
        "alternatives": [text] if text else [],
        "low_confidence_prefix": not low_prefixes.empty,
        "prefix_lane": prefix_lane,
        "missing_label_evidence": not has_label_evidence,
    }


def _drop_crop_icon_prefixes(words: pd.DataFrame) -> pd.DataFrame:
    """Drop a low-quality icon token only when a real word follows it."""
    words = words.sort_values("left").copy()
    while len(words) > 1:
        first = words.iloc[0]
        rest = words.iloc[1:]
        letters = "".join(character for character in str(first["text"]) if character.isalpha())
        lower_quality = float(first["conf"]) + 12.0 < float(rest["conf"].max())
        follows_word = any(
            sum(character.isalnum() for character in str(text)) >= 2
            for text in rest["text"]
        )
        follows_code = any(_is_structured_short_identifier(text) for text in rest["text"])
        # ``cPC`` is a genuine acronym with one OCR case error, whereas
        # ``os`` before Y3 is icon noise.  Do not discard the former.
        mostly_upper = bool(letters) and sum(char.isupper() for char in letters) >= len(letters) - 1
        if 0 < len(letters) <= 3 and not mostly_upper and lower_quality and (follows_word or follows_code):
            words = rest
            continue
        break
    return words


def _ocr_line_crop(image: np.ndarray, center_y: float, crop_x: int, variant: str) -> dict | None:
    height, width = image.shape[:2]
    top = max(0, int(round(center_y - 16)))
    bottom = min(height, int(round(center_y + 22)))
    right = max(crop_x + 2, width - TREE_PANEL_RIGHT_MARGIN)
    crop = image[top:bottom, crop_x:right]
    if crop.size == 0:
        return None

    code_variant = variant.startswith("code_")
    threshold = int(variant.split("_")[-1]) if (variant.startswith("binary_") or code_variant) else None
    if threshold is not None:
        gray = cv2.cvtColor(crop, cv2.COLOR_BGR2GRAY)
        _, crop = cv2.threshold(gray, threshold, 255, cv2.THRESH_BINARY)

    config = LINE_IMAGE_CONFIG
    if code_variant:
        config += " -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_-. "
    words = _ocr_data(crop, config)
    if words.empty:
        return None
    words = _drop_crop_icon_prefixes(words)
    if words.empty:
        return None
    text = _clean_text(" ".join(words.sort_values("left")["text"].astype(str)))
    if not text:
        return None
    confidence = float(words["conf"].mean())
    left = crop_x + int(words["left"].min())
    return {
        "text": text,
        "left": left,
        "top": top + int(words["top"].median()),
        "height": int(words["height"].max()),
        "width": int((words["left"] + words["width"]).max() - words["left"].min()),
        "conf": confidence,
        "score": confidence + _text_plausibility(text),
        "method": f"line_{variant}",
    }


def _requires_recovery(candidate: dict, image_median_confidence: float) -> bool:
    text = candidate["text"]
    compact = re.sub(r"[^A-Za-z0-9]", "", text)
    weak_relative_to_image = candidate["conf"] < image_median_confidence - 20.0
    ambiguous_short_code = bool(re.fullmatch(r"[A-Za-z][SZOIL]", compact))
    numeric_short = bool(re.fullmatch(r"\d{1,4}", compact))
    malformed = (
        len(compact) < 2
        or (len(compact) <= 3 and candidate["conf"] < 55.0 and not _is_structured_short_identifier(text))
        or ambiguous_short_code
        or numeric_short
        or "?" in text
        or "\ufffd" in text
    )
    clearly_long_and_plausible = (
        len(compact) >= 7
        and bool(re.search(r"[A-Za-z]{3,}", text))
        and candidate["conf"] >= 40.0
    )
    clearly_structured_identifier = (
        len(compact) >= 4
        and any(char.isalpha() for char in compact)
        and any(char.isdigit() for char in compact)
        and candidate["conf"] >= 40.0
    )
    return (
        candidate.get("low_confidence_prefix", False)
        and not clearly_long_and_plausible
        and not clearly_structured_identifier
    ) or malformed or (
        weak_relative_to_image and not clearly_long_and_plausible and not clearly_structured_identifier
    )


def _recover_candidate(
    image: np.ndarray,
    candidate: dict,
    center_y: float,
    lanes: list[int],
    indent_step: float,
) -> dict:
    """Recover a weak row without allowing an icon bbox to change its level.

    The crop lane is the structural evidence.  Tesseract's ``left`` inside a
    crop is only an OCR box and can begin on a branch/icon, so it is never
    reused to infer indentation.
    """
    current_lane = candidate["lane"]
    current_index = min(range(len(lanes)), key=lambda index: abs(lanes[index] - current_lane))
    lane_order = [current_lane]
    if candidate.get("prefix_lane") is not None:
        lane_order.append(int(candidate["prefix_lane"]))
    compact_original = re.sub(r"[^A-Za-z0-9]", "", candidate["text"])
    source_looks_like_short_code = len(compact_original) <= 3
    if len(compact_original) <= 5 and current_index + 1 < len(lanes):
        # An icon OCR box is often exactly one indentation before the label.
        lane_order.append(lanes[current_index + 1])
    lane_order = list(dict.fromkeys(lane_order))
    options: list[dict] = [candidate]

    def register_crop_option(option: dict, lane: int) -> dict:
        crop_lane, crop_level = _nearest_lane(lane, lanes)
        geometry_alignment = max(-24.0, 12.0 - 1.5 * abs(option["left"] - crop_lane))
        option.update(
            lane=crop_lane,
            level=crop_level,
            branch_x=int(round(crop_lane - indent_step)),
            crop_lane=crop_lane,
            geometry_alignment=geometry_alignment,
        )
        option["score"] += geometry_alignment
        if source_looks_like_short_code:
            compact_option = re.sub(r"[^A-Za-z0-9]", "", option["text"])
            if _is_structured_short_identifier(option["text"]):
                option["score"] += 14.0
            elif not any(char.isalpha() for char in compact_option):
                # A number fragment created by a branch/watermark must not
                # beat a weak but mergeable label solely on Tesseract's score.
                option["score"] -= 45.0
            else:
                option["score"] -= 8.0
        return option

    for lane in lane_order:
        raw = _ocr_line_crop(image, center_y, max(0, int(round(lane - 15))), "raw")
        if raw:
            options.append(register_crop_option(raw, lane))

    # Most rows need only the raw crop.  For short/weak labels we make a small
    # adaptive threshold sweep.  It is intentionally limited to this subset:
    # it fixes Y2/Y3 and labels hidden by the watermark without multiplying
    # the OCR time for every line of every capture.
    best_raw = max(options, key=lambda item: item["score"])
    # Threshold variants are reserved for genuinely short labels.  For a
    # long weak label, a raw crop is still useful evidence but four unrelated
    # threshold passes cost time and can manufacture a convincing fragment.
    short_or_weak = source_looks_like_short_code
    def has_reliable_code(items: list[dict]) -> bool:
        return any(
            _is_structured_short_identifier(option["text"])
            and option["conf"] >= 25.0
            # The original full-image candidate is not valid evidence after
            # its lane was inferred from neighbours if its OCR box still sits
            # on an icon one indentation to the left ("an 2)" before Y2).
            and ("crop_lane" in option or abs(float(option["left"]) - float(option["lane"])) <= 8.0)
            for option in items
        )

    if short_or_weak and not has_reliable_code(options):
        # Prefer the candidate's current structural lane.  A deeper lane is
        # tried only when the first crop has no viable result; this handles an
        # icon that was OCRed as the label while avoiding needless passes.
        recovery_lanes = [current_lane]
        for variant in ("binary_130", "code_110", "binary_100", "binary_160"):
            for lane in recovery_lanes:
                option = _ocr_line_crop(image, center_y, max(0, int(round(lane - 15))), variant)
                if option:
                    options.append(register_crop_option(option, lane))
            if has_reliable_code(options):
                break

        # If the current lane did not yield a structured code, one final
        # binary pass at the next lane handles the common icon-before-label
        # case without doing a full multi-threshold sweep twice.
        if not has_reliable_code(options) and len(lane_order) > 1:
            option = _ocr_line_crop(
                image,
                center_y,
                max(0, int(round(lane_order[1] - 15))),
                "binary_130",
            )
            if option:
                options.append(register_crop_option(option, lane_order[1]))

    winner = max(options, key=lambda item: item["score"])
    winner["alternatives"] = list(dict.fromkeys(option["text"] for option in options if option["text"]))
    # Unlike the old string-only alternatives, every alternative retains its
    # own score/confidence/method.  This prevents a bad threshold variant from
    # inheriting the score of a good OCR result during multi-capture fusion.
    evidence_by_key: dict[tuple[str, str, int], dict] = {}
    for option in options:
        cleaned = _clean_text(option.get("text", ""))
        if not cleaned:
            continue
        key = (cleaned, str(option.get("method", "")), int(option.get("crop_lane", option.get("lane", 0))))
        payload = {
            "text": cleaned,
            "conf": float(option.get("conf", 0.0)),
            "score": float(option.get("score", -math.inf)),
            "method": str(option.get("method", "")),
        }
        previous = evidence_by_key.get(key)
        if previous is None or payload["score"] > previous["score"]:
            evidence_by_key[key] = payload
    winner["candidate_evidence"] = list(evidence_by_key.values())
    return winner


def _batch_option(
    raw: dict,
    lane: int,
    lanes: list[int],
    indent_step: float,
    source_looks_like_short_code: bool,
) -> dict | None:
    """Convert one stacked-crop OCR result into a scored tree candidate."""
    option = dict(raw)
    text = _clean_text(option.get("text", ""))
    if not text:
        return None
    crop_lane, crop_level = _nearest_lane(lane, lanes)
    option["text"] = text
    option["method"] = f"line_batch_{option.get('method', 'raw')}"
    option["score"] = float(option.get("conf", 0.0)) + _text_plausibility(text)
    geometry_alignment = max(-24.0, 12.0 - 1.5 * abs(float(option["left"]) - crop_lane))
    option.update(
        lane=crop_lane,
        level=crop_level,
        branch_x=int(round(crop_lane - indent_step)),
        crop_lane=crop_lane,
        geometry_alignment=geometry_alignment,
    )
    option["score"] += geometry_alignment
    if source_looks_like_short_code:
        compact = re.sub(r"[^A-Za-z0-9]", "", text)
        if _is_structured_short_identifier(text):
            option["score"] += 14.0
        elif not any(character.isalpha() for character in compact):
            option["score"] -= 45.0
        else:
            option["score"] -= 8.0
    return option


def _finalize_recovery_options(options: list[dict]) -> dict:
    """Pick one option while retaining score-bearing audit evidence."""
    winner = max(options, key=lambda item: item["score"])
    winner["alternatives"] = list(dict.fromkeys(
        _clean_text(option.get("text", "")) for option in options if _clean_text(option.get("text", ""))
    ))
    evidence_by_key: dict[tuple[str, str, int], dict] = {}
    for option in options:
        cleaned = _clean_text(option.get("text", ""))
        if not cleaned:
            continue
        key = (cleaned, str(option.get("method", "")), int(option.get("crop_lane", option.get("lane", 0))))
        payload = {
            "text": cleaned,
            "conf": float(option.get("conf", 0.0)),
            "score": float(option.get("score", -math.inf)),
            "method": str(option.get("method", "")),
        }
        previous = evidence_by_key.get(key)
        if previous is None or payload["score"] > previous["score"]:
            evidence_by_key[key] = payload
    winner["candidate_evidence"] = list(evidence_by_key.values())
    return winner


def _batch_recover_candidates(
    image: np.ndarray,
    basic: list[dict],
    groups: list[pd.DataFrame],
    image_median_confidence: float,
    lanes: list[int],
    indent_step: float,
) -> tuple[dict[int, dict], set[int]]:
    """Recover all uncertain rows with one stacked Tesseract call per image.

    Individual crops were accurate but expensive because they launched
    Tesseract repeatedly.  This function submits all relevant line crops and
    preprocessing variants to :func:`batch_line_ocr` in a single call, then
    lets the same geometry/plausibility scoring decide among them.
    """
    options_by_row: dict[int, list[dict]] = {}
    request_info: dict[str, tuple[int, int, bool]] = {}
    requests: list[dict] = []
    recovered_indices: set[int] = set()

    for row_index, (group, original) in enumerate(zip(groups, basic)):
        candidate = dict(original)
        partial = bool(candidate["top"] <= 2 or candidate["top"] + candidate["height"] >= image.shape[0] - 2)
        icon_only_outside_tree = (
            candidate.get("missing_label_evidence", False)
            and int(candidate.get("left", 0)) < min(lanes) - 4
        )
        if partial or icon_only_outside_tree or not _requires_recovery(candidate, image_median_confidence):
            continue

        recovered_indices.add(row_index)
        options_by_row[row_index] = [candidate]
        center_y = float(np.median(group["center_y"]))
        compact = re.sub(r"[^A-Za-z0-9]", "", candidate["text"])
        source_short = len(compact) <= 3
        current_index = min(range(len(lanes)), key=lambda index: abs(lanes[index] - candidate["lane"]))
        crop_lanes = [int(candidate["lane"])]
        if (source_short or candidate.get("low_confidence_prefix", False)) and current_index + 1 < len(lanes):
            crop_lanes.append(int(lanes[current_index + 1]))
        crop_lanes = list(dict.fromkeys(crop_lanes))

        variants = ["raw"]
        if source_short:
            # These are independent evidence, not a confidence threshold.
            # Their small sweep solves different watermark/anti-alias cases.
            variants.extend(["binary_130", "code_110", "binary_100", "binary_160"])
        for lane in crop_lanes:
            for variant in variants:
                request_id = f"{row_index}:{lane}:{variant}"
                request_info[request_id] = (row_index, lane, source_short)
                requests.append({
                    "id": request_id,
                    "center_y": center_y,
                    "crop_x": max(0, int(round(lane - 15))),
                    "variant": variant,
                })

    if not requests:
        return {}, recovered_indices

    # PSM 11 keeps stacked crops independent; unlike PSM 6 it did not merge
    # the preceding CATIA branch with a short label in validation.
    batched = batch_line_ocr(image, requests, psm=11)
    for request_id, raw_candidates in batched.items():
        row_index, lane, source_short = request_info[request_id]
        for raw in raw_candidates:
            option = _batch_option(raw, lane, lanes, indent_step, source_short)
            if option is not None:
                options_by_row[row_index].append(option)

    recovered: dict[int, dict] = {}
    for row_index, options in options_by_row.items():
        # A blank batch response is still safe: retain the original candidate
        # and let cross-capture fusion/review decide rather than invent text.
        recovered[row_index] = _finalize_recovery_options(options)
    return recovered, recovered_indices


def read_capture(image_path: Path, capture_index: int) -> tuple[list[dict], np.ndarray]:
    image = cv2.imread(str(image_path))
    if image is None:
        raise FileNotFoundError(image_path)
    lanes, indent_step = _infer_text_lanes(image)
    words = _ocr_data(image, FULL_IMAGE_CONFIG)
    # Ignore the static scrollbar / off-panel OCR while keeping the first root.
    words = words[(words["left"] >= min(lanes) - 38) & (words["left"] < image.shape[1] - TREE_PANEL_RIGHT_MARGIN)]
    groups = _group_word_rows(words)
    basic = [_full_image_candidate(group, lanes, indent_step) for group in groups]
    confidence_values = [row["conf"] for row in basic if row["text"]]
    image_median = float(np.median(confidence_values)) if confidence_values else 60.0
    batch_recovered, batch_indices = _batch_recover_candidates(
        image,
        basic,
        groups,
        image_median,
        lanes,
        indent_step,
    )

    rows: list[dict] = []
    for row_index, (group, original_candidate) in enumerate(zip(groups, basic)):
        # Do not create a tree node from a connector/icon-only OCR group.
        # A genuine label may still be seen in an overlapping capture, while
        # retaining this group would insert a false parent/child row.
        icon_only_outside_tree = (
            original_candidate.get("missing_label_evidence", False)
            and int(original_candidate.get("left", 0)) < min(lanes) - 4
        )
        if icon_only_outside_tree:
            continue
        candidate = dict(batch_recovered.get(row_index, original_candidate))
        center = float(np.median(group["center_y"]))
        initial_partial = bool(
            original_candidate["top"] <= 2
            or original_candidate["top"] + original_candidate["height"] >= image.shape[0] - 2
        )
        needs_individual_fallback = (
            row_index not in batch_indices
            and _requires_recovery(candidate, image_median)
            and not initial_partial
        )
        if needs_individual_fallback:
            candidate = _recover_candidate(image, candidate, center, lanes, indent_step)
        if not candidate["text"]:
            continue
        candidate.setdefault("alternatives", [candidate["text"]])
        candidate.setdefault(
            "candidate_evidence",
            [{
                "text": candidate["text"],
                "conf": float(candidate["conf"]),
                "score": float(candidate["score"]),
                "method": str(candidate["method"]),
            }],
        )
        candidate.update(
            image=image_path.name,
            capture_index=capture_index,
            row_center=center,
            indent_step=indent_step,
            source_count=1,
            # A clipped row can support an existing node but must never create
            # a new tree node by itself (e.g. the fragment at image top).
            partial=bool(
                initial_partial
                or candidate["top"] <= 2
                or candidate["top"] + candidate["height"] >= image.shape[0] - 2
            ),
        )
        rows.append(candidate)
    return rows, image


def _registration_feature(image: np.ndarray) -> np.ndarray:
    """Scroll-moving foreground only: labels and short horizontal branches.

    CATIA's long vertical trunks are fixed at the same screen x coordinate and
    would otherwise make every small translation look correct.  They are
    removed before the correlation.
    """
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    h, w = gray.shape
    left = max(1, int(round(w * 0.08)))
    right = min(w, int(round(w * 0.93)))
    raw = ((gray[:, left:right] > 145).astype(np.uint8) * 255)
    vertical = cv2.morphologyEx(raw, cv2.MORPH_OPEN, np.ones((25, 1), np.uint8))
    foreground = ((raw > 0) & (vertical == 0)).astype(np.float32)
    # Preserve every vertical pixel but pool horizontal detail.  This keeps
    # the distinctive label/branch pattern while making the NCC scan cheap.
    return cv2.resize(
        foreground,
        (REGISTRATION_FEATURE_WIDTH, foreground.shape[0]),
        interpolation=cv2.INTER_AREA,
    ).astype(np.float32)


def _normalised_overlap(a: np.ndarray, b: np.ndarray) -> float:
    if a.shape[0] < 80 or b.shape[0] < 80:
        return -1.0
    if int(np.count_nonzero(a)) < 80 or int(np.count_nonzero(b)) < 80:
        return -1.0
    # ``ravel`` may be a view of the shared registration feature.  Centering
    # must never mutate it because the scan evaluates hundreds of shifts.
    first = a.ravel().astype(np.float32, copy=True)
    second = b.ravel().astype(np.float32, copy=True)
    first -= first.mean()
    second -= second.mean()
    denominator = float(np.linalg.norm(first) * np.linalg.norm(second))
    return float(np.dot(first, second) / denominator) if denominator else -1.0


def _estimate_shift_from_anchors(previous_rows: list[dict], current_rows: list[dict]) -> tuple[int | None, float, float]:
    """Estimate scroll from several independent, geometrically equal labels."""
    votes: list[tuple[int, float, str]] = []
    for previous in previous_rows:
        previous_key = _normalise_for_match(previous["text"])
        if len(previous_key) < 4 or previous["conf"] < 48.0:
            continue
        for current in current_rows:
            current_key = _normalise_for_match(current["text"])
            if len(current_key) < 4 or current["conf"] < 48.0:
                continue
            if previous["level"] != current["level"]:
                continue
            similarity = SequenceMatcher(None, previous_key, current_key).ratio()
            if similarity < 0.86:
                continue
            shift = int(round(previous["top"] - current["top"]))
            if shift <= REGISTRATION_MIN_SHIFT:
                continue
            weight = similarity * min(previous["conf"], current["conf"]) / 100.0
            votes.append((shift, weight, previous_key))
    if not votes:
        return None, -1.0, -1.0

    bins: dict[int, list[tuple[int, float, str]]] = {}
    for vote in votes:
        bins.setdefault(int(round(vote[0] / 5.0) * 5), []).append(vote)
    ranked = sorted(
        bins.items(),
        key=lambda item: (len({vote[2] for vote in item[1]}), sum(vote[1] for vote in item[1])),
        reverse=True,
    )
    best_bin, best_votes = ranked[0]
    best_distinct = len({vote[2] for vote in best_votes})
    best_weight = sum(vote[1] for vote in best_votes)
    second_weight = sum(vote[1] for _, votes_in_bin in ranked[1:] for vote in votes_in_bin)
    if best_distinct < 2 or best_weight < 0.8:
        return None, best_weight, second_weight
    weighted_shift = int(round(np.average([vote[0] for vote in best_votes], weights=[vote[1] for vote in best_votes])))
    return weighted_shift, best_weight, second_weight


def estimate_vertical_shift(
    previous: np.ndarray,
    current: np.ndarray,
    previous_rows: list[dict] | None = None,
    current_rows: list[dict] | None = None,
) -> tuple[int | None, float, float]:
    """Return previous_y - current_y for overlapping captures.

    The offset is calculated per pair.  It is deliberately not a constant:
    wheel settings and CATIA window sizes can change from one project to the
    next.
    """
    first = _registration_feature(previous)
    second = _registration_feature(current)
    height = min(first.shape[0], second.shape[0])
    minimum_overlap = max(120, int(round(height * 0.18)))
    maximum = height - minimum_overlap
    candidates: list[tuple[float, int]] = []
    for shift in range(REGISTRATION_MIN_SHIFT, maximum + 1):
        # Positive shift: current is scrolled down from previous.
        candidates.append((_normalised_overlap(first[shift:height], second[: height - shift]), shift))
        # Negative shift: support capture sequences made while scrolling up.
        candidates.append((_normalised_overlap(first[: height - shift], second[shift:height]), -shift))
    if not candidates:
        return None, -1.0, -1.0
    candidates.sort(reverse=True)
    best_score, best_shift = candidates[0]
    second_score = next(
        (score for score, shift in candidates[1:] if abs(shift - best_shift) > 10),
        -1.0,
    )
    # A weak or ambiguous peak is explicitly treated as a new segment; this
    # is safer than silently deleting a valid node.
    # Small score gaps are expected on sparse panels.  A visual peak above
    # 0.22 with a 0.015 margin proved stable over the capture sequence; lower
    # quality cases still require the independent text-anchor fallback below.
    if best_score >= 0.22 and best_score - second_score >= 0.015:
        return best_shift, best_score, second_score

    # Visual evidence has priority.  Text anchors are only a safe fallback
    # when visual registration is explicitly ambiguous.
    if previous_rows is not None and current_rows is not None:
        return _estimate_shift_from_anchors(previous_rows, current_rows)
    return None, best_score, second_score


@dataclass
class _NodeEvidence:
    segment: int
    global_y: float
    lane: int
    level: int
    branch_x: int
    observations: list[dict] = field(default_factory=list)

    def accepts(self, observation: dict, y_tolerance: float) -> bool:
        return (
            self.segment == observation["segment"]
            and abs(self.global_y - observation["global_y"]) <= y_tolerance
            and abs(self.lane - observation["lane"]) <= max(5, observation["indent_step"] * 0.30)
        )

    def add(self, observation: dict) -> None:
        self.observations.append(observation)
        self.global_y = float(np.median([item["global_y"] for item in self.observations]))


def _choose_text(observations: list[dict]) -> tuple[dict, float, bool]:
    """Pick a canonical string using real OCR evidence, then agreement.

    ``ocr_alternatives`` is an audit trail, not a set of equally credible
    answers.  Every candidate below therefore carries the score/confidence of
    the actual OCR pass that produced it.
    """
    candidates: list[dict] = []
    # A clipped top/bottom fragment can corroborate geometry but must not win
    # the spelling of a fully visible observation.
    spelling_observations = [item for item in observations if not item.get("partial", False)] or observations
    for observation in spelling_observations:
        evidence = observation.get("candidate_evidence")
        if not evidence:
            evidence = [{
                "text": observation["text"],
                "conf": observation.get("conf", 0.0),
                "score": observation.get("score", 0.0),
                "method": observation.get("method", "full_image"),
            }]
        for item in evidence:
            cleaned = _canonical_catia_ui_label(_clean_text(item.get("text", "")))
            if not cleaned:
                continue
            candidate = dict(observation)
            candidate["text"] = cleaned
            candidate["conf"] = float(item.get("conf", observation.get("conf", 0.0)))
            candidate["method"] = str(item.get("method", observation.get("method", "full_image")))
            candidate["candidate_score"] = float(item.get("score", observation.get("score", 0.0)))
            candidates.append(candidate)
    if not candidates:
        return dict(observations[0]), 0.0, True

    groups: dict[str, list[dict]] = {}
    for candidate in candidates:
        key = _normalise_for_match(candidate["text"])
        groups.setdefault(key, []).append(candidate)

    ranked: list[tuple[float, list[dict]]] = []
    for values in groups.values():
        best = max(item["candidate_score"] for item in values)
        source_support = len({item["image"] for item in values})
        ranked.append((best + min(18.0, 6.0 * (source_support - 1)), values))
    ranked.sort(key=lambda item: item[0], reverse=True)
    winner_score, winner_values = ranked[0]
    winner = max(winner_values, key=lambda item: item["candidate_score"])

    runner_score = ranked[1][0] if len(ranked) > 1 else -math.inf
    image_support = len({item["image"] for item in winner_values})
    final_confidence = min(99.0, float(winner["conf"]) + 5.0 * (image_support - 1))
    compact = re.sub(r"[^A-Za-z0-9]", "", winner["text"])
    reasons: list[str] = []
    if _text_plausibility(winner["text"]) < 0:
        reasons.append("implausible_text")
    fragile_single = (
        image_support == 1
        and winner["conf"] < 75.0
        and not _is_structured_short_identifier(winner["text"])
    )
    if image_support == 1 and (winner["conf"] < 58.0 or fragile_single):
        reasons.append("single_low_confidence")
    if image_support == 1 and winner["conf"] < 85.0:
        # Confidence becomes meaningful only together with evidence count.
        # A 82% one-off reading has already produced plausible-looking false
        # labels in this CATIA panel, whereas the same reading corroborated
        # by the overlap is much safer.  Keep it exportable, but never hide
        # it from the human feedback loop.
        reasons.append("single_unconfirmed_evidence")
    if runner_score > winner_score - 5.0:
        reasons.append("conflicting_candidates")
    # Two identical OCR mistakes can still agree across captures.  Short
    # uppercase labels ending in a common digit-confusion character (YS, ZO,
    # XI...) are retained but explicitly sent to review rather than silently
    # rewritten as a guessed part number.
    if re.fullmatch(r"[A-Za-z][SZOIL]", compact):
        reasons.append("ambiguous_short_code")
    if re.fullmatch(r"[a-z]{1,4}", winner["text"]):
        # A real CATIA code may be uppercase or alphanumeric; a short
        # lowercase token is more often an icon/anti-alias artefact.
        reasons.append("suspicious_short_lowercase")
    if re.match(r"^\d+\s+[A-Za-z]{3,}", winner["text"]):
        reasons.append("possible_icon_prefix")
    if (
        re.search(r"[_-]", winner["text"])
        and winner["conf"] < 85.0
        and any(character.isdigit() for character in winner["text"])
    ):
        # Identifiers often contain visually tiny underscores; do not claim a
        # fully exact file/part number from one uncertain OCR observation.
        reasons.append("identifier_format_uncertain")
    uncertain = bool(reasons)
    winner["uncertainty_reason"] = ";".join(reasons)
    return winner, final_confidence, uncertain


def _merge_observations(observations: list[dict]) -> pd.DataFrame:
    if not observations:
        return pd.DataFrame()
    nodes: list[_NodeEvidence] = []
    for observation in sorted(observations, key=lambda item: (item["segment"], item["global_y"], item["lane"])):
        # OCR row grouping can split one physical CATIA line into an icon
        # artefact and its real label about 10–16 px apart.  Eighteen pixels
        # remains safely below the normal 35–40 px tree-row spacing.
        tolerance = max(18.0, 0.40 * float(observation["height"]))
        matching = [node for node in nodes if node.accepts(observation, tolerance)]
        if not matching:
            # Two distinct CATIA nodes cannot occupy the same physical row.
            # When one screenshot OCRs an icon at the previous indentation,
            # use the aligned y coordinate to attach it as a weak observation
            # of the same node instead of exporting a duplicate.  The final
            # level is selected below from the strongest geometry evidence.
            matching = [
                node for node in nodes
                if node.segment == observation["segment"]
                and abs(node.global_y - observation["global_y"]) <= tolerance
            ]
        if matching:
            min(matching, key=lambda node: abs(node.global_y - observation["global_y"])).add(observation)
        else:
            node = _NodeEvidence(
                segment=observation["segment"],
                global_y=float(observation["global_y"]),
                lane=int(observation["lane"]),
                level=int(observation["level"]),
                branch_x=int(observation["branch_x"]),
            )
            node.add(observation)
            nodes.append(node)

    records = []
    for node in sorted(nodes, key=lambda item: (item.segment, item.global_y, item.level)):
        # A fragment clipped at the very top/bottom has no independent tree
        # identity.  It can support a normal node, but cannot create one.
        if all(bool(observation.get("partial", False)) for observation in node.observations):
            continue
        winner, final_confidence, uncertain = _choose_text(node.observations)
        def geometry_score(observation: dict) -> float:
            return (
                float(observation.get("conf", 0.0))
                + _text_plausibility(str(observation.get("text", "")))
                # A crop lane is useful structural evidence, but it cannot
                # override a clearly better full-image label merely because
                # the crop itself began on an icon ("LA Vues").
                + (3.0 if str(observation.get("method", "")).startswith("line_") else 0.0)
            )

        geometry_winner = max(node.observations, key=geometry_score)
        structural_observations = [
            observation for observation in node.observations if not observation.get("partial", False)
        ] or node.observations
        lanes_seen = {int(observation["lane"]) for observation in structural_observations}
        winner_norm = _normalise_for_match(str(geometry_winner.get("text", "")))
        significant_lane_conflict = False
        for observation in structural_observations:
            if int(observation["lane"]) == int(geometry_winner["lane"]):
                continue
            other_norm = _normalise_for_match(str(observation.get("text", "")))
            # A prefix such as "LA Vues" is the icon plus the same genuine
            # label, not independent hierarchy evidence.
            contaminated_prefix = bool(other_norm and winner_norm) and (
                other_norm.endswith(winner_norm) or winner_norm.endswith(other_norm)
            )
            if not contaminated_prefix and geometry_score(observation) >= geometry_score(geometry_winner) - 7.0:
                significant_lane_conflict = True
                break
        review_reason = winner.get("uncertainty_reason", "") if uncertain else ""
        if len(lanes_seen) > 1 and significant_lane_conflict:
            uncertain = True
            review_reason = ";".join(filter(None, [review_reason, "lane_conflict"]))
        alternatives = list(dict.fromkeys(
            option for observation in node.observations for option in observation.get("alternatives", []) if option
        ))
        records.append({
            "image": winner["image"],
            "capture_index": winner["capture_index"],
            "text": winner["text"],
            "left": int(winner["left"]),
            "top": int(winner["top"]),
            "width": int(winner["width"]),
            "height": int(winner["height"]),
            "conf": round(final_confidence, 3),
            "ocr_method": winner["method"],
            "ocr_alternatives": " | ".join(alternatives),
            "observations_count": len(node.observations),
            "global_y": round(node.global_y, 2),
            "branch_x": int(geometry_winner["branch_x"]),
            "level": int(geometry_winner["level"]),
            "review_needed": bool(uncertain),
            "review_reason": review_reason or "ambiguous_or_low_evidence" if uncertain else "",
            "partial_observations": sum(bool(observation.get("partial", False)) for observation in node.observations),
            "segment": node.segment,
        })
    return pd.DataFrame(records)


def write_review_queue(dataframe: pd.DataFrame, review_dir: Path = REVIEW_DIR) -> None:
    """Write the final, human-readable review queue.

    The queue may be written once immediately after OCR and again after tree
    reconstruction.  The latter is important: hierarchy or local-sequence
    inference can legitimately create a new review item after OCR has ended.
    """
    review_dir = Path(review_dir)
    review_dir.mkdir(parents=True, exist_ok=True)
    columns = [
        "node_id", "image", "capture_index", "top", "global_y", "text",
        "conf", "ocr_alternatives", "ocr_text_before_sequence", "sequence_inferred",
        "review_reason", "review_crop", "corrected_text",
        "validated_level", "validated_parent_id",
    ]
    review = dataframe[dataframe["review_needed"]].copy() if not dataframe.empty else dataframe.copy()
    for column in ("corrected_text", "validated_level", "validated_parent_id"):
        review[column] = ""
    review = review.reindex(columns=columns)
    review.to_csv(review_dir / "ocr_review.csv", index=False, encoding="utf-8-sig")

    # The correction file is deliberately an annotation dataset.  It is not
    # blindly converted into global string replacements for future parts.
    instructions = review_dir / "README_feedback.txt"
    instructions.write_text(
        "ocr_review.csv is a read-only queue for this run. Edit the durable "
        "corrections.csv template generated beside it: fill corrected_text and, when "
        "needed, validated_level / validated_parent_id. The template contains a digest "
        "of this exact capture set, so corrections are never silently reused for a "
        "different CATIA part. Keep the screenshots: applied corrections can then be "
        "exported as labelled OCR training samples.\n",
        encoding="utf-8",
    )


def save_review_crops(
    dataframe: pd.DataFrame,
    capture_paths: list[Path],
    review_dir: Path = REVIEW_DIR,
) -> pd.DataFrame:
    """Save visual evidence for review/training without touching source images."""
    result = dataframe.copy()
    result["review_crop"] = ""
    if result.empty or "review_needed" not in result:
        return result
    review_dir = Path(review_dir)
    crop_dir = review_dir / "crops"
    crop_dir.mkdir(parents=True, exist_ok=True)
    paths_by_name = {Path(path).name: Path(path) for path in capture_paths}
    cache: dict[str, np.ndarray] = {}

    for row_index, row in result[result["review_needed"]].iterrows():
        image_name = str(row.get("image", ""))
        source_path = paths_by_name.get(image_name)
        if source_path is None:
            continue
        if image_name not in cache:
            cache[image_name] = cv2.imread(str(source_path))
        image = cache[image_name]
        if image is None:
            continue
        height, width = image.shape[:2]
        left = max(0, int(row.get("left", 0)) - 36)
        top = max(0, int(row.get("top", 0)) - 10)
        right = min(width, int(row.get("left", 0)) + int(row.get("width", 0)) + 16)
        bottom = min(height, int(row.get("top", 0)) + int(row.get("height", 0)) + 12)
        if right - left < 2 or bottom - top < 2:
            continue
        filename = f"{row['node_id']}_{Path(image_name).stem}_y{int(round(float(row.get('global_y', 0))))}.png"
        destination = crop_dir / filename
        if cv2.imwrite(str(destination), image[top:bottom, left:right]):
            result.at[row_index, "review_crop"] = str(Path("crops") / filename)
    return result


def extract_capture_paths(paths: list[Path]) -> pd.DataFrame:
    """Extract an explicit ordered capture list (also useful for validation)."""
    if not paths:
        return pd.DataFrame()
    paths = sorted((Path(path) for path in paths), key=_natural_key)
    all_observations: list[dict] = []
    capture_metadata: list[dict] = []
    previous_image: np.ndarray | None = None
    previous_rows: list[dict] | None = None
    previous_path: Path | None = None
    offset = 0.0
    segment = 0

    for capture_index, path in enumerate(paths):
        rows, image = read_capture(path, capture_index)
        if not rows:
            # CATIA's scrolling routine can emit one or more empty frames at
            # the end.  They must not create an artificial segment or change
            # the offset of the last real capture.
            print(f"[SKIP] Empty capture: {path.name}")
            capture_metadata.append({
                "image": path.name,
                "capture_index": capture_index,
                "status": "empty",
                "offset": None,
                "segment": segment,
            })
            continue

        registration: dict[str, object] = {"status": "first", "shift": None, "score": None, "next_score": None}
        if previous_image is not None:
            shift, score, second_score = estimate_vertical_shift(previous_image, image, previous_rows, rows)
            if shift is None:
                segment += 1
                offset = 0.0
                print(f"[WARN] Ambiguous overlap: {path.name} (peak={score:.3f}, next={second_score:.3f})")
                registration = {"status": "new_segment", "shift": None, "score": score, "next_score": second_score}
            else:
                offset += float(shift)
                print(f"[OK] Overlap {previous_path.name} -> {path.name}: {shift}px (score={score:.3f})")
                registration = {"status": "registered", "shift": shift, "score": score, "next_score": second_score}
        capture_metadata.append({
            "image": path.name,
            "capture_index": capture_index,
            "offset": offset,
            "segment": segment,
            "rows": len(rows),
            **registration,
        })
        for row in rows:
            # The full-image row centre is a stable physical anchor.  OCR
            # boxes inside a crop can move upward/downward when a watermark
            # is read as a glyph, so ``top`` is kept for audit but must not
            # control overlap merging or the parent-child order.
            row["global_y"] = float(row["row_center"]) + offset
            row["segment"] = segment
            all_observations.append(row)
        previous_image = image
        previous_rows = rows
        previous_path = path

    result = _merge_observations(all_observations)
    if result.empty:
        return result
    result = result.sort_values(["segment", "global_y", "level", "left"]).reset_index(drop=True)
    result.insert(0, "node_id", [f"N{index:05d}" for index in range(1, len(result) + 1)])
    result = save_review_crops(result, paths)
    write_review_queue(result)

    with (REVIEW_DIR / "extraction_metadata.json").open("w", encoding="utf-8") as handle:
        json.dump({"captures": capture_metadata, "capture_count": len(paths)}, handle, indent=2)
    return result


def extract_all_captures(capture_folder: Path) -> pd.DataFrame:
    """Extract, register and merge every PNG screenshot in a capture folder."""
    paths = sorted(Path(capture_folder).glob("*.png"), key=_natural_key)
    return extract_capture_paths(paths)
