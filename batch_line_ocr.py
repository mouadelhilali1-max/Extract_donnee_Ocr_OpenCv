"""Batch OCR for already-localised CATIA tree rows.

This module deliberately does *not* decide which OCR spelling is correct.  It
only turns a number of small row crops into one Tesseract invocation and maps
the resulting boxes back to the coordinates of the original screenshot.

The intended caller has already inferred the text lane and the row centre.  A
caller may submit several variants for the same ``id`` (for example ``raw``
and ``code_110``); all resulting candidates are returned so that the tree
pipeline can compare them with overlap evidence rather than trusting a single
confidence score.

Example
-------
>>> requests = [
...     {"id": "row-42", "center_y": 120, "crop_x": 147, "variant": "raw"},
...     {"id": "row-42", "center_y": 120, "crop_x": 147, "variant": "code_110"},
... ]
>>> candidates = batch_line_ocr(image, requests)
>>> candidates["row-42"]  # one or more alternatives; no winner is selected
[{"text": "Y3", "conf": 86.1, ...}, ...]
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
from typing import Any, Iterable, Literal, Mapping

import cv2
import numpy as np
import pandas as pd
import pytesseract


Variant = Literal["raw", "binary", "code"]


@dataclass(frozen=True)
class LineCropRequest:
    """One localised CATIA label crop.

    ``center_y`` and ``crop_x`` are coordinates in the *source* OpenCV image.
    ``variant`` may also contain an explicit threshold, e.g. ``binary_130``
    or ``code_110``.  This is useful for a caller that wants independent OCR
    evidence from multiple preprocessing variants.

    ``top_padding`` and ``bottom_padding`` are intentionally asymmetric:
    CATIA's text normally sits slightly above its row centre in screenshots.
    They can be overridden per request without changing the batch API.
    """

    id: str
    center_y: float
    crop_x: int
    variant: str = "raw"
    right: int | None = None
    top_padding: int = 16
    bottom_padding: int = 22


@dataclass(frozen=True)
class _PreparedCrop:
    request: LineCropRequest
    image: np.ndarray
    source_left: int
    source_top: int
    source_right: int
    source_bottom: int
    stack_top: int


def _default_language() -> tuple[str, str]:
    """Use the project setting when available, while remaining reusable."""

    try:
        from ocr_config import OCR_LANGUAGE, TESSERACT_PATH, TESSDATA_CONFIG  # local project config

        if TESSERACT_PATH and Path(TESSERACT_PATH).exists():
            pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
        return str(OCR_LANGUAGE), str(TESSDATA_CONFIG)
    except (ImportError, OSError):
        return "eng", ""


def _coerce_request(value: LineCropRequest | Mapping[str, Any]) -> LineCropRequest:
    """Accept the public dataclass and simple ``dict`` requests alike."""

    if isinstance(value, LineCropRequest):
        return value
    if not isinstance(value, Mapping):
        raise TypeError("Each crop request must be a LineCropRequest or mapping")
    request_id = value.get("id", value.get("request_id"))
    if request_id is None:
        raise ValueError("A crop request needs an 'id'")
    try:
        return LineCropRequest(
            id=str(request_id),
            center_y=float(value["center_y"]),
            crop_x=int(value["crop_x"]),
            variant=str(value.get("variant", "raw")),
            right=None if value.get("right") is None else int(value["right"]),
            top_padding=int(value.get("top_padding", 16)),
            bottom_padding=int(value.get("bottom_padding", 22)),
        )
    except KeyError as error:
        raise ValueError(f"Crop request is missing {error.args[0]!r}") from error


def _as_bgr(image: np.ndarray) -> np.ndarray:
    if image.ndim == 2:
        return cv2.cvtColor(image, cv2.COLOR_GRAY2BGR)
    if image.ndim == 3 and image.shape[2] == 4:
        return cv2.cvtColor(image, cv2.COLOR_BGRA2BGR)
    if image.ndim == 3 and image.shape[2] == 3:
        return image
    raise ValueError("image must be a grayscale, BGR, or BGRA OpenCV image")


def _variant_threshold(variant: str) -> tuple[str, int | None]:
    """Parse raw / binary / code variants without silently changing them."""

    name = str(variant).strip().casefold() or "raw"
    match = re.fullmatch(r"(raw|binary|code)(?:_(\d{1,3}))?", name)
    if not match:
        raise ValueError(
            "variant must be raw, binary, code, binary_<0..255>, or code_<0..255>"
        )
    kind, explicit = match.groups()
    if kind == "raw":
        return kind, None
    threshold = int(explicit) if explicit is not None else (130 if kind == "binary" else 110)
    if not 0 <= threshold <= 255:
        raise ValueError("OCR threshold must be between 0 and 255")
    return kind, threshold


def _prepare_variant(crop: np.ndarray, variant: str) -> np.ndarray:
    """Apply only the explicit variant requested by the caller."""

    kind, threshold = _variant_threshold(variant)
    if kind == "raw":
        return _as_bgr(crop)
    gray = cv2.cvtColor(_as_bgr(crop), cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, int(threshold), 255, cv2.THRESH_BINARY)
    return cv2.cvtColor(binary, cv2.COLOR_GRAY2BGR)


def _crop_for_request(image: np.ndarray, request: LineCropRequest, right_margin: int) -> tuple[np.ndarray, int, int, int, int] | None:
    """Return the source crop and its exact source-image bounds."""

    height, width = image.shape[:2]
    if request.top_padding < 0 or request.bottom_padding < 0:
        raise ValueError("top_padding and bottom_padding must be non-negative")
    left = max(0, min(width, int(request.crop_x)))
    requested_right = width - int(right_margin) if request.right is None else int(request.right)
    right = max(left, min(width, requested_right))
    top = max(0, min(height, int(round(request.center_y - request.top_padding))))
    bottom = max(top, min(height, int(round(request.center_y + request.bottom_padding))))
    if right - left < 2 or bottom - top < 2:
        return None
    return image[top:bottom, left:right].copy(), left, top, right, bottom


def _background_colour(image: np.ndarray) -> tuple[int, int, int]:
    """A dark separator/background that does not look like a glyph to OCR."""

    if image.size == 0:
        return (0, 0, 0)
    corner_h = max(1, min(8, image.shape[0]))
    corner_w = max(1, min(8, image.shape[1]))
    colour = np.median(image[:corner_h, :corner_w].reshape(-1, 3), axis=0)
    return tuple(int(value) for value in colour)


def _stack_prepared_crops(
    image: np.ndarray,
    requests: Iterable[LineCropRequest],
    *,
    right_margin: int,
    separator_height: int,
) -> tuple[np.ndarray | None, list[_PreparedCrop]]:
    """Stack local crops with blank guard bands, retaining each y mapping."""

    source = _as_bgr(image)
    prepared_unstacked: list[tuple[LineCropRequest, np.ndarray, int, int, int, int]] = []
    for request in requests:
        bounds = _crop_for_request(source, request, right_margin)
        if bounds is None:
            continue
        crop, left, top, right, bottom = bounds
        prepared_unstacked.append(
            (request, _prepare_variant(crop, request.variant), left, top, right, bottom)
        )
    if not prepared_unstacked:
        return None, []

    max_width = max(item[1].shape[1] for item in prepared_unstacked)
    total_height = sum(item[1].shape[0] for item in prepared_unstacked)
    total_height += max(0, len(prepared_unstacked) - 1) * separator_height
    stack = np.empty((total_height, max_width, 3), dtype=np.uint8)
    stack[:, :] = _background_colour(source)

    prepared: list[_PreparedCrop] = []
    stack_top = 0
    for index, (request, crop, left, top, right, bottom) in enumerate(prepared_unstacked):
        crop_h, crop_w = crop.shape[:2]
        # Copy pixels without resizing: OCR coordinates can therefore be
        # mapped back by a simple offset, with no rounding error.
        stack[stack_top : stack_top + crop_h, :crop_w] = crop
        prepared.append(
            _PreparedCrop(
                request=request,
                image=crop,
                source_left=left,
                source_top=top,
                source_right=right,
                source_bottom=bottom,
                stack_top=stack_top,
            )
        )
        stack_top += crop_h
        if index < len(prepared_unstacked) - 1:
            stack_top += separator_height
    return stack, prepared


def _request_for_stack_y(prepared: list[_PreparedCrop], stack_y: float) -> _PreparedCrop | None:
    """Map a Tesseract box centre back to the crop that owns it."""

    for item in prepared:
        crop_height = item.image.shape[0]
        if item.stack_top <= stack_y < item.stack_top + crop_height:
            return item
    return None


def _is_edge_noise(word: str, left: int, width: int, crop_width: int) -> bool:
    """Discard only non-text marks touching a crop edge.

    This avoids inventing text clean-up rules.  A valid label beginning at the
    crop edge (``EXTRACT`` for example) is kept because it contains letters.
    """

    compact = re.sub(r"[^A-Za-z0-9_]", "", str(word))
    touches_edge = left <= 1 or left + width >= crop_width - 1
    return touches_edge and not compact


def _trim_unbalanced_crop_edge_noise(word: str, left: int, width: int, crop_width: int) -> str:
    """Trim only connector-like punctuation at the *outside* crop edges.

    ``)Z2`` is a common CATIA crop artefact: the closing branch shape sits at
    x=0, immediately before the real code.  We remove a leading closing mark
    and a trailing opening mark only when Tesseract's box touches that crop
    edge.  Balanced name punctuation such as ``(XY)`` is untouched.
    """

    value = str(word).replace("\x0c", "")
    if left <= 1:
        value = re.sub(r"^[\]\)}|]+(?=[A-Za-z0-9_])", "", value)
    if left + width >= crop_width - 1:
        value = re.sub(r"(?<=[A-Za-z0-9_])[\[\({|]+$", "", value)
    return value


def _clean_joined_text(words: list[dict[str, Any]]) -> str:
    """Remove OCR control whitespace only; do not autocorrect CATIA labels."""

    value = " ".join(str(word["text"]).replace("\x0c", " ") for word in words)
    return re.sub(r"\s+", " ", value).strip()


def _candidate_from_words(item: _PreparedCrop, words: list[dict[str, Any]], line_key: tuple[int, int, int]) -> dict[str, Any] | None:
    kept = [
        word.copy()
        for word in words
        if not _is_edge_noise(word["text"], word["local_left"], word["width"], item.image.shape[1])
    ]
    if not kept:
        return None
    for word in kept:
        word["raw_text"] = word["text"]
        word["text"] = _trim_unbalanced_crop_edge_noise(
            word["text"], word["local_left"], word["width"], item.image.shape[1]
        )
    kept.sort(key=lambda value: (value["local_left"], value["local_top"]))
    text = _clean_joined_text(kept)
    if not text:
        return None
    left = min(word["left"] for word in kept)
    top = min(word["top"] for word in kept)
    right = max(word["left"] + word["width"] for word in kept)
    bottom = max(word["top"] + word["height"] for word in kept)
    confidence_values = [float(word["conf"]) for word in kept if np.isfinite(word["conf"])]
    return {
        "id": item.request.id,
        "text": text,
        # Mean confidence describes this OCR observation; it is deliberately
        # not a validity threshold and is never used here to select a winner.
        "conf": float(np.mean(confidence_values)) if confidence_values else -1.0,
        "left": int(left),
        "top": int(top),
        "width": int(right - left),
        "height": int(bottom - top),
        "method": item.request.variant,
        "ocr_layout": "batch_psm6",
        "line_key": line_key,
        "crop_left": item.source_left,
        "crop_top": item.source_top,
        "crop_right": item.source_right,
        "crop_bottom": item.source_bottom,
        "words": [
            {
                "text": word["text"],
                "raw_text": word["raw_text"],
                "conf": float(word["conf"]),
                "left": int(word["left"]),
                "top": int(word["top"]),
                "width": int(word["width"]),
                "height": int(word["height"]),
            }
            for word in kept
        ],
    }


def batch_line_ocr(
    image: np.ndarray,
    requests: Iterable[LineCropRequest | Mapping[str, Any]],
    *,
    language: str | None = None,
    psm: int = 6,
    right_margin: int = 12,
    separator_height: int = 10,
    extra_config: str = "",
) -> dict[str, list[dict[str, Any]]]:
    """OCR all requested line crops using exactly one Tesseract call.

    Parameters
    ----------
    image:
        Source OpenCV image in grayscale, BGR or BGRA format.
    requests:
        ``LineCropRequest`` instances or mappings containing at least ``id``,
        ``center_y`` and ``crop_x``.  Submit multiple variants under the same
        id to obtain multiple independent candidates.
    language, psm, extra_config:
        Passed to Tesseract.  The default ``psm=6`` treats the stacked crops as
        a compact text block, avoiding one process launch per row.

    Returns
    -------
    dict[str, list[dict]]
        Every supplied id is present.  Each list contains every text line OCR
        returned for that crop; the function does not rank, merge or correct
        alternatives.  ``left/top/width/height`` are source-image coordinates.
    """

    source = _as_bgr(np.asarray(image))
    normalised = [_coerce_request(request) for request in requests]
    result: dict[str, list[dict[str, Any]]] = {request.id: [] for request in normalised}
    if not normalised:
        return result
    if separator_height < 1:
        raise ValueError("separator_height must be at least 1")
    stack, prepared = _stack_prepared_crops(
        source,
        normalised,
        right_margin=right_margin,
        separator_height=separator_height,
    )
    if stack is None or not prepared:
        return result

    if language is None:
        lang, tessdata_config = _default_language()
    else:
        lang, tessdata_config = str(language), ""
    config = f"{tessdata_config} --oem 3 --psm {int(psm)} {extra_config}".strip()
    # This is intentionally the only pytesseract invocation in this function.
    data = pytesseract.image_to_data(
        stack,
        lang=lang,
        config=config,
        output_type=pytesseract.Output.DATAFRAME,
    )
    if data is None or len(data) == 0:
        return result
    frame = pd.DataFrame(data).copy()
    needed = {"text", "conf", "left", "top", "width", "height"}
    if not needed.issubset(frame.columns):
        return result
    frame["text"] = frame["text"].fillna("").astype(str).str.strip()
    frame["conf"] = pd.to_numeric(frame["conf"], errors="coerce").fillna(-1.0)
    frame = frame[frame["text"] != ""]
    if frame.empty:
        return result
    for column in ("block_num", "par_num", "line_num"):
        if column not in frame.columns:
            frame[column] = 0

    grouped: dict[tuple[str, tuple[int, int, int]], tuple[_PreparedCrop, list[dict[str, Any]]]] = {}
    for _, row in frame.iterrows():
        local_top_in_stack = int(row["top"])
        box_height = max(1, int(row["height"]))
        item = _request_for_stack_y(prepared, local_top_in_stack + box_height / 2.0)
        if item is None:
            # The box belongs to the blank separator, so it must never be
            # attributed to either neighbouring CATIA row.
            continue
        local_left = int(row["left"])
        local_top = local_top_in_stack - item.stack_top
        word = {
            "text": str(row["text"]),
            "conf": float(row["conf"]),
            "local_left": local_left,
            "local_top": local_top,
            "left": item.source_left + local_left,
            "top": item.source_top + local_top,
            "width": int(row["width"]),
            "height": box_height,
        }
        line_key = (int(row["block_num"]), int(row["par_num"]), int(row["line_num"]))
        group_key = (item.request.id, line_key)
        if group_key not in grouped:
            grouped[group_key] = (item, [])
        grouped[group_key][1].append(word)

    for (_, line_key), (item, words) in grouped.items():
        candidate = _candidate_from_words(item, words, line_key)
        if candidate is not None:
            result[item.request.id].append(candidate)
    for candidates in result.values():
        candidates.sort(key=lambda candidate: (candidate["top"], candidate["left"], candidate["text"]))
    return result


__all__ = ["LineCropRequest", "batch_line_ocr"]
