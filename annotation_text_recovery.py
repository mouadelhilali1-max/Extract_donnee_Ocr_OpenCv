"""Evidence-based recovery for CATIA annotation OCR labels.

The CATIA tree repeats labels in overlapping screenshots.  Tesseract can read
an icon or a neighbouring line instead of a short identifier, even when a
correct reading exists among the OCR alternatives.  This module uses only that
recorded evidence and the local CATIA section structure; it never invents a
free-form part name.
"""

from __future__ import annotations

from collections import defaultdict
import re
import unicodedata

import pandas as pd

from annotation_visual_scope import normalise_label


# Standard CATIA tolerance/note IDs, e.g. 01B01 or 06A02.  The first/last two
# characters must be digits; the middle character is a letter.  OCR commonly
# confuses O/0, G/6, S/5, and 4/A in these narrow labels.
CODE_TOKEN = re.compile(
    r"(?<![A-Z0-9])([0-9OQDLIZSGBT]{2}[A-Z4][0-9OQDLIZSGBT]{2})(?![A-Z0-9])",
    re.IGNORECASE,
)
STRICT_CODE_TOKEN = re.compile(r"(?<![A-Z0-9])(\d{2}[A-Z]\d{2})(?![A-Z0-9])", re.IGNORECASE)

# Short labels displayed twice by CATIA, for example Z5 (Z5), X24 (X24), or
# Y1 (Y1).  The alternatives may contain 75 for Z5, ZA for Z4, or K24 for X24.
SHORT_TOKEN = re.compile(
    r"(?<![A-Z0-9])([XYZKVT7][0-9OQDLIZSGBTA]{1,2})(?![A-Z0-9])",
    re.IGNORECASE,
)

_DIGIT_CONFUSIONS = {
    "O": "0",
    "Q": "0",
    "D": "0",
    "I": "1",
    "L": "1",
    "Z": "2",
    "S": "5",
    "G": "6",
    "T": "7",
    "B": "8",
    "A": "4",
}
_SHORT_PREFIX_CONFUSIONS = {"K": "X", "V": "Y", "7": "Z", "T": "Z"}
_SEMANTIC_IDENTIFIER_PARENTS = {"notes", "tolerancegeometrique"}


def _string(value: object) -> str:
    return "" if value is None or pd.isna(value) else str(value).strip()


def _ascii_upper(value: object) -> str:
    text = _string(value).replace("â€™", "'").replace("â€˜", "'")
    plain = "".join(
        character
        for character in unicodedata.normalize("NFKD", text)
        if not unicodedata.combining(character)
    )
    return plain.upper()


def _normalise_code(value: str) -> str | None:
    """Turn one OCR-like code token into a strict ``00A00`` form."""
    compact = re.sub(r"[^A-Z0-9]", "", _ascii_upper(value))
    if len(compact) != 5:
        return None
    digits = "".join(
        _DIGIT_CONFUSIONS.get(character, character)
        for character in (compact[0], compact[1], compact[3], compact[4])
    )
    letter = "A" if compact[2] == "4" else compact[2]
    candidate = f"{digits[:2]}{letter}{digits[2:]}"
    return candidate if re.fullmatch(r"\d{2}[A-Z]\d{2}", candidate) else None


def _codes_in(value: object) -> list[str]:
    candidates: list[str] = []
    for match in CODE_TOKEN.finditer(_ascii_upper(value)):
        raw = match.group(1)
        # A word such as ``STELL`` can accidentally fit the permissive OCR
        # confusion alphabet (S/T/L/L).  A real CATIA ID always retains at
        # least one visible decimal in one of the OCR alternatives.  Reject
        # all-letter pseudo-codes before applying the confusion map.
        if not any(character.isdigit() for character in raw):
            continue
        candidate = _normalise_code(raw)
        if candidate:
            candidates.append(candidate)
    return candidates


def _contains_strict_code(value: object) -> bool:
    """Whether *value* visibly contains an unambiguous CATIA code."""
    return bool(
        re.search(
            r"(?<![A-Z0-9])\d{2}[A-Z]\d{2}(?![A-Z0-9])",
            _ascii_upper(value),
        )
    )


def _strict_codes_in(value: object) -> list[str]:
    """Return only codes that OCR printed in the exact CATIA grammar."""
    return [match.group(1).upper() for match in STRICT_CODE_TOKEN.finditer(_ascii_upper(value))]


def _has_strict_primary_code(value: object) -> bool:
    """Whether the first CATIA code display, before ``(``, is unambiguous.

    CATIA repeats a code inside parentheses.  The repeat may be clean even
    when the main visible token is corrupted (``92401 (02A01...)``).  Checking
    the entire string would incorrectly call that row "strict" because of the
    code in parentheses, so the head has to be considered separately.
    """
    head = _ascii_upper(_string(value).split("(", 1)[0])
    return bool(re.search(r"(?<![A-Z0-9])\d{2}[A-Z]\d{2}(?![A-Z0-9])", head))


def _normalise_short_identifier(value: str) -> str | None:
    compact = re.sub(r"[^A-Z0-9]", "", _ascii_upper(value))
    if len(compact) not in {2, 3}:
        return None
    prefix = _SHORT_PREFIX_CONFUSIONS.get(compact[0], compact[0])
    if prefix not in {"X", "Y", "Z"}:
        return None
    suffix = "".join(_DIGIT_CONFUSIONS.get(character, character) for character in compact[1:])
    candidate = f"{prefix}{suffix}"
    return candidate if re.fullmatch(r"[XYZ]\d{1,2}", candidate) else None


def _short_identifiers_in(value: object) -> list[str]:
    text = _ascii_upper(value)
    result: list[str] = []
    for match in SHORT_TOKEN.finditer(text):
        candidate = _normalise_short_identifier(match.group(1))
        if candidate:
            result.append(candidate)
    return result


def _alternatives(row: pd.Series) -> list[str]:
    """Return de-duplicated OCR readings; `` | `` is the export separator."""
    values = [_string(row.get("text", ""))]
    raw = _string(row.get("ocr_alternatives", ""))
    if raw:
        values.extend(part.strip() for part in re.split(r"\s+\|\s+", raw) if part.strip())
    return list(dict.fromkeys(value for value in values if value))


def _code_evidence(row: pd.Series) -> dict[str, float]:
    """Score code readings while preserving their position in a CATIA label.

    A code normally appears twice: once before ``(`` and once in the clipped
    repeat inside parentheses.  A global "this option contains a strict code"
    bonus is unsafe because it gives the same reward to a bad head token such
    as ``92401`` and to the clean repeated ``02A01``.  Score each occurrence
    in its own zone instead.
    """
    scores: dict[str, float] = defaultdict(float)
    for option_index, option in enumerate(_alternatives(row)):
        candidates = _codes_in(option)
        if not candidates:
            continue
        before, separator, after = option.partition("(")
        leading = _codes_in(before)
        inside = _codes_in(after) if separator else []
        strict_leading = set(_strict_codes_in(before))
        strict_inside = set(_strict_codes_in(after)) if separator else set()
        for candidate in dict.fromkeys(candidates):
            # Every independent OCR option counts.  A direct repeat in
            # parentheses is substantially stronger than a token accidentally
            # read from an icon or an adjacent row.
            score = 1.0
            if candidate in leading:
                score += 0.75
            if candidate in inside:
                score += 2.0
            if candidate in strict_leading:
                score += 3.0
            if candidate in strict_inside:
                score += 5.0
            if candidate in leading and candidate in inside:
                score += 3.0
            if option_index == 0:
                score += 0.75
            scores[candidate] += score
    return dict(scores)


def _first_code(value: object) -> str | None:
    """Return the first canonical CATIA code visible in *value*."""
    codes = _codes_in(value)
    return codes[0] if codes else None


def _looks_like_code_display(value: object) -> bool:
    """Guard code recovery against words which happen to fit the OCR alphabet.

    CATIA tolerance rows are short labels (sometimes prefixed by ``GE`` or a
    stray icon reading).  Long Notes strings such as ``Designations (TOL3D...)``
    are not code rows even though ``TOL3D`` can be normalised to a five-character
    code by the OCR confusion map.
    """
    plain = _ascii_upper(value)
    before = plain.split("(", 1)[0].strip()
    compact = re.sub(r"[^A-Z0-9]", "", before)
    return bool(_codes_in(value)) and len(compact) <= 9


def _canonical_code_display(value: object, code: str) -> str:
    """Render a recovered code while retaining CATIA's visible truncation.

    OCR can prepend ``GE`` from an icon and can independently misread the
    duplicate code in parentheses.  The logical CATIA label is the canonical
    code; when CATIA visibly repeats it in parentheses, keep that presentation
    but make both copies agree.
    """
    original = _string(value)
    if "(" not in original:
        return code
    tail = original[original.find("(") :]
    # Replace the first compact token after the opening parenthesis.  It is a
    # clipped repeat of the same code in CATIA, not a separate child label.
    tail = re.sub(r"(?<=\()[A-Z0-9]{3,8}", code, tail, count=1, flags=re.IGNORECASE)
    return f"{code} {tail.lstrip()}"


def _short_evidence(row: pd.Series) -> dict[str, float]:
    """Score short X/Y/Z IDs that are directly visible in OCR alternatives."""
    scores: dict[str, float] = defaultdict(float)
    for option_index, option in enumerate(_alternatives(row)):
        before, separator, after = option.partition("(")
        leading = _short_identifiers_in(before)
        inside = _short_identifiers_in(after) if separator else []
        for candidate in set(leading + inside):
            score = 0.0
            if candidate in leading:
                score += 1.0
            if candidate in inside:
                score += 2.0
            if candidate in leading and candidate in inside:
                score += 3.0
            if option_index == 0:
                score += 0.5
            scores[candidate] += score
    return dict(scores)


def _first_short_identifier(value: object) -> str | None:
    identifiers = _short_identifiers_in(value)
    return identifiers[0] if identifiers else None


def _is_clean_short_display(value: object, identifier: str) -> bool:
    escaped = re.escape(identifier)
    return bool(re.fullmatch(rf"\s*{escaped}\s*(?:\(\s*{escaped}\s*\))?\s*", _ascii_upper(value)))


def _looks_like_short_display(value: object) -> bool:
    """Guard against short-ID fragments found inside ordinary note text.

    OCR alternatives for long labels frequently contain an isolated ``Z7`` or
    ``X31`` read from a neighbouring row.  Such a fragment must not become a
    new node in the Notes sequence.  A genuine short-ID row has a compact
    first token (possibly one of the common OCR prefix confusions) before the
    optional repeated parenthesised token.
    """
    plain = _ascii_upper(value)
    first = plain.split("(", 1)[0].strip()
    first = re.sub(r"\s+", "", first)
    return bool(re.fullmatch(r"[XYZKVT7][0-9OQDLIZSGBTA]{1,2}", first))


def _render_short_identifier(identifier: str, original: object) -> str:
    # CATIA repeats these labels in parentheses. Preserve a no-parentheses
    # display only when the original truly had no parentheses.
    return f"{identifier} ({identifier})" if "(" in _string(original) else identifier


def _append_marker(result: pd.DataFrame, position: int, marker: str) -> None:
    existing = _string(result.at[position, "text_recovery"])
    values = [item for item in existing.split(";") if item]
    if marker not in values:
        values.append(marker)
    result.at[position, "text_recovery"] = ";".join(values)


def _append_review_reason(result: pd.DataFrame, position: int, reason: str) -> None:
    if "review_reason" not in result:
        result["review_reason"] = ""
    existing = _string(result.at[position, "review_reason"])
    values = [item for item in existing.split(";") if item]
    if reason not in values:
        values.append(reason)
    result.at[position, "review_reason"] = ";".join(values)
    if "review_needed" in result:
        result.at[position, "review_needed"] = True


def _identifier_context(row: pd.Series) -> bool:
    return normalise_label(row.get("parent", "")) in _SEMANTIC_IDENTIFIER_PARENTS


def _recover_codes(result: pd.DataFrame) -> None:
    """Recover malformed tolerance/notes codes from their alternative evidence."""
    for position, row in result.iterrows():
        original = _string(row.get("text", ""))
        evidence = _code_evidence(row)
        if not evidence:
            continue
        current_codes = _codes_in(original)
        current = current_codes[0] if current_codes else None
        # CATIA renders these short labels twice, for example
        # ``92401 (02A01...)``.  The second occurrence is a direct reading of
        # the same on-screen row, not a neighbouring OCR alternative.  When
        # the first occurrence is malformed and the repeated code differs,
        # prefer that direct visual evidence.
        inside = original.split("(", 1)[1] if "(" in original else ""
        repeated_codes = _codes_in(inside)
        repeated = repeated_codes[0] if repeated_codes else None
        strict_primary = _has_strict_primary_code(original)
        # Even when both occurrences normalise to the same candidate (for
        # example ``01401 (01401...)`` -> ``01A01``), the parenthesised copy
        # must win over the all-digit head heuristic.  It is direct evidence
        # that this row is the code being displayed, not a neighbour.
        use_repeated_display = bool(current and repeated and not strict_primary)
        best_score = max(evidence.values())
        candidates = [candidate for candidate, score in evidence.items() if score == best_score]
        # Prefer the already visible canonical form only for a real tie.
        # A five-digit, all-numeric OCR token is especially unreliable: the
        # middle ``A`` and often the leading zero have both been mistaken for
        # digits (for example ``93403`` for the visible ``03A03``).  In that
        # case do not automatically favour the normalised reading of the main
        # OCR field; an unambiguous alternative is better evidence.  For every
        # other malformed token, retain the current reading only on a genuine
        # score tie so neighbouring rows cannot overwrite a usable label.
        raw_head = re.sub(r"[^A-Z0-9]", "", _ascii_upper(original.split("(", 1)[0]))
        # Icon OCR can leave a short alphabetic prefix (``GE93403``).  Look
        # at the trailing five-character token so that the same all-digit
        # safeguard still applies in that form.
        numeric_token = re.search(r"\d{5}$", raw_head)
        current_is_all_numeric = bool(numeric_token)
        ranked_candidates = sorted(evidence, key=lambda item: -evidence[item])
        if use_repeated_display:
            candidate = repeated
        elif current_is_all_numeric:
            # Keep the OCR alternative order as a tie-breaker.  The first
            # explicit code after an all-digit main reading is normally the
            # same label seen by the next overlapping frame; alphabetical
            # sorting would arbitrarily select a neighbouring code.
            candidate = next((item for item in ranked_candidates if item != current), ranked_candidates[0])
        else:
            candidate = current if current in candidates else ranked_candidates[0]
        # Only the first displayed token describes this row.  CATIA often
        # repeats the same code in parentheses, so testing the entire label
        # would turn ``OGA02 (06A02...)`` into a ten-character string and
        # prevent recovery in a small (one- or two-row) section.
        display_head = _string(original).split("(", 1)[0]
        compact = re.sub(r"[^A-Z0-9]", "", _ascii_upper(display_head))
        # Ignore an icon prefix when the meaningful token is a trailing run of
        # five digits (``GE93403`` or ``7 93403``).
        shape_compact = numeric_token.group(0) if numeric_token else compact
        looks_malformed = bool(re.fullmatch(r"[A-Z0-9]{2,6}", shape_compact)) and (
            (
                any(character.isalpha() for character in shape_compact)
                and any(character.isdigit() for character in shape_compact)
            )
            # A middle ``A`` is often read as ``4``; an all-digit five
            # character token can therefore still be a malformed 00A00 ID.
            or (len(shape_compact) == 5 and shape_compact.isdigit())
        )
        strict_current = strict_primary
        # Never replace a code which is already printed unambiguously.  The
        # alternatives of a clean row often contain the neighbouring row's
        # code; choosing their highest score would silently change a correct
        # value (for example 05B01 -> 05A01).
        if strict_current:
            # Do not change the code itself, but remove a stray icon reading
            # such as ``GE`` and make CATIA's repeated parenthesised code
            # agree.  This also works when the section has fewer than three
            # code rows, where the sequence decoder below intentionally does
            # not run.
            if current and _identifier_context(row) and _looks_like_code_display(original):
                cleaned = _canonical_code_display(original, current)
                if cleaned != original:
                    result.at[position, "text"] = cleaned
                    _append_marker(result, position, "identifier_normalised")
            continue
        # Recovery is intentionally limited to a compact, code-like display.
        # Long Notes labels may contain strings such as ``TOL3D`` which fit
        # the OCR confusion pattern but are not CATIA IDs.  They must remain
        # untouched even when an alternative gives them a high score.
        if not _identifier_context(row) or not (looks_malformed or _looks_like_code_display(original)):
            continue
        if current and current == candidate:
            # Normalise an OCR spelling such as OGA02 even if no different
            # alternative happened to win.  The renderer also removes any
            # prefix read from the CATIA icon (for example ``GE 06A02``).
            recovered = _canonical_code_display(original, candidate)
            if recovered != original:
                result.at[position, "text"] = recovered
                _append_marker(result, position, "identifier_normalised")
            continue
        if current or len(evidence) == 1 or best_score >= 3.0:
            # This row has already passed the compact-display guard.  Render
            # it from the recovered logical code rather than retaining an OCR
            # prefix such as ``7 `` or ``GE`` that belongs to the CATIA icon.
            recovered = _canonical_code_display(original, candidate)
            if recovered != original:
                result.at[position, "text"] = recovered
                _append_marker(
                    result,
                    position,
                    "identifier_from_repeated_display" if use_repeated_display else "identifier_from_alternatives",
                )


def _repair_capture_labels(result: pd.DataFrame) -> None:
    """Restore known CATIA capture-label syntax only when its context proves it."""
    reference_shape = re.compile(
        r"^(?:REF\s+)?([A-Z])\s*(?:\|\s*)?([A-Z]-[A-Z])\s*(?:\|\s*)?([A-Z]-[A-Z])$"
    )
    for position, row in result.iterrows():
        original = _string(row.get("text", ""))
        plain = _ascii_upper(original)
        parent_key = normalise_label(row.get("parent", ""))
        replacement = ""
        # In the tree, a child displayed only as SPRINGBACK under CPC is the
        # clipped label CPC SPRINGBACK, as verified by the wider screenshots.
        if normalise_label(original) == "springback":
            if parent_key in {"cpc", "captures"}:
                replacement = "CPC SPRINGBACK"
            elif parent_key == "ctf":
                replacement = "CTF SPRINGBACK"
        else:
            match = reference_shape.fullmatch(plain)
            if match:
                replacement = f"REF {match.group(1)}|{match.group(2)}|{match.group(3)}"
        if replacement and replacement != original:
            result.at[position, "text"] = replacement
            _append_marker(result, position, "capture_label_from_structure")


def _repair_reference_labels(result: pd.DataFrame) -> None:
    """Correct the small, deterministic OCR error in reference element IDs."""
    # CATIA prints these labels as ``Elément de référence.1 (A)`` through
    # ``.5 (E)``.  In the narrow font, the final ``5`` is often read as ``S``
    # (or the number is lost altogether).  The letter in parentheses is a
    # reliable structural check, so use it to restore the number.
    pattern = re.compile(r"^ELEMENT DE REFERENCE\.?\s*([0-9S]?)\s*\(([A-E])\)$")
    expected_by_letter = {letter: str(index) for index, letter in enumerate("ABCDE", start=1)}
    for position, row in result.iterrows():
        if normalise_label(row.get("parent", "")) != "references":
            continue
        original = _string(row.get("text", ""))
        match = pattern.fullmatch(_ascii_upper(original))
        if not match:
            continue
        expected = expected_by_letter[match.group(2)]
        if match.group(1) == expected:
            continue
        replacement = f"Elément de référence.{expected} ({match.group(2)})"
        result.at[position, "text"] = replacement
        _append_marker(result, position, "reference_label_normalised")


def _promote_capture_siblings(result: pd.DataFrame) -> None:
    """Undo an overlap merge that made root capture labels children of CPC.

    The same ``CPC SPRINGBACK`` and ``REF A|B-C|D-E`` labels occur in the
    Captures branch and again as root-level siblings.  A geometry-only merge
    can attach the latter observation below the root-level ``CPC`` node.  Once
    the label has been recovered, the surrounding tree structure identifies
    the correct parent unambiguously.
    """
    if "node_id" not in result or "parent_id" not in result:
        return
    by_id = result.set_index("node_id", drop=False).to_dict("index")
    for position, row in result.iterrows():
        parent_id = _string(row.get("parent_id", ""))
        parent = by_id.get(parent_id)
        if not parent or normalise_label(parent.get("text", "")) != "cpc":
            continue
        # Only promote a child of a root-level CPC.  A CPC branch inside
        # Captures is a legitimate parent and must remain untouched.
        if int(parent.get("level", 99) or 99) != 1:
            continue
        label_key = normalise_label(row.get("text", ""))
        if label_key not in {"cpcspringback", "refabcde", "refabcdef"} and not label_key.startswith("ref"):
            continue
        grandparent_id = _string(parent.get("parent_id", ""))
        if not grandparent_id:
            continue
        if int(row.get("level", 99) or 99) == 1 and parent_id == grandparent_id:
            continue
        result.at[position, "parent_id"] = grandparent_id
        result.at[position, "level"] = int(parent.get("level", 1) or 1)
        result.at[position, "parent"] = _string(parent.get("parent", "ROOT"))
        _append_marker(result, position, "parent_recovered_from_visual_structure")
        _append_review_reason(result, position, "parent_recovered_from_visual_structure")


def _remove_orphan_duplicate_parents(result: pd.DataFrame) -> pd.DataFrame:
    """Drop a childless duplicate parent created by an overlapping capture."""
    if result.empty or "node_id" not in result or "parent_id" not in result:
        return result
    child_ids = {_string(value) for value in result["parent_id"].tolist() if _string(value)}
    groups: dict[tuple[str, str, int], list[int]] = defaultdict(list)
    for position, row in result.iterrows():
        key = (
            _string(row.get("parent_id", "")),
            normalise_label(row.get("text", "")),
            int(row.get("level", 0) or 0),
        )
        groups[key].append(position)

    dropped: set[int] = set()
    for positions in groups.values():
        if len(positions) < 2:
            continue
        # Keep the instance which has descendants.  If both instances have
        # descendants they may be genuine separate branches, so do not guess.
        with_children = [p for p in positions if _string(result.at[p, "node_id"]) in child_ids]
        without_children = [p for p in positions if _string(result.at[p, "node_id"]) not in child_ids]
        if with_children and without_children:
            dropped.update(without_children)
    if not dropped:
        return result
    return result.drop(index=sorted(dropped)).reset_index(drop=True)


def _recover_short_identifier_sequences(result: pd.DataFrame) -> pd.DataFrame:
    """Resolve overlap duplicates in Notes/Tolerance X/Y/Z identifier lists.

    CATIA uses unique short labels in each of these flat sections. A shifted
    overlap can make a prior row (for example X19) win over the direct Z5
    observation. We only replace a repeated export value with a previously
    unseen value that exists in its own OCR alternatives. If all alternatives
    are already represented, the row is an overlap duplicate and is omitted
    from the final tree rather than exported as a false extra node.
    """
    if result.empty:
        return result

    group_columns = [column for column in ("segment", "parent_id", "parent", "level") if column in result]
    if not group_columns:
        return result
    dropped: set[int] = set()
    grouped = result.groupby(group_columns, sort=False, dropna=False)
    for _, indexes in grouped.groups.items():
        # ``groupby`` usually preserves the input order, but an index created
        # by an earlier merge/export can be non-monotonic.  The CATIA order is
        # the dataframe order, so make it explicit before decoding a flat
        # sequence.
        positions = sorted(indexes)
        if not positions or not _identifier_context(result.loc[positions[0]]):
            continue
        # A short series must be a real list, not one isolated identifier.
        series_positions = [
            position
            for position in positions
            if _first_short_identifier(result.at[position, "text"]) or _short_evidence(result.loc[position])
        ]
        if len(series_positions) < 4:
            continue

        seen: set[str] = set()
        recent: list[str] = []
        for position in positions:
            row = result.loc[position]
            current = _first_short_identifier(row.get("text", ""))
            evidence = _short_evidence(row)
            if not current and not evidence:
                continue
            # Do not promote an accidental X/Y/Z fragment found in a long
            # label (e.g. ``ID PLM`` or a tolerance-code alternative) into a
            # Notes node.  If the row itself is not a compact short display,
            # it cannot be a member of this flat identifier sequence.
            if current is None and not _looks_like_short_display(row.get("text", "")):
                continue
            available = [
                (score, candidate)
                for candidate, score in evidence.items()
                if candidate not in seen
            ]
            available.sort(key=lambda item: (-item[0], item[1]))
            is_recent_duplicate = current is not None and current in recent[-8:]

            replacement: str | None = None
            if current is None and available:
                replacement = available[0][1]
            elif current is not None and is_recent_duplicate:
                if available:
                    # A new alternative is safer than exporting a duplicate
                    # within a flat, unique CATIA note/tolerance sequence.
                    replacement = available[0][1]
                elif len(evidence) > 1 or "|" in _string(row.get("ocr_alternatives", "")):
                    dropped.add(position)
                    _append_marker(result, position, "overlap_duplicate_omitted")
                    _append_review_reason(result, position, "overlap_duplicate_omitted")
                    continue

            chosen = replacement or current
            if chosen:
                original = _string(row.get("text", ""))
                if replacement or not _is_clean_short_display(original, chosen):
                    result.at[position, "text"] = _render_short_identifier(chosen, original)
                    _append_marker(result, position, "identifier_from_alternatives")
                    _append_review_reason(result, position, "identifier_recovered_from_overlap")
                seen.add(chosen)
                recent.append(chosen)

    if not dropped:
        return result
    return result.drop(index=sorted(dropped)).reset_index(drop=True)


def _recover_code_sequences(result: pd.DataFrame) -> pd.DataFrame:
    """Remove overlap duplicates and recover missing tolerance-code rows.

    A long CATIA tree is captured with overlapping frames.  At a frame
    boundary the OCR graph can therefore contain the same tolerance code two
    or three times, while the next code is present only as an alternative
    reading of one of those rows.  Tolerance codes are flat children of one
    parent, so a code already emitted in that parent is a reliable signal of
    an overlap duplicate.  We keep the strongest unseen alternative and omit
    a duplicate only when no unseen alternative is available.
    """
    if result.empty:
        return result

    group_columns = [column for column in ("segment", "parent_id", "parent", "level") if column in result]
    if not group_columns:
        return result

    dropped: set[int] = set()
    grouped = result.groupby(group_columns, sort=False, dropna=False)
    for _, indexes in grouped.groups.items():
        positions = sorted(indexes)
        if not positions:
            continue
        parent_key = normalise_label(result.loc[positions[0]].get("parent", ""))
        if parent_key not in _SEMANTIC_IDENTIFIER_PARENTS:
            continue

        # Notes also contains long metadata labels.  Restrict those groups to
        # rows which actually look like a compact code display; all children
        # of Tolérance géométrique are code candidates by definition.
        candidate_positions = []
        for position in positions:
            row = result.loc[position]
            if parent_key == "notes" and not _looks_like_code_display(row.get("text", "")):
                continue
            if _first_code(row.get("text", "")) or _code_evidence(row):
                candidate_positions.append(position)
        if len(candidate_positions) < 3:
            continue

        seen: set[str] = set()
        for position in candidate_positions:
            row = result.loc[position]
            current = _first_code(row.get("text", ""))
            evidence = _code_evidence(row)
            available = sorted(
                ((score, candidate) for candidate, score in evidence.items() if candidate not in seen),
                key=lambda item: (-item[0], item[1]),
            )

            replacement: str | None = None
            if current is None:
                if available:
                    replacement = available[0][1]
                else:
                    # OCR did not produce a usable code for this row; retain
                    # it for the audit rather than silently deleting data.
                    continue
            elif current in seen:
                if available:
                    replacement = available[0][1]
                else:
                    dropped.add(position)
                    _append_marker(result, position, "overlap_duplicate_omitted")
                    _append_review_reason(result, position, "overlap_duplicate_omitted")
                    continue

            chosen = replacement or current
            if not chosen:
                continue
            if replacement:
                result.at[position, "text"] = _canonical_code_display(row.get("text", ""), chosen)
                _append_marker(result, position, "identifier_from_alternatives")
                _append_review_reason(result, position, "identifier_recovered_from_overlap")
            else:
                # Even when the main code is clean, CATIA's icon can be read
                # as a ``GE`` prefix and the repeated, truncated code in
                # parentheses can be read differently.  Normalise only this
                # compact code display; descriptive Notes remain untouched.
                original = _string(row.get("text", ""))
                cleaned = _canonical_code_display(original, chosen)
                if cleaned != original:
                    result.at[position, "text"] = cleaned
                    _append_marker(result, position, "identifier_normalised")
            seen.add(chosen)

    if not dropped:
        return result
    return result.drop(index=sorted(dropped)).reset_index(drop=True)


def _refresh_tree_columns(dataframe: pd.DataFrame) -> pd.DataFrame:
    """Refresh direct parent labels, paths, and display line numbers."""
    result = dataframe.copy().reset_index(drop=True)
    if "node_id" not in result:
        return result
    by_id = result.set_index("node_id", drop=False).to_dict("index")
    paths: list[str] = []
    parents: list[str] = []
    for _, row in result.iterrows():
        chain = [str(row.get("text", ""))]
        parent_id = _string(row.get("parent_id", ""))
        seen = {_string(row.get("node_id", ""))}
        direct_parent = "ROOT" if not parent_id else _string(row.get("parent", ""))
        while parent_id and parent_id in by_id and parent_id not in seen:
            parent = by_id[parent_id]
            parent_text = _string(parent.get("text", ""))
            if len(chain) == 1:
                direct_parent = parent_text
            chain.append(parent_text)
            seen.add(parent_id)
            parent_id = _string(parent.get("parent_id", ""))
        parents.append(direct_parent)
        paths.append(" > ".join(reversed(chain)))
    if "parent" in result:
        result["parent"] = parents
    result["full_path"] = paths
    if "line" in result:
        result["line"] = range(1, len(result) + 1)
    return result


def recover_annotation_text(dataframe: pd.DataFrame) -> pd.DataFrame:
    """Return an auditable, corrected annotation subtree before Excel export."""
    result = dataframe.copy().reset_index(drop=True)
    if result.empty or "text" not in result:
        return result
    result["ocr_text_before_recovery"] = result["text"].astype(str)
    result["text_recovery"] = ""

    _recover_codes(result)
    _repair_capture_labels(result)
    _repair_reference_labels(result)
    _promote_capture_siblings(result)
    result = _remove_orphan_duplicate_parents(result)
    result = _recover_code_sequences(result)
    result = _recover_short_identifier_sequences(result)
    return _refresh_tree_columns(result)


__all__ = ["recover_annotation_text"]
