"""Build one CATIA tree from registered OCR node observations.

Levels and duplicates are resolved before this class (in
``catia_tree_pipeline``).  This class only creates parent relationships once,
from stable node identifiers.  It deliberately never creates a silent ROOT
when an intermediate level is not visible in a screenshot.
"""

from __future__ import annotations

import re

import pandas as pd


class TreeBuilder:
    def __init__(self) -> None:
        self.df = pd.DataFrame()

    def load_dataframe(self, dataframe: pd.DataFrame) -> None:
        self.df = dataframe.copy()
        if self.df.empty:
            return
        if "segment" not in self.df:
            self.df["segment"] = 0
        if "global_y" not in self.df:
            self.df["global_y"] = self.df.get("top", range(len(self.df)))
        if "level" not in self.df:
            self._infer_levels_from_left()
        self.df = self.df.sort_values(["segment", "global_y", "level", "left"], kind="stable").reset_index(drop=True)

    def _infer_levels_from_left(self) -> None:
        """Compatibility fallback for callers that provide old OCR data."""
        positions = sorted(int(value) for value in self.df["left"].dropna().unique())
        clusters: list[list[int]] = []
        for value in positions:
            if not clusters or value - clusters[-1][-1] > 18:
                clusters.append([value])
            else:
                clusters[-1].append(value)
        centers = [sum(cluster) / len(cluster) for cluster in clusters] or [0]
        self.df["level"] = self.df["left"].apply(
            lambda value: min(range(len(centers)), key=lambda index: abs(centers[index] - value))
        )

    def _ensure_node_ids(self) -> None:
        if "node_id" not in self.df:
            self.df.insert(0, "node_id", [f"N{index:05d}" for index in range(1, len(self.df) + 1)])
        else:
            self.df["node_id"] = self.df["node_id"].astype(str)

    def _assign_parents(self) -> None:
        parent_ids: list[str] = []
        parent_names: list[str] = []
        hierarchy_status: list[str] = []

        stack: dict[int, tuple[str, str]] = {}
        known_nodes: dict[str, tuple[str, int]] = {}
        current_segment = None
        root_seen_in_segment = False
        for _, row in self.df.iterrows():
            segment = row["segment"]
            if current_segment != segment:
                stack = {}
                known_nodes = {}
                current_segment = segment
                root_seen_in_segment = False

            level = max(0, int(row["level"]))
            raw_explicit_parent = row.get("validated_parent_id", "")
            explicit_parent_id = "" if pd.isna(raw_explicit_parent) else str(raw_explicit_parent).strip()
            if explicit_parent_id:
                # A human may validate a parent after reviewing an ambiguous
                # screenshot.  It is honoured only when the declared node is
                # an earlier node of the same segment; never accept a stale
                # ID from another run as though it were a real relation.
                if explicit_parent_id in known_nodes:
                    explicit_text, _ = known_nodes[explicit_parent_id]
                    parent_ids.append(explicit_parent_id)
                    parent_names.append(explicit_text)
                    hierarchy_status.append("human_validated")
                else:
                    parent_ids.append("")
                    parent_names.append("PARENT_UNKNOWN")
                    hierarchy_status.append("invalid_validated_parent")
            elif level == 0:
                if not root_seen_in_segment:
                    parent_ids.append("")
                    parent_names.append("ROOT")
                    hierarchy_status.append("root")
                    root_seen_in_segment = True
                else:
                    # A later level-0 row can be a true second root, but in a
                    # scrolling CATIA panel it more often means the actual
                    # document root is above the visible crop.  Do not invent
                    # a second independent tree without a human decision.
                    parent_ids.append("")
                    parent_names.append("PARENT_UNKNOWN")
                    hierarchy_status.append("root_level_unconfirmed")
            else:
                possible_depths = [depth for depth in stack if depth < level]
                if not possible_depths:
                    # The capture begins in the middle of an expanded tree.
                    # Mark the missing context rather than inventing ROOT.
                    parent_ids.append("")
                    parent_names.append("PARENT_UNKNOWN")
                    hierarchy_status.append("parent_unknown")
                else:
                    parent_level = max(possible_depths)
                    parent_id, parent_text = stack[parent_level]
                    parent_ids.append(parent_id)
                    parent_names.append(parent_text)
                    hierarchy_status.append(
                        "direct" if parent_level == level - 1 else "parent_inferred_gap"
                    )

            # A sibling closes every deeper branch before it is stored.
            stack = {depth: item for depth, item in stack.items() if depth < level}
            stack[level] = (str(row["node_id"]), str(row["text"]))
            known_nodes[str(row["node_id"])] = (str(row["text"]), level)

        self.df["parent_id"] = parent_ids
        self.df["parent"] = parent_names
        self.df["hierarchy_status"] = hierarchy_status

    def _build_paths(self) -> None:
        by_id = self.df.set_index("node_id", drop=False).to_dict("index")
        paths: list[str] = []
        for _, row in self.df.iterrows():
            chain = [str(row["text"])]
            parent_id = str(row["parent_id"] or "")
            seen = {str(row["node_id"])}
            while parent_id and parent_id in by_id and parent_id not in seen:
                parent = by_id[parent_id]
                chain.append(str(parent["text"]))
                seen.add(parent_id)
                parent_id = str(parent.get("parent_id") or "")
            if row["hierarchy_status"] in {"parent_unknown", "root_level_unconfirmed"}:
                chain.append("PARENT_UNKNOWN")
            paths.append(" > ".join(reversed(chain)))
        self.df["full_path"] = paths

    def _infer_sequential_codes(self) -> None:
        """Use a local sibling sequence as evidence for one OCR digit confusion.

        This is intentionally not a global replacement such as ``YS -> Y5``.
        It activates only when consecutive siblings already establish a prefix
        and the immediately following value is an OCR-confusable digit.  The
        inferred value stays in the review queue for human validation.
        """
        self.df["sequence_inferred"] = False
        self.df["ocr_text_before_sequence"] = ""
        if "review_reason" not in self.df:
            self.df["review_reason"] = ""

        ambiguity_to_digit = {"O": "0", "I": "1", "L": "1", "Z": "2", "S": "5"}
        grouped = self.df.groupby(["segment", "parent_id", "level"], sort=False, dropna=False)
        for _, subset in grouped:
            indexes = list(subset.index)
            for position, index in enumerate(indexes):
                value = str(self.df.at[index, "text"])
                ambiguous = re.fullmatch(r"([A-Za-z]+)([OILSZ])", value)
                if ambiguous is None or position < 2:
                    continue
                prefix, ambiguous_character = ambiguous.groups()
                expected_digit = ambiguity_to_digit[ambiguous_character.upper()]
                previous_values: list[int] = []
                for previous_index in indexes[max(0, position - 4):position]:
                    previous = re.fullmatch(rf"{re.escape(prefix)}(\d+)", str(self.df.at[previous_index, "text"]))
                    if previous is not None:
                        previous_values.append(int(previous.group(1)))
                if len(previous_values) < 2:
                    continue
                expected_value = previous_values[-1] + 1
                if previous_values[-2] + 1 != previous_values[-1] or str(expected_value) != expected_digit:
                    continue

                inferred = f"{prefix}{expected_value}"
                self.df.at[index, "ocr_text_before_sequence"] = value
                self.df.at[index, "text"] = inferred
                self.df.at[index, "sequence_inferred"] = True
                existing = self.df.at[index, "review_reason"]
                existing_text = "" if pd.isna(existing) else str(existing).strip()
                suffix = "sequence_inferred_needs_validation"
                self.df.at[index, "review_reason"] = ";".join(
                    part for part in (existing_text, suffix) if part
                )

    def build(self) -> pd.DataFrame:
        if self.df.empty:
            return self.df
        self._ensure_node_ids()
        self._assign_parents()
        self._infer_sequential_codes()
        self._build_paths()
        self.df["line"] = range(1, len(self.df) + 1)

        # A structurally inferred relation is exportable but remains auditable.
        if "review_needed" not in self.df:
            self.df["review_needed"] = False
        structural_review = self.df["hierarchy_status"].isin([
            "parent_unknown", "parent_inferred_gap", "invalid_validated_parent",
            "root_level_unconfirmed",
        ])
        self.df["review_needed"] = self.df["review_needed"].astype(bool) | structural_review
        self.df.loc[self.df["sequence_inferred"], "review_needed"] = True
        return self.df
