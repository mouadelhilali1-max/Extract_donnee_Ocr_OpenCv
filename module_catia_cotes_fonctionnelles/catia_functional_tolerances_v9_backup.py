"""
catia_functional_tolerances.py
==============================

Module spécialisé pour CATIA V5 — version 9.0 FRAME FIRST / PAROIS LSD / OCR LOCAL :
- lecture directe de l'arbre du CATPart, sans capture d'écran ;
- détection de la hiérarchie REF > groupe fonctionnel > série ;
- lecture des annotations FTA/TPS exposées par l'Automation CATIA ;
- extraction de la multiplicité, du symbole, de la tolérance et des références A..E ;
- rapprochement entre les séries de l'arbre et les annotations ;
- export vers un classeur Excel .xlsx grâce à Microsoft Excel COM ;
- conservation d'un journal détaillé pour les cas partiellement lisibles.

Le fichier peut être :
1. importé depuis le main.py existant avec la fonction run() ;
2. exécuté directement depuis PyCharm.

Dépendance Python :
    pip install pywin32

Pré-requis :
- Windows ;
- CATIA V5 ouvert avec un CATPart actif ;
- Microsoft Excel installé pour produire le fichier .xlsx.

Remarque importante :
La lecture directe fonctionne lorsque les informations visibles sont de vrais objets
FTA/TPS CATIA. Si certaines indications sont uniquement graphiques ou non exposées
par l'Automation de votre version CATIA, le module les marque dans le journal et peut
recevoir des lignes issues de votre ancien OCR via le paramètre ocr_fallback_provider.
"""

from __future__ import annotations

import csv
import datetime as _dt
import json
import os
import re
import sys
import traceback
import unicodedata
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Callable, Iterable, Optional, Sequence

_CURRENT_DIR = Path(__file__).resolve().parent
if str(_CURRENT_DIR) not in sys.path:
    sys.path.insert(0, str(_CURRENT_DIR))

try:
    import pythoncom
    import win32com.client
except ImportError as exc:  # pragma: no cover - seulement sur la machine CATIA
    raise ImportError(
        "Le module 'pywin32' est obligatoire. Installez-le avec : pip install pywin32"
    ) from exc


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

SERIES_PATTERN = re.compile(r"(?<![A-Z0-9])(\d{2}[A-Z]\d{2})(?![A-Z0-9])", re.I)
GROUP_PATTERN = re.compile(r"^\s*\d{2}\s*[-_]", re.I)
REF_PARENT_PATTERN = re.compile(r"^\s*REF(?:\s|$|[_:/|\\-])", re.I)
MULTIPLICITY_PATTERN = re.compile(
    r"(?<![A-Z0-9])(\d+)\s*[X×]\s*(?=\d{2}[A-Z]\d{2})", re.I
)
NUMBER_PATTERN = re.compile(r"(?<![A-Z0-9])([+-]?\d+(?:[.,]\d+)?)(?![A-Z0-9])")
DATUM_LETTER_PATTERN = re.compile(r"(?<![A-Z0-9])([A-E])(?![A-Z0-9])", re.I)

# Noms de collections qui peuvent réellement contenir des enfants dans un CATPart.
ROOT_COLLECTIONS = (
    "HybridBodies",
    "Bodies",
    "OrderedGeometricalSets",
    "AxisSystems",
)

CHILD_COLLECTIONS = (
    "HybridBodies",
    "Bodies",
    "OrderedGeometricalSets",
    "HybridShapes",
    "Shapes",
    "Sketches",
    "AxisSystems",
)

# Traduction du type Automation CATIA vers une désignation compréhensible.
# Le type CATIA brut reste toujours conservé dans Excel comme source principale.
TYPE_LABELS = {
    "FTA_TRUEPOSITION": ("Position / localisation", "⌖"),
    "FTA_PATTERNTRUEPOS": ("Position d'un motif", "⌖"),
    "FTA_CONCENTRICITY": ("Concentricité", "◎"),
    "FTA_SYMMETRY": ("Symétrie", ""),
    "FTA_PARALLELISM": ("Parallélisme", "∥"),
    "FTA_PERPENDICULARITY": ("Perpendicularité", "⟂"),
    "FTA_ANGULARITY": ("Angularité", "∠"),
    "FTA_FLATNESS": ("Planéité", "▱"),
    "FTA_STRAIGHTNESS": ("Rectitude", "—"),
    "FTA_CIRCULARITY": ("Circularité", "○"),
    "FTA_CYLINDRICITY": ("Cylindricité", "⌭"),
    "FTA_PROFILEOFANYLINE": ("Profil d'une ligne", "⌒"),
    "FTA_PROFILEOFASURFACE": ("Profil d'une surface", "⌓"),
    "FTA_TOTALRUNOUT": ("Battement total", ""),
    "FTA_CIRCULARRUNOUT": ("Battement circulaire", ""),
    "FTA_LINEARdimension".upper(): ("Cote linéaire", ""),
    "FTA_ANGULARDIMENSION": ("Cote angulaire", ""),
    "FTA_NONSEMANTICGDT": ("Tolérance géométrique non sémantique", ""),
    "FTA_NONSEMANTICDIMENSION": ("Cote non sémantique", ""),
    "FTA_TEXT": ("Texte 3D", ""),
    "FTA_NOA": ("Note / NOA", ""),
    "FTA_FLAGNOTE": ("Note repérée", ""),
    "FTA_REFERENCEFRAME": ("Système de références", ""),
    "FTA_DATUMSIMPLE": ("Référence simple", ""),
}


# ---------------------------------------------------------------------------
# Modèles de données
# ---------------------------------------------------------------------------

@dataclass
class TreeNode:
    order: int
    name: str
    object_type: str
    depth: int
    path: str
    parent_path: str
    ref_parent: str = ""
    functional_group: str = ""
    series_code: str = ""


@dataclass
class AnnotationRecord:
    order: int
    set_name: str = ""
    set_standard: str = ""
    set_kind: str = ""
    annotation_name: str = ""
    catia_type: str = ""
    catia_super_type: str = ""
    tps_status: str = ""
    raw_text: str = ""
    text_candidates: list[str] = field(default_factory=list)
    series_code: str = ""
    multiplicity: Optional[int] = None
    symbol_label: str = ""
    symbol_character: str = ""
    tolerance_value: Optional[float] = None
    tolerance_lower: Optional[float] = None
    tolerance_upper: Optional[float] = None
    tolerance_source: str = ""
    datum_raw: str = ""
    datum_a: bool = False
    datum_b: bool = False
    datum_c: bool = False
    datum_d: bool = False
    datum_e: bool = False
    surface_count: Optional[int] = None
    associated_geometry: str = ""
    read_status: str = "LECTURE_PARTIELLE"
    diagnostic: str = ""


@dataclass
class FunctionalRow:
    order: int
    ref_parent: str = ""
    functional_group: str = ""
    series_code: str = ""
    multiplicity: Optional[int] = None
    symbol_character: str = ""
    symbol_label: str = ""
    catia_type: str = ""
    annotation_layout: str = ""
    condition_text: str = ""
    tolerance_value: Optional[float] = None
    tolerance_lower: Optional[float] = None
    tolerance_upper: Optional[float] = None
    tolerance_source: str = ""
    datum_a: bool = False
    datum_b: bool = False
    datum_c: bool = False
    datum_d: bool = False
    datum_e: bool = False
    datum_raw: str = ""
    annotation_raw: str = ""
    tree_path: str = ""
    associated_geometry: str = ""
    symbol_image_path: str = ""
    capture_source: str = ""
    ocr_confidence: Optional[float] = None
    ocr_rotation: Optional[float] = None
    read_status: str = ""
    comment: str = ""


@dataclass
class LogEntry:
    level: str
    phase: str
    message: str
    details: str = ""


# ---------------------------------------------------------------------------
# Utilitaires généraux et COM
# ---------------------------------------------------------------------------

class CatiaFunctionalExportError(RuntimeError):
    """Erreur explicite du module d'export CATIA."""


def _now_stamp() -> str:
    return _dt.datetime.now().strftime("%Y%m%d_%H%M%S")


def _clean_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        text = str(value)
    except Exception:
        return ""
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("\u00a0", " ").replace("\r", " ").replace("\n", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _unique_texts(values: Iterable[Any]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        text = _clean_text(value)
        if not text:
            continue
        key = text.casefold()
        if key not in seen:
            seen.add(key)
            result.append(text)
    return result


def _safe_get(obj: Any, attribute: str, default: Any = None) -> Any:
    if obj is None:
        return default
    try:
        return getattr(obj, attribute)
    except Exception:
        return default


def _safe_call(obj: Any, method: str, *args: Any, default: Any = None) -> Any:
    if obj is None:
        return default
    try:
        member = getattr(obj, method)
        return member(*args) if callable(member) else member
    except Exception:
        return default


def _safe_bool_call(obj: Any, method: str) -> bool:
    result = _safe_call(obj, method, default=False)
    try:
        return bool(result)
    except Exception:
        return False


def _safe_float(value: Any) -> Optional[float]:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = _clean_text(value).replace(",", ".")
    try:
        return float(text)
    except (TypeError, ValueError):
        return None


def _safe_int(value: Any) -> Optional[int]:
    if value is None or isinstance(value, bool):
        return None
    try:
        result = int(float(str(value).replace(",", ".")))
    except (TypeError, ValueError):
        return None
    return result if result >= 1 else None


def _object_name(obj: Any, fallback: str = "") -> str:
    for attr in ("Name", "DisplayName", "Label", "PartNumber"):
        value = _safe_get(obj, attr)
        text = _clean_text(value)
        if text:
            return text
    return fallback


def _object_type(obj: Any) -> str:
    for attr in ("Type", "TypeName"):
        value = _safe_get(obj, attr)
        text = _clean_text(value)
        if text:
            return text
    try:
        return type(obj).__name__
    except Exception:
        return "Objet CATIA"


def _iter_collection(collection: Any) -> Iterable[Any]:
    if collection is None:
        return
    count = _safe_get(collection, "Count", 0)
    try:
        count = int(count)
    except Exception:
        count = 0
    for index in range(1, count + 1):
        item = _safe_call(collection, "Item", index)
        if item is not None:
            yield item


def _find_series(text: str) -> str:
    match = SERIES_PATTERN.search(_clean_text(text).upper())
    return match.group(1).upper() if match else ""


def _natural_key(text: str) -> tuple:
    parts = re.split(r"(\d+)", _clean_text(text).upper())
    return tuple(int(part) if part.isdigit() else part for part in parts)


def _nearest_matching_ancestor(path_names: Sequence[str], pattern: re.Pattern[str]) -> str:
    for value in reversed(path_names):
        if pattern.search(value):
            return value
    return ""


def _connect_catia() -> Any:
    try:
        return win32com.client.GetActiveObject("CATIA.Application")
    except Exception as exc:
        raise CatiaFunctionalExportError(
            "CATIA V5 n'est pas détecté. Ouvrez CATIA, ouvrez le CATPart puis relancez le programme."
        ) from exc


def _get_active_part_document(catia: Any) -> tuple[Any, Any]:
    document = _safe_get(catia, "ActiveDocument")
    if document is None:
        raise CatiaFunctionalExportError("Aucun document CATIA actif.")

    part = _safe_get(document, "Part")
    if part is None:
        doc_name = _object_name(document, "document inconnu")
        raise CatiaFunctionalExportError(
            f"Le document actif '{doc_name}' n'est pas un CATPart exploitable."
        )
    return document, part


# ---------------------------------------------------------------------------
# Lecture directe de l'arbre CATIA
# ---------------------------------------------------------------------------

def read_part_tree(part: Any, logs: list[LogEntry]) -> list[TreeNode]:
    """
    Parcourt les collections principales du Part.
    Le parcours conserve l'ordre CATIA et calcule les parents REF/groupe.
    """
    nodes: list[TreeNode] = []
    order = 0
    visited_paths: set[str] = set()

    root_name = _object_name(part, "Part")
    root_path = root_name

    def add_node(
        obj: Any,
        depth: int,
        path_names: list[str],
        parent_path: str,
    ) -> None:
        nonlocal order

        name = _object_name(obj, f"Objet_{order + 1}")
        current_names = [*path_names, name]
        current_path = " > ".join(current_names)

        # Protection contre des doubles références ou cycles COM.
        path_key = current_path.casefold()
        if path_key in visited_paths:
            return
        visited_paths.add(path_key)

        order += 1
        series_code = _find_series(name)
        ref_parent = _nearest_matching_ancestor(current_names[:-1], REF_PARENT_PATTERN)
        functional_group = _nearest_matching_ancestor(current_names[:-1], GROUP_PATTERN)

        nodes.append(
            TreeNode(
                order=order,
                name=name,
                object_type=_object_type(obj),
                depth=depth,
                path=current_path,
                parent_path=parent_path,
                ref_parent=ref_parent,
                functional_group=functional_group,
                series_code=series_code,
            )
        )

        # Ne pas descendre indéfiniment dans une structure COM anormale.
        if depth >= 30:
            logs.append(
                LogEntry(
                    "AVERTISSEMENT",
                    "ARBRE",
                    f"Profondeur maximale atteinte pour {current_path}.",
                )
            )
            return

        for collection_name in CHILD_COLLECTIONS:
            collection = _safe_get(obj, collection_name)
            for child in _iter_collection(collection):
                add_node(
                    child,
                    depth + 1,
                    current_names,
                    current_path,
                )

    # Racine explicite dans la feuille Arbre_complet.
    nodes.append(
        TreeNode(
            order=0,
            name=root_name,
            object_type="Part",
            depth=0,
            path=root_path,
            parent_path="",
        )
    )
    visited_paths.add(root_path.casefold())

    for collection_name in ROOT_COLLECTIONS:
        collection = _safe_get(part, collection_name)
        for child in _iter_collection(collection):
            add_node(child, 1, [root_name], root_path)

    series_count = sum(1 for node in nodes if node.series_code)
    logs.append(
        LogEntry(
            "INFO",
            "ARBRE",
            f"{len(nodes)} éléments lus directement dans l'arbre ; {series_count} série(s) détectée(s).",
        )
    )
    return nodes


# ---------------------------------------------------------------------------
# Lecture directe des annotations FTA/TPS
# ---------------------------------------------------------------------------

def _extract_drawing_dimension_values(annotation: Any) -> tuple[list[str], dict[str, Any]]:
    texts: list[str] = []
    values: dict[str, Any] = {}

    dim3d = _safe_call(annotation, "Dimension3D")
    if dim3d is None:
        return texts, values

    dim2d = _safe_call(dim3d, "Get2dAnnot")
    if dim2d is None:
        return texts, values

    # Valeur principale.
    dim_value = _safe_call(dim2d, "GetValue")
    if dim_value is not None:
        raw_value = _safe_get(dim_value, "Value")
        numeric = _safe_float(raw_value)
        if numeric is not None:
            values["dimension_value"] = numeric
            texts.append(str(numeric))

        for fake_index in (1, 2):
            fake = _safe_call(dim_value, "GetFakeDimValue", fake_index)
            if fake:
                texts.append(fake)

        # Texte avant/après valeur lorsqu'il est accessible.
        bault_result = _safe_call(dim_value, "GetBaultText")
        if isinstance(bault_result, (tuple, list)):
            texts.extend(_clean_text(item) for item in bault_result)

    # Tolérances d'une DrawingDimension : selon le binding COM, les sorties
    # peuvent être retournées dans un tuple.
    tolerances = _safe_call(dim2d, "GetTolerances")
    if isinstance(tolerances, (tuple, list)):
        values["drawing_tolerances_raw"] = list(tolerances)
        for item in tolerances:
            if isinstance(item, str):
                texts.append(item)

        numerics = [
            number
            for number in (_safe_float(item) for item in tolerances)
            if number is not None
        ]
        if len(numerics) >= 2:
            values["tolerance_lower"] = numerics[-2]
            values["tolerance_upper"] = numerics[-1]

    return texts, values


def _extract_dimension_limit(annotation: Any) -> dict[str, Any]:
    result: dict[str, Any] = {}
    has_limit = _safe_bool_call(annotation, "HasDimensionLimit")
    if not has_limit:
        # Certaines versions ne répondent pas à HasDimensionLimit mais
        # permettent quand même l'appel DimensionLimit.
        limit = _safe_call(annotation, "DimensionLimit")
    else:
        limit = _safe_call(annotation, "DimensionLimit")

    if limit is None:
        return result

    nominal = _safe_float(_safe_get(limit, "Nominalvalue"))
    if nominal is not None:
        result["nominal"] = nominal

    result["limit_type"] = _clean_text(_safe_get(limit, "DimensionLimitType"))
    result["tabulated_limit"] = _clean_text(_safe_get(limit, "TabulatedLimit"))

    limits = _safe_call(limit, "Limits")
    if isinstance(limits, (tuple, list)):
        numerics = [
            number for number in (_safe_float(item) for item in limits) if number is not None
        ]
        if len(numerics) >= 2:
            result["lower"] = numerics[-2]
            result["upper"] = numerics[-1]

    return result


def _extract_text_interfaces(annotation: Any) -> list[str]:
    candidates: list[str] = []

    annotation_name = _object_name(annotation)
    if annotation_name:
        candidates.append(annotation_name)

    # Texte TPS.
    text_interface = _safe_call(annotation, "Text")
    if text_interface is not None:
        candidates.append(_safe_get(text_interface, "Text"))
        drawing_text = _safe_call(text_interface, "Get2dAnnot")
        if drawing_text is not None:
            candidates.extend(
                [
                    _safe_get(drawing_text, "Text"),
                    _safe_get(drawing_text, "Name"),
                ]
            )

    # Note NOA.
    noa = _safe_call(annotation, "Noa")
    if noa is not None:
        candidates.extend(
            [
                _safe_get(noa, "Text"),
                _safe_get(noa, "Name"),
            ]
        )

    # Flag note.
    flag = _safe_call(annotation, "FlagNote")
    if flag is not None:
        candidates.extend(
            [
                _safe_get(flag, "Text"),
                _safe_get(flag, "FlagNoteText"),
                _safe_get(flag, "Name"),
            ]
        )

    dimension_texts, _ = _extract_drawing_dimension_values(annotation)
    candidates.extend(dimension_texts)

    return _unique_texts(candidates)


def _extract_reference_frame_text(annotation: Any) -> tuple[str, list[str]]:
    """
    Cherche les références A..E dans le ReferenceFrame associé à la tolérance.
    La structure exacte varie entre les releases CATIA ; les noms bruts sont
    donc conservés pour diagnostic.
    """
    candidates: list[str] = []
    diagnostics: list[str] = []

    ref_annotation = None
    if _safe_bool_call(annotation, "IsAnAssociatedRefFrame"):
        associated = _safe_call(annotation, "AssociatedRefFrame")
        ref_annotation = _safe_get(associated, "ReferenceFrame")
        if ref_annotation is not None:
            candidates.append(_object_name(ref_annotation))

    if ref_annotation is None and _clean_text(_safe_get(annotation, "Type")).upper() == "FTA_REFERENCEFRAME":
        ref_annotation = annotation

    if ref_annotation is None:
        return "", diagnostics

    ref_frame = _safe_call(ref_annotation, "ReferenceFrame")
    if ref_frame is None:
        diagnostics.append("ReferenceFrame associé trouvé, mais interface ReferenceFrame non accessible.")
        return " | ".join(_unique_texts(candidates)), diagnostics

    candidates.extend(
        [
            _safe_get(ref_frame, "Name"),
            _safe_get(ref_frame, "Frame"),
        ]
    )

    all_datums = _safe_get(ref_frame, "AllDatumsSimple")
    for datum_annotation in _iter_collection(all_datums):
        candidates.append(_object_name(datum_annotation))
        datum_simple = _safe_call(datum_annotation, "DatumSimple")
        if datum_simple is not None:
            candidates.extend(
                [
                    _safe_get(datum_simple, "Name"),
                    _safe_get(datum_simple, "Label"),
                    _safe_get(datum_simple, "Text"),
                ]
            )

    return " | ".join(_unique_texts(candidates)), diagnostics


def _extract_surface_information(annotation: Any) -> tuple[Optional[int], str]:
    count = _safe_call(annotation, "GetSurfacesCount")
    try:
        surface_count = int(count) if count is not None else None
    except Exception:
        surface_count = None

    # GetSurfaces est une méthode SAFEARRAY à paramètres de sortie. Son
    # comportement dépend du wrapper COM généré. On tente une lecture sans
    # empêcher l'export si le binding local ne la prend pas en charge.
    surfaces = _safe_call(annotation, "GetSurfaces")
    names: list[str] = []
    if isinstance(surfaces, (tuple, list)):
        for surface in surfaces:
            name = _object_name(surface)
            if name:
                names.append(name)

    return surface_count, " | ".join(_unique_texts(names))


def _parse_multiplicity(text: str) -> Optional[int]:
    """Retourne uniquement une multiplicité explicitement écrite, ex. 5X."""
    match = MULTIPLICITY_PATTERN.search(_clean_text(text).upper())
    if match:
        try:
            return max(1, int(match.group(1)))
        except ValueError:
            pass
    return None


def _parse_datums(*texts: str) -> tuple[str, dict[str, bool]]:
    combined = " | ".join(_unique_texts(texts)).upper()
    found = {letter: False for letter in "ABCDE"}

    # Les références peuvent apparaître sous A, B-C, D-E, A|B-C|D-E, etc.
    for letter in DATUM_LETTER_PATTERN.findall(
        combined.replace("|", " ").replace("/", " ").replace("\\", " ")
    ):
        found[letter.upper()] = True

    # Le motif seul est trop permissif sur les noms de séries. On retire
    # d'abord les codes 01A01, 06B01, etc. puis on refait un passage.
    without_series = SERIES_PATTERN.sub(" ", combined)
    found = {letter: False for letter in "ABCDE"}
    for letter in DATUM_LETTER_PATTERN.findall(without_series):
        found[letter.upper()] = True

    return combined, found


def _parse_tolerance_from_text(text: str, series_code: str) -> Optional[float]:
    normalized = _clean_text(text).upper()
    if not normalized:
        return None

    # Retire la multiplicité et le code série ; le premier nombre restant
    # après le symbole est généralement l'intervalle de tolérance.
    normalized = MULTIPLICITY_PATTERN.sub(" ", normalized)
    if series_code:
        normalized = re.sub(re.escape(series_code), " ", normalized, flags=re.I)

    # Enlève les références A..E et séparateurs classiques.
    normalized = re.sub(r"(?<![A-Z0-9])[A-E](?![A-Z0-9])", " ", normalized)
    numbers = NUMBER_PATTERN.findall(normalized)
    for token in numbers:
        value = _safe_float(token)
        if value is not None:
            return value
    return None


def _build_annotation_record(
    annotation: Any,
    annotation_order: int,
    annotation_set: Any,
) -> AnnotationRecord:
    catia_type = _clean_text(_safe_get(annotation, "Type"))
    super_type = _clean_text(_safe_get(annotation, "SuperType"))
    type_label, type_symbol = TYPE_LABELS.get(
        catia_type.upper(),
        (catia_type.replace("FTA_", "").replace("_", " ") or "Type inconnu", ""),
    )

    text_candidates = _extract_text_interfaces(annotation)
    dimension_texts, dimension_values = _extract_drawing_dimension_values(annotation)
    text_candidates = _unique_texts([*text_candidates, *dimension_texts])

    dimension_limit = _extract_dimension_limit(annotation)
    ref_frame_text, ref_diagnostics = _extract_reference_frame_text(annotation)
    surface_count, associated_geometry = _extract_surface_information(annotation)

    # Texte brut le plus complet disponible.
    raw_text = " | ".join(text_candidates)
    series_code = _find_series(raw_text) or _find_series(_object_name(annotation))
    multiplicity = _parse_multiplicity(raw_text)

    tolerance_value = None
    tolerance_lower = None
    tolerance_upper = None
    tolerance_source = ""

    if "nominal" in dimension_limit:
        tolerance_value = _safe_float(dimension_limit.get("nominal"))
        tolerance_source = "CATIA.DimensionLimit.Nominalvalue"
    elif "dimension_value" in dimension_values:
        tolerance_value = _safe_float(dimension_values.get("dimension_value"))
        tolerance_source = "CATIA.Dimension3D.Get2dAnnot.GetValue"
    else:
        tolerance_value = _parse_tolerance_from_text(raw_text, series_code)
        if tolerance_value is not None:
            tolerance_source = "Texte annotation"

    tolerance_lower = _safe_float(
        dimension_limit.get("lower", dimension_values.get("tolerance_lower"))
    )
    tolerance_upper = _safe_float(
        dimension_limit.get("upper", dimension_values.get("tolerance_upper"))
    )

    datum_raw, datums = _parse_datums(raw_text, ref_frame_text)

    diagnostics: list[str] = []
    diagnostics.extend(ref_diagnostics)
    if not series_code:
        diagnostics.append("Code série non trouvé dans le nom ou le texte de l'annotation.")
    if tolerance_value is None:
        diagnostics.append("Valeur de tolérance non exposée ou non reconnue.")
    if not any(datums.values()):
        diagnostics.append("Références A..E non exposées ou non reconnues.")
    if not raw_text:
        diagnostics.append("Aucun texte brut exposé par les interfaces testées.")

    if series_code and tolerance_value is not None:
        read_status = "OK"
        if diagnostics:
            read_status = "LECTURE_PARTIELLE"
    else:
        read_status = "LECTURE_PARTIELLE"

    set_name = _object_name(annotation_set)
    set_standard = _clean_text(_safe_get(annotation_set, "Standard"))
    set_kind = _clean_text(_safe_get(annotation_set, "KindOfSet"))

    return AnnotationRecord(
        order=annotation_order,
        set_name=set_name,
        set_standard=set_standard,
        set_kind=set_kind,
        annotation_name=_object_name(annotation),
        catia_type=catia_type,
        catia_super_type=super_type,
        tps_status=_clean_text(_safe_get(annotation, "TPSStatus")),
        raw_text=raw_text,
        text_candidates=text_candidates,
        series_code=series_code,
        multiplicity=multiplicity,
        symbol_label=type_label,
        symbol_character=type_symbol,
        tolerance_value=tolerance_value,
        tolerance_lower=tolerance_lower,
        tolerance_upper=tolerance_upper,
        tolerance_source=tolerance_source,
        datum_raw=datum_raw,
        datum_a=datums["A"],
        datum_b=datums["B"],
        datum_c=datums["C"],
        datum_d=datums["D"],
        datum_e=datums["E"],
        surface_count=surface_count,
        associated_geometry=associated_geometry,
        read_status=read_status,
        diagnostic="; ".join(diagnostics),
    )


def read_annotations(
    part: Any,
    logs: list[LogEntry],
    ocr_fallback_provider: Optional[Callable[[], Iterable[dict[str, Any]]]] = None,
) -> list[AnnotationRecord]:
    records: list[AnnotationRecord] = []

    annotation_sets = _safe_get(part, "AnnotationSets")
    if annotation_sets is None:
        logs.append(
            LogEntry(
                "AVERTISSEMENT",
                "ANNOTATIONS",
                "La propriété Part.AnnotationSets n'est pas accessible.",
                "Le document peut ne pas contenir de vraies annotations FTA/TPS ou la licence/API peut être limitée.",
            )
        )
    else:
        # Utile dans certains cas où les sets ne sont pas encore chargés.
        _safe_call(annotation_sets, "LoadAnnotationSetsList")

        annotation_order = 0
        for annotation_set in _iter_collection(annotation_sets):
            annotations = _safe_get(annotation_set, "Annotations")
            for annotation in _iter_collection(annotations):
                annotation_order += 1
                try:
                    records.append(
                        _build_annotation_record(
                            annotation,
                            annotation_order,
                            annotation_set,
                        )
                    )
                except Exception as exc:
                    records.append(
                        AnnotationRecord(
                            order=annotation_order,
                            set_name=_object_name(annotation_set),
                            annotation_name=_object_name(annotation),
                            read_status="ERREUR_ANNOTATION",
                            diagnostic=f"{type(exc).__name__}: {exc}",
                        )
                    )
                    logs.append(
                        LogEntry(
                            "ERREUR",
                            "ANNOTATIONS",
                            f"Erreur pendant la lecture de l'annotation #{annotation_order}.",
                            traceback.format_exc(),
                        )
                    )

    logs.append(
        LogEntry(
            "INFO",
            "ANNOTATIONS",
            f"{len(records)} annotation(s) FTA/TPS lue(s) directement.",
        )
    )

    # Crochet facultatif vers l'ancien système OCR.
    if not records and ocr_fallback_provider is not None:
        logs.append(
            LogEntry(
                "AVERTISSEMENT",
                "OCR",
                "Aucune annotation directe trouvée : activation du fournisseur OCR existant.",
            )
        )
        try:
            ocr_rows = list(ocr_fallback_provider() or [])
            records.extend(_convert_ocr_rows(ocr_rows))
            logs.append(
                LogEntry(
                    "INFO",
                    "OCR",
                    f"{len(ocr_rows)} ligne(s) reçue(s) depuis le fallback OCR.",
                )
            )
        except Exception:
            logs.append(
                LogEntry(
                    "ERREUR",
                    "OCR",
                    "Le fallback OCR a échoué.",
                    traceback.format_exc(),
                )
            )

    return records


def _convert_ocr_rows(rows: Iterable[dict[str, Any]]) -> list[AnnotationRecord]:
    converted: list[AnnotationRecord] = []
    for index, row in enumerate(rows, start=1):
        raw = _clean_text(
            row.get("raw_text")
            or row.get("text")
            or row.get("annotation")
            or row.get("ocr_text")
        )
        series = _clean_text(row.get("series_code")) or _find_series(raw)
        tolerance = _safe_float(row.get("tolerance_value"))
        if tolerance is None:
            tolerance = _parse_tolerance_from_text(raw, series)
        datum_raw, datums = _parse_datums(
            raw,
            _clean_text(row.get("datum_raw") or row.get("references")),
        )
        converted.append(
            AnnotationRecord(
                order=index,
                set_name="OCR fallback",
                annotation_name=_clean_text(row.get("name")),
                catia_type=_clean_text(row.get("catia_type") or "OCR"),
                raw_text=raw,
                text_candidates=[raw] if raw else [],
                series_code=series,
                multiplicity=(_safe_int(row.get("multiplicity")) or _parse_multiplicity(raw)),
                symbol_label=_clean_text(row.get("symbol_label") or row.get("symbol")),
                symbol_character=_clean_text(row.get("symbol_character")),
                tolerance_value=tolerance,
                tolerance_lower=_safe_float(row.get("tolerance_lower")),
                tolerance_upper=_safe_float(row.get("tolerance_upper")),
                tolerance_source="OCR fallback",
                datum_raw=datum_raw,
                datum_a=datums["A"],
                datum_b=datums["B"],
                datum_c=datums["C"],
                datum_d=datums["D"],
                datum_e=datums["E"],
                read_status="OCR",
                diagnostic="Donnée issue du fournisseur OCR existant.",
            )
        )
    return converted


# ---------------------------------------------------------------------------
# Rapprochement arbre / annotations
# ---------------------------------------------------------------------------

def merge_tree_and_annotations(
    tree_nodes: Sequence[TreeNode],
    annotations: Sequence[AnnotationRecord],
    logs: list[LogEntry],
) -> list[FunctionalRow]:
    series_nodes = [node for node in tree_nodes if node.series_code]

    annotations_by_series: dict[str, list[AnnotationRecord]] = {}
    annotations_without_series: list[AnnotationRecord] = []
    for annotation in annotations:
        if annotation.series_code:
            annotations_by_series.setdefault(annotation.series_code, []).append(annotation)
        else:
            annotations_without_series.append(annotation)

    rows: list[FunctionalRow] = []
    output_order = 0
    matched_annotation_ids: set[int] = set()

    # Une ligne par série de l'arbre, ou plusieurs si plusieurs annotations
    # portent explicitement le même code.
    for node in series_nodes:
        matches = annotations_by_series.get(node.series_code, [])
        if not matches:
            output_order += 1
            rows.append(
                FunctionalRow(
                    order=output_order,
                    ref_parent=node.ref_parent,
                    functional_group=node.functional_group,
                    series_code=node.series_code,
                    tree_path=node.path,
                    read_status="SERIE_SANS_ANNOTATION",
                    comment=(
                        "La série existe dans l'arbre, mais aucune annotation FTA/TPS "
                        "portant ce code n'a été trouvée directement."
                    ),
                )
            )
            continue

        for annotation in matches:
            matched_annotation_ids.add(id(annotation))
            output_order += 1
            comment_parts = []
            if annotation.diagnostic:
                comment_parts.append(annotation.diagnostic)
            rows.append(
                FunctionalRow(
                    order=output_order,
                    ref_parent=node.ref_parent,
                    functional_group=node.functional_group,
                    series_code=node.series_code,
                    multiplicity=annotation.multiplicity,
                    symbol_character=annotation.symbol_character,
                    symbol_label=annotation.symbol_label,
                    catia_type=annotation.catia_type,
                    tolerance_value=annotation.tolerance_value,
                    tolerance_lower=annotation.tolerance_lower,
                    tolerance_upper=annotation.tolerance_upper,
                    tolerance_source=annotation.tolerance_source,
                    datum_a=annotation.datum_a,
                    datum_b=annotation.datum_b,
                    datum_c=annotation.datum_c,
                    datum_d=annotation.datum_d,
                    datum_e=annotation.datum_e,
                    datum_raw=annotation.datum_raw,
                    annotation_raw=annotation.raw_text,
                    tree_path=node.path,
                    associated_geometry=annotation.associated_geometry,
                    read_status=annotation.read_status,
                    comment="; ".join(comment_parts),
                )
            )

    # Annotations avec un code qui n'existe pas dans l'arbre parcouru.
    for annotation in annotations:
        if id(annotation) in matched_annotation_ids:
            continue
        output_order += 1
        status = (
            "ANNOTATION_SANS_SERIE_ARBRE"
            if annotation.series_code
            else "ANNOTATION_SANS_CODE_SERIE"
        )
        rows.append(
            FunctionalRow(
                order=output_order,
                series_code=annotation.series_code,
                multiplicity=annotation.multiplicity,
                symbol_character=annotation.symbol_character,
                symbol_label=annotation.symbol_label,
                catia_type=annotation.catia_type,
                tolerance_value=annotation.tolerance_value,
                tolerance_lower=annotation.tolerance_lower,
                tolerance_upper=annotation.tolerance_upper,
                tolerance_source=annotation.tolerance_source,
                datum_a=annotation.datum_a,
                datum_b=annotation.datum_b,
                datum_c=annotation.datum_c,
                datum_d=annotation.datum_d,
                datum_e=annotation.datum_e,
                datum_raw=annotation.datum_raw,
                annotation_raw=annotation.raw_text,
                associated_geometry=annotation.associated_geometry,
                read_status=status,
                comment=annotation.diagnostic,
            )
        )

    logs.append(
        LogEntry(
            "INFO",
            "RAPPROCHEMENT",
            f"{len(rows)} ligne(s) fonctionnelle(s) préparée(s) pour Excel.",
            (
                f"{len(series_nodes)} série(s) dans l'arbre ; "
                f"{len(annotations)} annotation(s) ; "
                f"{len(annotations_without_series)} annotation(s) sans code série."
            ),
        )
    )
    return rows



def export_tree_series_manifest(
    project_root: Path,
    tree_nodes: Sequence[TreeNode],
    document_name: str,
) -> Path:
    """
    Sauvegarde le résultat de la première étape : extraction de la barre gauche.

    Cette liste devient la référence officielle pour l'OCR des captures.
    Aucun code série OCR qui n'existe pas dans cette liste n'est accepté.
    """
    series_rows = [
        {
            "order": node.order,
            "series_code": node.series_code,
            "ref_parent": node.ref_parent,
            "functional_group": node.functional_group,
            "tree_path": node.path,
        }
        for node in tree_nodes
        if node.series_code
    ]

    output_dir = project_root / "results" / "tree_extraction"
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / "tree_series_latest.json"

    payload = {
        "document": document_name,
        "generated_at": _dt.datetime.now().isoformat(timespec="seconds"),
        "method": "CATIA_COM_direct_tree",
        "series_count": len(series_rows),
        "series": series_rows,
    }
    output_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return output_path



# ---------------------------------------------------------------------------
# Complément visuel OCR : cadres rectangulaires, IT, multiplicité et références
# ---------------------------------------------------------------------------

def enrich_rows_with_visual_ocr(
    functional_rows: Sequence[FunctionalRow],
    tree_nodes: Sequence[TreeNode],
    logs: list[LogEntry],
    project_root: Path,
) -> list[dict[str, Any]]:
    """
    V9 FRAME FIRST : la liste extraite de la barre gauche reste la référence
    officielle. OpenCV inventorie les cadres, LSD mesure leurs parois physiques
    puis Tesseract ne lit que les cellules et libellés locaux.
    """
    known_series = sorted(
        {node.series_code for node in tree_nodes if node.series_code},
        key=_natural_key,
    )
    if not known_series:
        logs.append(
            LogEntry(
                "AVERTISSEMENT",
                "OCR_VISUEL",
                "Aucune série connue dans l'arbre : OCR visuel ignoré.",
            )
        )
        return []

    try:
        try:
            from .visual_annotation_scanner import scan_annotation_captures
        except Exception:
            try:
                from module_catia_cotes_fonctionnelles.visual_annotation_scanner import scan_annotation_captures
            except Exception:
                from visual_annotation_scanner import scan_annotation_captures
    except Exception:
        logs.append(
            LogEntry(
                "ERREUR",
                "OCR_VISUEL",
                "Impossible d'importer visual_annotation_scanner.py.",
                traceback.format_exc(),
            )
        )
        print(
            "      ERREUR : visual_annotation_scanner.py est absent ou non importable.",
            flush=True,
        )
        return []

    capture_dir = project_root / "captures_annotations"
    capture_count = 0
    if capture_dir.exists():
        capture_count = sum(
            1
            for path in capture_dir.rglob("*")
            if path.suffix.lower() in {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"}
        )

    print(
        f"      Dossier OCR : {capture_dir} ({capture_count} capture(s))",
        flush=True,
    )

    series_groups = {
        node.series_code: node.functional_group
        for node in tree_nodes
        if node.series_code
    }
    try:
        visual_rows = scan_annotation_captures(
            project_root=project_root,
            known_series=known_series,
            series_groups=series_groups,
            capture_dirs=[capture_dir],
            capture_current_view=False,
            # Recherche série-guidée tous les 5 degrés.
            rotations=(),
            upscale=3.0,
        )
    except Exception:
        logs.append(
            LogEntry(
                "ERREUR",
                "OCR_VISUEL",
                "L'analyse des captures a échoué.",
                traceback.format_exc(),
            )
        )
        print("      ERREUR pendant l'OCR visuel. Consultez le Journal.", flush=True)
        return []

    by_series: dict[str, dict[str, Any]] = {
        _clean_text(item.get("series_code")).upper(): item
        for item in visual_rows
        if _clean_text(item.get("series_code"))
    }

    enriched = 0
    for row in functional_rows:
        visual = by_series.get(row.series_code.upper())
        if not visual:
            # Ne jamais laisser l'ancienne valeur artificielle 1.
            if row.multiplicity == 1 and not row.annotation_raw:
                row.multiplicity = None
            continue

        multiplicity = _safe_int(visual.get("multiplicity"))
        tolerance = _safe_float(visual.get("tolerance_value"))
        confidence = _safe_float(visual.get("confidence"))
        rotation = _safe_float(visual.get("rotation_angle"))
        annotation_layout = _clean_text(visual.get("annotation_layout"))
        condition_text = _clean_text(visual.get("condition_text"))

        # Les données visuelles ne sont injectées que si le scanner les a déjà
        # validées physiquement. Une confiance basse reste vide : mieux vaut une
        # cellule Excel vide qu'une valeur fausse.
        if confidence is not None and confidence < 0.78:
            logs.append(LogEntry("AVERTISSEMENT", "OCR_VISUEL",
                f"{row.series_code}: lecture visuelle rejetée (confiance {confidence:.3f})."))
            continue

        # Les données visuelles sont prioritaires pour les champs explicitement
        # visibles dans le cadre. Une valeur absente ne remplace pas une valeur COM.
        if multiplicity is not None:
            row.multiplicity = multiplicity
        elif row.multiplicity == 1 and not row.annotation_raw:
            row.multiplicity = None

        if tolerance is not None:
            row.tolerance_value = tolerance
            row.tolerance_source = "OCR visuel — cadre CATIA"

        # Le symbole sera traité dans une phase ultérieure.
        # Pour le moment, les trois champs restent volontairement vides.
        row.symbol_character = ""
        row.symbol_label = ""
        row.symbol_image_path = ""

        # Références : compléter sans effacer les informations directes.
        for letter in "abcde":
            field_name = f"datum_{letter}"
            if bool(visual.get(field_name)):
                setattr(row, field_name, True)
        if _clean_text(visual.get("datum_raw")):
            row.datum_raw = _clean_text(visual.get("datum_raw"))

        # Les références appartiennent déjà à la hiérarchie CATIA (REF_parent).
        # Pour un cadre standard, la donnée arbre est plus fiable qu'un OCR
        # partiel de A | B-C | D-E. Les conditionnels à 2 cellules restent sans
        # références lorsque le cadre n'en affiche pas.
        is_conditional = bool(
            annotation_layout.startswith("CONDITIONNEL")
            or condition_text
        )
        if not is_conditional and row.ref_parent:
            _parent_raw, _parent_datums = _parse_datums(row.ref_parent)
            for _letter in "abcde":
                if _parent_datums.get(_letter.upper(), False):
                    setattr(row, f"datum_{_letter}", True)
            if any(_parent_datums.values()) and not row.datum_raw:
                row.datum_raw = row.ref_parent

        visual_raw = _clean_text(visual.get("raw_text"))
        if visual_raw:
            row.annotation_raw = visual_raw
        row.capture_source = _clean_text(visual.get("source_image"))
        row.ocr_confidence = confidence
        row.ocr_rotation = rotation
        row.catia_type = row.catia_type or "OCR_VISUEL"
        row.annotation_layout = annotation_layout
        row.condition_text = condition_text

        missing: list[str] = []
        if row.tolerance_value is None:
            missing.append("IT")
        references_required = not (
            row.annotation_layout.startswith("CONDITIONNEL")
            or bool(row.condition_text)
        )
        if references_required and not any(
            [
                row.datum_a,
                row.datum_b,
                row.datum_c,
                row.datum_d,
                row.datum_e,
            ]
        ):
            missing.append("références")

        vis_status = _clean_text(visual.get("read_status"))
        if vis_status in ("CONFLIT_IT_MULTI_CAPTURES", "AMBIGU_CANDIDATS_MULTIPLES"):
            row.read_status = vis_status
            extra = f"Statut OCR : {vis_status}."
        elif missing:
            row.read_status = "OCR_PARTIEL"
            extra = "Cadre reconnu, mais lecture incomplète : " + ", ".join(missing)
        else:
            row.read_status = "OCR_OK"
            if row.annotation_layout.startswith("CONDITIONNEL") or row.condition_text:
                extra = "IT extrait du cadre conditionnel et associé à la série."
                if row.condition_text:
                    extra += f" Condition : {row.condition_text}."
            else:
                extra = "IT et références extraites du cadre rectangulaire et associées par série."

        if row.multiplicity is None:
            extra += " Multiplicité non affichée ou non lue ; cellule laissée vide."

        diagnostic = _clean_text(visual.get("diagnostic"))
        additions = [part for part in (extra, diagnostic) if part]
        row.comment = "; ".join(
            part for part in [row.comment, *additions] if _clean_text(part)
        )
        enriched += 1

    # Le symbole est volontairement reporté à la prochaine étape du projet.
    for functional_row in functional_rows:
        functional_row.symbol_character = ""
        functional_row.symbol_label = ""
        functional_row.symbol_image_path = ""

    logs.append(
        LogEntry(
            "INFO" if visual_rows else "AVERTISSEMENT",
            "OCR_VISUEL",
            f"{len(visual_rows)} série(s) reconnue(s) dans les captures ; {enriched} ligne(s) Excel complétée(s).",
            f"Captures analysées : {capture_count}. Diagnostic : results/frame_inventory_ocr/frame_inventory_latest.json",
        )
    )
    return visual_rows


# ---------------------------------------------------------------------------
# Export Excel COM
# ---------------------------------------------------------------------------

def _excel_safe(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "Oui" if value else ""
    if isinstance(value, (int, float, str)):
        return value
    if isinstance(value, (list, tuple, dict)):
        return json.dumps(value, ensure_ascii=False)
    return _clean_text(value)


def _matrix(headers: Sequence[str], rows: Sequence[Sequence[Any]]) -> tuple:
    return tuple(
        [tuple(headers)]
        + [tuple(_excel_safe(value) for value in row) for row in rows]
    )


def _write_excel_sheet(
    workbook: Any,
    name: str,
    headers: Sequence[str],
    rows: Sequence[Sequence[Any]],
    column_widths: Optional[dict[int, float]] = None,
    autofilter: bool = True,
) -> Any:
    sheet = workbook.Worksheets.Add()
    sheet.Name = name[:31]

    data = _matrix(headers, rows)
    row_count = len(data)
    col_count = len(headers)
    target = sheet.Range(
        sheet.Cells(1, 1),
        sheet.Cells(max(1, row_count), max(1, col_count)),
    )
    target.Value = data

    # En-tête professionnel.
    header = sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, col_count))
    header.Font.Bold = True
    header.Font.Color = 0xFFFFFF
    header.Interior.Color = 0x8B5A2B  # Bleu/ocre lisible selon conversion BGR Excel.
    header.HorizontalAlignment = -4108  # xlCenter
    header.VerticalAlignment = -4108
    header.WrapText = True
    header.RowHeight = 32

    # Corps.
    if row_count > 1:
        body = sheet.Range(sheet.Cells(2, 1), sheet.Cells(row_count, col_count))
        body.VerticalAlignment = -4160  # xlTop
        body.WrapText = True
        body.Borders.LineStyle = 1
        body.Borders.Weight = 2

    target.Borders.LineStyle = 1
    target.Borders.Weight = 2

    # Filtres et volets figés.
    if autofilter and row_count >= 1:
        header.AutoFilter()
    sheet.Activate()
    sheet.Application.ActiveWindow.SplitRow = 1
    sheet.Application.ActiveWindow.FreezePanes = True

    # Largeurs d'abord automatiques, puis plafonnées.
    target.Columns.AutoFit()
    for column in range(1, col_count + 1):
        current_width = sheet.Columns(column).ColumnWidth
        if current_width and current_width > 42:
            sheet.Columns(column).ColumnWidth = 42
        elif current_width and current_width < 9:
            sheet.Columns(column).ColumnWidth = 9

    if column_widths:
        for index, width in column_widths.items():
            if 1 <= index <= col_count:
                sheet.Columns(index).ColumnWidth = width

    # Mettre le statut en couleur.
    status_col = None
    for index, header_name in enumerate(headers, start=1):
        if header_name == "Statut_lecture":
            status_col = index
            break
    if status_col and row_count > 1:
        for row_index in range(2, row_count + 1):
            cell = sheet.Cells(row_index, status_col)
            status = _clean_text(cell.Value).upper()
            if status == "OK":
                cell.Interior.Color = 0xC6EFCE
                cell.Font.Color = 0x006100
            elif "ERREUR" in status:
                cell.Interior.Color = 0xFFC7CE
                cell.Font.Color = 0x9C0006
            elif status:
                cell.Interior.Color = 0xFFEB9C
                cell.Font.Color = 0x9C6500

    return sheet


def export_to_excel(
    output_path: Path,
    functional_rows: Sequence[FunctionalRow],
    tree_nodes: Sequence[TreeNode],
    annotations: Sequence[AnnotationRecord],
    logs: Sequence[LogEntry],
    document_name: str,
) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Add()

        # Supprime les feuilles vierges après création de nos feuilles.
        default_sheets = [workbook.Worksheets.Item(i) for i in range(1, workbook.Worksheets.Count + 1)]

        functional_headers = [
            "Ordre",
            "REF_parent",
            "Groupe_fonctionnel",
            "Serie",
            "Multiplicite",
            "Symbole",
            "Designation_symbole",
            "Type_CATIA",
            "Type_annotation",
            "Condition_associee",
            "IT",
            "Tolerance_min",
            "Tolerance_max",
            "Source_tolerance",
            "Ref_A",
            "Ref_B",
            "Ref_C",
            "Ref_D",
            "Ref_E",
            "References_brutes",
            "Annotation_brute",
            "Chemin_arbre",
            "Geometrie_associee",
            "Capture_source",
            "Confiance_OCR",
            "Rotation_OCR",
            "Image_symbole",
            "Statut_lecture",
            "Commentaire",
        ]
        functional_data = [
            [
                row.order,
                row.ref_parent,
                row.functional_group,
                row.series_code,
                row.multiplicity,
                row.symbol_character,
                row.symbol_label,
                row.catia_type,
                row.annotation_layout,
                row.condition_text,
                row.tolerance_value,
                row.tolerance_lower,
                row.tolerance_upper,
                row.tolerance_source,
                row.datum_a,
                row.datum_b,
                row.datum_c,
                row.datum_d,
                row.datum_e,
                row.datum_raw,
                row.annotation_raw,
                row.tree_path,
                row.associated_geometry,
                row.capture_source,
                row.ocr_confidence,
                row.ocr_rotation,
                row.symbol_image_path,
                row.read_status,
                row.comment,
            ]
            for row in functional_rows
        ]
        functional_sheet = _write_excel_sheet(
            workbook,
            "Cotes_fonctionnelles",
            functional_headers,
            functional_data,
            column_widths={
                2: 24,
                3: 27,
                4: 12,
                5: 12,
                6: 10,
                7: 24,
                8: 18,
                9: 27,
                10: 24,
                11: 12,
                20: 30,
                21: 42,
                22: 42,
                23: 28,
                24: 42,
                25: 14,
                26: 13,
                27: 42,
                28: 24,
                29: 42,
            },
        )

        # Insère le recadrage exact du symbole dans la cellule Symbole.
        # La valeur texte reste disponible si l'image ne peut pas être ajoutée.
        symbol_column = functional_headers.index("Symbole") + 1
        image_path_column = functional_headers.index("Image_symbole") + 1
        for excel_row, functional_row in enumerate(functional_rows, start=2):
            image_path = Path(functional_row.symbol_image_path) if functional_row.symbol_image_path else None
            if image_path is None or not image_path.exists():
                continue
            try:
                cell = functional_sheet.Cells(excel_row, symbol_column)
                functional_sheet.Rows(excel_row).RowHeight = 48
                functional_sheet.Shapes.AddPicture(
                    str(image_path.resolve()),
                    False,
                    True,
                    float(cell.Left + 2),
                    float(cell.Top + 2),
                    float(max(18, cell.Width - 4)),
                    float(max(18, cell.Height - 4)),
                )
                functional_sheet.Cells(excel_row, image_path_column).Value = str(image_path.resolve())
            except Exception:
                # L'export ne doit pas échouer pour une image individuelle.
                pass

        tree_headers = [
            "Ordre",
            "Nom",
            "Type_objet",
            "Profondeur",
            "Parent",
            "Chemin_complet",
            "REF_parent",
            "Groupe_fonctionnel",
            "Serie_detectee",
        ]
        tree_data = [
            [
                node.order,
                node.name,
                node.object_type,
                node.depth,
                node.parent_path,
                node.path,
                node.ref_parent,
                node.functional_group,
                node.series_code,
            ]
            for node in tree_nodes
        ]
        _write_excel_sheet(
            workbook,
            "Arbre_complet",
            tree_headers,
            tree_data,
            column_widths={2: 28, 3: 24, 5: 40, 6: 55, 7: 27, 8: 27},
        )

        annotation_headers = [
            "Ordre",
            "Set_annotation",
            "Standard",
            "Type_set",
            "Nom_annotation",
            "Type_CATIA",
            "SuperType_CATIA",
            "TPS_Status",
            "Serie",
            "Multiplicite",
            "Symbole",
            "Designation",
            "Tolerance",
            "Tolerance_min",
            "Tolerance_max",
            "Source_tolerance",
            "Ref_A",
            "Ref_B",
            "Ref_C",
            "Ref_D",
            "Ref_E",
            "References_brutes",
            "Texte_brut",
            "Textes_candidats",
            "Nb_surfaces",
            "Geometrie_associee",
            "Statut_lecture",
            "Diagnostic",
        ]
        annotation_data = [
            [
                record.order,
                record.set_name,
                record.set_standard,
                record.set_kind,
                record.annotation_name,
                record.catia_type,
                record.catia_super_type,
                record.tps_status,
                record.series_code,
                record.multiplicity,
                record.symbol_character,
                record.symbol_label,
                record.tolerance_value,
                record.tolerance_lower,
                record.tolerance_upper,
                record.tolerance_source,
                record.datum_a,
                record.datum_b,
                record.datum_c,
                record.datum_d,
                record.datum_e,
                record.datum_raw,
                record.raw_text,
                " || ".join(record.text_candidates),
                record.surface_count,
                record.associated_geometry,
                record.read_status,
                record.diagnostic,
            ]
            for record in annotations
        ]
        _write_excel_sheet(
            workbook,
            "Annotations_brutes",
            annotation_headers,
            annotation_data,
            column_widths={
                2: 25,
                5: 28,
                6: 25,
                22: 34,
                23: 46,
                24: 46,
                26: 30,
                27: 24,
                28: 46,
            },
        )

        log_headers = ["Date_heure", "Niveau", "Phase", "Message", "Details"]
        timestamp = _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_data = [
            [timestamp, entry.level, entry.phase, entry.message, entry.details]
            for entry in logs
        ]
        log_data.insert(
            0,
            [
                timestamp,
                "INFO",
                "DOCUMENT",
                f"Document CATIA analysé : {document_name}",
                f"Fichier généré : {output_path}",
            ],
        )
        _write_excel_sheet(
            workbook,
            "Journal",
            log_headers,
            log_data,
            column_widths={1: 20, 2: 18, 3: 20, 4: 55, 5: 70},
            autofilter=True,
        )

        # Supprimer les feuilles Excel vierges initiales.
        for sheet in default_sheets:
            try:
                sheet.Delete()
            except Exception:
                pass

        # Mettre la feuille principale en première position.
        main_sheet = workbook.Worksheets("Cotes_fonctionnelles")
        main_sheet.Move(Before=workbook.Worksheets(1))
        main_sheet.Activate()

        # 51 = xlOpenXMLWorkbook (.xlsx)
        workbook.SaveAs(str(output_path.resolve()), FileFormat=51)
        workbook.Close(SaveChanges=True)
        workbook = None
        excel.Quit()
        excel = None
        return output_path.resolve()

    except Exception as exc:
        try:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        raise CatiaFunctionalExportError(
            "Impossible de créer le fichier Excel. Vérifiez que Microsoft Excel "
            "est installé et qu'aucune boîte de dialogue Excel n'est bloquée."
        ) from exc


def export_diagnostics_json(
    output_path: Path,
    tree_nodes: Sequence[TreeNode],
    annotations: Sequence[AnnotationRecord],
    functional_rows: Sequence[FunctionalRow],
    logs: Sequence[LogEntry],
) -> Path:
    """
    Fichier de diagnostic facultatif. Très utile pour corriger une différence
    d'API entre deux releases CATIA sans refaire des captures d'écran.
    """
    payload = {
        "generated_at": _dt.datetime.now().isoformat(timespec="seconds"),
        "tree": [asdict(item) for item in tree_nodes],
        "annotations": [asdict(item) for item in annotations],
        "functional_rows": [asdict(item) for item in functional_rows],
        "logs": [asdict(item) for item in logs],
    }
    output_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return output_path


# ---------------------------------------------------------------------------
# API publique destinée au main.py
# ---------------------------------------------------------------------------

def run(
    catia: Any = None,
    output_dir: str | os.PathLike[str] | None = None,
    output_name: str | None = None,
    ocr_fallback_provider: Optional[Callable[[], Iterable[dict[str, Any]]]] = None,
    create_diagnostic_json: bool = True,
) -> Path:
    """
    Point d'entrée recommandé pour le main.py.

    Paramètres
    ----------
    catia:
        Objet CATIA.Application déjà connecté par main.py. Si None, le module
        se connecte à l'instance CATIA ouverte.
    output_dir:
        Dossier de sortie. Par défaut : sous-dossier "exports" du projet.
    output_name:
        Nom .xlsx facultatif.
    ocr_fallback_provider:
        Fonction facultative retournant une liste de dictionnaires OCR.
        Elle n'est appelée que si aucune annotation FTA/TPS directe n'est lue.
    create_diagnostic_json:
        Génère un fichier JSON à côté d'Excel pour faciliter le réglage initial.

    Retour
    ------
    pathlib.Path
        Chemin absolu du classeur Excel produit.
    """
    pythoncom.CoInitialize()
    logs: list[LogEntry] = []

    try:
        catia = catia or _connect_catia()
        document, part = _get_active_part_document(catia)
        document_name = _object_name(document, "CATPart")

        logs.append(
            LogEntry(
                "INFO",
                "DEMARRAGE",
                "Lecture directe CATIA démarrée.",
                f"Document : {document_name}",
            )
        )

        project_root = Path(__file__).resolve().parent.parent

        print("[1/5] Extraction directe de la barre gauche CATIA...", flush=True)
        tree_nodes = read_part_tree(part, logs)
        series_nodes = [node for node in tree_nodes if node.series_code]
        tree_manifest = export_tree_series_manifest(
            project_root,
            tree_nodes,
            document_name,
        )
        print(
            f"[1/5] Barre gauche terminée : {len(series_nodes)} série(s).",
            flush=True,
        )
        print(f"      Manifeste arbre : {tree_manifest}", flush=True)

        print("[2/5] Préparation de la hiérarchie REF > groupe > série...", flush=True)

        print("[3/5] Lecture directe des annotations FTA/TPS...", flush=True)
        annotations = read_annotations(part, logs, ocr_fallback_provider)
        print(
            f"[3/5] Lecture directe terminée : {len(annotations)} annotation(s).",
            flush=True,
        )
        functional_rows = merge_tree_and_annotations(tree_nodes, annotations, logs)
        print(
            f"[2/5] Hiérarchie préparée : {len(functional_rows)} ligne(s).",
            flush=True,
        )
        print("[4/5] OCR V9.0 : cadres OpenCV -> parois LSD -> cellule IT locale vérifiée...", flush=True)
        visual_rows = enrich_rows_with_visual_ocr(
            functional_rows,
            tree_nodes,
            logs,
            project_root,
        )
        print(
            f"[4/5] OCR terminé : {len(visual_rows)} série(s) reconnue(s).",
            flush=True,
        )

        destination = Path(output_dir) if output_dir else project_root / "exports"
        destination.mkdir(parents=True, exist_ok=True)

        if output_name:
            filename = output_name
            if not filename.lower().endswith(".xlsx"):
                filename += ".xlsx"
        else:
            safe_document = re.sub(
                r"[^A-Za-z0-9._-]+",
                "_",
                Path(document_name).stem,
            ).strip("_") or "CATPart"
            filename = f"{safe_document}_cotes_fonctionnelles_{_now_stamp()}.xlsx"

        excel_path = destination / filename
        print(
            f"[5/5] Création Excel : {len(functional_rows)} ligne(s)...",
            flush=True,
        )
        excel_path = export_to_excel(
            excel_path,
            functional_rows,
            tree_nodes,
            annotations,
            logs,
            document_name,
        )
        print("[5/5] Fichier Excel créé.", flush=True)

        if create_diagnostic_json:
            diagnostic_path = excel_path.with_suffix(".diagnostic.json")
            export_diagnostics_json(
                diagnostic_path,
                tree_nodes,
                annotations,
                functional_rows,
                logs,
            )

        return excel_path

    finally:
        pythoncom.CoUninitialize()


def main() -> int:
    """Permet aussi d'exécuter ce fichier directement."""
    try:
        output = run()
        diagnostic = output.with_suffix(".diagnostic.json")
        missing_series: list[str] = []
        if diagnostic.exists():
            try:
                payload = json.loads(diagnostic.read_text(encoding="utf-8"))
                for item in payload.get("functional_rows", []):
                    if item.get("series_code") and item.get("tolerance_value") is None:
                        missing_series.append(str(item.get("series_code")))
            except Exception:
                pass

        if missing_series:
            print("\nExport Excel créé, mais couverture incomplète :")
            print(output)
            print("Séries sans IT : " + ", ".join(sorted(set(missing_series), key=_natural_key)))
        else:
            print("\nExport terminé avec succès — couverture complète :")
            print(output)
        if diagnostic.exists():
            print("Diagnostic :")
            print(diagnostic)
        return 0
    except CatiaFunctionalExportError as exc:
        print(f"\nERREUR : {exc}", file=sys.stderr)
        return 1
    except Exception:
        print("\nERREUR INATTENDUE :", file=sys.stderr)
        traceback.print_exc()
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
