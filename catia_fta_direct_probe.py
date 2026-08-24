"""
CATIA V5 — Diagnostic FTA/TPS direct
=====================================

Objectif :
- vérifier si le CATPart expose réellement Part.AnnotationSets ;
- lister tous les AnnotationSet / Annotation ;
- enregistrer les propriétés accessibles sans OCR ;
- tester plusieurs interfaces TPS connues ;
- produire un JSON exploitable pour construire le lecteur direct final.

Ce script ne modifie PAS la pièce CATIA.
"""

from __future__ import annotations

import json
import time
import traceback
from pathlib import Path
from typing import Any

try:
    import pythoncom
    import win32com.client
except ImportError as exc:
    raise SystemExit(
        "pywin32 manque. Installez-le avec : "
        r".\.venv\Scripts\python.exe -m pip install pywin32"
    ) from exc


def safe_get(obj: Any, name: str, default=None):
    if obj is None:
        return default
    try:
        return getattr(obj, name)
    except Exception:
        return default


def safe_call(obj: Any, name: str, *args):
    if obj is None:
        return None
    try:
        member = getattr(obj, name)
    except Exception:
        return None
    try:
        return member(*args) if callable(member) else member
    except Exception:
        return None


def object_name(obj: Any) -> str:
    for name in ("Name", "name"):
        value = safe_get(obj, name)
        if value not in (None, ""):
            return str(value)
    return ""


def collection_count(collection: Any) -> int:
    value = safe_get(collection, "Count")
    try:
        return int(value)
    except Exception:
        return 0


def iter_collection(collection: Any):
    count = collection_count(collection)
    if count <= 0:
        return
    for index in range(1, count + 1):
        item = safe_call(collection, "Item", index)
        if item is not None:
            yield item


def scalar(value: Any):
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    return str(value)


def probe_properties(obj: Any, names: list[str]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for name in names:
        try:
            value = getattr(obj, name)
            if callable(value):
                continue
            if isinstance(value, (str, int, float, bool)) or value is None:
                result[name] = value
        except Exception:
            continue
    return result


def probe_interface(annotation: Any, method_name: str) -> dict[str, Any] | None:
    interface = safe_call(annotation, method_name)
    if interface is None:
        return None

    properties = probe_properties(
        interface,
        [
            "Name",
            "Type",
            "Value",
            "Text",
            "Tolerance",
            "ToleranceValue",
            "NominalValue",
            "UpperTolerance",
            "LowerTolerance",
            "Length",
            "Diameter",
            "Status",
        ],
    )

    # Certains objets exposent Text comme méthode/interface.
    text_interface = safe_call(interface, "Text")
    if text_interface is not None and text_interface is not interface:
        properties["Text_interface"] = probe_properties(
            text_interface,
            ["Text", "Value", "Name"],
        )

    return {
        "interface": method_name,
        "properties": properties,
    }


def main() -> int:
    started = time.perf_counter()
    pythoncom.CoInitialize()

    report: dict[str, Any] = {
        "status": "STARTED",
        "annotation_sets": [],
        "errors": [],
    }

    try:
        catia = win32com.client.GetActiveObject("CATIA.Application")
        document = safe_get(catia, "ActiveDocument")
        if document is None:
            raise RuntimeError("Aucun document CATIA actif.")

        part = safe_get(document, "Part")
        if part is None:
            raise RuntimeError(
                f"Le document actif {object_name(document)!r} n'est pas un CATPart."
            )

        report["document"] = object_name(document)
        report["part"] = object_name(part)

        annotation_sets = safe_get(part, "AnnotationSets")
        report["part_annotation_sets_accessible"] = annotation_sets is not None
        report["annotation_set_count"] = collection_count(annotation_sets)

        # Quelques versions CATIA chargent la liste paresseusement.
        if annotation_sets is not None:
            safe_call(annotation_sets, "LoadAnnotationSetsList")
            report["annotation_set_count_after_load"] = collection_count(annotation_sets)

        total_annotations = 0

        for set_index, annotation_set in enumerate(
            iter_collection(annotation_sets) or [],
            start=1,
        ):
            set_row: dict[str, Any] = {
                "index": set_index,
                "name": object_name(annotation_set),
                "properties": probe_properties(
                    annotation_set,
                    [
                        "Name",
                        "Standard",
                        "KindOfSet",
                        "ActiveView",
                    ],
                ),
                "annotations": [],
            }

            annotations = safe_get(annotation_set, "Annotations")
            set_row["annotation_count"] = collection_count(annotations)

            captures = safe_get(annotation_set, "Captures")
            set_row["capture_count"] = collection_count(captures)

            tps_views = safe_get(annotation_set, "TPSViews")
            set_row["tps_view_count"] = collection_count(tps_views)

            for ann_index, annotation in enumerate(
                iter_collection(annotations) or [],
                start=1,
            ):
                total_annotations += 1

                row: dict[str, Any] = {
                    "index": ann_index,
                    "name": object_name(annotation),
                    "properties": probe_properties(
                        annotation,
                        [
                            "Name",
                            "Type",
                            "SuperType",
                            "TPSStatus",
                            "Value",
                            "Tolerance",
                            "ToleranceValue",
                            "Text",
                        ],
                    ),
                    "interfaces": [],
                }

                # Interfaces documentées ou fréquemment disponibles en FTA/TPS.
                for method_name in (
                    "Text",
                    "Dimension3D",
                    "CompositeTolerance",
                    "ToleranceZone",
                    "ReferenceFrame",
                    "AssociatedRefFrame",
                    "DatumSimple",
                    "DatumTarget",
                    "ToleranceUnitBasisValue",
                    "TolerancePerUnitBasisRestrictiveValue",
                    "ProjectedToleranceZone",
                    "MaterialCondition",
                    "EnvelopCondition",
                    "DefaultAnnotation",
                    "Noa",
                    "FlagNote",
                ):
                    info = probe_interface(annotation, method_name)
                    if info is not None:
                        row["interfaces"].append(info)

                # Nombre de géométries référencées.
                surface_count = safe_call(annotation, "GetSurfacesCount")
                if surface_count is not None:
                    try:
                        row["surface_count"] = int(surface_count)
                    except Exception:
                        row["surface_count"] = scalar(surface_count)

                set_row["annotations"].append(row)

            report["annotation_sets"].append(set_row)

        report["total_annotations"] = total_annotations
        report["status"] = (
            "DIRECT_FTA_AVAILABLE"
            if total_annotations > 0
            else "NO_DIRECT_ANNOTATIONS"
        )

    except Exception as exc:
        report["status"] = "ERROR"
        report["errors"].append(
            {
                "type": type(exc).__name__,
                "message": str(exc),
                "traceback": traceback.format_exc(),
            }
        )
    finally:
        report["elapsed_seconds"] = round(time.perf_counter() - started, 3)
        pythoncom.CoUninitialize()

    # Le fichier est écrit dans le projet courant lorsque le script est lancé
    # depuis Projet_analyse.
    result_dir = Path.cwd() / "results" / "fta_direct_probe"
    result_dir.mkdir(parents=True, exist_ok=True)
    output = result_dir / "fta_direct_probe_latest.json"
    output.write_text(
        json.dumps(report, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print()
    print("CATIA FTA/TPS — DIAGNOSTIC DIRECT")
    print("--------------------------------")
    print("Statut :", report.get("status"))
    print("AnnotationSets :", report.get("annotation_set_count_after_load",
                                          report.get("annotation_set_count", 0)))
    print("Annotations directes :", report.get("total_annotations", 0))
    print("Temps :", report.get("elapsed_seconds"), "s")
    print("Diagnostic :", output)

    if report.get("status") == "DIRECT_FTA_AVAILABLE":
        print()
        print("BON RESULTAT : la pièce expose les annotations FTA/TPS.")
        print("Le futur export peut privilégier la lecture directe et éviter l'OCR global.")
    elif report.get("status") == "NO_DIRECT_ANNOTATIONS":
        print()
        print("La pièce n'expose pas d'annotations FTA/TPS sémantiques via cette route.")
        print("Le moteur devra utiliser l'inventaire physique des cadres + OCR local.")
    else:
        print()
        print("Le diagnostic direct a rencontré une erreur ; voir le JSON.")

    return 0 if report.get("status") != "ERROR" else 1


if __name__ == "__main__":
    raise SystemExit(main())
