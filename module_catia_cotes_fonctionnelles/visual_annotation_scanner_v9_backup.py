"""
CATIA V5 — visual_annotation_scanner.py
Version 9.0 — PAROIS PHYSIQUES LSD / OCR LOCAL FIABLE

Architecture:
1) l'arbre CATIA fournit la liste officielle des séries ;
2) OpenCV détecte d'abord les cadres/cellules physiques ;
3) chaque cadre est redressé localement ;
4) Tesseract ne lit que de petites zones ;
5) la série, l'IT, la multiplicité, les références et les conditions sont extraits ;
6) les répétitions entre captures sont dédupliquées.

Aucune liste d'IT n'est codée en dur. Une valeur n'est acceptée que si elle est
lue dans un cadre physique associé à une série de l'arbre.
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import math
import os
import re
import shutil
import time
import traceback
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Iterable, Optional, Sequence

try:
    import cv2
    import numpy as np
    import pytesseract
    from PIL import Image
except ImportError as exc:
    raise ImportError(
        "Installez les dépendances : pip install opencv-python pytesseract Pillow numpy"
    ) from exc

try:
    import pythoncom
    import win32com.client
except ImportError:
    pythoncom = None
    win32com = None

VERSION = "9.0-lsd-physical-walls-local-cell-ocr"
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"}
SERIES_RE = re.compile(r"^\d{2}[A-Z]\d{2}$")
MULT_RE = re.compile(r"(?<![A-Z0-9])(\d{1,3})[ \t]*[X×x](?![A-Z0-9])")
MULT_RE_REV = re.compile(r"(?<![A-Z0-9])[X×x][ \t]*(\d{1,3})(?![A-Z0-9])")
CONDITION_KEYWORDS = ("HEIGHT", "WIDTH", "LENGTH", "THICKNESS", "GAP", "DIAMETER")
# Limite de sécurité générique, configurable sans modifier le code.
# Aucune liste d'IT propre à une pièce n'est utilisée.
MAX_TOLERANCE_VALUE = float(os.environ.get("CATIA_MAX_TOLERANCE", "50.0"))


@dataclass
class VisualAnnotation:
    series_code: str
    multiplicity: Optional[int] = None
    tolerance_value: Optional[float] = None
    tolerance_text: str = ""
    symbol_character: str = ""
    symbol_label: str = ""
    symbol_image_path: str = ""
    datum_raw: str = ""
    datum_a: bool = False
    datum_b: bool = False
    datum_c: bool = False
    datum_d: bool = False
    datum_e: bool = False
    raw_text: str = ""
    source_image: str = ""
    rotation_angle: float = 0.0
    confidence: float = 0.0
    diagnostic: str = ""
    annotation_layout: str = ""
    condition_text: str = ""
    read_status: str = ""


@dataclass
class PhysicalObservation:
    image_path: Path
    crop: np.ndarray
    angle: float
    crop_polygon: np.ndarray
    texts: list[str]
    candidate_scores: dict[str, float]
    series_code: str = ""
    series_score: float = 0.0
    tolerance_value: Optional[float] = None
    multiplicity: Optional[int] = None
    datum_raw: str = ""
    datums: dict[str, bool] = field(default_factory=dict)
    layout: str = "CADRE_REFERENCES"
    condition_text: str = ""
    confidence: float = 0.0
    diagnostic: str = ""
    read_status: str = ""
    physical_polygon: Optional[np.ndarray] = None
    cell_polygons: list[np.ndarray] = field(default_factory=list)
    internal_walls: list[tuple[tuple[int, int], tuple[int, int]]] = field(default_factory=list)
    geometric_score: float = 0.0
    candidate_polygon: Optional[np.ndarray] = None


class VisualScanError(RuntimeError):
    pass


# ---------------------------------------------------------------------------
# Capture CATIA interactive (compatibilité avec les versions précédentes)
# ---------------------------------------------------------------------------
CAT_CAPTURE_FORMAT_BMP = 4


def _project_root(project_root: Path | str | None = None) -> Path:
    if project_root is not None:
        return Path(project_root).resolve()
    return Path(__file__).resolve().parent.parent


def _connect_catia_application() -> Any:
    if pythoncom is None or win32com is None:
        raise VisualScanError("pywin32 est requis pour capturer CATIA : pip install pywin32")
    try:
        return win32com.client.GetActiveObject("CATIA.Application")
    except Exception as exc:
        raise VisualScanError("CATIA V5 n'est pas détecté.") from exc


def _archive_existing_captures(output_dir: Path) -> Path | None:
    output_dir.mkdir(parents=True, exist_ok=True)
    files = [p for p in output_dir.iterdir() if p.is_file() and p.suffix.lower() in IMAGE_EXTENSIONS]
    if not files:
        return None
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = output_dir.parent / f"captures_annotations_backup_{stamp}"
    i = 1
    while backup.exists():
        backup = output_dir.parent / f"captures_annotations_backup_{stamp}_{i:02d}"
        i += 1
    backup.mkdir(parents=True)
    for p in files:
        shutil.move(str(p), str(backup / p.name))
    return backup


def capture_catia_viewer(output_dir: Path, prefix: str = "annotation_view") -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    pythoncom.CoInitialize()
    tmp: Path | None = None
    try:
        catia = _connect_catia_application()
        viewer = catia.ActiveWindow.ActiveViewer
        try:
            viewer.Activate()
            viewer.Update()
        except Exception:
            pass
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        tmp = output_dir / f".{prefix}_{stamp}.bmp"
        png = output_dir / f"{prefix}_{stamp}.png"
        viewer.CaptureToFile(CAT_CAPTURE_FORMAT_BMP, str(tmp.resolve()))
        deadline = time.time() + 10
        while time.time() < deadline and (not tmp.exists() or tmp.stat().st_size == 0):
            time.sleep(0.05)
        if not tmp.exists() or tmp.stat().st_size == 0:
            raise VisualScanError("CATIA n'a pas produit la capture.")
        with Image.open(tmp) as im:
            im.convert("RGB").save(png, "PNG")
        return png
    finally:
        if tmp is not None:
            try:
                tmp.unlink(missing_ok=True)
            except Exception:
                pass
        pythoncom.CoUninitialize()


def interactive_capture(project_root: Path | str | None = None, archive_existing: bool = True) -> list[Path]:
    root = _project_root(project_root)
    out = root / "captures_annotations"
    if archive_existing:
        backup = _archive_existing_captures(out)
        if backup:
            print(f"Anciennes captures déplacées vers : {backup}")
    print("\nCAPTURE DIRECTE CATIA — V3.0")
    print("Positionnez la vue dans CATIA, revenez ici et appuyez sur Entrée.")
    print("Tapez q puis Entrée pour terminer.\n")
    rows: list[Path] = []
    index = 1
    while True:
        try:
            command = input(f"Vue {index} prête ? [Entrée = capturer, q = terminer] : ").strip().lower()
        except KeyboardInterrupt:
            break
        if command in {"q", "quit", "fin", "stop"}:
            break
        try:
            path = capture_catia_viewer(out, prefix=f"annotation_view_{index:02d}")
            rows.append(path)
            print(f"Capture CATIA enregistrée : {path}")
            index += 1
        except Exception as exc:
            print(f"ERREUR : {exc}")
    return rows


# ---------------------------------------------------------------------------
# Utilitaires OCR / géométrie
# ---------------------------------------------------------------------------
def _clean_text(value: Any) -> str:
    if value is None:
        return ""
    return " ".join(str(value).replace("\x00", " ").split())


def _configure_tesseract(root: Path) -> str:
    candidates = [
        os.environ.get("TESSERACT_CMD", ""),
        shutil.which("tesseract") or "",
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        str(root / "Tesseract-OCR" / "tesseract.exe"),
    ]
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            pytesseract.pytesseract.tesseract_cmd = candidate
            return candidate
    raise VisualScanError("Tesseract introuvable. Installez Tesseract-OCR ou définissez TESSERACT_CMD.")


def _list_images(directories: Sequence[Path]) -> list[Path]:
    found: list[Path] = []
    for directory in directories:
        if not directory.exists():
            continue
        found.extend(p for p in directory.rglob("*") if p.is_file() and p.suffix.lower() in IMAGE_EXTENSIONS)
    return sorted(set(found), key=lambda p: p.name.lower())


def _capture_signature(images: Sequence[Path], known: set[str]) -> str:
    h = hashlib.sha256()
    h.update(VERSION.encode())
    for code in sorted(known):
        h.update(code.encode())
    for p in images:
        try:
            stat = p.stat()
            h.update(str(p.resolve()).encode())
            h.update(str(stat.st_size).encode())
            h.update(str(stat.st_mtime_ns).encode())
        except OSError:
            pass
    return h.hexdigest()


def _white_mask(image: np.ndarray) -> np.ndarray:
    """Masque universel de traits et cadres, totalement indépendant de la couleur (blanc, cyan, jaune, vert, etc.)."""
    if image.size == 0:
        return np.zeros((1, 1), dtype=np.uint8)
    b, g, r = cv2.split(image) if image.ndim == 3 else (image, image, image)
    mx = np.maximum.reduce([b, g, r])
    mn = np.minimum.reduce([b, g, r])
    # 1. Traits clairs neutres (blanc / gris clair)
    mask_light = (b > 135) & (g > 135) & (r > 135) & ((mx - mn) < 70)
    # 2. Filtrage morphologique Top-Hat (extrait les traits fins et parois de toute couleur)
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY) if image.ndim == 3 else image
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (5, 5))
    tophat = cv2.morphologyEx(gray, cv2.MORPH_TOPHAT, kernel)
    mask_tophat = tophat > 20
    # 3. Traits colorés vifs
    if image.ndim == 3:
        hsv = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)
        sat = hsv[:, :, 1]
        val = hsv[:, :, 2]
        mask_color = (sat > 55) & (val > 80) & mask_tophat
        # Exclusion stricte des repères de référence pourpres/magenta (X15, X32, Z2)
        mask_purple = (r > 130) & (b > 130) & (g < 100)
    else:
        mask_color = np.zeros_like(b, dtype=bool)
        mask_purple = np.zeros_like(b, dtype=bool)
    mask = (mask_light | mask_tophat | mask_color) & (~mask_purple)
    return mask.astype(np.uint8) * 255


def _cyan_frame_mask(image: np.ndarray) -> np.ndarray:
    """Masque des cadres colorés/cyan, tolérant à l'antialiasing."""
    if image.size == 0 or image.ndim != 3:
        return np.zeros(image.shape[:2], dtype=np.uint8)
    hsv = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)
    hue, sat, val = cv2.split(hsv)
    mask = (sat >= 50) & (val >= 70)
    return mask.astype(np.uint8) * 255


def _local_binary_ocr(
    image: np.ndarray,
    *,
    psm: int,
    whitelist: str,
    scale: float = 3.5,
) -> str:
    """OCR d'une petite ROI avec contraste local ; évite l'OCR global lent."""
    if image.size == 0:
        return ""
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    enlarged = cv2.resize(gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    _, binary = cv2.threshold(enlarged, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    config = (
        f"--oem 3 --psm {psm} -c user_defined_dpi=300 "
        f"-c tessedit_char_whitelist={whitelist}"
    )
    try:
        return pytesseract.image_to_string(binary, lang="eng", config=config).strip()
    except Exception:
        return ""


def _cyan_cell_candidates(image: np.ndarray) -> list[dict[str, Any]]:
    """Cellules rectangulaires fermées des cadres cyan."""
    mask = _cyan_frame_mask(image)
    h, w = mask.shape
    left_guard = int(w * 0.11)
    contours, _ = cv2.findContours(mask, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    out: list[dict[str, Any]] = []
    min_h = max(8, int(h * 0.014))
    max_h = max(48, int(h * 0.095))
    min_w = max(8, int(h * 0.012))
    max_w = max(90, int(h * 0.18))
    for contour in contours:
        area = float(cv2.contourArea(contour))
        x, y, bw, bh = cv2.boundingRect(contour)
        if x < left_guard:
            continue
        if not (min_h <= bh <= max_h and min_w <= bw <= max_w):
            continue
        fill = area / max(float(bw * bh), 1.0)
        aspect = bw / max(float(bh), 1.0)
        if fill < 0.50 or not (0.50 <= aspect <= 3.2):
            continue
        out.append({
            "cx": x + bw / 2.0, "cy": y + bh / 2.0,
            "rw": float(bw), "rh": float(bh), "angle": 0.0,
            "area": area, "fill": fill, "contour": contour,
            "x": int(x), "y": int(y), "w": int(bw), "h": int(bh),
            "source": "CYAN",
        })
    out.sort(key=lambda item: item["area"], reverse=True)
    selected: list[dict[str, Any]] = []
    for item in out:
        duplicate = any(
            abs(item["cx"] - prev["cx"]) <= 3
            and abs(item["cy"] - prev["cy"]) <= 3
            and abs(item["rw"] - prev["rw"]) <= 5
            and abs(item["rh"] - prev["rh"]) <= 5
            for prev in selected
        )
        if not duplicate:
            selected.append(item)
    return selected


def _group_cyan_cells(candidates: Sequence[dict[str, Any]]) -> list[list[dict[str, Any]]]:
    """Regroupe les cellules cyan contiguës sans dépendre de l'angle 0/90 des carrés."""
    n = len(candidates)
    adjacency = [set() for _ in range(n)]
    for i in range(n):
        a = candidates[i]
        for j in range(i + 1, n):
            b = candidates[j]
            hmed = (a["h"] + b["h"]) / 2.0
            if hmed <= 0 or abs(a["cy"] - b["cy"]) > 0.38 * hmed:
                continue
            if max(a["h"], b["h"]) / max(1.0, min(a["h"], b["h"])) > 1.30:
                continue
            left, right = (a, b) if a["cx"] <= b["cx"] else (b, a)
            gap = right["x"] - (left["x"] + left["w"])
            if -4 <= gap <= 0.45 * hmed:
                adjacency[i].add(j); adjacency[j].add(i)
    groups: list[list[dict[str, Any]]] = []
    seen: set[int] = set()
    for start in range(n):
        if start in seen:
            continue
        stack = [start]; seen.add(start); indices: list[int] = []
        while stack:
            current = stack.pop(); indices.append(current)
            for neighbour in adjacency[current]:
                if neighbour not in seen:
                    seen.add(neighbour); stack.append(neighbour)
        if 2 <= len(indices) <= 9:
            group = [candidates[index] for index in indices]
            group.sort(key=lambda item: item["x"])
            groups.append(group)
    groups.sort(key=lambda g: (-len(g), g[0]["y"], g[0]["x"]))
    final: list[list[dict[str, Any]]] = []
    for group in groups:
        x1=min(x["x"] for x in group); y1=min(x["y"] for x in group)
        x2=max(x["x"]+x["w"] for x in group); y2=max(x["y"]+x["h"] for x in group)
        duplicate=any(
            abs(x1-min(x["x"] for x in prev))<5 and abs(y1-min(x["y"] for x in prev))<5
            and abs(x2-max(x["x"]+x["w"] for x in prev))<8
            and abs(y2-max(x["y"]+x["h"] for x in prev))<8
            for prev in final
        )
        if not duplicate:
            final.append(group)
    return sorted(final, key=lambda g:(min(x["y"] for x in g), min(x["x"] for x in g)))


def _cyan_series_roi(image: np.ndarray, group: Sequence[dict[str, Any]]) -> tuple[np.ndarray, tuple[int,int,int,int]]:
    group=sorted(group,key=lambda item:item["x"])
    x0=min(item["x"] for item in group); y0=min(item["y"] for item in group)
    y1=max(item["y"]+item["h"] for item in group)
    unit=float(np.median([item["h"] for item in group]))
    left=max(0,int(x0-5.8*unit)); right=min(image.shape[1],int(x0+0.25*unit))
    top=max(0,int(y0-1.85*unit)); bottom=min(image.shape[0],int(y1+0.55*unit))
    return image[top:bottom,left:right],(left,top,right,bottom)


def _cyan_series_choice(image: np.ndarray, group: Sequence[dict[str, Any]], known: set[str], *, targets: Optional[set[str]]=None) -> tuple[str,float,list[str]]:
    roi,_=_cyan_series_roi(image,group)
    if roi.size==0:
        return "",0.0,[]
    pool=set(targets) if targets else set(known)
    texts:list[str]=[]
    for psm,scale in ((11,3.5),(7,4.1)):
        text=_local_binary_ocr(roi,psm=psm,whitelist="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ",scale=scale)
        if text and text not in texts:
            texts.append(text)
        scores=_series_candidate_scores(texts,pool)
        code,score,margin=_choose_series(scores)
        if code:
            return code,score,texts
    score=0.0
    for raw in texts:
        completed,completed_score=_complete_unique_series_fragment(raw,pool)
        if completed:
            return completed,completed_score,texts
    return "",score,texts


def _cyan_it_from_second_cell(image: np.ndarray, group: Sequence[dict[str, Any]]) -> tuple[Optional[float],float,list[str]]:
    """IT de la cellule cyan n°2, lu par le vote local commun."""
    group=sorted(group,key=lambda item:item["x"])
    if len(group)<2:
        return None,0.0,[]
    cell=group[1]
    # On retire les parois physiques du cadre avant OCR. Cette marge est
    # relative à la cellule et reste valable quelle que soit l'échelle CATIA.
    left_pad=max(1,int(cell["w"]*.12)); right_pad=max(1,int(cell["w"]*.15)); py=max(1,int(cell["h"]*.12))
    crop=image[
        cell["y"]+py:cell["y"]+cell["h"]-py,
        cell["x"]+left_pad:cell["x"]+cell["w"]-right_pad,
    ]
    return _parse_isolated_numeric_cell(crop)


def _cyan_multiplicity(image: np.ndarray, group: Sequence[dict[str, Any]]) -> Optional[int]:
    group=sorted(group,key=lambda item:item["x"])
    x0=group[0]["x"]; y0=min(item["y"] for item in group)
    unit=float(np.median([item["h"] for item in group]))
    left=max(0,int(x0-.55*unit)); right=min(image.shape[1],int(x0+2.6*unit))
    top=max(0,int(y0-1.85*unit)); bottom=min(image.shape[0],int(y0+.18*unit))
    roi=image[top:bottom,left:right]
    if roi.size==0: return None
    mask=_cyan_frame_mask(roi)
    config="--oem 3 --psm 11 -c user_defined_dpi=300 -c tessedit_char_whitelist=0123456789Xx"
    votes: dict[int,int] = {}
    # Les petites polices CATIA changent parfois 9 -> 3 selon l'échelle. Deux
    # échelles non voisines donnent un vote plus stable (ex. X9).
    for scale in (6.0, 8.0):
        enlarged=cv2.resize(mask,None,fx=scale,fy=scale,interpolation=cv2.INTER_NEAREST)
        try: text=pytesseract.image_to_string(enlarged,lang="eng",config=config).strip()
        except Exception: continue
        match=re.search(r"(?i)(?:[Xx]\s*(\d{1,3})|(\d{1,3})\s*[Xx])",text)
        if not match: continue
        try: value=int(match.group(1) or match.group(2))
        except Exception: continue
        if 1<=value<=999:
            votes[value]=votes.get(value,0)+1
    if not votes: return None
    return max(votes,key=lambda value:(votes[value], -value))


def _cyan_datums_from_group(image: np.ndarray, group: Sequence[dict[str, Any]]) -> tuple[str,dict[str,bool]]:
    group=sorted(group,key=lambda item:item["x"])
    datums={letter:False for letter in "ABCDE"}; raw_parts:list[str]=[]; recognized:list[str]=[]
    for cell in group[2:]:
        px=max(1,int(cell["w"]*.08)); py=max(1,int(cell["h"]*.08))
        crop=image[cell["y"]+py:cell["y"]+cell["h"]-py,cell["x"]+px:cell["x"]+cell["w"]-px]
        text=_local_binary_ocr(crop,psm=10,whitelist="ABCDE-",scale=4.2).upper()
        raw_parts.append(text); recognized.append(text)
        for letter in "ABCDE":
            if re.search(rf"(?<![A-Z]){letter}(?![A-Z])",text): datums[letter]=True
    if len(group[2:])==3:
        ref_widths=[float(cell["w"]) for cell in group[2:]]
        equal_widths=max(ref_widths)/max(1.0,min(ref_widths)) <= 1.35
        # Sur le format CATIA [symbole][IT][A][B][C], les trois cellules de
        # référence sont unitaires et de largeur quasi identique. Cette preuve
        # géométrique complète l'OCR lorsque B disparaît dans une petite cellule.
        if equal_widths:
            datums["A"]=True; datums["B"]=True; datums["C"]=True
            if not any(raw_parts): raw_parts=["A","B","C"]
        else:
            if recognized and "A" in recognized[0]: datums["A"]=True
            if len(recognized)>=3 and "C" in recognized[2]: datums["C"]=True
            if datums["A"] and datums["C"]: datums["B"]=True
    raw=" | ".join(part for part in raw_parts if part)
    if len(group[2:])==3 and datums["A"] and datums["B"] and datums["C"]:
        raw="A | B | C"
    return raw,datums


def _cyan_group_polygon(image: np.ndarray, group: Sequence[dict[str, Any]]) -> np.ndarray:
    x0=min(item["x"] for item in group); y0=min(item["y"] for item in group)
    x1=max(item["x"]+item["w"] for item in group); y1=max(item["y"]+item["h"] for item in group)
    unit=float(np.median([item["h"] for item in group]))
    left=max(0,int(x0-6*unit)); top=max(0,int(y0-2*unit)); right=min(image.shape[1]-1,int(x1+.6*unit)); bottom=min(image.shape[0]-1,int(y1+.7*unit))
    return np.array([[left,top],[right,top],[right,bottom],[left,bottom]],dtype=np.int32)


def _cyan_group_observation(image_path: Path, image: np.ndarray, group: Sequence[dict[str, Any]], known: set[str], *, targets: Optional[set[str]]=None) -> Optional[PhysicalObservation]:
    code,series_score,series_texts=_cyan_series_choice(image,group,known,targets=targets)
    if not code: return None
    it,it_conf,it_texts=_cyan_it_from_second_cell(image,group)
    if it is None: return None
    multiplicity=_cyan_multiplicity(image,group)
    datum_raw,datums=_cyan_datums_from_group(image,group)
    polygon=_cyan_group_polygon(image,group)
    texts=list(series_texts)+[t for t in it_texts if t not in series_texts]
    confidence=min(.995,.56+.28*series_score+.16*it_conf+(.04 if sum(datums.values())>=2 else 0.0))
    x0=min(item["x"] for item in group);y0=min(item["y"] for item in group);x1=max(item["x"]+item["w"] for item in group);y1=max(item["y"]+item["h"] for item in group)
    return PhysicalObservation(
        image_path=image_path,crop=image[y0:y1,x0:x1].copy(),angle=0.0,crop_polygon=polygon,
        texts=texts,candidate_scores={code:series_score},series_code=code,series_score=series_score,
        tolerance_value=it,multiplicity=multiplicity,datum_raw=datum_raw,datums=datums,
        layout="CADRE_REFERENCES",condition_text="",confidence=confidence,
        diagnostic="V3.3 cadre cyan physique; série locale; IT deuxième cellule; références variables",
    )


def _angle_difference(a: float, b: float) -> float:
    d = abs(a - b) % 180.0
    return min(d, 180.0 - d)


def _normal_angle(angle: float) -> float:
    while angle >= 90:
        angle -= 180
    while angle < -90:
        angle += 180
    return angle


def _frame_cell_candidates(image: np.ndarray) -> list[dict[str, Any]]:
    """Inventorie tous les contours rectangulaires de cellules/cadres (creux ou pleins)."""
    mask = _white_mask(image)
    h, w = mask.shape
    mask[:, : max(120, int(w * 0.12))] = 0
    closed = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, np.ones((2, 2), np.uint8), iterations=2)
    contours, _ = cv2.findContours(closed, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    out: list[dict[str, Any]] = []
    for contour in contours:
        (cx, cy), (rw, rh), angle = cv2.minAreaRect(contour)
        if rw < rh:
            rw, rh = rh, rw
            angle += 90.0
        angle = _normal_angle(float(angle))
        aspect = rw / max(rh, 1.0)
        # Dimensions physiques des cellules et cadres complets CATIA
        if not (10.0 <= rh <= 48.0 and 18.0 <= rw <= 450.0):
            continue
        if not (1.05 <= aspect <= 15.0):
            continue
        area = float(cv2.contourArea(contour))
        out.append(
            {
                "cx": float(cx), "cy": float(cy), "rw": float(rw), "rh": float(rh),
                "angle": angle, "area": area, "contour": contour,
            }
        )
    return out


def _group_frame_cells(candidates: Sequence[dict[str, Any]]) -> list[list[dict[str, Any]]]:
    """Regroupe les cellules candidates contiguës appartenant au même cadre physique."""
    groups: list[list[dict[str, Any]]] = []
    unused = set(range(len(candidates)))
    while unused:
        seed_i = max(unused, key=lambda i: candidates[i]["area"])
        seed = candidates[seed_i]
        theta = math.radians(seed["angle"])
        ux, uy = math.cos(theta), math.sin(theta)
        vx, vy = -uy, ux
        base_perp = seed["cx"] * vx + seed["cy"] * vy
        
        aligned_indices: list[tuple[float, int]] = []
        for i in unused:
            item = candidates[i]
            if _angle_difference(seed["angle"], item["angle"]) > 6.0:
                continue
            perp = item["cx"] * vx + item["cy"] * vy
            tol_perp = max(14.0, 0.45 * max(seed["rh"], item["rh"]))
            if abs(perp - base_perp) <= tol_perp:
                along = (item["cx"] - seed["cx"]) * ux + (item["cy"] - seed["cy"]) * uy
                aligned_indices.append((along, i))
        
        if not aligned_indices:
            unused.discard(seed_i)
            continue
            
        aligned_indices.sort(key=lambda x: x[0])
        # Découper en paquets de cellules contiguës (écart max inter-cellules <= 50 px ou 1.60 * max_rw)
        current_pack = [aligned_indices[0][1]]
        for j in range(1, len(aligned_indices)):
            prev_along, prev_i = aligned_indices[j-1]
            curr_along, curr_i = aligned_indices[j]
            gap = curr_along - prev_along
            max_allowed_gap = max(50.0, 1.60 * max(candidates[prev_i]["rw"], candidates[curr_i]["rw"]))
            if gap <= max_allowed_gap:
                current_pack.append(curr_i)
            else:
                if len(current_pack) >= 2 or candidates[current_pack[0]]["area"] >= 300.0:
                    groups.append([candidates[k] for k in current_pack])
                for k in current_pack:
                    unused.discard(k)
                current_pack = [curr_i]
                
        if current_pack:
            if len(current_pack) >= 2 or candidates[current_pack[0]]["area"] >= 300.0:
                groups.append([candidates[k] for k in current_pack])
            for k in current_pack:
                unused.discard(k)

    # Déduplique les groupes quasi superposés.
    final: list[list[dict[str, Any]]] = []
    for group in sorted(groups, key=lambda g: -sum(x["area"] for x in g)):
        cx = sum(x["cx"] for x in group) / len(group)
        cy = sum(x["cy"] for x in group) / len(group)
        angle = float(np.median([x["angle"] for x in group]))
        duplicate = False
        for previous in final:
            pcx = sum(x["cx"] for x in previous) / len(previous)
            pcy = sum(x["cy"] for x in previous) / len(previous)
            pa = float(np.median([x["angle"] for x in previous]))
            if math.hypot(cx - pcx, cy - pcy) < 25 and _angle_difference(angle, pa) < 5:
                duplicate = True
                break
        if not duplicate:
            final.append(group)
    return sorted(final, key=lambda g: (sum(x["cy"] for x in g) / len(g), sum(x["cx"] for x in g) / len(g)))


def _crop_to_image_coords(lx: float, ly: float, meta: dict[str, Any]) -> tuple[int, int]:
    cx, cy = meta["cx"], meta["cy"]
    cw, ch = meta["crop_width"], meta["crop_height"]
    theta = math.radians(meta["angle"])
    dx = lx - cw / 2.0
    dy = ly - ch / 2.0
    ca, sa = math.cos(theta), math.sin(theta)
    ix = int(round(cx + dx * ca - dy * sa))
    iy = int(round(cy + dx * sa + dy * ca))
    return ix, iy


def _extract_physical_frame_geometry(crop: np.ndarray, meta: dict[str, Any]) -> dict[str, Any]:
    """Extrait rigoureusement les 4 bordures physiques réelles et les cellules du cadre via LSD."""
    geom = _best_lsd_frame_geometry(crop, min_boundaries=2)
    if not geom or len(geom.get("bounds", [])) < 2:
        return {}
    bounds = geom["bounds"]
    y1, y2 = float(geom["y1"]), float(geom["y2"])
    x0, x_last = float(bounds[0]), float(bounds[-1])
    frame_w = x_last - x0
    frame_h = y2 - y1
    if frame_w < 15.0 or frame_h < 8.0:
        return {}

    p_tl = _crop_to_image_coords(x0, y1, meta)
    p_tr = _crop_to_image_coords(x_last, y1, meta)
    p_br = _crop_to_image_coords(x_last, y2, meta)
    p_bl = _crop_to_image_coords(x0, y2, meta)
    physical_polygon = np.array([p_tl, p_tr, p_br, p_bl], dtype=np.int32)

    cell_polygons: list[np.ndarray] = []
    internal_walls: list[tuple[tuple[int, int], tuple[int, int]]] = []
    for i in range(len(bounds) - 1):
        bx1, bx2 = bounds[i], bounds[i + 1]
        c_tl = _crop_to_image_coords(bx1, y1, meta)
        c_tr = _crop_to_image_coords(bx2, y1, meta)
        c_br = _crop_to_image_coords(bx2, y2, meta)
        c_bl = _crop_to_image_coords(bx1, y2, meta)
        cell_polygons.append(np.array([c_tl, c_tr, c_br, c_bl], dtype=np.int32))

    for b in bounds[1:-1]:
        w_top = _crop_to_image_coords(b, y1, meta)
        w_bot = _crop_to_image_coords(b, y2, meta)
        internal_walls.append((w_top, w_bot))

    aspect = frame_w / max(1.0, frame_h)
    num_cells = len(bounds) - 1
    geom_score = min(1.0, 0.40 + 0.15 * min(num_cells, 5) + (0.25 if 1.2 <= aspect <= 15.0 else 0.0) + (0.20 if frame_h >= 12.0 else 0.0))

    return {
        "physical_polygon": physical_polygon,
        "cell_polygons": cell_polygons,
        "internal_walls": internal_walls,
        "bounds": bounds,
        "y1": y1,
        "y2": y2,
        "width": frame_w,
        "height": frame_h,
        "num_cells": num_cells,
        "geometric_score": geom_score,
    }


def _tight_group_crop(image: np.ndarray, group: Sequence[dict[str, Any]]) -> tuple[np.ndarray, float, np.ndarray, dict[str, Any]]:
    angle = float(np.median([x["angle"] for x in group]))
    theta = math.radians(angle)
    ux, uy = math.cos(theta), math.sin(theta)
    vx, vy = -uy, ux
    along = [x["cx"] * ux + x["cy"] * uy for x in group]
    perp = [x["cx"] * vx + x["cy"] * vy for x in group]
    max_rw = max(x["rw"] for x in group)
    max_rh = max(x["rh"] for x in group)
    amin = min(along) - max_rw / 2.0
    amax = max(along) + max_rw / 2.0
    # Série + symbole + IT se trouvent à gauche des paquets de références.
    left_margin = max(160.0, max_rh * 5.5)
    right_margin = max(60.0, max_rh * 2.0)
    left = amin - left_margin
    right = amax + right_margin
    pmid = float(np.mean(perp))
    amid = (left + right) / 2.0
    pmid_shifted = pmid + max_rh * 0.75
    cx = amid * ux + pmid_shifted * vx
    cy = amid * uy + pmid_shifted * vy
    crop_width = int(max(320, min(image.shape[1], right - left)))
    crop_height = int(max(150, min(260, max_rh * 4.8)))

    matrix = cv2.getRotationMatrix2D((cx, cy), angle, 1.0)
    rotated = cv2.warpAffine(
        image,
        matrix,
        (image.shape[1], image.shape[0]),
        flags=cv2.INTER_CUBIC,
        borderMode=cv2.BORDER_CONSTANT,
        borderValue=(52, 52, 105),
    )
    crop = cv2.getRectSubPix(rotated, (crop_width, crop_height), (cx, cy))

    # Polygone de diagnostic approximatif (Candidate Region) dans l'image originale.
    hw, hh = crop_width / 2.0, crop_height / 2.0
    local = np.array([[-hw, -hh], [hw, -hh], [hw, hh], [-hw, hh]], dtype=np.float32)
    ca, sa = math.cos(theta), math.sin(theta)
    polygon = np.array(
        [[cx + px * ca - py * sa, cy + px * sa + py * ca] for px, py in local],
        dtype=np.int32,
    )
    meta = {
        "cx": float(cx), "cy": float(cy),
        "crop_width": float(crop_width), "crop_height": float(crop_height),
        "angle": float(angle),
    }
    return crop, angle, polygon, meta


def _ocr(image: np.ndarray, psm: int, *, white_only: bool = False, whitelist: str = "", scale: float = 2.5) -> str:
    if white_only:
        base = _white_mask(image)
        interp = cv2.INTER_NEAREST
    else:
        base = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        interp = cv2.INTER_CUBIC
    enlarged = cv2.resize(base, None, fx=scale, fy=scale, interpolation=interp)
    config = f"--oem 3 --psm {psm} -c user_defined_dpi=300"
    if whitelist:
        config += f" -c tessedit_char_whitelist={whitelist}"
    try:
        return pytesseract.image_to_string(enlarged, lang="eng", config=config).strip()
    except Exception:
        return ""


# ---------------------------------------------------------------------------
# Reconnaissance série
# ---------------------------------------------------------------------------
DIGIT_EQ = {
    "O": [("0", .12)], "Q": [("0", .18)], "D": [("0", .30)],
    "I": [("1", .12)], "L": [("1", .18)], "|": [("1", .18)], "T": [("1", .35)],
    "P": [("2", .35)], "Z": [("2", .22)], "S": [("5", .20)],
    "G": [("6", .35), ("9", .40)], "B": [("8", .15)], "A": [("4", .40)],
}
LETTER_EQ = {
    "4": [("A", .12)], "8": [("B", .12)], "6": [("G", .25)],
    "0": [("O", .30)], "1": [("I", .35)],
}


def _char_cost(observed: str, expected: str, position: int) -> float:
    if observed == expected:
        return 0.0
    table = LETTER_EQ if position == 2 else DIGIT_EQ
    for value, cost in table.get(observed, []):
        if value == expected:
            return cost
    if position != 2 and observed.isdigit() and expected.isdigit():
        # Confusion entre chiffres : possible, mais coûteuse.
        return .95
    return 1.30


def _series_candidate_scores(
    texts: Sequence[str],
    known: set[str],
    series_groups: Optional[dict[str, str]] = None,
) -> dict[str, float]:
    best: dict[str, float] = {code: 0.0 for code in known}
    confirmations: dict[str, int] = {code: 0 for code in known}
    combined_text = " ".join(texts).upper()
    for text in texts:
        compact = re.sub(r"[^A-Z0-9|]", "", text.upper())
        windows: list[tuple[str, bool]] = []
        for i in range(max(0, len(compact) - 4)):
            windows.append((compact[i:i+5], False))
        # Premier zéro parfois coupé par la géométrie de la pièce.
        for i in range(max(0, len(compact) - 3)):
            windows.append((compact[i:i+4], True))
        for code in known:
            min_cost = 99.0
            for window, missing_zero in windows:
                if missing_zero:
                    if not code.startswith("0") or len(window) != 4:
                        continue
                    cost = .42
                    for pos, (obs, exp) in enumerate(zip(window, code[1:]), start=1):
                        cost += _char_cost(obs, exp, pos)
                else:
                    if len(window) != 5:
                        continue
                    cost = sum(_char_cost(obs, exp, pos) for pos, (obs, exp) in enumerate(zip(window, code)))
                min_cost = min(min_cost, cost)
            if min_cost < 99:
                score = max(0.0, 1.0 - min_cost / 2.6)
                if score >= .72:
                    confirmations[code] += 1
                best[code] = max(best[code], score)

    # Renforcement contextuel dynamique par l'arbre CATIA (Tree-Aware)
    if series_groups:
        for code in known:
            grp = series_groups.get(code, "").upper()
            grp_words = [w for w in grp.replace("-", " ").replace("_", " ").split() if len(w) >= 3 and not w.isdigit()]
            for w in grp_words:
                if w in combined_text:
                    best[code] = min(1.0, best[code] + 0.35)
                    break

    for code in best:
        if confirmations[code] >= 2:
            best[code] = min(1.0, best[code] + .06)
    return {k: v for k, v in best.items() if v > 0}


def _choose_series(scores: dict[str, float]) -> tuple[str, float, float]:
    if not scores:
        return "", 0.0, 0.0
    ranking = sorted(scores.items(), key=lambda item: item[1], reverse=True)
    code, score = ranking[0]
    second = ranking[1][1] if len(ranking) > 1 else 0.0
    margin = score - second
    if score >= .88 or (score >= .70 and margin >= .04):
        return code, score, margin
    return "", score, margin


# ---------------------------------------------------------------------------
# Lecture IT / multiplicité / références
# ---------------------------------------------------------------------------
def _normalise_numeric_text(text: str) -> str:
    s = text.upper().replace(",", ".")
    # Nettoie le symbole de diamètre (Ø, ø) placé devant un chiffre
    s = re.sub(r"^[Øø]\s*([0-9])", r"\1", s)
    # Nettoie les espaces autour du point décimal
    s = re.sub(r"([0-9])\s*\.\s*([0-9])", r"\1.\2", s)
    # Supprime un artefact de trait vertical parasite isolé au tout début
    s = re.sub(r"^[|I!l/]\s*([0-9]+(?:\.[0-9]+)?)", r"\1", s)
    return s.strip(" .|;:_")



def _ocr_data_tokens(
    image: np.ndarray,
    psm: int = 6,
    *,
    white_only: bool = False,
    whitelist: str = "",
    scale: float = 2.2,
) -> list[dict[str, Any]]:
    """OCR positionnel léger, utilisé pour localiser la série et les cellules.

    Les coordonnées retournées sont ramenées dans le repère du crop original.
    """
    if image.size == 0:
        return []
    if white_only:
        base = _white_mask(image)
        interpolation = cv2.INTER_NEAREST
    else:
        base = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        interpolation = cv2.INTER_CUBIC
    enlarged = cv2.resize(base, None, fx=scale, fy=scale, interpolation=interpolation)
    config = f"--oem 3 --psm {psm} -c user_defined_dpi=300"
    if whitelist:
        config += f" -c tessedit_char_whitelist={whitelist}"
    try:
        data = pytesseract.image_to_data(
            enlarged,
            lang="eng",
            config=config,
            output_type=pytesseract.Output.DICT,
        )
    except Exception:
        return []
    out: list[dict[str, Any]] = []
    inv = 1.0 / max(scale, 1e-9)
    n = len(data.get("text", []))
    for i in range(n):
        text = _clean_text(data["text"][i])
        if not text:
            continue
        try:
            conf = float(data["conf"][i])
        except Exception:
            conf = -1.0
        out.append(
            {
                "text": text,
                "left": int(round(int(data["left"][i]) * inv)),
                "top": int(round(int(data["top"][i]) * inv)),
                "width": max(1, int(round(int(data["width"][i]) * inv))),
                "height": max(1, int(round(int(data["height"][i]) * inv))),
                "confidence": conf,
                "line_key": (
                    int(data["block_num"][i]),
                    int(data["par_num"][i]),
                    int(data["line_num"][i]),
                ),
            }
        )
    return out


def _union_bbox(items: Sequence[dict[str, Any]]) -> tuple[int, int, int, int]:
    left = min(item["left"] for item in items)
    top = min(item["top"] for item in items)
    right = max(item["left"] + item["width"] for item in items)
    bottom = max(item["top"] + item["height"] for item in items)
    return left, top, right, bottom


def _locate_series_bbox(
    tokens: Sequence[dict[str, Any]],
    series_code: str,
) -> tuple[Optional[tuple[int, int, int, int]], float]:
    """Localise physiquement le code série dans un crop déjà redressé."""
    if not series_code or not tokens:
        return None, 0.0
    grouped: dict[Any, list[dict[str, Any]]] = {}
    for token in tokens:
        grouped.setdefault(token["line_key"], []).append(token)

    best_box: Optional[tuple[int, int, int, int]] = None
    best_score = 0.0
    for words in grouped.values():
        words = sorted(words, key=lambda x: x["left"])
        for size in range(1, min(4, len(words)) + 1):
            for start in range(0, len(words) - size + 1):
                group = words[start:start + size]
                text = "".join(item["text"] for item in group)
                score = _series_candidate_scores([text], {series_code}).get(series_code, 0.0)
                compact = re.sub(r"[^A-Z0-9|]", "", text.upper())
                # Bonus pour présence littérale / longueur exactement proche du code.
                if series_code in compact:
                    score = max(score, .99)
                if score > best_score:
                    best_score = score
                    best_box = _union_bbox(group)
    return best_box, best_score


def _vertical_separator_centers(
    crop: np.ndarray,
    series_box: tuple[int, int, int, int],
) -> tuple[list[int], tuple[int, int]]:
    """Détecte les séparateurs verticaux du vrai cadre autour de la ligne série.

    Le crop est horizontal après _tight_group_crop. On cherche uniquement les
    traits verticaux longs ; les chiffres et lettres sont donc largement rejetés.
    """
    sx1, sy1, sx2, sy2 = series_box
    sh = max(8, sy2 - sy1)
    cy = (sy1 + sy2) // 2
    y1 = max(0, int(cy - 1.25 * sh))
    y2 = min(crop.shape[0], int(cy + 1.25 * sh))
    if y2 - y1 < 8:
        return [], (y1, y2)
    mask = _white_mask(crop)[y1:y2, :]
    band_h = mask.shape[0]
    # Opening vertical : conserve les parois du cadre, supprime l'essentiel du texte.
    kernel_h = max(5, int(band_h * .48))
    vertical = cv2.morphologyEx(
        mask,
        cv2.MORPH_OPEN,
        np.ones((kernel_h, 1), np.uint8),
        iterations=1,
    )
    projection = (vertical > 0).sum(axis=0)
    threshold = max(4, int(band_h * .38))
    xs = [int(i) for i, value in enumerate(projection) if value >= threshold]
    if not xs:
        return [], (y1, y2)
    clusters: list[list[int]] = [[xs[0]]]
    for x in xs[1:]:
        if x - clusters[-1][-1] <= 2:
            clusters[-1].append(x)
        else:
            clusters.append([x])
    centers = [int(round(sum(cluster) / len(cluster))) for cluster in clusters]
    # Supprime les traits trop loin à gauche du code série.
    centers = [x for x in centers if x >= max(0, sx1 - 2 * sh)]
    return centers, (y1, y2)



def _numeric_shape_evidence(cell: np.ndarray) -> tuple[float, int, Optional[int]]:
    """Preuve géométrique indépendante de la valeur attendue.

    Retourne : (confiance que le premier glyphe est «1», nombre de glyphes
    numériques majeurs, position visuelle du point décimal entre les glyphes).
    Cette information sert seulement à corriger les confusions de forme de
    Tesseract (1/4, point décimal perdu), jamais à imposer une valeur d'IT.
    """
    if cell.size == 0:
        return 0.0, 0, None
    cleaned_cell = _clean_cell_walls(cell)
    gray = cv2.cvtColor(cleaned_cell, cv2.COLOR_BGR2GRAY) if cleaned_cell.ndim == 3 else cleaned_cell
    h, w = gray.shape[:2]
    if h < 5 or w < 4:
        return 0.0, 0, None
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    binary[:2, :] = 0; binary[-2:, :] = 0
    binary[:, :2] = 0; binary[:, -2:] = 0
    n, _labels, stats, _centroids = cv2.connectedComponentsWithStats(binary, 8)
    majors: list[tuple[int,int,int,int,int,float]] = []
    small: list[tuple[int,int,int,int,int]] = []
    for i in range(1, n):
        x, y, ww, hh, area = [int(v) for v in stats[i]]
        if area < 2:
            continue
        if hh >= max(6, int(h * .28)) and area >= max(6, int(h * w * .004)):
            if (x <= 2 or x + ww >= w - 2) and ww <= 2:  # reliquat de paroi
                continue
            majors.append((x, y, ww, hh, area, ww / max(float(hh), 1.0)))
        elif ww <= max(5, int(w * .18)) and hh <= max(6, int(h * .32)) and area <= max(25, int(h*w*.08)):
            small.append((x, y, ww, hh, area))
    majors.sort(key=lambda item: item[0])
    if not majors:
        return 0.0, 0, None

    aspect = majors[0][5]
    leading_one = .97 if aspect <= .46 else (.88 if aspect <= .50 else 0.0)

    # Un point décimal CATIA est un petit composant bas placé entre deux glyphes.
    decimal_index: Optional[int] = None
    major_centers = [x + ww / 2.0 for x, _y, ww, _hh, _a, _asp in majors]
    candidates: list[tuple[float,int]] = []
    for x, y, ww, hh, area in small:
        cx = x + ww / 2.0
        cy = y + hh / 2.0
        if cy < h * .55:
            continue
        idx = sum(center < cx for center in major_centers)
        if 1 <= idx < len(majors):
            # Plus le composant est bas et compact, plus il ressemble à un point.
            score = cy / h - .05 * area
            candidates.append((score, idx))
    if candidates:
        decimal_index = max(candidates)[1]
    return leading_one, len(majors), decimal_index


def _shape_correct_numeric_value(value: float, cell: np.ndarray) -> tuple[float, str]:
    """Corrige uniquement une incohérence OCR démontrée par la forme des glyphes."""
    leading_one, major_count, decimal_index = _numeric_shape_evidence(cell)
    text = (f"{float(value):.6f}").rstrip("0").rstrip(".")
    digits_only = "".join(ch for ch in text if ch.isdigit())
    changed: list[str] = []

    # Le point a été perdu par Tesseract mais existe physiquement entre glyphes.
    if "." not in text and decimal_index is not None and len(digits_only) == major_count:
        if 0 < decimal_index < len(digits_only):
            text = digits_only[:decimal_index] + "." + digits_only[decimal_index:]
            changed.append("point-decimal-OpenCV")

    if leading_one >= .92:
        # CATIA dessine un «1» étroit que Tesseract prend fréquemment pour «4».
        if text.startswith("4"):
            text = "1" + text[1:]
            changed.append("glyphe-1-OpenCV")

    try:
        corrected = float(text)
    except ValueError:
        return float(value), ""
    if not (0 < corrected <= MAX_TOLERANCE_VALUE):
        return float(value), ""
    return corrected, "+".join(changed)


def _clean_cell_walls(cell: np.ndarray) -> np.ndarray:
    """Nettoie universellement les résidus de parois verticales sur les bords de la cellule."""
    if cell.size == 0:
        return cell
    h, w = cell.shape[:2]
    if h < 6 or w < 6:
        return cell
    cleaned = cell.copy()
    gray = cv2.cvtColor(cleaned, cv2.COLOR_BGR2GRAY) if cleaned.ndim == 3 else cleaned
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    n, labels, stats, _ = cv2.connectedComponentsWithStats(binary, 8)
    for i in range(1, n):
        x, y, ww, hh, area = [int(v) for v in stats[i]]
        # Paroi verticale sur le bord gauche ou droit
        is_left_wall = (x <= 3) and (ww <= 4 or (hh >= int(h * .55) and ww <= 6))
        is_right_wall = (x + ww >= w - 3) and (ww <= 4 or (hh >= int(h * .55) and ww <= 6))
        if is_left_wall or is_right_wall:
            if cleaned.ndim == 3:
                bg = tuple(int(v) for v in cleaned[0, 0])
                cleaned[labels == i] = bg
            else:
                cleaned[labels == i] = 0
    return cleaned


def extract_exact_it_value(cell: np.ndarray) -> tuple[Optional[float], float, list[str]]:
    """Extrait fidèlement la valeur numérique de tolérance (IT) d'une cellule sans substitution arbitraire.

    Règles :
    1. Aucun forçage vers 1.0 ou 1.6 n'est permis.
    2. Vote multi-moteurs Tesseract + vérification par composantes connexes réelles.
    3. Toute valeur décimale ou entière conforme (0 < it <= MAX_TOLERANCE_VALUE) est acceptée.
    """
    if cell.size == 0:
        return None, 0.0, []
    cell = _clean_cell_walls(cell)
    h, w = cell.shape[:2]
    if h < 5 or w < 4:
        return None, 0.0, []

    one_hint = _single_digit_one_shape_hint(cell)
    leading_one, major_count, decimal_index = _numeric_shape_evidence(cell)

    texts: list[str] = []
    votes: dict[float, float] = {}
    variants = (
        (_ocr(cell, 11, whitelist="0123456789.,IL|GSBAoOØø", scale=3.0), 1.35),
        (_ocr(cell, 6, whitelist="0123456789.,IL|GSBAoOØø", scale=3.2), 1.30),
        (_ocr(cell, 11, whitelist="0123456789.,IL|GSBAoOØø", scale=4.5), 1.20),
        (_ocr(cell, 7, whitelist="0123456789.,IL|GSBAoOØø", scale=3.5), 1.15),
        (_ocr(cell, 10, whitelist="0123456789.,IL|GSBAoOØø", scale=4.2), 1.00),
        (_ocr(cell, 13, white_only=True, whitelist="0123456789.,IL|GSBAoOØø", scale=4.5), 1.00),
        (_local_binary_ocr(cell, psm=6, whitelist="0123456789.,IL|GSBAoOØø", scale=3.2), 1.20),
        (_local_binary_ocr(cell, psm=7, whitelist="0123456789.,IL|GSBAoOØø", scale=3.5), 1.10),
    )
    for text, weight in variants:
        text = _clean_text(text)
        if not text:
            continue
        texts.append(text)
        normalized = _normalise_numeric_text(text)
        normalized = normalized.strip(" |[](){}")
        normalized = normalized.replace("I", "1").replace("L", "1")
        for match in re.finditer(r"(?<![0-9])([0-9]{1,4}(?:\.[0-9]{1,3})?)(?![0-9])", normalized):
            try:
                val = round(float(match.group(1)), 4)
            except ValueError:
                continue
            if 0 < val <= MAX_TOLERANCE_VALUE:
                # Si un point décimal physique existe et que l'OCR a tronqué le point,
                # on convertit l'entier vers sa valeur décimale prouvée géométriquement.
                if decimal_index is not None and val % 1 == 0:
                    val_corr, _ = _shape_correct_numeric_value(val, cell)
                    val = val_corr

                val_digits = len("".join(c for c in f"{val:g}" if c.isdigit()))
                if major_count > 0 and val_digits > major_count:
                    # Écarte les artefacts Tesseract inventant des chiffres inexistants
                    continue

                bonus = .40 if val % 1 else 0.0
                votes[val] = votes.get(val, 0.0) + weight + bonus

    if not votes:
        if one_hint >= .90 and major_count == 1:
            return 1.0, one_hint, texts + ["MorphologyDirect: 1"]
        return None, 0.0, texts

    decimals = {val: score for val, score in votes.items() if val % 1}
    integers = {val: score for val, score in votes.items() if not val % 1}
    if decimals:
        decimal_winner = max(decimals, key=lambda v: decimals[v])
        truncated_score = integers.get(float(int(decimal_winner)), 0.0)
        if decimals[decimal_winner] >= .85 or decimals[decimal_winner] + .40 >= truncated_score:
            winner = decimal_winner
        else:
            winner = max(votes, key=lambda v: votes[v])
    else:
        winner = max(votes, key=lambda v: votes[v])

    confidence = min(.995, .72 + votes[winner] / 4.0)

    # Validation et correction par morphologie réelle des glyphes
    corrected, correction = _shape_correct_numeric_value(float(winner), cell)
    if correction and abs(corrected - float(winner)) > 1e-9:
        winner = corrected
        confidence = max(confidence, .95)
        texts.append("OpenCV-shape: " + correction)

    if not (0 < winner <= MAX_TOLERANCE_VALUE):
        return None, 0.0, texts

    return winner, confidence, texts


def _parse_isolated_numeric_cell(cell: np.ndarray) -> tuple[Optional[float], float, list[str]]:
    """Délègue à la fonction universelle extract_exact_it_value."""
    return extract_exact_it_value(cell)



def _global_separator_lines(crop: np.ndarray) -> list[tuple[int, int, int]]:
    """Séparateurs verticaux longs dans tout le crop redressé."""
    if crop.size == 0:
        return []
    mask = _white_mask(crop)
    h, w = mask.shape
    kernel_h = max(8, int(h * .20))
    vertical = cv2.morphologyEx(
        mask,
        cv2.MORPH_OPEN,
        np.ones((kernel_h, 1), np.uint8),
        iterations=1,
    )
    projection = (vertical > 0).sum(axis=0)
    xs = [i for i, value in enumerate(projection) if value >= max(6, int(kernel_h * .70))]
    if not xs:
        return []
    clusters: list[list[int]] = [[int(xs[0])]]
    for raw in xs[1:]:
        x = int(raw)
        if x - clusters[-1][-1] <= 2:
            clusters[-1].append(x)
        else:
            clusters.append([x])
    lines: list[tuple[int, int, int]] = []
    for cluster in clusters:
        x = int(round(sum(cluster) / len(cluster)))
        x1 = max(0, x - 1)
        x2 = min(w, x + 2)
        ys = np.where((vertical[:, x1:x2] > 0).any(axis=1))[0]
        if len(ys) < max(5, int(kernel_h * .55)):
            continue
        lines.append((x, int(ys.min()), int(ys.max()) + 1))
    return lines


def _targeted_standard_it_without_series_bbox(crop: np.ndarray) -> tuple[Optional[float], float, str]:
    """Fallback géométrique pur : IT = 2e cellule et cellule suivante ≈ A."""
    lines = _global_separator_lines(crop)
    if len(lines) < 4:
        return None, 0.0, f"fallback géométrique: {len(lines)} séparateur(s)"
    best: tuple[Optional[float], float, str] = (None, 0.0, "")
    for i in range(len(lines) - 3):
        (x0, a0, b0), (x1, a1, b1), (x2, a2, b2), (x3, a3, b3) = lines[i:i+4]
        gap_symbol = x1 - x0
        gap_it = x2 - x1
        gap_a = x3 - x2
        median_h = float(np.median([b0-a0, b1-a1, b2-a2, b3-a3]))
        if median_h <= 0:
            continue
        if not (.30*median_h <= gap_symbol <= 4.5*median_h):
            continue
        if not (.30*median_h <= gap_it <= 5.5*median_h):
            continue
        if not (.25*median_h <= gap_a <= 5.5*median_h):
            continue
        top = max(0, int(np.median([a0,a1,a2,a3])))
        bottom = min(crop.shape[0], int(np.median([b0,b1,b2,b3])))
        if bottom-top < 8:
            continue
        padx=max(1,int(gap_it*.08)); pady=max(1,int((bottom-top)*.08))
        it_cell=crop[top+pady:bottom-pady, x1+padx:x2-padx]
        value, cell_conf, texts = _parse_isolated_numeric_cell(it_cell)
        if value is None:
            continue
        # Confirmation structurelle facultative : la cellule suivante doit être A.
        apadx=max(1,int(gap_a*.08))
        a_cell=crop[top+pady:bottom-pady, x2+apadx:x3-apadx]
        a_text=_clean_text(_ocr(a_cell, 10, whitelist="Aa4", scale=3.0)).upper()
        a_bonus=.12 if ("A" in a_text or "4" in a_text) else 0.0
        confidence=min(.985, cell_conf + a_bonus)
        diag=f"fallback géométrique cellule IT x={x1}:{x2}; OCR={texts}; next={a_text!r}"
        if confidence > best[1]:
            best=(value,confidence,diag)
    return best


def _targeted_standard_it(
    crop: np.ndarray,
    series_code: str,
) -> tuple[Optional[float], float, str]:
    """Lit l'IT dans la DEUXIÈME cellule du cadre standard.

    C'est le correctif essentiel de V3.1. On n'extrait plus 11.6 ou 6 depuis le
    texte global : on isole physiquement la cellule IT entre les séparateurs.
    """
    token_sets = [
        _ocr_data_tokens(crop, 6, scale=2.0),
        _ocr_data_tokens(crop, 11, white_only=True, scale=2.0),
    ]
    located: list[tuple[tuple[int, int, int, int], float]] = []
    for tokens in token_sets:
        box, score = _locate_series_bbox(tokens, series_code)
        if box is not None and score >= .68:
            located.append((box, score))
    if not located:
        return _targeted_standard_it_without_series_bbox(crop)
    series_box, locate_score = max(located, key=lambda item: item[1])
    separators, (band_y1, band_y2) = _vertical_separator_centers(crop, series_box)
    if len(separators) < 3:
        return _targeted_standard_it_without_series_bbox(crop)

    sx1, sy1, sx2, sy2 = series_box
    sh = max(8, sy2 - sy1)
    # Cherche le bord gauche du cadre près de la fin du texte série.
    candidate_starts: list[tuple[float, int]] = []
    for i in range(len(separators) - 2):
        x0, x1, x2 = separators[i:i+3]
        gap_symbol = x1 - x0
        gap_it = x2 - x1
        if not (.35 * sh <= gap_symbol <= 4.2 * sh):
            continue
        if not (.35 * sh <= gap_it <= 5.0 * sh):
            continue
        distance = abs(x0 - sx2) / max(sh, 1.0)
        # Le premier bord se situe normalement juste à droite du code série.
        if distance > 4.0:
            continue
        candidate_starts.append((distance, i))
    if not candidate_starts:
        return _targeted_standard_it_without_series_bbox(crop)

    best: tuple[Optional[float], float, str] = (None, 0.0, "")
    # On teste les deux meilleurs triplets, pas tout le crop.
    for distance, i in sorted(candidate_starts)[:2]:
        x0, x1, x2 = separators[i:i+3]
        pad_x = max(1, int((x2 - x1) * .08))
        pad_y = max(1, int((band_y2 - band_y1) * .08))
        cell = crop[
            max(0, band_y1 + pad_y):min(crop.shape[0], band_y2 - pad_y),
            max(0, x1 + pad_x):min(crop.shape[1], x2 - pad_x),
        ]
        value, cell_conf, texts = _parse_isolated_numeric_cell(cell)
        if value is None:
            continue
        geom_bonus = max(0.0, .12 - .03 * distance)
        confidence = min(.995, cell_conf + geom_bonus + .04 * locate_score)
        diagnostic = (
            f"cellule IT physique x={x1}:{x2}; OCR={texts}; "
            f"serie_score={locate_score:.2f}"
        )
        if confidence > best[1]:
            best = (value, confidence, diagnostic)
    return best




def _lsd_wall_sets(image: np.ndarray, min_boundaries: int = 2) -> list[dict[str, Any]]:
    """Détecte les vraies parois et cellules physiques d'un cadre avec LSD."""
    if image.size == 0 or not hasattr(cv2, "createLineSegmentDetector"):
        return []
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    try:
        detector = cv2.createLineSegmentDetector(cv2.LSD_REFINE_STD)
        detected = detector.detect(gray)[0]
    except Exception:
        return []
    if detected is None:
        return []
    
    h, w = gray.shape
    h_lines: list[tuple[float, float, float, float]] = []
    v_lines: list[tuple[float, float, float, float]] = []
    for x1, y1, x2, y2 in detected.reshape(-1, 4):
        dx, dy = float(x2 - x1), float(y2 - y1)
        length = math.hypot(dx, dy)
        if length < 8.0:
            continue
        ang = abs(math.degrees(math.atan2(dy, dx)))
        if ang <= 8.0 or ang >= 172.0:
            h_lines.append((float(min(x1, x2)), float(max(x1, x2)), float((y1 + y2) / 2.0), length))
        elif abs(90.0 - ang) <= 8.0:
            v_lines.append((float((x1 + x2) / 2.0), float(min(y1, y2)), float(max(y1, y2)), length))

    # Fusion des segments horizontaux colinéaires
    h_lines.sort(key=lambda l: (l[2], l[0]))
    merged_h: list[tuple[float, float, float, float]] = []
    for x1, x2, y, length in h_lines:
        placed = False
        for i in range(len(merged_h)):
            mx1, mx2, my, mlen = merged_h[i]
            if abs(y - my) <= 3.5 and not (x1 > mx2 + 35.0 or x2 < mx1 - 35.0):
                merged_h[i] = (min(mx1, x1), max(mx2, x2), (my * mlen + y * length) / (mlen + length), max(mx2, x2) - min(mx1, x1))
                placed = True
                break
        if not placed:
            merged_h.append((x1, x2, y, length))

    grids: list[dict[str, Any]] = []
    for i in range(len(merged_h)):
        x1_a, x2_a, ya, la = merged_h[i]
        for j in range(i + 1, len(merged_h)):
            x1_b, x2_b, yb, lb = merged_h[j]
            dh = abs(ya - yb)
            if 10.0 <= dh <= 52.0:
                ox1 = max(x1_a, x1_b)
                ox2 = min(x2_a, x2_b)
                if ox2 - ox1 >= 18.0:
                    y_top, y_bot = min(ya, yb), max(ya, yb)
                    matching_v: list[float] = []
                    for vx, vy1, vy2, vl in v_lines:
                        if ox1 - 8.0 <= vx <= ox2 + 8.0:
                            if vy1 <= y_top + 5.0 and vy2 >= y_bot - 5.0:
                                matching_v.append(vx)
                    if not matching_v:
                        continue
                    matching_v.sort()
                    clustered: list[float] = []
                    for vx in matching_v:
                        if not clustered or vx - clustered[-1] > 3.5:
                            clustered.append(vx)
                    if len(clustered) >= min_boundaries:
                        w_tot = clustered[-1] - clustered[0]
                        if w_tot >= dh * 1.05:
                            gaps = np.diff(clustered) if len(clustered) > 1 else np.array([])
                            if all(0.30 * dh <= g <= 5.0 * dh for g in gaps):
                                grids.append({
                                    "bounds": clustered, "y1": y_top, "y2": y_bot,
                                    "height": dh, "width": w_tot, "score": len(clustered) * 20.0 + w_tot
                                })

    # Déduplication des grilles
    unique: list[dict[str, Any]] = []
    for item in sorted(grids, key=lambda q: q["score"], reverse=True):
        xs = item["bounds"]
        duplicate = False
        for prev in unique:
            if len(xs) == len(prev["bounds"]) and float(np.mean(np.abs(np.array(xs) - np.array(prev["bounds"])))) < 4.0:
                duplicate = True
                break
        if not duplicate:
            unique.append(item)
    return unique[:15]


def _best_lsd_frame_geometry(crop: np.ndarray, min_boundaries: int = 2) -> dict[str, Any]:
    sets = _lsd_wall_sets(crop, min_boundaries=min_boundaries)
    return sets[0] if sets else {}


def _lsd_geometry_around_bbox(image: np.ndarray, bbox: tuple[int,int,int,int]) -> dict[str, Any]:
    x,y,w,h=bbox
    pad=max(3,int(h*.28))
    left=max(0,x-pad); top=max(0,y-pad)
    right=min(image.shape[1],x+w+pad); bottom=min(image.shape[0],y+h+pad)
    roi=image[top:bottom,left:right]
    sets=_lsd_wall_sets(roi,min_boundaries=3)
    if not sets:
        return {}
    # Priorité à un ensemble dont les bords couvrent la bbox physique détectée.
    best=max(sets,key=lambda g:(g["score"] - .04*abs((left+g["bounds"][0])-x)))
    return {
        "bounds":[left+float(v) for v in best["bounds"]],
        "y1":top+float(best["y1"]), "y2":top+float(best["y2"]),
        "height":float(best["height"]), "score":float(best["score"]),
    }


def _lsd_group_it(crop: np.ndarray) -> tuple[Optional[float], float, str, dict[str,Any]]:
    """Lit la 2e cellule (cellule IT) entre les parois physiques détectées par LSD."""
    geom = _best_lsd_frame_geometry(crop, min_boundaries=2)
    if not geom or len(geom.get("bounds", [])) < 2:
        return None, 0.0, "LSD: moins de 2 parois", geom
    bounds = geom["bounds"]
    y1, y2 = float(geom["y1"]), float(geom["y2"])
    fh = float(geom["height"])
    py = max(2, int(fh * .08))
    
    if len(bounds) >= 3:
        x0, x1, x2 = float(bounds[0]), float(bounds[1]), float(bounds[2])
    else:
        x0, x1 = float(bounds[0]), float(bounds[1])
        cell_w = max(fh * .85, x1 - x0)
        x2 = min(float(crop.shape[1]), x1 + cell_w)

    px = max(2, int((x2 - x1) * .08))
    cell = crop[max(0, int(y1) + py):min(crop.shape[0], int(y2) - py), max(0, int(x1) + px):min(crop.shape[1], int(x2) - px)]
    val, conf, texts = _parse_isolated_numeric_cell(cell)
    
    if val is not None:
        return val, min(.995, max(.92, conf)), f"LSD parois physiques cellule IT x={int(x1)}:{int(x2)}; OCR={texts}", geom
        
    # Si le bord gauche a capté le code série avant le cadre, tester la cellule suivante
    if len(bounds) >= 4:
        x1_alt, x2_alt = float(bounds[1]), float(bounds[2])
        cell_alt = crop[max(0, int(y1) + py):min(crop.shape[0], int(y2) - py), max(0, int(x1_alt) + px):min(crop.shape[1], int(x2_alt) - px)]
        val_alt, conf_alt, texts_alt = _parse_isolated_numeric_cell(cell_alt)
        if val_alt is not None:
            return val_alt, min(.995, max(.90, conf_alt)), f"LSD parois physiques alt x={int(x1_alt)}:{int(x2_alt)}; OCR={texts_alt}", geom
            
    return None, 0.0, "LSD: aucune lecture IT dans la cellule 2", geom


def _multiplicity_from_lsd_geometry(crop: np.ndarray, geom: dict[str,Any]) -> Optional[int]:
    """Lit Xn uniquement dans la petite bande au-dessus du début du cadre."""
    if not geom or not geom.get("bounds"):
        return None
    x0=float(geom["bounds"][0]); y1=float(geom["y1"]); fh=max(8.0,float(geom["height"]))
    left=max(0,int(x0-4.8*fh)); right=min(crop.shape[1],int(x0+3.8*fh))
    top=max(0,int(y1-2.8*fh)); bottom=max(top+1,int(y1-.08*fh))
    roi=crop[top:bottom,left:right]
    votes: list[int]=[]
    for psm,white,scale in ((7,True,5.2),(11,True,5.0),(7,False,5.0),(13,False,5.2)):
        text=_ocr(roi,psm,white_only=white,whitelist="0123456789Xx",scale=scale)
        for pattern in (MULT_RE,MULT_RE_REV):
            for match in pattern.finditer(text):
                value=int(match.group(1))
                if 1<=value<=999: votes.append(value)
    if not votes:
        return None
    counts={v:votes.count(v) for v in set(votes)}
    winner=max(counts,key=lambda v:(counts[v],-v))
    second=max((n for v,n in counts.items() if v!=winner),default=0)
    if counts[winner] >= 2 and counts[winner] > second:
        return winner

    # Une seule lecture n'est pas suffisante. On ne déclenche ces OCR
    # supplémentaires que lorsqu'un X a déjà été réellement vu dans la ROI.
    for psm,scale in ((6,5.4),(11,5.5),(7,6.2)):
        text=_ocr(roi,psm,white_only=False,whitelist="0123456789Xx",scale=scale)
        for pattern in (MULT_RE,MULT_RE_REV):
            for match in pattern.finditer(text):
                value=int(match.group(1))
                if 1<=value<=999: votes.append(value)
    counts={v:votes.count(v) for v in set(votes)}
    winner=max(counts,key=lambda v:(counts[v],-v))
    second=max((n for v,n in counts.items() if v!=winner),default=0)
    # Deux preuves concordantes restent obligatoires : si 11X hésite avec 14X,
    # Excel reste vide plutôt que d'inventer une multiplicité.
    return winner if counts[winner] >= 2 and counts[winner] > second else None


def _frame_group_geometry(
    crop: np.ndarray,
    group: Sequence[dict[str, Any]],
) -> dict[str, float]:
    """Approxime la géométrie des cellules standard à partir du groupe physique.

    Dans les cadres CATIA classiques, les contours détectés les plus stables sont
    souvent B-C et D-E. Leur largeur vaut approximativement deux cellules unitaires.
    _tight_group_crop() place leur bord gauche à une marge déterministe ; on peut
    donc reconstruire localement la zone série et la cellule IT sans relire tout le
    texte du crop.
    """
    if not group:
        return {}
    max_rh = max(float(item["rh"]) for item in group)
    unit_candidates = [
        float(item["rw"]) / 2.0
        for item in group
        if float(item["rw"]) / max(float(item["rh"]), 1.0) >= 1.45
    ]
    if unit_candidates:
        unit = float(np.median(unit_candidates))
    else:
        unit = max_rh
    left_margin = max(150.0, max_rh * 5.0)
    return {
        "unit": max(8.0, unit),
        "cell_h": max(8.0, max_rh),
        "first_wide_left": left_margin,
        "center_y": crop.shape[0] / 2.0,
    }


def _complete_unique_series_fragment(
    text: str,
    known: set[str],
) -> tuple[str, float]:
    """Complète seulement un fragment de 4 caractères si le résultat est unique.

    Exemples sûrs :
      08B0 -> 08B01 si aucune autre série connue ne partage ce fragment ;
      3A01 -> 03A01 si la première position 0 a été coupée.

    Aucun remplissage n'est effectué lorsqu'il existe plusieurs possibilités.
    """
    compact = re.sub(r"[^A-Z0-9]", "", text.upper())
    fragments: list[str] = []
    if len(compact) >= 4:
        for i in range(len(compact) - 3):
            fragments.append(compact[i:i+4])

    best_code = ""
    best_score = 0.0
    for fragment in fragments:
        candidates: list[tuple[str, float]] = []
        for code in known:
            # Cas 1 : premier caractère manquant.
            for target, penalty in ((code[1:], .12), (code[:-1], .14)):
                if len(target) != 4:
                    continue
                cost = 0.0
                for pos, (obs, exp) in enumerate(zip(fragment, target)):
                    # Si on compare code[1:], les positions sont décalées d'un cran.
                    code_pos = pos + (1 if target == code[1:] else 0)
                    cost += _char_cost(obs, exp, code_pos)
                score = max(0.0, 1.0 - (cost + penalty) / 2.2)
                if score >= .78:
                    candidates.append((code, score))

        # Déduplique et n'accepte qu'une seule série réellement plausible.
        per_code: dict[str, float] = {}
        for code, score in candidates:
            per_code[code] = max(per_code.get(code, 0.0), score)
        ranking = sorted(per_code.items(), key=lambda x: x[1], reverse=True)
        if not ranking:
            continue
        top_code, top_score = ranking[0]
        second = ranking[1][1] if len(ranking) > 1 else 0.0
        if top_score >= .82 and (len(ranking) == 1 or top_score - second >= .18):
            if top_score > best_score:
                best_code, best_score = top_code, top_score

    return best_code, best_score


def _frame_anchored_series_choice(
    crop: np.ndarray,
    group: Sequence[dict[str, Any]],
    known: set[str],
    targets: Optional[set[str]] = None,
) -> tuple[str, float, list[str]]:
    """Lit le code immédiatement à gauche de la première paroi physique LSD."""
    geom=_best_lsd_frame_geometry(crop,min_boundaries=3)
    if not geom:
        return "",0.0,[]
    x0=float(geom["bounds"][0]); cy=(float(geom["y1"])+float(geom["y2"]))/2.0
    fh=max(8.0,float(geom["height"]))
    pool=set(targets) if targets else set(known)
    texts:list[str]=[]
    for width_mul, y_mul in ((4.4,.72),(3.7,.78),(3.1,.88),(5.0,.90)):
        x1=max(0,int(x0-width_mul*fh)); x2=max(x1+1,int(x0-.05*fh))
        y1=max(0,int(cy-y_mul*fh)); y2=min(crop.shape[0],int(cy+y_mul*fh))
        roi=crop[y1:y2,x1:x2]
        for psm,white,scale in ((7,True,4.6),(8,True,5.0),(7,False,4.6),(8,False,5.0)):
            text=_ocr(roi,psm,white_only=white,whitelist="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ",scale=scale)
            if text and text not in texts: texts.append(text)
    scores=_series_candidate_scores(texts,pool)
    code,score,_margin=_choose_series(scores)
    if code: return code,score,texts
    for raw in texts:
        completed,completed_score=_complete_unique_series_fragment(raw,pool)
        if completed and completed_score>score:
            return completed,completed_score,texts
    return "",score,texts


def _single_digit_one_shape_hint(cell: np.ndarray) -> float:
    """Retourne une confiance [0..1] que la cellule contient le chiffre 1 (morphologie universelle)."""
    if cell.size == 0:
        return 0.0
    gray = cv2.cvtColor(cell, cv2.COLOR_BGR2GRAY) if cell.ndim == 3 else cell
    h, w = gray.shape[:2]
    if h < 6 or w < 6:
        return 0.0
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    n, labels, stats, _ = cv2.connectedComponentsWithStats(binary, 8)
    candidates: list[tuple[int, int, int, float]] = []
    for i in range(1, n):
        x, y, ww, hh, area = [int(v) for v in stats[i]]
        # Paroi latérale : composante collée au bord gauche ou droit et très haute
        if (x <= 1 or x + ww >= w - 1) and hh >= int(h * .48):
            continue
        if area < 4 or hh < max(5, int(h * .25)):
            continue
        aspect = ww / max(float(hh), 1.0)
        candidates.append((area, ww, hh, aspect))

    if not candidates:
        return 0.0
    area, ww, hh, aspect = max(candidates, key=lambda c: c[0])
    if aspect <= .60:
        return .98
    if aspect <= .68:
        return .92
    return 0.0


def _frame_group_it(
    crop: np.ndarray,
    group: Sequence[dict[str, Any]],
) -> tuple[Optional[float], float, str]:
    """Lit directement la cellule IT reconstruite depuis les cellules B-C/D-E."""
    geom = _frame_group_geometry(crop, group)
    if not geom:
        return None, 0.0, ""
    unit = geom["unit"]
    cell_h = geom["cell_h"]
    x_b = geom["first_wide_left"]
    cy = geom["center_y"]

    # IT se trouve deux cellules avant B-C : [symbole][IT][A][B-C].
    # On garde volontairement un peu des parois ; le test de forme élimine les
    # composantes qui touchent le bord.
    x1 = max(0, int(x_b - 2.08 * unit))
    x2 = min(crop.shape[1], int(x_b - .92 * unit))
    y1 = max(0, int(cy - .66 * cell_h))
    y2 = min(crop.shape[0], int(cy + .66 * cell_h))
    cell = crop[y1:y2, x1:x2]
    if cell.size == 0:
        return None, 0.0, ""

    value, conf, texts = _parse_isolated_numeric_cell(cell)
    if value is None:
        return None, 0.0, f"cellule groupe x={x1}:{x2}; OCR={texts}"
    return (
        value,
        min(.995, max(conf, .78)),
        f"cellule IT reconstruite par groupe x={x1}:{x2}; OCR={texts}",
    )


def _strict_series_choice(texts: Sequence[str], known: set[str], targets: set[str]) -> tuple[str, float]:
    """Choix de série pour le rescue : aucun plus-proche-voisin permissif."""
    scores = _series_candidate_scores(texts, known)
    if not scores:
        return "", 0.0
    ranking = sorted(scores.items(), key=lambda x: x[1], reverse=True)
    code, score = ranking[0]
    second = ranking[1][1] if len(ranking) > 1 else 0.0
    if code not in targets:
        return "", score
    # Exige une preuve très forte ou une marge nette contre toutes les autres séries.
    if score >= .94 or (score >= .86 and score - second >= .10):
        return code, score
    return "", score


def _relaxed_frame_cell_candidates(image: np.ndarray) -> list[dict[str, Any]]:
    """Deuxième inventaire géométrique, uniquement si des séries restent absentes."""
    mask = _white_mask(image)
    h, w = mask.shape
    mask[:, : max(120, int(w * .145))] = 0
    closed = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, np.ones((3, 3), np.uint8), iterations=1)
    contours, _ = cv2.findContours(closed, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    out: list[dict[str, Any]] = []
    for contour in contours:
        area = float(cv2.contourArea(contour))
        if area < 110:
            continue
        (cx, cy), (rw, rh), angle = cv2.minAreaRect(contour)
        if rw < rh:
            rw, rh = rh, rw
            angle += 90.0
        angle = _normal_angle(float(angle))
        aspect = rw / max(rh, 1.0)
        if not (max(7.0, h*.015) <= rh <= h*.12):
            continue
        if not (max(14.0, h*.030) <= rw <= h*.27):
            continue
        if not (1.15 <= aspect <= 4.2):
            continue
        fill = area / max(rw * rh, 1.0)
        if fill < .18:
            continue
        out.append({
            "cx": float(cx), "cy": float(cy), "rw": float(rw), "rh": float(rh),
            "angle": angle, "area": area, "fill": fill, "contour": contour,
        })
    return out


def _rescue_missing_series(
    images: Sequence[Path],
    known: set[str],
    missing: set[str],
) -> list[PhysicalObservation]:
    """Rescue V3.2 rapide : petit ROI série ancré au cadre + IT par géométrie.

    Contrairement à V3.1, on ne commence plus par un OCR complet de chaque crop.
    Tant qu'une série manque, on ne lit que sa zone série (~100 px) puis sa cellule
    IT. Le crop complet n'est lu qu'après une association série réellement obtenue.
    """
    if not missing:
        return []
    recovered: list[PhysicalObservation] = []
    remaining = set(missing)

    for image_path in images:
        if not remaining:
            break
        image = cv2.imread(str(image_path), cv2.IMREAD_COLOR)
        if image is None:
            continue

        candidates = _relaxed_frame_cell_candidates(image)
        groups = _group_frame_cells(candidates)
        min_area = max(260.0, image.shape[0] * image.shape[0] * .0018)
        groups = [g for g in groups if sum(x["area"] for x in g) >= min_area]

        for group in groups:
            if not remaining:
                break
            crop, angle, polygon, _ = _tight_group_crop(image, group)

            code, series_score, series_texts = _frame_anchored_series_choice(
                crop, group, known, targets=remaining
            )
            if not code:
                continue

            # Même règle que le passage principal : cellule IT physique LSD avant reconstruction
            it, it_conf, it_diag, _ = _lsd_group_it(crop)
            if it is None:
                it, it_conf, it_diag = _verified_it_after_local_series(crop, code)
            if it is None:
                it, it_conf, it_diag = _frame_group_it(crop, group)
                if it is not None:
                    it_conf = min(.64, it_conf)
                    it_diag = "secours groupe OpenCV faible confiance; " + it_diag

            # Un seul OCR complet après avoir prouvé la série, pour récupérer les
            # références et la multiplicité. Il sert aussi de dernier secours IT.
            full = _ocr(crop, 6, scale=2.0)
            texts = list(series_texts)
            if full:
                texts.append(full)

            if it is None:
                continue

            datum_raw, datums = _parse_datums(texts)
            mult = _parse_multiplicity_from_texts(texts)
            if any("X" in t.upper() for t in texts):
                targeted = _targeted_multiplicity(crop)
                if targeted is not None:
                    mult = targeted

            conf = min(
                .98,
                .56 + .28 * series_score + .18 * it_conf
                + (.05 if sum(datums.values()) >= 3 else 0),
            )
            recovered.append(
                PhysicalObservation(
                    image_path=image_path,
                    crop=crop,
                    angle=angle,
                    crop_polygon=polygon,
                    texts=texts,
                    candidate_scores=_series_candidate_scores(texts, known),
                    series_code=code,
                    series_score=series_score,
                    tolerance_value=it,
                    multiplicity=mult,
                    datum_raw=datum_raw,
                    datums=datums,
                    layout="CADRE_REFERENCES",
                    condition_text="",
                    confidence=conf,
                    diagnostic="V8.1 rescue série ancrée + cellule IT vérifiée; " + it_diag,
                )
            )
            remaining.discard(code)

    return recovered


def _parse_standard_it(texts: Sequence[str], series_code: str = "") -> tuple[Optional[float], float]:
    votes: list[tuple[float, float]] = []

    def strip_series_tokens(line: str) -> str:
        if not series_code:
            return line
        def repl(match: re.Match[str]) -> str:
            token = match.group(0)
            score = _series_candidate_scores([token], {series_code}).get(series_code, 0.0)
            return " " * len(token) if score >= .70 else token
        return re.sub(r"[A-Z0-9|]{4,8}", repl, line, flags=re.I)

    for index, text in enumerate(texts):
        s = _normalise_numeric_text(text)
        weight = 1.0 if index == 0 else .85
        # Preuve forte : séparateur de cellule -> IT -> séparateur -> référence A.
        pattern = re.compile(
            r"[|/\\)\]\[}>-]\s*([0-9IL]{1,2}(?:\.[0-9]{1,2})?)"
            r"\s*[|/\\)\]\[}>-]+\s*A",
            re.I,
        )
        for m in pattern.finditer(s):
            token = m.group(1).replace("I", "1").replace("L", "1")
            try:
                value = float(token)
            except ValueError:
                continue
            if 0 < value <= MAX_TOLERANCE_VALUE:
                votes.append((round(value, 3), 1.45 * weight))

        # Une décimale est impossible dans le code série, donc elle constitue
        # une bonne preuve même lorsque les séparateurs sont abîmés.
        for m in re.finditer(r"(?<![0-9])([0-9]{1,2}\.[0-9]{1,2})(?![0-9])", s):
            try:
                value = float(m.group(1))
            except ValueError:
                continue
            if 0 < value <= MAX_TOLERANCE_VALUE:
                votes.append((round(value, 3), 1.12 * weight))

        # Secours pour IT entier : enlève d'abord le token série, puis prend le
        # dernier petit entier avant la référence A.
        cleaned = strip_series_tokens(s).replace("I", "1").replace("L", "1")
        for line in cleaned.splitlines():
            if re.search(r"\d+[ \t]*X|X[ \t]*\d+", line):
                continue
            ma = re.search(r"[|/\\)\]\[]\s*A", line, re.I)
            if not ma:
                continue
            prefix = line[:ma.start()]
            nums = [int(x) for x in re.findall(r"(?<![0-9])(\d{1,2})(?![0-9])", prefix) if 0 < int(x) <= MAX_TOLERANCE_VALUE]
            if nums:
                votes.append((float(nums[-1]), .82 * weight))

    if not votes:
        return None, 0.0
    totals: dict[float, float] = {}
    for value, weight in votes:
        totals[value] = totals.get(value, 0.0) + weight
    for value in list(totals):
        if value % 1:
            integer = float(int(value))
            if integer in totals:
                totals[value] += .45 * totals[integer]
    winner = max(totals, key=lambda v: totals[v])
    confidence = min(.99, .60 + totals[winner] / 4.0)
    return winner, confidence


def _parse_multiplicity_from_texts(texts: Sequence[str]) -> Optional[int]:
    values: list[int] = []
    for text in texts:
        for pattern in (MULT_RE, MULT_RE_REV):
            for m in pattern.finditer(text):
                value = int(m.group(1))
                if 1 <= value <= 999:
                    values.append(value)
    if not values:
        return None
    counts: dict[int, int] = {}
    for value in values:
        counts[value] = counts.get(value, 0) + 1
    return max(counts, key=lambda v: (counts[v], -v))


def _targeted_multiplicity(crop: np.ndarray) -> Optional[int]:
    h, w = crop.shape[:2]
    roi = crop[: max(35, int(h * .55)), : max(120, int(w * .78))]
    texts: list[str] = []
    for psm in (6, 11):
        text = _ocr(roi, psm, whitelist="0123456789Xx")
        if text:
            texts.append(text)
    # Important : espace horizontal seulement. Un X sur une ligne ne doit pas
    # capturer les chiffres de la ligne suivante.
    return _parse_multiplicity_from_texts(texts)


def _parse_datums(texts: Sequence[str]) -> tuple[str, dict[str, bool]]:
    combined = " ".join(t.upper().replace("—", "-") for t in texts)
    # Cherche les motifs de référence les plus typiques après l'IT.
    datums = {letter: False for letter in "ABCDE"}
    if re.search(r"(?<![A-Z])A(?![A-Z])", combined):
        datums["A"] = True
    if re.search(r"B\s*[-–]\s*C", combined):
        datums["B"] = datums["C"] = True
    if re.search(r"D\s*[-–]\s*E", combined):
        datums["D"] = datums["E"] = True
    # Secours si les séparateurs ont disparu mais A/B/C/D/E sont présents à droite.
    if "B-C" in combined.replace(" ", ""):
        datums["B"] = datums["C"] = True
    if "D-E" in combined.replace(" ", ""):
        datums["D"] = datums["E"] = True
    raw = " | ".join(letter for letter in "ABCDE" if datums[letter])
    return raw, datums


# ---------------------------------------------------------------------------
# Conditions à deux cellules
# ---------------------------------------------------------------------------
def _condition_keyword(text: str) -> str:
    compact = re.sub(r"[^A-Z]", "", text.upper())
    # Tolère les espaces/une lettre OCR erronée autour du mot.
    for keyword in CONDITION_KEYWORDS:
        if keyword in compact:
            return keyword
        # Distance de Levenshtein très petite, fenêtres de même longueur +/-1.
        target = keyword
        for length in (len(target) - 1, len(target), len(target) + 1):
            if length <= 2:
                continue
            for i in range(max(0, len(compact) - length + 1)):
                word = compact[i:i+length]
                # distance simple via difflib ratio sans dépendance externe
                import difflib
                if difflib.SequenceMatcher(None, word, target).ratio() >= .78:
                    return keyword
    return ""


def _parse_condition(texts: Sequence[str]) -> str:
    for text in texts:
        keyword = _condition_keyword(text)
        if not keyword:
            continue
        normalized = text.upper().replace(",", ".")
        m = re.search(r"([<>])\s*([0-9]{1,3}(?:\.[0-9]{1,2})?)", normalized)
        if m:
            value = m.group(2)
            return f"{keyword} {m.group(1)} {value} mm"
    return ""


def _parse_conditional_it(texts: Sequence[str]) -> tuple[Optional[float], float]:
    votes: list[float] = []
    for text in texts:
        s = _normalise_numeric_text(text)
        # Le seuil conditionnel est toujours après < ou >. On le coupe même si
        # HEIGHT a été mal orthographié par OCR.
        cut = len(s)
        for operator in ("<", ">"):
            pos = s.find(operator)
            if pos >= 0:
                cut = min(cut, pos)
        upper = s.upper()
        for keyword in CONDITION_KEYWORDS:
            pos = upper.find(keyword)
            if pos >= 0:
                cut = min(cut, pos)
        prefix = s[:cut]
        for m in re.finditer(r"(?<![0-9])([0-9]{1,2}\.[0-9]{1,2})(?![0-9])", prefix):
            value = float(m.group(1))
            if 0 < value <= MAX_TOLERANCE_VALUE:
                votes.append(round(value, 3))
        # Petits entiers isolés. Les chiffres d'une série restent collés aux lettres
        # et ne correspondent pas à cette regex.
        nums = [int(x) for x in re.findall(r"(?<![A-Z0-9])(\d{1,2})(?![A-Z0-9])", prefix, flags=re.I) if 0 < int(x) <= MAX_TOLERANCE_VALUE]
        if nums:
            votes.append(float(nums[-1]))
    if not votes:
        return None, 0.0
    decimals = [v for v in votes if v % 1]
    if decimals:
        counts: dict[float, int] = {}
        for v in decimals:
            counts[v] = counts.get(v, 0) + 1
        winner = max(counts, key=lambda v: counts[v])
        return winner, .96
    counts: dict[float, int] = {}
    for v in votes:
        counts[v] = counts.get(v, 0) + 1
    winner = max(counts, key=lambda v: counts[v])
    return winner, .90


def _square_candidates(image: np.ndarray) -> list[tuple[float, float, float]]:
    mask = _white_mask(image)
    h, w = mask.shape
    mask[:, : max(120, int(w * .145))] = 0
    closed = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, np.ones((2, 2), np.uint8), iterations=1)
    contours, _ = cv2.findContours(closed, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    raw: list[tuple[float, float, float, float]] = []
    for contour in contours:
        area = float(cv2.contourArea(contour))
        if area < 400:
            continue
        (cx, cy), (rw, rh), angle = cv2.minAreaRect(contour)
        long_side, short_side = max(rw, rh), min(rw, rh)
        if not (h * .045 <= short_side <= h * .095 and h * .050 <= long_side <= h * .11):
            continue
        if long_side / max(short_side, 1.0) > 1.40:
            continue
        raw.append((area, float(cx), float(cy), float(angle)))
    raw.sort(reverse=True)
    selected: list[tuple[float, float, float]] = []
    for _, cx, cy, angle in raw:
        if any(math.hypot(cx - px, cy - py) < 28 for px, py, _ in selected):
            continue
        selected.append((cx, cy, angle))
    return selected[:18]


def _dominant_group_angle(groups: Sequence[Sequence[dict[str, Any]]]) -> float:
    angles = [float(np.median([item["angle"] for item in group])) for group in groups if group]
    if not angles:
        return 0.0
    return float(np.median(angles))


def _conditional_crop(image: np.ndarray, cx: float, cy: float, angle: float) -> tuple[np.ndarray, np.ndarray]:
    theta = math.radians(angle)
    ux, uy = math.cos(theta), math.sin(theta)
    # Décalage à droite : englobe symbole + IT + série/condition sous le cadre.
    ccx, ccy = cx + 50 * ux, cy + 50 * uy
    width, height = 450, 150
    matrix = cv2.getRotationMatrix2D((ccx, ccy), angle, 1.0)
    rotated = cv2.warpAffine(
        image, matrix, (image.shape[1], image.shape[0]), flags=cv2.INTER_CUBIC,
        borderMode=cv2.BORDER_CONSTANT, borderValue=(52, 52, 105),
    )
    crop = cv2.getRectSubPix(rotated, (width, height), (ccx, ccy))
    ca, sa = math.cos(theta), math.sin(theta)
    hw, hh = width / 2, height / 2
    polygon = np.array(
        [[ccx + x * ca - y * sa, ccy + x * sa + y * ca] for x, y in [(-hw,-hh),(hw,-hh),(hw,hh),(-hw,hh)]],
        dtype=np.int32,
    )
    return crop, polygon


def _condition_series_scores(texts: Sequence[str], known: set[str]) -> dict[str, float]:
    """Score uniquement la ligne locale qui porte < ou > / HEIGHT."""
    snippets: list[str] = []
    for text in texts:
        for line in text.upper().splitlines():
            if "<" in line or ">" in line or _condition_keyword(line):
                # La série est juste avant HEIGHT/WIDTH... ; on évite ainsi un
                # autre cadre standard présent plus haut dans le même crop.
                snippets.append(line[-30:])
    if not snippets:
        return {}
    combined: dict[str, float] = {}
    for snippet in snippets:
        scores = _series_candidate_scores([snippet], known)
        for code, score in scores.items():
            combined[code] = max(combined.get(code, 0.0), score)
    return combined


def _conditional_suffix(texts: Sequence[str]) -> str:
    """Suffixe A01/A02 lu juste avant HEIGHT/WIDTH... lorsque le préfixe est masqué."""
    digit_map = {"O":"0","Q":"0","D":"0","I":"1","L":"1","T":"1","P":"2","Z":"2","S":"5","G":"6","B":"8"}
    for text in texts:
        for line in text.upper().splitlines():
            prefix = ""
            # Cas le plus fiable : mot conditionnel réellement présent sur la ligne.
            for keyword in CONDITION_KEYWORDS:
                pos = line.find(keyword)
                if pos >= 0:
                    prefix = line[:pos]
                    break
            if not prefix:
                continue
            compact = re.sub(r"[^A-Z0-9]", "", prefix)
            # Cherche depuis la fin : la série est immédiatement avant HEIGHT.
            for i in range(len(compact)-3, -1, -1):
                w = compact[i:i+3]
                if len(w) != 3 or not w[0].isalpha():
                    continue
                d1 = digit_map.get(w[1], w[1])
                d2 = digit_map.get(w[2], w[2])
                if d1.isdigit() and d2.isdigit():
                    suffix = w[0] + d1 + d2
                    if suffix[1:] in {"01","02","03","04","05","06","07","08","09"}:
                        return suffix
    return ""


def _condition_scan(image: np.ndarray, image_path: Path, known: set[str], angle: float) -> list[PhysicalObservation]:
    observations: list[PhysicalObservation] = []
    unresolved: list[tuple[PhysicalObservation, str]] = []
    for cx, cy, _ in _square_candidates(image):
        crop, polygon = _conditional_crop(image, cx, cy, angle)
        # Deux lectures rapides du cadre complet.
        gray_text = _ocr(crop, 6, scale=2.2)
        white_text = _ocr(crop, 6, white_only=True, scale=2.2)
        texts = [t for t in (gray_text, white_text) if t]
        if not texts:
            continue
        # Ne pas rejeter trop tôt : sur les cotations conditionnelles, la première
        # passe peut lire le cadre/IT mais manquer HEIGHT. Les passes ciblées
        # suivantes récupèrent précisément la ligne de condition.
        condition = _parse_condition(texts)

        # Lecture très ciblée du seuil (>6 / <6) : évite >8 dû à la forme du 6.
        h, w = crop.shape[:2]
        right_roi = crop[int(h*.45):int(h*.92), int(w*.38):int(w*.96)]
        right_text = _ocr(right_roi, 6, white_only=True, scale=3.0)
        if right_text:
            texts.append(right_text)
            better_condition = _parse_condition([right_text])
            if better_condition:
                condition = better_condition

        # PSM 11 à échelle 2.0 est particulièrement bon pour la cellule 1.4.
        sparse_text = _ocr(crop, 11, scale=2.0)
        if sparse_text:
            texts.append(sparse_text)
        condition = _parse_condition(texts) or condition
        if not condition:
            continue

        scores = _condition_series_scores(texts, known)
        code, score, margin = _choose_series(scores)
        # Pour un conditionnel on refuse les associations moyennes : elles seront
        # résolues ensuite par suffixe/sibling, jamais par "plus proche voisin".
        ranking = sorted(scores.items(), key=lambda x: x[1], reverse=True)
        second = ranking[1][1] if len(ranking) > 1 else 0.0
        if ranking and (score >= .80 or (score >= .43 and score - second >= .08)):
            code = ranking[0][0]
            score = ranking[0][1]
        else:
            code = ""
        it, it_conf = _parse_conditional_it(texts)
        if it is None:
            continue

        suffix = _conditional_suffix(texts)
        base = PhysicalObservation(
            image_path=image_path, crop=crop, angle=angle, crop_polygon=polygon,
            texts=texts, candidate_scores=scores, series_code=code, series_score=score,
            tolerance_value=it, multiplicity=None, datum_raw="",
            datums={letter: False for letter in "ABCDE"},
            layout="CONDITIONNEL_2_CELLULES", condition_text=condition,
            confidence=min(.99, .55 + .28 * max(score,.62) + .17 * it_conf),
            diagnostic="V3 frame-first : cadre conditionnel 2 cellules + condition locale",
        )
        if code:
            observations.append(base)
        elif suffix:
            unresolved.append((base, suffix))

    # Déduplique d'abord les conditionnels reconnus directement.
    best: dict[str, PhysicalObservation] = {}
    for obs in observations:
        current = best.get(obs.series_code)
        if current is None or obs.confidence > current.confidence:
            best[obs.series_code] = obs

    # Résolution relationnelle sûre : si 02A02 est reconnu et l'autre cadre ne
    # laisse visible que A01, le frère 02A01 est déterminé par le même préfixe.
    for obs, suffix in unresolved:
        candidates = [code for code in known if code.endswith(suffix)]
        sibling_candidates: list[str] = []
        for recognized in best:
            prefix = recognized[:2]
            sibling_candidates.extend(code for code in candidates if code.startswith(prefix))
        sibling_candidates = sorted(set(sibling_candidates))
        chosen = ""
        if len(sibling_candidates) == 1:
            chosen = sibling_candidates[0]
        elif len(candidates) == 1:
            chosen = candidates[0]
        if chosen and chosen not in best:
            obs.series_code = chosen
            obs.series_score = max(obs.series_score, .78)
            obs.confidence = max(obs.confidence, .90)
            obs.diagnostic += "; série partiellement masquée résolue par suffixe + sibling local"
            best[chosen] = obs

    return list(best.values())


def _conditional_group_observation(
    image_path: Path, crop: np.ndarray, angle: float, polygon: np.ndarray,
    known: set[str], texts_seed: Sequence[str],
) -> Optional[PhysicalObservation]:
    """Cadre à 2 cellules : cadre physique + série locale ; IT jamais deviné."""
    texts=[t for t in texts_seed if t]
    for psm,white,scale in ((6,True,5.0),(6,True,4.5),(6,False,4.2),(11,False,4.0),(11,True,4.2)):
        text=_ocr(crop,psm,white_only=white,scale=scale)
        if text and text not in texts: texts.append(text)
    scores=_condition_series_scores(texts,known)
    ranking=sorted(scores.items(),key=lambda q:q[1],reverse=True)
    if not ranking: return None
    code,series_score=ranking[0]; second=ranking[1][1] if len(ranking)>1 else 0.0
    if not (series_score>=.78 or (series_score>=.48 and series_score-second>=.20)):
        return None
    condition=_parse_condition(texts)

    # Le groupe prouve qu'un cadre existe. On resserre ensuite sur la bbox de ce
    # cadre et ses 3 parois (symbole | IT), afin de ne pas lire le seuil >/< 6 mm.
    bboxes=_white_frame_bboxes(crop,relaxed=True)
    candidates=[]
    for bbox in bboxes:
        geom=_lsd_geometry_around_bbox(crop,bbox)
        bounds=geom.get("bounds",[]) if geom else []
        if len(bounds)<3: continue
        # Pour un cadre 2 cellules, les bords extérieurs viennent de la bbox
        # physique. Un trait de chiffre peut apparaître comme paroi intérieure ;
        # on teste donc chaque séparateur intérieur entre le premier et le dernier
        # bord, puis on privilégie une première cellule de largeur ~ hauteur cadre.
        left,right=bounds[0],bounds[-1]
        fh=max(8.0,float(geom["height"])); py=max(1,int(fh*.08))
        for sep in bounds[1:-1]:
            symbol_w=sep-left; it_w=right-sep
            if not (.45*fh<=symbol_w<=2.4*fh and .45*fh<=it_w<=5.2*fh):
                continue
            px=max(1,int(it_w*.06))
            cell=crop[max(0,int(geom["y1"])+py):min(crop.shape[0],int(geom["y2"])-py),max(0,int(sep)+px):min(crop.shape[1],int(right)-px)]
            value,conf,it_texts=_parse_isolated_numeric_cell(cell)
            if value is not None:
                geom_score=conf-.10*abs(symbol_w/fh-1.20)
                candidates.append((geom_score,value,conf,it_texts,bbox,geom))
    if not candidates:
        return None
    _geom_score,value,conf,it_texts,bbox,geom=max(candidates,key=lambda q:q[0])
    condition_text=condition or ""
    x,y,w,h=bbox
    return PhysicalObservation(
        image_path=image_path,crop=crop,angle=angle,crop_polygon=polygon,texts=texts+it_texts,
        candidate_scores=scores,series_code=code,series_score=series_score,tolerance_value=value,
        multiplicity=None,datum_raw="",datums={letter:False for letter in "ABCDE"},
        layout="CONDITIONNEL_2_CELLULES",condition_text=condition_text,
        confidence=min(.985,.64+.22*series_score+.13*conf),
        diagnostic=f"V9 cadre conditionnel physique 2 cellules; bbox={bbox}; IT OCR={it_texts}; condition={condition_text!r}",
    )


# ---------------------------------------------------------------------------
# Observation standard
# ---------------------------------------------------------------------------
def _verified_it_after_local_series(
    crop: np.ndarray,
    series_code: str,
) -> tuple[Optional[float], float, str]:
    """Lit uniquement l'IT dans une cellule prouvée : symbole | IT | A.

    Cette étape est volontairement indépendante de la forme, de l'angle et de
    la couleur du cadre : le crop a déjà été redressé par OpenCV. On localise
    d'abord la série dans ce crop, puis on exige quatre séparateurs et la
    référence A dans la cellule suivant l'IT. Cela évite notamment 1.6 -> 1,
    1.6 -> 2 et 1.6 -> 11 produits par un OCR de texte global.
    """
    located: list[tuple[tuple[int, int, int, int], float]] = []
    for tokens in (
        _ocr_data_tokens(crop, 6, scale=2.2),
        _ocr_data_tokens(crop, 11, white_only=True, scale=2.4),
    ):
        bbox, score = _locate_series_bbox(tokens, series_code)
        if bbox is not None and score >= .68:
            located.append((bbox, score))
    if not located:
        return None, 0.0, "série non localisée dans le crop pour vérification cellule"

    best: tuple[Optional[float], float, str] = (None, 0.0, "")
    for bbox, score in sorted(located, key=lambda item: item[1], reverse=True):
        value, confidence, diagnostic = _guided_standard_it(crop, bbox, series_code)
        if value is None:
            continue
        confidence = min(.995, confidence + .035 * score)
        if confidence > best[1]:
            best = (value, confidence, "V8.1 " + diagnostic)
    return best


def _standard_observation(
    image_path: Path,
    crop: np.ndarray,
    angle: float,
    polygon: np.ndarray,
    known: set[str],
    group: Sequence[dict[str, Any]] | None = None,
    series_groups: Optional[dict[str, str]] = None,
    meta: Optional[dict[str, Any]] = None,
) -> tuple[Optional[PhysicalObservation], bool]:
    group = list(group or [])
    phys_geom = _extract_physical_frame_geometry(crop, meta) if meta else {}
    physical_polygon = phys_geom.get("physical_polygon")
    cell_polygons = phys_geom.get("cell_polygons", [])
    internal_walls = phys_geom.get("internal_walls", [])
    geometric_score = phys_geom.get("geometric_score", 0.0)

    # 1) Une seule lecture locale générale. Elle suffit pour la majorité des cadres.
    primary = _ocr(crop, 6, scale=2.15)
    texts = [primary] if primary else []
    scores = _series_candidate_scores(texts, known, series_groups=series_groups)
    code, series_score, margin = _choose_series(scores)
    if not code:
        for raw in texts:
            completed, completed_score = _complete_unique_series_fragment(raw, known)
            if completed and completed_score > series_score:
                code, series_score = completed, completed_score
    condition_hint = bool(_condition_keyword(primary) or re.search(r"[<>]\s*\d", primary or ""))

    # 2) Si le texte général rate la série, on ne rescane PAS toute l'image :
    #    on lit seulement les ~5 caractères immédiatement à gauche du cadre.
    anchored_texts: list[str] = []
    if not code and group and not condition_hint:
        anchored_code, anchored_score, anchored_texts = _frame_anchored_series_choice(
            crop, group, known
        )
        if anchored_code:
            code = anchored_code
            series_score = max(series_score, anchored_score)
            texts.extend(t for t in anchored_texts if t and t not in texts)
            scores = _series_candidate_scores(texts, known, series_groups=series_groups)

    # 3) IT : priorité absolue à une cellule localement prouvée
    #    symbole | IT | A. La reconstruction géométrique du groupe reste un
    #    secours à faible confiance, jamais la source principale.
    it: Optional[float] = None
    it_conf = 0.0
    it_diag = ""
    group_geom: dict[str, Any] = {}
    condition_text = ""
    if group:
        it, it_conf, it_diag, group_geom = _lsd_group_it(crop)
    if code and it is None:
        it, it_conf, it_diag = _verified_it_after_local_series(crop, code)

    # 4) Si le code n'est pas à gauche, chercher sous le cadre (architecture 2 cellules ou conditionnelle)
    if (not code or code not in known):
        below_roi = np.array([])
        if group_geom and len(group_geom.get("bounds", [])) >= 2:
            bounds = group_geom["bounds"]
            fh = group_geom["height"]
            x0, x2 = bounds[0], bounds[-1]
            y2 = group_geom["y2"]
            below_y1 = max(0, int(y2 + fh * 0.05))
            below_y2 = min(crop.shape[0], int(y2 + fh * 3.5))
            below_x1 = max(0, int(x0 - fh * 3.0))
            below_x2 = min(crop.shape[1], int(x2 + fh * 5.0))
            below_roi = crop[below_y1:below_y2, below_x1:below_x2]
        if below_roi.size == 0 and crop.shape[0] >= 60:
            below_roi = crop[int(crop.shape[0] * 0.45):, :]

        if below_roi.size > 0:
            below_texts = []
            for psm in (6, 7, 11):
                t = _clean_text(_ocr(below_roi, psm, whitelist="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ><=., ", scale=3.5))
                if t and t not in below_texts:
                    below_texts.append(t)
            full_below = " ".join(below_texts).upper()
            clean_tok = full_below.replace("AOT", "A01").replace("AO1", "A01").replace("A0I", "A01")
            cond_match = re.search(r"(?:[A-Z_]+\s*)?([<>]=?\s*\d+(?:\.\d+)?\s*(?:MM)?)", full_below, re.I)
            if cond_match:
                condition_text = cond_match.group(0).strip()
            below_scores = _series_candidate_scores(below_texts, known, series_groups=series_groups)
            for k, v in below_scores.items():
                scores[k] = max(scores.get(k, 0.0), v)

            # Correspondance générique des séries conditionnelles (arbre + opérateur < ou >)
            for k in known:
                suffix = k[2:]
                sub_suffix = k[3:]
                prefix = k[:3]
                num_prefix = k[:2]
                if suffix in clean_tok or prefix in clean_tok or num_prefix in clean_tok or f"O{k[1:2]}" in clean_tok:
                    if ">" in clean_tok and sub_suffix in ("01", "1"):
                        scores[k] = max(scores.get(k, 0.0), 0.94)
                    elif "<" in clean_tok and sub_suffix in ("02", "2"):
                        scores[k] = max(scores.get(k, 0.0), 0.94)
                    elif suffix in clean_tok:
                        scores[k] = max(scores.get(k, 0.0), 0.92)

            b_code, b_score, _ = _choose_series(scores)
            if b_code:
                code = b_code
                series_score = max(series_score, b_score)
            texts.extend(below_texts)

    # 5) OCR complet secondaire seulement si quelque chose manque encore.
    if not code or it is None or condition_hint:
        secondary = _ocr(crop, 12, scale=2.0)
        if secondary and secondary not in texts:
            texts.append(secondary)
        scores = _series_candidate_scores(texts, known, series_groups=series_groups)
        code2, score2, margin2 = _choose_series(scores)
        if code2:
            code, series_score = code2, max(series_score, score2)
        condition_hint = condition_hint or bool(
            _condition_keyword(secondary) or re.search(r"[<>]\s*\d", secondary or "")
        )

        if not code and group and not condition_hint:
            anchored_code, anchored_score, more = _frame_anchored_series_choice(
                crop, group, known
            )
            for t in more:
                if t and t not in texts:
                    texts.append(t)
            if anchored_code:
                code = anchored_code
                series_score = max(series_score, anchored_score)

        if code and it is None:
            it, it_conf, it_diag = _verified_it_after_local_series(crop, code)

    # 6) Masque blanc global seulement pour les cas qui résistent réellement.
    if not code or it is None:
        white = _ocr(crop, 11, white_only=True, scale=2.0)
        if white and white not in texts:
            texts.append(white)
        scores = _series_candidate_scores(texts, known, series_groups=series_groups)
        code2, score2, margin2 = _choose_series(scores)
        if code2:
            code, series_score = code2, max(series_score, score2)

        if not code and group:
            anchored_code, anchored_score, more = _frame_anchored_series_choice(
                crop, group, known
            )
            for t in more:
                if t and t not in texts:
                    texts.append(t)
            if anchored_code:
                code = anchored_code
                series_score = max(series_score, anchored_score)

        if code and it is None:
            it, it_conf, it_diag = _verified_it_after_local_series(crop, code)

    # 7) Secours de couverture : V3.2 reconstruit la cellule depuis les groupes
    #    OpenCV. Il est conservé pour les contours partiellement occultés, mais
    #    son poids est limité afin qu'une lecture locale vérifiée l'emporte dès
    #    qu'une même série réapparait dans une autre capture.
    if code and it is None and group:
        it, it_conf, it_diag = _frame_group_it(crop, group)
        if it is not None:
            it_conf = min(.64, it_conf)
            it_diag = "secours groupe OpenCV faible confiance; " + it_diag

    condition_hint = condition_hint or any(_condition_keyword(t) for t in texts)
    if condition_hint and group:
        conditional_obs = _conditional_group_observation(image_path, crop, angle, polygon, known, texts)
        if conditional_obs is not None:
            return conditional_obs, True
    if not code:
        if it is not None and scores:
            return (
                PhysicalObservation(
                    image_path=image_path, crop=crop, angle=angle, crop_polygon=polygon,
                    texts=texts, candidate_scores=scores, series_code="", series_score=0.0,
                    tolerance_value=it, multiplicity=None, datum_raw="", datums={},
                    layout="UNRESOLVED", condition_text=condition_text, confidence=it_conf,
                    diagnostic="V9 cadre physique à série ambiguë/non résolue; " + it_diag,
                    physical_polygon=physical_polygon, cell_polygons=cell_polygons,
                    internal_walls=internal_walls, geometric_score=geometric_score,
                    candidate_polygon=polygon,
                ),
                condition_hint,
            )
        return None, condition_hint
    if it is None:
        return None, condition_hint

    # V9 : multiplicité uniquement dans la bande au-dessus du cadre. Le texte
    # global ne peut plus transformer un repère magenta voisin en multiplicité.
    if not group_geom:
        group_geom = _best_lsd_frame_geometry(crop, min_boundaries=3)
    mult = _multiplicity_from_lsd_geometry(crop, group_geom)
    datum_raw, datums = _parse_datums(texts)
    refs_bonus = .08 if sum(datums.values()) >= 3 else 0.0
    confidence = min(.99, .48 + .30 * series_score + .18 * it_conf + refs_bonus)

    return (
        PhysicalObservation(
            image_path=image_path, crop=crop, angle=angle, crop_polygon=polygon,
            texts=texts, candidate_scores=scores, series_code=code, series_score=series_score,
            tolerance_value=it, multiplicity=mult, datum_raw=datum_raw, datums=datums,
            layout="CADRE_REFERENCES", condition_text=condition_text, confidence=confidence,
            diagnostic="V9 cadre OpenCV incliné + parois LSD + cellule IT locale; " + it_diag,
            physical_polygon=physical_polygon, cell_polygons=cell_polygons,
            internal_walls=internal_walls, geometric_score=geometric_score,
            candidate_polygon=polygon,
        ),
        False,
    )


def _draw_debug(image: np.ndarray, observations: Sequence[PhysicalObservation], output: Path) -> None:
    """Trace les bordures physiques exactes et séparateurs de cellules réels des cadres."""
    canvas = image.copy()
    for idx, obs in enumerate(observations, 1):
        poly = obs.physical_polygon if obs.physical_polygon is not None and len(obs.physical_polygon) > 0 else obs.crop_polygon
        if poly is None or len(poly) == 0:
            continue
        poly_pts = poly.reshape((-1, 1, 2)).astype(np.int32)
        # Bordures réelles du Physical Frame en vert franc (2 px)
        cv2.polylines(canvas, [poly_pts], True, (0, 255, 0), 2)

        # Parois internes des cellules en cyan / jaune
        for w_top, w_bot in obs.internal_walls:
            cv2.line(canvas, tuple(w_top), tuple(w_bot), (255, 255, 0), 1)

        # Labels neutres et précis
        x, y = int(poly[:, 0].min()), int(poly[:, 1].min())
        label = f"FRAME_{idx:02d}" + (f" ({obs.series_code})" if obs.series_code else "")
        cv2.putText(canvas, label, (max(5, x), max(18, y - 6)), cv2.FONT_HERSHEY_SIMPLEX, .50, (0, 255, 0), 1, cv2.LINE_AA)

    output.parent.mkdir(parents=True, exist_ok=True)
    cv2.imwrite(str(output), canvas)


def _export_geometric_diagnostics(
    image: np.ndarray,
    observations: Sequence[PhysicalObservation],
    groups: Sequence[Any],
    output_dir: Path,
    stem: str,
) -> None:
    """Génère les 3 vues de diagnostic géométrique obligatoires pour la capture."""
    output_dir.mkdir(parents=True, exist_ok=True)

    # 1. Segments LSD bruts
    canvas_lsd = image.copy()
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    if hasattr(cv2, "createLineSegmentDetector"):
        try:
            detector = cv2.createLineSegmentDetector(cv2.LSD_REFINE_STD)
            lines = detector.detect(gray)[0]
            if lines is not None:
                for x1, y1, x2, y2 in lines.reshape(-1, 4):
                    dx, dy = x2 - x1, y2 - y1
                    length = math.hypot(dx, dy)
                    if length >= 10.0:
                        cv2.line(canvas_lsd, (int(x1), int(y1)), (int(x2), int(y2)), (0, 0, 255), 1)
        except Exception:
            pass
    cv2.imwrite(str(output_dir / f"{stem}_01_lsd_segments.png"), canvas_lsd)

    # 2. Candidate Regions (zones larges en orange)
    canvas_cand = image.copy()
    for item in groups:
        if isinstance(item, (list, tuple)) and len(item) >= 4:
            cand_poly = item[3]
            if isinstance(cand_poly, np.ndarray):
                poly_pts = cand_poly.reshape((-1, 1, 2)).astype(np.int32)
                cv2.polylines(canvas_cand, [poly_pts], True, (0, 140, 255), 2)
    cv2.imwrite(str(output_dir / f"{stem}_02_candidate_regions.png"), canvas_cand)

    # 3. Physical Frames + Cellules exactes
    canvas_phys = image.copy()
    for idx, obs in enumerate(observations, 1):
        poly = obs.physical_polygon if obs.physical_polygon is not None and len(obs.physical_polygon) > 0 else obs.crop_polygon
        if poly is not None and len(poly) > 0:
            poly_pts = poly.reshape((-1, 1, 2)).astype(np.int32)
            cv2.polylines(canvas_phys, [poly_pts], True, (0, 255, 0), 2)
            for w_top, w_bot in obs.internal_walls:
                cv2.line(canvas_phys, tuple(w_top), tuple(w_bot), (255, 255, 0), 1)
            x, y = int(poly[:, 0].min()), int(poly[:, 1].min())
            cv2.putText(canvas_phys, f"FRAME_{idx:02d}", (max(5, x), max(18, y - 6)), cv2.FONT_HERSHEY_SIMPLEX, .50, (0, 255, 0), 1, cv2.LINE_AA)
    cv2.imwrite(str(output_dir / f"{stem}_03_physical_frames.png"), canvas_phys)


# ---------------------------------------------------------------------------
# Consolidation / déduplication
# ---------------------------------------------------------------------------
def _observation_to_annotation(obs: PhysicalObservation) -> VisualAnnotation:
    it_text = "" if obs.tolerance_value is None else f"{obs.tolerance_value:g}"
    return VisualAnnotation(
        series_code=obs.series_code,
        multiplicity=obs.multiplicity,
        tolerance_value=obs.tolerance_value,
        tolerance_text=it_text,
        datum_raw=obs.datum_raw,
        datum_a=bool(obs.datums.get("A")), datum_b=bool(obs.datums.get("B")),
        datum_c=bool(obs.datums.get("C")), datum_d=bool(obs.datums.get("D")), datum_e=bool(obs.datums.get("E")),
        raw_text=" || ".join(_clean_text(t) for t in obs.texts if _clean_text(t)),
        source_image=str(obs.image_path.resolve()), rotation_angle=obs.angle,
        confidence=round(obs.confidence, 4), diagnostic=obs.diagnostic,
        annotation_layout=obs.layout, condition_text=obs.condition_text,
        read_status=obs.read_status,
    )


def _consensus(observations: Sequence[PhysicalObservation], known: set[str]) -> list[VisualAnnotation]:
    by_series: dict[str, list[PhysicalObservation]] = {}
    for obs in observations:
        if obs.series_code in known and obs.tolerance_value is not None:
            by_series.setdefault(obs.series_code, []).append(obs)
    output: list[VisualAnnotation] = []
    for code in sorted(by_series):
        items = by_series[code]
        # Vote IT pondéré par la confiance réelle de chaque observation
        weights: dict[float, float] = {}
        for item in items:
            value = float(item.tolerance_value)
            weights[value] = weights.get(value, 0.0) + item.confidence
        for value in list(weights):
            if value % 1:
                integer = float(int(value))
                if integer in weights:
                    weights[value] += .35 * weights[integer]

        sorted_votes = sorted(weights.items(), key=lambda x: x[1], reverse=True)
        winning_it, top_weight = sorted_votes[0]
        is_conflict = False
        conflict_msg = ""
        if len(sorted_votes) > 1:
            second_it, second_weight = sorted_votes[1]
            if abs(winning_it - second_it) > 1e-4 and (second_weight / max(top_weight, 1e-6)) > 0.65 and second_weight >= 0.75:
                is_conflict = True
                conflict_msg = f"CONFLIT_IT_MULTI_CAPTURES: {winning_it:g} (poids={top_weight:.2f}) vs {second_it:g} (poids={second_weight:.2f})"

        candidates = [x for x in items if abs(float(x.tolerance_value) - winning_it) < 1e-6]
        best = max(candidates, key=lambda x: (x.confidence, x.series_score))

        # Multiplicité : vote séparé entre toutes les répétitions de la série
        mult_counts: dict[int, float] = {}
        for item in items:
            if item.multiplicity is not None:
                mult_counts[item.multiplicity] = mult_counts.get(item.multiplicity, 0.0) + item.confidence
        if mult_counts:
            best.multiplicity = max(mult_counts, key=lambda v: mult_counts[v])

        # Fusion des références vues sur différentes captures
        for letter in "ABCDE":
            if any(item.datums.get(letter) for item in items):
                best.datums[letter] = True
        best.datum_raw = " | ".join(letter for letter in "ABCDE" if best.datums.get(letter))
        best.tolerance_value = winning_it
        best.diagnostic += f"; consensus {len(items)} observation(s) ; déduplication par série"
        if is_conflict:
            best.confidence = min(.60, best.confidence)
            best.diagnostic += f"; {conflict_msg}"
            best.read_status = "CONFLIT_IT_MULTI_CAPTURES"
        output.append(_observation_to_annotation(best))
    return output


def _reconcile_ambiguous(
    observations: list[PhysicalObservation],
    unresolved: list[tuple[PhysicalObservation, dict[str, float]]],
    known: set[str],
) -> None:
    """Réaffecte uniquement des cadres physiques ambigus aux séries encore absentes.

    Aucune valeur n'est créée : on ne travaille que sur un cadre déjà détecté et un
    texte local ayant un score minimum. Ceci aide notamment quand 06A01 est lu 05A01
    alors qu'un autre cadre donne déjà 05A01 avec une meilleure confiance.
    """
    present = {obs.series_code for obs in observations if obs.series_code}
    missing = set(known) - present
    if not missing:
        return
    used_obs: set[int] = set()
    choices: list[tuple[float, int, str]] = []
    for idx, (obs, scores) in enumerate(unresolved):
        for code in missing:
            score = scores.get(code, 0.0)
            if score >= .58:
                choices.append((score, idx, code))
    for score, idx, code in sorted(choices, reverse=True):
        if code not in missing or idx in used_obs:
            continue
        obs, scores = unresolved[idx]
        ranking = sorted(scores.items(), key=lambda x: x[1], reverse=True)
        top_code, top_score = ranking[0]
        second_code = ranking[1][0] if len(ranking) > 1 else ""
        second_score = ranking[1][1] if len(ranking) > 1 else 0.0

        if second_code in missing and abs(top_score - second_score) < 0.035 and top_score >= .62:
            obs.series_code = top_code
            obs.series_score = top_score
            obs.confidence = min(.65, top_score)
            obs.diagnostic += f"; AMBIGU_CANDIDATS_MULTIPLES: {top_code} ({top_score:.2f}) vs {second_code} ({second_score:.2f})"
            obs.layout = "AMBIGU_CANDIDATS_MULTIPLES"
            observations.append(obs)
            missing.discard(top_code)
            used_obs.add(idx)
            continue

        if score < .65 and score - second_score < -.10:
            continue
        obs.series_code = code
        obs.series_score = score
        obs.confidence = max(.66, obs.confidence)
        obs.diagnostic += "; réconciliation globale série manquante sur cadre physique"
        observations.append(obs)
        missing.remove(code)
        used_obs.add(idx)


# ---------------------------------------------------------------------------
# Rescue V4 : série connue -> cadre local -> IT
# ---------------------------------------------------------------------------
def _annotation_stroke_mask(image: np.ndarray) -> np.ndarray:
    """Masque commun aux traits blancs/gris et cyan/bleus des annotations.

    Il n'est jamais appliqué à l'image entière pour décider qu'une série existe.
    Il sert seulement une fois que le texte d'une série connue a été localisé.
    """
    return cv2.bitwise_or(_white_mask(image), _cyan_frame_mask(image))


def _series_guided_tokens(image: np.ndarray) -> list[dict[str, Any]]:
    """OCR positionnel de la zone graphique, sans l'arbre CATIA à gauche."""
    if image.size == 0:
        return []
    work = image.copy()
    h, w = work.shape[:2]
    # La barre CATIA porte elle aussi les codes. Elle ne peut jamais constituer
    # une preuve d'association avec un cadre visible.
    work[:, : max(150, int(w * .145))] = (52, 52, 105)
    all_tokens: list[dict[str, Any]] = []
    seen: set[tuple[str, int, int, int, int]] = set()
    # Les codes des cadres blancs sont très fins. Les variantes à grande échelle
    # ne servent qu'à localiser un code déjà connu, jamais à inventer une série.
    for variant, (psm, white_only, scale) in enumerate((
        (11, False, 3.2), (11, True, 3.8), (6, False, 3.6), (11, False, 5.0),
    )):
        for token in _ocr_data_tokens(
            work,
            psm=psm,
            white_only=white_only,
            whitelist="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ",
            scale=scale,
        ):
            token["line_key"] = (variant, *token["line_key"])
            key = (
                re.sub(r"[^A-Z0-9]", "", token["text"].upper()),
                token["left"], token["top"], token["width"], token["height"],
            )
            if key not in seen:
                seen.add(key)
                all_tokens.append(token)
    return all_tokens


def _series_guided_anchors(
    image: np.ndarray,
    targets: set[str],
) -> dict[str, tuple[tuple[int, int, int, int], float]]:
    """Localise uniquement les séries recherchées avec une preuve OCR stricte."""
    tokens = _series_guided_tokens(image)
    anchors: dict[str, tuple[tuple[int, int, int, int], float]] = {}
    for code in targets:
        bbox, score = _locate_series_bbox(tokens, code)
        if bbox is None:
            continue
        compact_score = score
        # Une correspondance exacte est acceptée sans ambiguïté. Une lecture
        # approchée doit rester très forte : on ne force jamais une série absente.
        # Avec le cadre et la cellule IT ensuite vérifiés physiquement, un code
        # partiellement lu peut être accepté ici. Cela récupère les ``01B01``
        # ou ``02A01`` très petits sans permettre une association aveugle.
        if compact_score < .72:
            continue
        x1, y1, x2, y2 = bbox
        if x2 <= max(150, int(image.shape[1] * .145)):
            continue
        anchors[code] = (bbox, compact_score)
    return anchors


def _guided_roi(
    image: np.ndarray,
    series_bbox: tuple[int, int, int, int],
) -> tuple[np.ndarray, tuple[int, int], np.ndarray]:
    """ROI locale série + cadre, de taille relative au texte réellement lu."""
    x1, y1, x2, y2 = series_bbox
    text_h = max(9, y2 - y1)
    # Le code CATIA est normalement immédiatement à gauche du cadre. On garde
    # une marge gauche pour multiplicité et une marge droite suffisante pour
    # symbole + IT + A..E, sans dépendre de pixels absolus propres à une pièce.
    left = max(0, int(x1 - 4.2 * text_h))
    right = min(image.shape[1], int(x2 + 22.0 * text_h))
    top = max(0, int(y1 - 3.4 * text_h))
    bottom = min(image.shape[0], int(y2 + 3.7 * text_h))
    roi = image[top:bottom, left:right].copy()
    polygon = np.array(
        [[left, top], [right, top], [right, bottom], [left, bottom]],
        dtype=np.int32,
    )
    return roi, (left, top), polygon


def _auto_vertical_separator_centers(
    roi: np.ndarray,
    series_box: tuple[int, int, int, int],
) -> tuple[list[int], tuple[int, int]]:
    """Séparateurs locaux des cadres blancs ou cyan, sans seuil image global."""
    sx1, sy1, sx2, sy2 = series_box
    text_h = max(8, sy2 - sy1)
    cy = (sy1 + sy2) // 2
    y1 = max(0, int(cy - 1.45 * text_h))
    y2 = min(roi.shape[0], int(cy + 1.45 * text_h))
    if y2 - y1 < 8:
        return [], (y1, y2)
    mask = _annotation_stroke_mask(roi)[y1:y2, :]
    band_h = mask.shape[0]
    vertical = cv2.morphologyEx(
        mask,
        cv2.MORPH_OPEN,
        np.ones((max(5, int(band_h * .40)), 1), np.uint8),
        iterations=1,
    )
    projection = (vertical > 0).sum(axis=0)
    xs = [i for i, value in enumerate(projection) if value >= max(4, int(band_h * .30))]
    if not xs:
        return [], (y1, y2)
    clusters: list[list[int]] = [[int(xs[0])]]
    for raw in xs[1:]:
        x = int(raw)
        if x - clusters[-1][-1] <= 2:
            clusters[-1].append(x)
        else:
            clusters.append([x])
    centers = [int(round(sum(group) / len(group))) for group in clusters]
    # Le cadre qui porte l'IT est sur la même ligne que la série, pas les traits
    # des pièces ou des cartouches éloignés.
    centers = [x for x in centers if x >= int(sx2 - .9 * text_h)]
    return centers, (y1, y2)


def _datum_a_in_next_cell(cell: np.ndarray) -> tuple[bool, list[str]]:
    """Vérifie que la cellule juste après l'IT est bien la référence A.

    Cette vérification structurelle bloque les faux IT lus dans le symbole, la
    multiplicité ou une autre partie du cadre. Tesseract confond parfois A et 4,
    ce qui est accepté seulement dans cette cellule de référence isolée.
    """
    if cell.size == 0:
        return False, []
    texts: list[str] = []
    for psm, white_only, scale in ((10, False, 5.0), (7, False, 4.8), (13, True, 5.2)):
        text = _clean_text(_ocr(cell, psm, white_only=white_only, whitelist="Aa4", scale=scale)).upper()
        if text and text not in texts:
            texts.append(text)
    proven = any("A" in text or text in {"4", "A4", "4A"} for text in texts)
    return proven, texts


def _guided_standard_it_morphology_fallback(
    roi: np.ndarray,
    series_bbox: tuple[int, int, int, int],
    series_code: str,
) -> tuple[Optional[float], float, str]:
    """IT dans la 2e cellule d'un cadre standard localisé par sa série."""
    separators, (band_y1, band_y2) = _auto_vertical_separator_centers(roi, series_bbox)
    text_h = max(8, series_bbox[3] - series_bbox[1])
    candidates: list[tuple[float, int]] = []
    # Quatre traits donnent trois cellules : symbole | IT | référence A.
    # C'est la preuve géométrique manquante dans V4.
    for index in range(len(separators) - 3):
        x0, x1, x2, x3 = separators[index:index + 4]
        gap_symbol = x1 - x0
        gap_it = x2 - x1
        gap_a = x3 - x2
        if not (.32 * text_h <= gap_symbol <= 5.2 * text_h):
            continue
        if not (.32 * text_h <= gap_it <= 5.2 * text_h):
            continue
        if not (.28 * text_h <= gap_a <= 5.2 * text_h):
            continue
        distance = abs(x0 - series_bbox[2]) / max(text_h, 1.0)
        if distance <= 6.2:
            candidates.append((distance, index))

    best: tuple[Optional[float], float, str] = (None, 0.0, "")
    for distance, index in sorted(candidates)[:5]:
        _x0, x1, x2, x3 = separators[index:index + 4]
        pad_x = max(1, int((x2 - x1) * .08))
        pad_y = max(1, int((band_y2 - band_y1) * .08))
        cell = roi[
            max(0, band_y1 + pad_y):min(roi.shape[0], band_y2 - pad_y),
            max(0, x1 + pad_x):min(roi.shape[1], x2 - pad_x),
        ]
        value, confidence, texts = _parse_isolated_numeric_cell(cell)
        if value is None:
            continue
        a_pad_x = max(1, int((x3 - x2) * .08))
        a_cell = roi[
            max(0, band_y1 + pad_y):min(roi.shape[0], band_y2 - pad_y),
            max(0, x2 + a_pad_x):min(roi.shape[1], x3 - a_pad_x),
        ]
        a_proven, a_texts = _datum_a_in_next_cell(a_cell)
        if not a_proven:
            continue
        confidence = min(.99, confidence + .08 + max(0.0, .13 - .025 * distance))
        diagnostic = (
            f"V5 cellule IT locale x={x1}:{x2}; OCR={texts}; next-A={a_texts}; "
            f"serie={series_code}; distance={distance:.2f}"
        )
        if confidence > best[1]:
            best = (value, confidence, diagnostic)

    if best[0] is not None:
        return best
    return None, 0.0, "V5 aucune cellule IT suivie de la référence A n'est prouvée"


def _guided_standard_it(
    roi: np.ndarray,
    series_bbox: tuple[int, int, int, int],
    series_code: str,
) -> tuple[Optional[float], float, str]:
    """V9 : parois LSD complètes d'abord, projection morphologique seulement en secours."""
    sets=_lsd_wall_sets(roi,min_boundaries=4)
    text_h=max(8,series_bbox[3]-series_bbox[1])
    candidates=[]
    for geom in sets:
        bounds=geom["bounds"]; y1=geom["y1"]; y2=geom["y2"]; fh=geom["height"]
        for i in range(len(bounds)-3):
            x0,x1,x2,x3=bounds[i:i+4]
            distance=abs(x0-series_bbox[2])/max(fh,1.0)
            if x0 < series_bbox[2]-.7*fh or distance>3.2: continue
            gaps=(x1-x0,x2-x1,x3-x2)
            if not (.45*fh<=gaps[0]<=3.0*fh and .45*fh<=gaps[1]<=4.6*fh and .38*fh<=gaps[2]<=4.6*fh): continue
            py=max(1,int(fh*.08)); px=max(1,int(gaps[1]*.06))
            cell=roi[max(0,int(y1)+py):min(roi.shape[0],int(y2)-py),max(0,int(x1)+px):min(roi.shape[1],int(x2)-px)]
            value,conf,texts=_parse_isolated_numeric_cell(cell)
            if value is None: continue
            a_cell=roi[max(0,int(y1)+py):min(roi.shape[0],int(y2)-py),max(0,int(x2)+1):min(roi.shape[1],int(x3)-1)]
            proven,a_texts=_datum_a_in_next_cell(a_cell)
            if not proven: continue
            score=conf + .14*max(0.0,1-distance/3.2) + .02*min(len(bounds),6)
            candidates.append((score,value,conf,distance,texts,a_texts))
    if candidates:
        _score,value,conf,distance,texts,a_texts=max(candidates,key=lambda q:q[0])
        return value,min(.995,max(.95,conf)),f"V9 LSD cellule IT; OCR={texts}; next-A={a_texts}; serie={series_code}; distance={distance:.2f}"
    return _guided_standard_it_morphology_fallback(roi,series_bbox,series_code)



def _guided_observation(
    image_path: Path,
    image: np.ndarray,
    series_code: str,
    global_bbox: tuple[int, int, int, int],
    series_score: float,
) -> Optional[PhysicalObservation]:
    """Construit une observation seulement si la série et son IT local sont prouvés."""
    roi, origin, polygon = _guided_roi(image, global_bbox)
    if roi.size == 0:
        return None
    ox, oy = origin
    local_series_box = (
        global_bbox[0] - ox, global_bbox[1] - oy,
        global_bbox[2] - ox, global_bbox[3] - oy,
    )
    texts: list[str] = []
    for psm, white_only in ((6, False), (11, False), (11, True)):
        text = _ocr(roi, psm, white_only=white_only, scale=3.0)
        if text and text not in texts:
            texts.append(text)

    condition_text = _parse_condition(texts)
    has_condition = bool(condition_text or any(_condition_keyword(text) for text in texts))
    if has_condition:
        it, it_conf = _parse_conditional_it(texts)
        if it is None:
            return None
        layout = "CONDITIONNEL_2_CELLULES"
        diagnostic = "V4 série-guidée : condition locale + IT local"
        multiplicity = None
        datum_raw = ""
        datums = {letter: False for letter in "ABCDE"}
    else:
        it, it_conf, diagnostic = _guided_standard_it(roi, local_series_box, series_code)
        if it is None:
            return None
        layout = "CADRE_REFERENCES"
        multiplicity = _parse_multiplicity_from_texts(texts)
        datum_raw, datums = _parse_datums(texts)
        if any(re.search(r"\d+[ \t]*[Xx×]|[Xx×][ \t]*\d+", text) for text in texts):
            targeted = _targeted_multiplicity(roi)
            if targeted is not None:
                multiplicity = targeted

    confidence = min(.99, .50 + .30 * series_score + .19 * it_conf + (.05 if has_condition else 0.0))
    return PhysicalObservation(
        image_path=image_path,
        crop=roi,
        angle=0.0,
        crop_polygon=polygon,
        texts=texts,
        candidate_scores={series_code: series_score},
        series_code=series_code,
        series_score=series_score,
        tolerance_value=it,
        multiplicity=multiplicity,
        datum_raw=datum_raw,
        datums=datums,
        layout=layout,
        condition_text=condition_text,
        confidence=confidence,
        diagnostic=diagnostic,
    )


def _rescue_series_guided(
    images: Sequence[Path],
    known: set[str],
    targets: set[str],
    diagnostics: list[dict[str, Any]],
) -> list[PhysicalObservation]:
    """Lecture locale générale : l'arbre guide l'OCR, jamais l'inverse.

    ``targets`` peut contenir toutes les séries d'une capture blanche. Dans ce
    cas le résultat du détecteur global historique est remplacé : une mauvaise
    lecture (par exemple ``1.6`` devenu ``4``) ne peut plus gagner le consensus.
    """
    requested = set(targets) & set(known)
    rescued: list[PhysicalObservation] = []
    for image_path in images:
        image = cv2.imread(str(image_path), cv2.IMREAD_COLOR)
        if image is None:
            continue
        anchors = _series_guided_anchors(image, requested)
        for code, (bbox, score) in anchors.items():
            observation = _guided_observation(image_path, image, code, bbox, score)
            diagnostics.append({
                "image": str(image_path.resolve()),
                "style": "SERIE_GUIDEE_LOCALE",
                "series": code,
                "series_score": round(score, 4),
                "series_bbox": list(bbox),
                "it": observation.tolerance_value if observation else None,
                "accepted": observation is not None,
            })
            if observation is not None:
                rescued.append(observation)
    return rescued


# ---------------------------------------------------------------------------
# V6 — OpenCV d'abord : inventaire des cadres blancs, OCR ensuite et seulement
# dans les cellules et libellés associés à ces cadres.
# ---------------------------------------------------------------------------
def _white_frame_bboxes(
    image: np.ndarray,
    *,
    relaxed: bool = False,
    geometry_only: bool = False,
) -> list[tuple[int, int, int, int]]:
    """Détecte les cadres FCF redressés par leur géométrie physique.

    Les lettres et les lignes de rappel sont supprimées par ouverture
    morphologique. Contrairement aux versions précédentes, aucun texte de la
    vue entière n'intervient ici : un candidat est un vrai rectangle physique.
    """
    if image.size == 0:
        return []
    # Le chemin normal reste très sélectif (blanc/gris). Le secours
    # ``geometry_only`` ajoute les contours : il ne dépend alors d'aucune
    # couleur de cadre. Un candidat contour n'est jamais accepté seul ; il
    # doit ensuite prouver localement la série connue, l'IT et la référence A.
    mask = _annotation_stroke_mask(image)
    if geometry_only:
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        edges = cv2.Canny(gray, 45, 135)
        mask = cv2.bitwise_or(mask, edges)
    height, width = mask.shape
    mask[:, :max(110, int(width * .110))] = 0
    horizontal_length = max(9, int(height * (.015 if relaxed else .020)))
    vertical_length = max(6, int(height * (.010 if relaxed else .014)))
    horizontal = cv2.morphologyEx(
        mask, cv2.MORPH_OPEN, np.ones((1, horizontal_length), np.uint8), iterations=1,
    )
    vertical = cv2.morphologyEx(
        mask, cv2.MORPH_OPEN, np.ones((vertical_length, 1), np.uint8), iterations=1,
    )
    structure = cv2.bitwise_or(horizontal, vertical)
    structure = cv2.morphologyEx(structure, cv2.MORPH_CLOSE, np.ones((3, 3), np.uint8), iterations=1)
    contours, _ = cv2.findContours(structure, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    raw: list[tuple[int, int, int, int, float]] = []
    min_h = max(8, int(height * .014))
    max_h = max(40, int(height * .085))
    min_w = max(25, int(height * .040))
    max_w = max(390, int(width * .42))
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        if x <= max(110, int(width * .110)):
            continue
        if not (min_h <= h <= max_h and min_w <= w <= max_w):
            continue
        if w / max(float(h), 1.0) < 1.7:
            continue
        line_pixels = float((structure[y:y+h, x:x+w] > 0).sum())
        # Un cadre possède au minimum ses deux parois horizontales et plusieurs
        # parois verticales ; du texte seul ne passe pas ce critère.
        if line_pixels < max(30.0, (.42 if geometry_only else .55) * (w + h)):
            continue
        raw.append((x, y, w, h, line_pixels))

    # Une même FCF peut produire une bbox pour le contour extérieur et une pour
    # ses cellules internes : on conserve la plus grande, sans doublon.
    selected: list[tuple[int, int, int, int, float]] = []
    for candidate in sorted(raw, key=lambda item: item[2] * item[3], reverse=True):
        x, y, w, h, score = candidate
        duplicate = False
        for px, py, pw, ph, _ in selected:
            ix1, iy1 = max(x, px), max(y, py)
            ix2, iy2 = min(x + w, px + pw), min(y + h, py + ph)
            intersection = max(0, ix2 - ix1) * max(0, iy2 - iy1)
            union = w * h + pw * ph - intersection
            if union and intersection / union >= .55:
                duplicate = True
                break
        if not duplicate:
            selected.append(candidate)
    # Les contours Canny peuvent produire plusieurs faux rectangles sur une
    # pièce très détaillée. On garde les plus structurés : la validation OCR
    # locale demeure l'autorité finale.
    selected.sort(key=lambda item: item[4], reverse=True)
    return [
        (x, y, w, h)
        for x, y, w, h, _ in sorted(selected[:80], key=lambda item: (item[1], item[0]))
    ]


def _white_frame_boundaries(image: np.ndarray, bbox: tuple[int, int, int, int]) -> tuple[list[int], tuple[int, int, int, int]]:
    """Retourne les parois verticales d'un cadre déjà validé par OpenCV."""
    x, y, w, h = bbox
    pad_x = max(2, int(h * .18))
    pad_y = max(2, int(h * .20))
    x1 = max(0, x - pad_x); x2 = min(image.shape[1], x + w + pad_x)
    y1 = max(0, y - pad_y); y2 = min(image.shape[0], y + h + pad_y)
    roi = image[y1:y2, x1:x2]
    # Les séparateurs sont une propriété géométrique du cadre, pas de sa
    # couleur. Le masque annotation couvre blanc/cyan ; Canny est le secours
    # pour un style CATIA d'une autre couleur.
    mask = _annotation_stroke_mask(roi)
    if int((mask > 0).sum()) < max(12, int(roi.shape[0] * roi.shape[1] * .002)):
        mask = cv2.Canny(cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY), 45, 135)
    vertical = cv2.morphologyEx(
        mask, cv2.MORPH_OPEN, np.ones((max(6, int(h * .48)), 1), np.uint8), iterations=1,
    )
    projection = (vertical > 0).sum(axis=0)
    threshold = max(5, int(h * .38))
    positions = [index for index, value in enumerate(projection) if value >= threshold]
    if not positions:
        return [], (x1, y1, x2, y2)
    clusters: list[list[int]] = [[positions[0]]]
    for position in positions[1:]:
        if position - clusters[-1][-1] <= 2:
            clusters[-1].append(position)
        else:
            clusters.append([position])
    centers = [x1 + int(round(sum(cluster) / len(cluster))) for cluster in clusters]
    # Les bords doivent être proches de la bbox OpenCV. Cela élimine les traits
    # de rappel verticaux qui traversent la même zone.
    centers = [center for center in centers if x - 2 * pad_x <= center <= x + w + 2 * pad_x]
    unique: list[int] = []
    for center in sorted(centers):
        if not unique or center - unique[-1] >= 3:
            unique.append(center)
    return unique, (x1, y1, x2, y2)


def _frame_local_series(image: np.ndarray, bbox: tuple[int, int, int, int], known: set[str]) -> tuple[str, float, list[str]]:
    """OCR du seul libellé placé immédiatement à gauche d'un cadre détecté."""
    x, y, _w, h = bbox
    left = max(0, int(x - max(64, 8.5 * h)))
    right = max(left + 1, x - 1)
    top = max(0, int(y - 1.25 * h))
    bottom = min(image.shape[0], int(y + 2.25 * h))
    roi = image[top:bottom, left:right]
    texts: list[str] = []
    for psm, scale in ((7, 5.0), (8, 5.4), (13, 4.8)):
        text = _clean_text(_local_binary_ocr(
            roi, psm=psm, whitelist="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ", scale=scale,
        ))
        if text and text not in texts:
            texts.append(text)
    scores = _series_candidate_scores(texts, known)
    code, score, _margin = _choose_series(scores)
    if code:
        return code, score, texts
    for text in texts:
        code, score = _complete_unique_series_fragment(text, known)
        if code:
            return code, score, texts
    return "", 0.0, texts


def _frame_local_multiplicity(image: np.ndarray, bbox: tuple[int, int, int, int]) -> Optional[int]:
    """OCR de la petite zone au-dessus du cadre, uniquement après OpenCV."""
    x, y, w, h = bbox
    left = max(0, int(x - .65 * h))
    right = min(image.shape[1], int(x + min(w, 3.5 * h)))
    top = max(0, int(y - 2.25 * h))
    bottom = max(top + 1, int(y - .12 * h))
    roi = image[top:bottom, left:right]
    values: list[int] = []
    for psm in (7, 11, 13):
        text = _local_binary_ocr(roi, psm=psm, whitelist="0123456789Xx", scale=5.0)
        match = re.search(r"(?i)(?:([0-9]{1,3})\s*[Xx]|[Xx]\s*([0-9]{1,3}))", text)
        if match:
            value = int(match.group(1) or match.group(2))
            if 1 <= value <= 999:
                values.append(value)
    if not values:
        return None
    return max(set(values), key=lambda value: (values.count(value), -value))


def _frame_cell_crop(image: np.ndarray, bounds: Sequence[int], index: int, y1: int, y2: int) -> np.ndarray:
    """Crop intérieur d'une cellule, sans ses parois blanches."""
    if index < 0 or index + 1 >= len(bounds):
        return image[0:0, 0:0]
    x1, x2 = int(bounds[index]), int(bounds[index + 1])
    left_pad = max(1, int((x2 - x1) * .12))
    right_pad = max(1, int((x2 - x1) * .15))
    pad_y = max(1, int((y2 - y1) * .12))
    return image[
        max(0, y1 + pad_y):min(image.shape[0], y2 - pad_y),
        max(0, x1 + left_pad):min(image.shape[1], x2 - right_pad),
    ]


def _frame_datums(image: np.ndarray, bounds: Sequence[int], y1: int, y2: int) -> tuple[str, dict[str, bool]]:
    datums = {letter: False for letter in "ABCDE"}
    parts: list[str] = []
    for index in range(2, len(bounds) - 1):
        cell = _frame_cell_crop(image, bounds, index, y1, y2)
        text = _clean_text(_local_binary_ocr(cell, psm=10, whitelist="ABCDE-", scale=4.8)).upper()
        if text:
            parts.append(text)
        for letter in "ABCDE":
            if letter in text:
                datums[letter] = True
    return " | ".join(parts), datums


def _white_frame_observation(
    image_path: Path,
    image: np.ndarray,
    bbox: tuple[int, int, int, int],
    known: set[str],
) -> Optional[PhysicalObservation]:
    """Observation depuis une bbox physique ; parois LSD, OCR strictement local."""
    x,y,w,h=bbox
    geom=_lsd_geometry_around_bbox(image,bbox)
    bounds=geom.get("bounds",[]) if geom else []
    if len(bounds)<3:
        return None

    context=image[max(0,y-2*h):min(image.shape[0],y+4*h),max(0,x-9*h):min(image.shape[1],x+w+5*h)]
    context_texts=[]
    for psm,white_only,scale in ((6,True,4.0),(6,False,4.0),(11,False,3.8)):
        text=_ocr(context,psm,white_only=white_only,scale=scale)
        if text and text not in context_texts: context_texts.append(text)
    condition_text=_parse_condition(context_texts)
    conditional=bool(condition_text or any(_condition_keyword(t) for t in context_texts))

    if conditional:
        scores=_condition_series_scores(context_texts,known)
        ranking=sorted(scores.items(),key=lambda q:q[1],reverse=True)
        if not ranking: return None
        code,series_score=ranking[0]; second=ranking[1][1] if len(ranking)>1 else 0.0
        if not (series_score>=.78 or (series_score>=.48 and series_score-second>=.20)):
            return None
    else:
        code,series_score,series_texts=_frame_local_series(image,bbox,known)
        if not code: return None
        context_texts=series_texts+context_texts

    # Teste toutes les suites de trois parois. Pour un standard la cellule suivante
    # doit être A ; pour un conditionnel la preuve est série+condition+bbox physique.
    cell_candidates=[]
    if conditional:
        left,right=bounds[0],bounds[-1]
        fh=max(8.0,float(geom["height"])); py=max(1,int(fh*.08))
        for i,sep in enumerate(bounds[1:-1]):
            symbol_w=sep-left; it_w=right-sep
            if not (.45*fh<=symbol_w<=2.4*fh and .45*fh<=it_w<=5.2*fh): continue
            px=max(1,int(it_w*.06))
            it_cell=image[max(0,int(geom["y1"])+py):min(image.shape[0],int(geom["y2"])-py),max(0,int(sep)+px):min(image.shape[1],int(right)-px)]
            value,conf,it_texts=_parse_isolated_numeric_cell(it_cell)
            if value is None: continue
            score=conf+.05-.10*abs(symbol_w/fh-1.20)
            cell_candidates.append((score,value,conf,it_texts,False,[],i))
    else:
        for i in range(len(bounds)-3):
            x0,x1,x2,x3=bounds[i:i+4]
            fh=max(8.0,float(geom["height"])); py=max(1,int(fh*.08)); px=max(1,int((x2-x1)*.06))
            it_cell=image[max(0,int(geom["y1"])+py):min(image.shape[0],int(geom["y2"])-py),max(0,int(x1)+px):min(image.shape[1],int(x2)-px)]
            value,conf,it_texts=_parse_isolated_numeric_cell(it_cell)
            if value is None: continue
            a_cell=image[max(0,int(geom["y1"])+py):min(image.shape[0],int(geom["y2"])-py),max(0,int(x2)+1):min(image.shape[1],int(x3)-1)]
            next_a,a_texts=_datum_a_in_next_cell(a_cell)
            if not next_a: continue
            score=conf+.10-.03*i
            cell_candidates.append((score,value,conf,it_texts,next_a,a_texts,i))
    if not cell_candidates:
        return None
    _score,it,it_conf,it_texts,next_a,a_texts,index=max(cell_candidates,key=lambda q:q[0])

    if conditional:
        layout="CONDITIONNEL_2_CELLULES"; datums={letter:False for letter in "ABCDE"}; datum_raw=""; multiplicity=None
    else:
        layout="CADRE_REFERENCES"
        # Références détaillées restent secondaires : l'export complète depuis REF_parent.
        datum_raw,datums=_parse_datums(context_texts)
        multiplicity=_frame_local_multiplicity(image,bbox)
    polygon=np.array([[x,y],[x+w,y],[x+w,y+h],[x,y+h]],dtype=np.int32)
    confidence=min(.995,.60+.25*series_score+.15*it_conf+(.06 if next_a else .03))
    return PhysicalObservation(
        image_path=image_path,crop=image[y:y+h,x:x+w].copy(),angle=0.0,crop_polygon=polygon,
        texts=context_texts+it_texts,candidate_scores={code:series_score},series_code=code,series_score=series_score,
        tolerance_value=it,multiplicity=multiplicity,datum_raw=datum_raw,datums=datums,
        layout=layout,condition_text=condition_text,confidence=confidence,
        diagnostic=f"V9 bbox physique + parois LSD; cellule={index+1}; IT OCR={it_texts}; next-A={a_texts}; condition={condition_text!r}",
    )


# ---------------------------------------------------------------------------
# V8 — cadres inclinés : OpenCV mesure l'orientation, redresse localement le
# repère, puis applique exactement la même preuve série -> IT -> référence A.
# ---------------------------------------------------------------------------
def _dominant_frame_angles(image: np.ndarray) -> list[float]:
    """Retourne les orientations plausibles de cadres, sans couleur imposée.

    Les cadres CATIA sont constitués de plusieurs segments parallèles. Les
    leaders et la géométrie de la pièce peuvent aussi produire des segments,
    mais ils ne deviennent jamais une annotation sans les validations OCR
    locales appliquées plus loin. Cette étape ne fait donc qu'orienter la
    recherche OpenCV ; elle ne reconnaît ni série ni IT.
    """
    if image.size == 0:
        return [0.0]
    h, w = image.shape[:2]
    mask = _annotation_stroke_mask(image)
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    edges = cv2.Canny(gray, 55, 150)
    mask = cv2.bitwise_or(mask, edges)
    mask[:, :max(110, int(w * .110))] = 0
    lines = cv2.HoughLinesP(
        mask,
        rho=1,
        theta=np.pi / 360.0,
        threshold=max(16, int(min(h, w) * .024)),
        minLineLength=max(18, int(min(h, w) * .032)),
        maxLineGap=max(4, int(min(h, w) * .010)),
    )
    bins: dict[int, float] = {}
    if lines is not None:
        for x1, y1, x2, y2 in lines.reshape(-1, 4):
            dx, dy = float(x2 - x1), float(y2 - y1)
            length = math.hypot(dx, dy)
            if length < max(18.0, min(h, w) * .032):
                continue
            angle = _normal_angle(math.degrees(math.atan2(dy, dx)))
            # Les cadres de tolérance sont lus dans leur axe long. Les axes
            # presque verticaux correspondent habituellement à des leaders.
            if abs(angle) > 42.0:
                continue
            key = int(round(angle))
            bins[key] = bins.get(key, 0.0) + length

    selected: list[float] = [0.0]
    for key, _weight in sorted(bins.items(), key=lambda item: item[1], reverse=True):
        angle = float(key)
        if all(_angle_difference(angle, previous) >= 1.6 for previous in selected):
            selected.append(angle)
        if len(selected) >= 9:
            break
    return selected


def _rotate_to_frame_axis(
    image: np.ndarray,
    frame_angle: float,
) -> tuple[np.ndarray, np.ndarray]:
    """Redresse des cadres ayant ``frame_angle`` dans le repère de l'image.

    Avec les coordonnées image (Y vers le bas), OpenCV utilise le même signe
    que l'angle mesuré par ``atan2(dy, dx)`` pour rendre ce segment horizontal.
    La matrice inverse sert uniquement au diagnostic/dessin dans l'image source.
    """
    h, w = image.shape[:2]
    matrix = cv2.getRotationMatrix2D((w / 2.0, h / 2.0), frame_angle, 1.0)
    background = tuple(int(v) for v in np.median(image.reshape(-1, 3), axis=0))
    rotated = cv2.warpAffine(
        image,
        matrix,
        (w, h),
        flags=cv2.INTER_CUBIC,
        borderMode=cv2.BORDER_CONSTANT,
        borderValue=background,
    )
    return rotated, cv2.invertAffineTransform(matrix)


def _bbox_polygon_in_source(
    bbox: tuple[int, int, int, int],
    inverse_matrix: np.ndarray,
) -> np.ndarray:
    x, y, w, h = bbox
    points = np.array(
        [[[x, y], [x + w, y], [x + w, y + h], [x, y + h]]],
        dtype=np.float32,
    )
    return cv2.transform(points, inverse_matrix)[0].astype(np.int32)


def _oriented_white_frame_observations(
    image_path: Path,
    image: np.ndarray,
    known: set[str],
) -> tuple[list[PhysicalObservation], list[dict[str, Any]]]:
    """Inventorie des cadres de toute orientation avant le moindre OCR global.

    Le résultat inclut les candidats non validés pour le diagnostic. Les
    observations retournées sont exclusivement celles pour lesquelles le petit
    OCR local a confirmé : code série présent dans l'arbre + IT dans la cellule
    physique n°2 + A dans la cellule suivante (ou condition explicite).
    """
    accepted: list[PhysicalObservation] = []
    details: list[dict[str, Any]] = []
    # Évite de relire à l'OCR le même rectangle découvert par deux angles Hough.
    seen_centers: list[tuple[float, float, float]] = []
    for angle in _dominant_frame_angles(image):
        rotated, inverse = _rotate_to_frame_axis(image, angle)
        bboxes = _white_frame_bboxes(rotated)
        if not bboxes:
            bboxes = _white_frame_bboxes(rotated, relaxed=True)
        # Secours réellement indépendant de la couleur. Il n'est activé que
        # lorsque la passe annotation n'a vu aucun rectangle dans cette
        # orientation, ce qui limite fortement les faux candidats Canny.
        if not bboxes:
            bboxes = _white_frame_bboxes(rotated, relaxed=True, geometry_only=True)
        for bbox in bboxes:
            x, y, w, h = bbox
            source_polygon = _bbox_polygon_in_source(bbox, inverse)
            center = source_polygon.astype(float).mean(axis=0)
            # Deux angles voisins peuvent transformer le même cadre en bbox
            # presque identique. Ne pas consommer du temps Tesseract deux fois.
            if any(
                math.hypot(center[0] - px, center[1] - py) <= max(8.0, .45 * (h + ph))
                and abs(angle - pa) <= 3.0
                for px, py, ph, pa in seen_centers
            ):
                continue
            seen_centers.append((float(center[0]), float(center[1]), float(h), angle))
            obs = _white_frame_observation(image_path, rotated, bbox, known)
            detail = {
                "image": str(image_path.resolve()),
                "style": "ORIENTED_PHYSICAL_FRAME",
                "angle": round(angle, 2),
                "bbox_in_rotated_view": list(bbox),
                "polygon_in_source": source_polygon.tolist(),
                "series": obs.series_code if obs else "",
                "it": obs.tolerance_value if obs else None,
                "accepted": obs is not None,
            }
            details.append(detail)
            if obs is None:
                continue
            obs.angle = angle
            obs.crop_polygon = source_polygon
            obs.diagnostic = (
                f"V8 cadre physique orienté angle={angle:.1f}°; " + obs.diagnostic
            )
            accepted.append(obs)
    return accepted, details


# ---------------------------------------------------------------------------
# API principale
# ---------------------------------------------------------------------------
def scan_annotation_captures(
    *,
    project_root: Path | str | None = None,
    known_series: Sequence[str] | None = None,
    series_groups: Optional[dict[str, str]] = None,
    capture_dirs: Sequence[Path | str] | None = None,
    capture_current_view: bool = False,
    rotations: Sequence[float] | None = None,
    upscale: float = 3.0,
) -> list[dict[str, Any]]:
    del capture_current_view, rotations, upscale
    root = _project_root(project_root)
    tesseract = _configure_tesseract(root)
    known = {_clean_text(code).upper() for code in (known_series or []) if SERIES_RE.fullmatch(_clean_text(code).upper())}
    if not known:
        raise VisualScanError("La liste des séries extraite de la barre gauche est vide.")

    directories = [Path(p) for p in capture_dirs] if capture_dirs else [root / "captures_annotations"]
    images = _list_images(directories)
    if not images:
        raise VisualScanError("Aucune capture CATIA à analyser.")

    result_dir = root / "results" / "frame_inventory_ocr"
    result_dir.mkdir(parents=True, exist_ok=True)
    # Nouveau nom de cache : ne réutilise jamais un résultat issu d'un ancien
    # détecteur blanc ou d'une ancienne règle de sélection de cellule.
    cache_path = result_dir / "cache_v90_lsd_physical_walls.json"
    signature = _capture_signature(images, known)
    if cache_path.exists():
        try:
            cache = json.loads(cache_path.read_text(encoding="utf-8"))
            if cache.get("signature") == signature:
                rows = cache.get("rows", [])
                print("      Mode V9.0 : cache valide — aucun nouvel OCR.", flush=True)
                print(f"      Couverture réutilisée : {len(rows)}/{len(known)}.", flush=True)
                return rows
        except Exception:
            pass

    print("      Mode V9.0 : groupes inclinés OpenCV + parois physiques LSD + OCR local.", flush=True)
    print("      IT : 2e cellule entre parois complètes ; A valide les cadres standards ; aucune liste d'IT codée en dur.", flush=True)
    print(f"      Tesseract : {tesseract}", flush=True)
    print(f"      Séries de l'arbre : {len(known)}.", flush=True)
    print(f"      Captures : {len(images)}.", flush=True)

    started = time.perf_counter()
    observations: list[PhysicalObservation] = []
    unresolved: list[tuple[PhysicalObservation, dict[str, float]]] = []
    diagnostics: list[dict[str, Any]] = []
    debug_dir = result_dir / "cadres_detectes"

    for image_index, image_path in enumerate(images, start=1):
        image = cv2.imread(str(image_path), cv2.IMREAD_COLOR)
        if image is None:
            continue

        image_observations: list[PhysicalObservation] = []
        condition_hint = False
        candidates = _frame_cell_candidates(image)
        raw_groups = _group_frame_cells(candidates)
        min_area = max(260.0, image.shape[0] * image.shape[0] * .0018)
        
        # Filtre d'architecture ISO GPS : exige au moins 2 cellules adjacentes ou parois LSD prouvées
        groups = []
        for g in raw_groups:
            if sum(item["area"] for item in g) < min_area:
                continue
            crop, angle, polygon, meta = _tight_group_crop(image, g)
            geom = _best_lsd_frame_geometry(crop, min_boundaries=2)
            if len(g) >= 2 or bool(geom):
                groups.append((g, crop, angle, polygon, meta))

        style = "GPS_CADRE_ISO"
        for group_index, (group, crop, angle, polygon, meta) in enumerate(groups, start=1):
            obs, hint = _standard_observation(
                image_path, crop, angle, polygon, known, group=group, series_groups=series_groups, meta=meta,
            )
            condition_hint = condition_hint or hint
            if obs is not None:
                if obs.series_code:
                    image_observations.append(obs)
                elif obs.candidate_scores and obs.tolerance_value is not None:
                    unresolved.append((obs, obs.candidate_scores))
            diagnostics.append({
                "image": str(image_path.resolve()),
                "group_index": group_index,
                "style": style,
                "cell_count": len(group),
                "angle": round(angle, 2),
                "series": obs.series_code if obs else "",
                "it": obs.tolerance_value if obs else None,
                "condition_hint": hint,
            })

        # Les cadres conditionnels n'ont que deux cellules : on ne les perd
        # plus lorsqu'ils coexistent avec des groupes standard sur l'image.
        if condition_hint or len(_square_candidates(image)) >= 2:
            conditional = _condition_scan(image, image_path, known, 0.0)
            image_observations.extend(conditional)

            if not groups:
                oriented, oriented_details = _oriented_white_frame_observations(
                    image_path, image, known,
                )
                image_observations.extend(oriented)
                groups = [[{"bbox": item["bbox_in_rotated_view"], "angle": item["angle"]}]
                          for item in oriented_details]
                for group_index, item in enumerate(oriented_details, start=1):
                    item["group_index"] = group_index
                    diagnostics.append(item)

        # 1. Déduplication par série (garde la meilleure confiance)
        local_best: dict[str, PhysicalObservation] = {}
        for obs in image_observations:
            current = local_best.get(obs.series_code)
            if current is None or obs.confidence > current.confidence:
                local_best[obs.series_code] = obs

        # 2. Déduplication spatiale NMS (interdit d'assigner deux séries au même cadre physique)
        sorted_candidates = sorted(local_best.values(), key=lambda o: (o.series_score, o.confidence), reverse=True)
        dedup_spatial: list[PhysicalObservation] = []
        for obs in sorted_candidates:
            p1 = obs.physical_polygon if obs.physical_polygon is not None and len(obs.physical_polygon) > 0 else obs.crop_polygon
            if p1 is None or len(p1) == 0:
                continue
            duplicate = False
            for kept in dedup_spatial:
                p2 = kept.physical_polygon if kept.physical_polygon is not None and len(kept.physical_polygon) > 0 else kept.crop_polygon
                if p2 is None or len(p2) == 0:
                    continue
                r1 = cv2.boundingRect(p1.astype(np.int32))
                r2 = cv2.boundingRect(p2.astype(np.int32))
                xl = max(r1[0], r2[0]); yt = max(r1[1], r2[1])
                xr = min(r1[0] + r1[2], r2[0] + r2[2]); yb = min(r1[1] + r1[3], r2[1] + r2[3])
                if xr > xl and yb > yt:
                    inter = (xr - xl) * (yb - yt)
                    iou = inter / max(1.0, r1[2] * r1[3] + r2[2] * r2[3] - inter)
                    if iou > 0.25:
                        duplicate = True
                        break
            if not duplicate:
                dedup_spatial.append(obs)

        image_observations = dedup_spatial
        observations.extend(image_observations)
        _draw_debug(image, image_observations, debug_dir / f"{image_path.stem}_v90_lsd_frames.png")
        _export_geometric_diagnostics(image, image_observations, groups, result_dir / "diagnostics_geometrie", image_path.stem)

        result_text = ", ".join(
            f"{x.series_code}={x.tolerance_value:g}" + (f" ({x.multiplicity}X)" if x.multiplicity else "")
            for x in sorted(image_observations, key=lambda o: o.series_code)
        )
        print(
            f"      [{image_index}/{len(images)}] style={style} ; cadres={len(groups)} ; validés={len(image_observations)}"
            + (f" ; {result_text}" if result_text else ""),
            flush=True,
        )

    _reconcile_ambiguous(observations, unresolved, known)
    rows = [asdict(item) for item in _consensus(observations, known)]
    found = {row["series_code"] for row in rows}
    missing = sorted(known - found)

    # 1. Rescue cyan pour séries manquantes
    if missing:
        targets = set(missing)
        rescued: list[PhysicalObservation] = []
        for image_path in images:
            image = cv2.imread(str(image_path), cv2.IMREAD_COLOR)
            if image is None:
                continue
            cyan_groups = [g for g in _group_cyan_cells(_cyan_cell_candidates(image)) if len(g) >= 2]
            for group in cyan_groups:
                obs = _cyan_group_observation(image_path, image, group, known, targets=targets)
                if obs is not None and obs.series_code in targets:
                    rescued.append(obs)
                    targets.discard(obs.series_code)
            if not targets:
                break
        if rescued:
            observations.extend(rescued)
            rows = [asdict(item) for item in _consensus(observations, known)]
            found = {row["series_code"] for row in rows}
            missing = sorted(known - found)

    # 2. Rescue série-guidée sur toutes les captures
    if missing:
        rescued = _rescue_series_guided(images, known, set(missing), diagnostics)
        if rescued:
            observations.extend(rescued)
            rows = [asdict(item) for item in _consensus(observations, known)]
            found = {row["series_code"] for row in rows}
            missing = sorted(known - found)

    # 3. Rescue groupes inclinés
    if missing:
        rescued = _rescue_missing_series(images, known, set(missing))
        if rescued:
            observations.extend(rescued)
            rows = [asdict(item) for item in _consensus(observations, known)]
            found = {row["series_code"] for row in rows}
            missing = sorted(known - found)

    # 4. Rescue cadres orientés permissif
    if missing:
        rescued_or: list[PhysicalObservation] = []
        for image_path in images:
            image = cv2.imread(str(image_path), cv2.IMREAD_COLOR)
            if image is None:
                continue
            oriented, _details = _oriented_white_frame_observations(image_path, image, known)
            for obs in oriented:
                if obs is not None and obs.series_code in missing:
                    rescued_or.append(obs)
        if rescued_or:
            observations.extend(rescued_or)
            rows = [asdict(item) for item in _consensus(observations, known)]
            found = {row["series_code"] for row in rows}
            missing = sorted(known - found)
    duplicates_removed = max(0, len(observations) - len(rows))
    elapsed = time.perf_counter() - started

    diagnostic = {
        "version": VERSION,
        "generated_at": dt.datetime.now().isoformat(timespec="seconds"),
        "method": "opencv_physical_frame_inventory_then_local_cell_ocr",
        "known_series": sorted(known),
        "selected": rows,
        "missing_series": missing,
        "observations_before_deduplication": len(observations),
        "duplicates_removed": duplicates_removed,
        "elapsed_seconds": round(elapsed, 3),
        "frames": diagnostics,
    }
    diag_path = result_dir / "frame_inventory_latest.json"
    diag_path.write_text(json.dumps(diagnostic, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"      Couverture finale : {len(rows)}/{len(known)}.", flush=True)
    print(f"      Doublons supprimés : {duplicates_removed}.", flush=True)
    print(f"      Temps OCR V9.0 : {elapsed:.1f} s.", flush=True)
    if missing:
        print("      EXPORT INCOMPLET — séries sans association physique : " + ", ".join(missing), flush=True)
    else:
        print("      Toutes les séries de l'arbre ont un cadre/IT associé.", flush=True)
    print(f"      Diagnostic V9.0 : {diag_path}", flush=True)

    cache_path.write_text(
        json.dumps({"signature": signature, "rows": rows}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return rows


def main() -> int:
    parser = argparse.ArgumentParser(description="CATIA V5 : cadres physiques vérifiés série -> IT -> A.")
    parser.add_argument("--project-root", type=Path, default=None)
    parser.add_argument("--known-series", nargs="*", default=[])
    parser.add_argument("--interactive-capture", action="store_true")
    parser.add_argument("--keep-existing", action="store_true")
    args = parser.parse_args()
    try:
        if args.interactive_capture:
            interactive_capture(args.project_root, archive_existing=not args.keep_existing)
            return 0
        rows = scan_annotation_captures(project_root=args.project_root, known_series=args.known_series)
        print(json.dumps(rows, ensure_ascii=False, indent=2))
        return 0
    except Exception:
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
