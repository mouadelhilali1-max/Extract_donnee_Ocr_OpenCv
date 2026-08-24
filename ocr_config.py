"""
=========================================================
OCR CONFIGURATION
Projet : CATIA Tree Extractor
=========================================================
"""

from pathlib import Path
import pytesseract

# =========================================================
# PROJECT PATHS
# =========================================================

# Racine du projet
PROJECT_ROOT = Path(__file__).parent

# Dossier contenant les captures de l'arbre
CAPTURES_DIR = PROJECT_ROOT / "captures"

# Dossier des résultats
RESULTS_DIR = PROJECT_ROOT / "results"

# Sous-dossiers
PREPROCESS_DIR = RESULTS_DIR / "preprocessed_images"
EXCEL_DIR = RESULTS_DIR / "excel"
CSV_DIR = RESULTS_DIR / "csv"
JSON_DIR = RESULTS_DIR / "json"
LOG_DIR = RESULTS_DIR / "logs"
REVIEW_DIR = RESULTS_DIR / "review"
# Dedicated outputs for the requested CATIA annotation-results subtree.  They
# are deliberately separate from the historical complete-tree exports and
# feedback queue, so switching modes cannot mix human corrections.
ANNOTATION_REVIEW_DIR = REVIEW_DIR / "annotations"

# Création automatique des dossiers
PREPROCESS_DIR.mkdir(parents=True, exist_ok=True)
EXCEL_DIR.mkdir(parents=True, exist_ok=True)
CSV_DIR.mkdir(parents=True, exist_ok=True)
JSON_DIR.mkdir(parents=True, exist_ok=True)
LOG_DIR.mkdir(parents=True, exist_ok=True)
REVIEW_DIR.mkdir(parents=True, exist_ok=True)
ANNOTATION_REVIEW_DIR.mkdir(parents=True, exist_ok=True)

# =========================================================
# TESSERACT
# =========================================================

# Modifier ce chemin uniquement si nécessaire
TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

# =========================================================
# OCR PARAMETERS
# =========================================================

# Français + Anglais
# Seul le modÃ¨le anglais est installÃ© avec le Tesseract local. Demander
# "fra+eng" provoquait une erreur Ã  chaque image puis un repli incomplet.
# Installer fra.traineddata permet ensuite de remettre "fra+eng".
TESSDATA_DIR = PROJECT_ROOT / "models" / "tessdata"
OCR_LANGUAGE = "fra+eng"
# The project path contains no spaces; leave it unquoted because pytesseract
# already applies Windows quoting and Tesseract would otherwise see quote
# characters as part of the tessdata directory.
TESSDATA_CONFIG = f"--tessdata-dir {TESSDATA_DIR}"

# PSM :
# 6 = bloc uniforme de texte (idéal pour l'arbre CATIA)
OCR_PSM = 11

# OEM :
# 3 = meilleur moteur disponible
OCR_OEM = 3

# Commande OCR complète
OCR_CONFIG = (
    f"{TESSDATA_CONFIG} --oem {OCR_OEM} "
    f"--psm {OCR_PSM}"
)

# The full-image pass detects candidate rows.  The line pass is used only on
# uncertain rows after cropping away CATIA icons.
FULL_IMAGE_CONFIG = f"{TESSDATA_CONFIG} --oem {OCR_OEM} --psm 11"
LINE_IMAGE_CONFIG = f"{TESSDATA_CONFIG} --oem {OCR_OEM} --psm 7"

# Robust geometry defaults.  They are estimates, not fixed coordinates: the
# extractor recalculates the indentation grid for every capture.
DEFAULT_INDENT_STEP = 28
TREE_PANEL_LEFT = 35
TREE_PANEL_RIGHT_MARGIN = 12
ROW_GROUP_TOLERANCE = 10
REGISTRATION_MIN_SHIFT = 10
REGISTRATION_MAX_SHIFT_MARGIN = 70

# =========================================================
# IMAGE PROCESSING
# =========================================================

# Niveau de gris
USE_GRAYSCALE = False

# Contraste
ENHANCE_CONTRAST = False

# Binarisation
USE_THRESHOLD = False

# Suppression du bruit
REMOVE_NOISE = False

# Taille minimale d'un texte accepté
MIN_TEXT_LENGTH = 2

# Seuil de confiance OCR
MIN_CONFIDENCE = 35

# CATIA icons and branch lines regularly become very tall OCR boxes.  Tree
# labels are compact, so reject those boxes before line reconstruction.
MIN_WORD_HEIGHT = 8
MAX_WORD_HEIGHT = 34
LINE_TOP_TOLERANCE = 7
OVERLAP_LOOKBACK = 80

# Indentation tolerance (pixels) pour déterminer les niveaux parents/enfants
# Augmenter si les captures ont des indentations larges. Ajuster si nécessaire.
INDENT_TOLERANCE = 18

# Détermination des niveaux :
# - PER_IMAGE_LEVELING = True : calcule les niveaux séparément pour chaque capture
# - PER_IMAGE_LEVELING = False : calcule globalement sur toutes les captures (recommandé pour arbres qui se découpent en plusieurs images)
PER_IMAGE_LEVELING = False

# Comportement pour la construction des parents :
# - RESET_PARENT_PER_IMAGE = True : réinitialise la pile de parents à chaque image
# - RESET_PARENT_PER_IMAGE = False : conserve la pile entre images (utile si l'arbre se prolonge d'une capture à l'autre)
RESET_PARENT_PER_IMAGE = False

# =========================================================
# EXPORT
# =========================================================

EXCEL_FILE = EXCEL_DIR / "tree_text.xlsx"
CSV_FILE = CSV_DIR / "tree_text.csv"
JSON_FILE = JSON_DIR / "tree_text.json"

# The annotation-only mode is the default export requested for this project.
# Change only this label if a future CATIA configuration uses another name.
ANNOTATION_TARGET_LABEL = "Résultat d'un ensemble d'annotations"
ANNOTATION_EXCEL_FILE = EXCEL_DIR / "annotation_tree.xlsx"
ANNOTATION_CSV_FILE = CSV_DIR / "annotation_tree.csv"
ANNOTATION_JSON_FILE = JSON_DIR / "annotation_tree.json"
ANNOTATION_CORRECTIONS_FILE = ANNOTATION_REVIEW_DIR / "corrections.csv"
ANNOTATION_TRAINING_FILE = ANNOTATION_REVIEW_DIR / "validated_ocr_training.jsonl"
ANNOTATION_FEEDBACK_APPLICATION_FILE = ANNOTATION_REVIEW_DIR / "feedback_application.csv"

# =========================================================
# DEBUG
# =========================================================

SAVE_PREPROCESSED_IMAGES = True

VERBOSE = True
