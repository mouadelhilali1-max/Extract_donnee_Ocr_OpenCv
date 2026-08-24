# Extraction et Analyse Automatique des Tolérances Fonctionnelles CATIA V5 (GD&T / ISO GPS)

Bienvenue dans le système automatisé d'extraction, d'analyse géométrique et d'exportation des cotations fonctionnelles et tolérances géométriques depuis **CATIA V5** vers **Microsoft Excel (.xlsx)**.

---

## 📖 Sommaire
1. [Architecture Générale](#-architecture-générale)
2. [Pré-requis Système de A à Z](#-pré-requis-système-de-a-à-z)
3. [Installation Pas à Pas](#-installation-pas-à-pas)
4. [Guide d'Exécution de A à Z](#-guide-dexécution-de-a-à-z)
5. [Structure des Fichiers et Résultats Produits](#-structure-des-fichiers-et-résultats-produits)
6. [Dépannage et Questions Fréquentes](#-dépannage-et-questions-fréquentes)

---

## 🎯 Architecture Générale

Le projet repose sur un pipeline hybride **« Arbre CATIA + Vision par Ordinateur (LSD & OCR Local) »** :

```
┌──────────────────────────┐        ┌───────────────────────────────┐
│     CATIA V5 ACTIF       │        │     CAPTURES D'ANNOTATIONS    │
│  (Session Windows COM)   │        │     (captures_annotations/)   │
└────────────┬─────────────┘        └───────────────┬───────────────┘
             │                                      │
             ▼                                      ▼
   read_part_tree()                     scan_annotation_captures()
  (Extraction de l'arbre                (Détection OpenCV + Parois LSD
    des séries 01B01...)                  + Isolation Cellule 2 IT)
             │                                      │
             └──────────────────┬───────────────────┘
                                │
                                ▼
                   enrich_rows_with_visual_ocr()
                    (Rapprochement & Consensus)
                                │
                                ▼
                       export_to_excel()
             (Génération du classeur Excel .xlsx
                 + Rapport .diagnostic.json)
```

1. **Extraction de l'Arbre CATIA (Automation COM)** : Récupération de la hiérarchie officielle `REF` > `Groupes fonctionnels` > `Séries`.
2. **Détection Géométrique des Cadres (OpenCV + LSD)** :
   - Détection des cadres par contours fermés (`minAreaRect`).
   - Mesure exacte des parois verticales et horizontales avec l'algorithme LSD (*Line Segment Detector*).
   - **Isolation stricte de la Cellule 2 ($IT$)** et confirmation par la Référence $A$ en Cellule 3.
   - Détection des cadres conditionnels à 2 cellules (`[ ⌓ | 1.4 ]` avec `HEIGHT > 6mm`).
   - Extraction des multiplicateurs ($Xn$) et des références ($A\dots E$).
3. **Export Excel Structuré** : Génération d'un classeur Excel formaté avec styles professionnels et d'un fichier diagnostic complet.

---

## 🛠️ Pré-requis Système de A à Z

Pour exécuter le projet sur n'importe quel ordinateur Windows, vérifiez les pré-requis suivants :

1. **Système d'exploitation** : Windows 10 ou Windows 11 (64 bits).
2. **Python** : Version 3.10, 3.11 ou 3.12 installée (avec `pip` et ajouté au `PATH`).
3. **Tesseract OCR pour Windows** :
   - Téléchargez et installez l'installeur officiel : [Tesseract OCR Windows Installer](https://github.com/UB-Mannheim/tesseract/wiki).
   - Installez-le dans le chemin par défaut : `C:\Program Files\Tesseract-OCR\`.
4. **CATIA V5** : Installé sur la machine (avec un document `.CATPart` ouvert pour le mode extraction directe).
5. **Microsoft Excel** : Installé sur la machine (requis par le module d'automatisation COM `win32com`).

---

## 📦 Installation Pas à Pas

Ouvrez un terminal **PowerShell** et suivez ces étapes :

### 1. Cloner le dépôt GitHub
```powershell
git clone https://github.com/mouadelhilali1-max/Extract_donnee_Ocr_OpenCv.git
cd Extract_donnee_Ocr_OpenCv
```

### 2. Créer et activer l'environnement virtuel Python
```powershell
# Création de l'environnement virtuel (.venv)
python -m venv .venv

# Activation sous PowerShell
.\.venv\Scripts\Activate.ps1
```
*(Si PowerShell bloque l'activation, autorisez les scripts locaux avec : `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned`)*

### 3. Installer les dépendances Python
```powershell
pip install --upgrade pip
pip install -r requirements.txt
```

---

## 🚀 Guide d'Exécution de A à Z

Le projet propose 3 modes d'exécution adaptés à chaque cas d'usage :

### Mode 1 : Pipeline Complet de Production (Recommandé avec CATIA V5 ouvert)
1. Ouvrez votre logiciel **CATIA V5**.
2. Ouvrez le fichier pièce (`.CATPart`) contenant les annotations et tolérances à analyser.
3. Dans votre terminal avec l'environnement virtuel activé, lancez :
   ```powershell
   python module_catia_cotes_fonctionnelles/catia_functional_tolerances.py
   ```
4. **Résultat** : Le programme se connecte à CATIA, lit l'arbre, scanne les annotations visuelles, réalise le consensus et crée automatiquement le fichier Excel `.xlsx` ainsi que le fichier `.diagnostic.json` dans `results/excel/`.

---

### Mode 2 : Évaluation Visuelle Autonome (Sans avoir besoin d'ouvrir CATIA)
Si vous souhaitez tester ou valider uniquement la détection géométrique et l'OCR sur les captures d'écran déjà enregistrées :
```powershell
python scratch/evaluate_geometry_detection.py
```
*Ce mode traite l'ensemble des images du dossier `captures_annotations/`, affiche les métriques de précision/rappel en temps réel et génère les images de debug dans `results/frame_inventory_ocr/`.*

---

### Mode 3 : Exécution des Tests Unitaires & Non-Régression
Pour vérifier l'intégrité de tous les algorithmes (tests synthétiques et non-régression) :
```powershell
python -m unittest discover -s tests -p "test_*.py"
```

---

## 📂 Structure des Fichiers et Résultats Produits

```
Extract_donnee_Ocr_OpenCv/
│
├── module_catia_cotes_fonctionnelles/
│   ├── catia_functional_tolerances.py   # Chef d'orchestre : connexion CATIA COM & export Excel
│   └── visual_annotation_scanner.py     # Moteur de vision : OpenCV, parois LSD, OCR localisé
│
├── captures_annotations/                # Jeu de captures réelles pour le scan visuel
├── models/tessdata/                     # Fichiers de données linguistiques OCR (fra, eng)
├── tests/                               # Tests automatisés (synthétiques et réels)
│   ├── test_genericity_synthetic.py     # Tests de généricité sur tolérances arbitraires
│   └── test_regression_valeurs_connues.py # Tests de non-régression
│
├── results/                             # Dossier de sortie généré automatiquement
│   ├── excel/                           # Classeurs finaux (.xlsx) avec styles professionnels
│   └── frame_inventory_ocr/             # Diagnostics JSON et images de debug des cadres
│
├── requirements.txt                     # Liste des dépendances Python (pywin32, opencv, pytesseract...)
├── .gitignore                           # Exclusions Git (fichiers temporaires, caches)
└── README.md                            # Documentation complète du projet
```

---

## ❓ Dépannage et Questions Fréquentes

- **Erreur `TesseractNotFoundError`** :
  Assurez-vous que Tesseract est bien installé dans `C:\Program Files\Tesseract-OCR\tesseract.exe`. Si vous l'avez installé ailleurs, définissez la variable d'environnement :
  `$env:TESSERACT_CMD = "C:\MonChemin\tesseract.exe"`.
- **Erreur `CATIA.Application introuvable`** :
  Vérifiez que CATIA V5 est bien lancé et qu'un document `.CATPart` est actif au premier plan.
- **Vider le cache d'OCR** :
  Pour forcer un recalcul brut complet sans cache, supprimez les fichiers `results/frame_inventory_ocr/cache_*.json`.


