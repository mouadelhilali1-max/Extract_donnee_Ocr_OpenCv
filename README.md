# Extraction et Analyse Automatique des Tolérances Fonctionnelles CATIA V5

Ce projet automatise l'extraction, la détection visuelle et l'exportation des cotations fonctionnelles et tolérances géométriques (GD&T / GPS ISO) depuis **CATIA V5** vers un tableau **Microsoft Excel (.xlsx)** structuré.

---

## 🎯 Architecture Générale du Système

Le pipeline repose sur une architecture robuste à double canal :

1. **Extraction de l'Arbre CATIA (Automation COM)** :
   - Connexion directe à la session active de CATIA V5 via l'API Windows COM (`pywin32`).
   - Parcours de l'arbre de spécification (`Specification Tree`) pour identifier la hiérarchie `REF` > `Groupes fonctionnels` > `Séries` (ex: `01B01`, `02A01`, `06B01`...).
2. **Détection Visuelle par Vision par Ordinateur & OCR Local** :
   - Détection des cadres physiques réels via **OpenCV** (contours fermés, masques, `minAreaRect`).
   - Mesure de la structure interne des cellules via l'algorithme **LSD (Line Segment Detector)** pour isoler les parois verticales et horizontales.
   - Extraction ciblée de la **Cellule 2** pour la valeur $IT$ (Tolérance), avec validation par la présence de la référence $A$ en Cellule 3.
   - Détection des cadres conditionnels (2 cellules avec condition `HEIGHT > 6mm`, `HEIGHT < 6mm`).
   - Extraction des multiplicateurs ($Xn$) et des références ($A\dots E$).
3. **Rapprochement & Export Excel** :
   - Réconciliation Arbre ↔ Vision.
   - Génération d'un classeur Excel formaté avec styles professionnels et d'un rapport de diagnostic `.diagnostic.json`.

---

## 🛠️ Pré-requis Système

- **Système d'exploitation** : Windows 10 ou 11 (64 bits).
- **Python** : Version 3.10, 3.11 ou 3.12.
- **CATIA V5** : Installé et ouvert avec le fichier pièce (`.CATPart`) actif.
- **Microsoft Excel** : Installé sur la machine pour générer les classeurs `.xlsx`.
- **Tesseract OCR** :
  - Télécharger et installer [Tesseract OCR pour Windows](https://github.com/UB-Mannheim/tesseract/wiki).
  - Emplacement standard : `C:\Program Files\Tesseract-OCR\tesseract.exe`.

---

## 📦 Installation

1. **Cloner le dépôt Git** :
   ```powershell
   git clone <URL_DU_DEPOT_GITHUB>
   cd Projet_analyse
   ```

2. **Créer et activer un environnement virtuel Python** :
   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```

3. **Installer les dépendances Python** :
   ```powershell
   pip install --upgrade pip
   pip install -r requirements.txt
   ```

---

## 🚀 Exécution du Projet

### Option 1 : Lancement du Pipeline Complet de Production (avec CATIA ouvert)
Ouvrez CATIA V5 avec votre pièce active, puis lancez :
```powershell
python module_catia_cotes_fonctionnelles/catia_functional_tolerances.py
```
*Le programme lit l'arbre CATIA, analyse les captures d'annotations, réconcilie les données et génère le fichier Excel final dans `results/excel/`.*

### Option 2 : Évaluation & Test Visuel Autonome (sans CATIA ouvert)
Pour tester uniquement la détection géométrique et l'OCR sur le jeu de captures existant :
```powershell
python scratch/evaluate_geometry_detection.py
```

### Option 3 : Lancement des Tests Automatisés
```powershell
python -m unittest discover -s tests -p "test_*.py"
```

---

## 📂 Structure du Répertoire

```
Projet_analyse/
├── module_catia_cotes_fonctionnelles/
│   ├── catia_functional_tolerances.py   # Chef d'orchestre : lecture arbre COM & export Excel
│   └── visual_annotation_scanner.py     # Moteur de vision : OpenCV, parois LSD, OCR local
├── captures_annotations/                # Captures d'écrans CATIA des cadres de tolérances
├── tests/                               # Tests unitaires et de non-régression
├── results/                             # Fichiers Excel, JSON et diagnostics graphiques générés
├── requirements.txt                     # Liste des bibliothèques Python requises
├── .gitignore                           # Fichiers et dossiers exclus du suivi Git
└── README.md                            # Documentation du projet
```

