# Extraction et Analyse Automatique des Tolérances Fonctionnelles CATIA V5 (GD&T / ISO GPS)

Bienvenue dans le système automatisé d'extraction, d'analyse géométrique et d'exportation des cotations fonctionnelles et tolérances géométriques depuis **CATIA V5** vers **Microsoft Excel (.xlsx)**.

---

## 📖 Sommaire
1. [Architecture Générale](#-architecture-générale)
2. [Pré-requis Système de A à Z](#-pré-requis-système-de-a-à-z)
3. [Installation Pas à Pas](#-installation-pas-à-pas)
4. [Toutes les Méthodes d'Exécution de A à Z](#-toutes-les-méthodes-dexécution-de-a-à-z)
5. [Procédure pour Traiter une Nouvelle Pièce (Vidage du Cache)](#-procédure-pour-traiter-une-nouvelle-pièce-vidage-du-cache)
6. [Structure des Fichiers et Résultats Produits](#-structure-des-fichiers-et-résultats-produits)
7. [Dépannage et Questions Fréquentes](#-dépannage-et-questions-fréquentes)

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

## 🚀 Toutes les Méthodes d'Exécution de A à Z

Le projet propose différentes commandes selon vos besoins :

### Méthode 1 : Pipeline Complet de Production (Arbre CATIA + Scan Visuel + Export Excel)
À utiliser lorsque **CATIA V5 est ouvert** avec votre pièce (`.CATPart`) active :
```powershell
python module_catia_cotes_fonctionnelles/catia_functional_tolerances.py
```
- **Ce que fait la commande** :
  1. Se connecte à CATIA V5 et lit l'arbre des spécifications.
  2. Lance automatiquement le scanner visuel sur les captures.
  3. Rapproche chaque série avec son cadre physique.
  4. Crée le classeur Excel final formaté dans `results/excel/`.
  5. Génère le rapport de diagnostic JSON `.diagnostic.json`.

---

### Méthode 2 : Lancement Direct du Scanner Visuel d'Annotations
Pour exécuter directement le moteur de vision et d'OCR sans passer par CATIA :
```powershell
python module_catia_cotes_fonctionnelles/visual_annotation_scanner.py
```
- **Ce que fait la commande** :
  - Scanne toutes les captures présentes dans `captures_annotations/`.
  - Effectue la détection géométrique LSD, l'isolation de la cellule $IT$, l'extraction des références et des conditions.
  - Déduplique et génère le fichier `results/frame_inventory_ocr/frame_inventory_latest.json` ainsi que les images de debug dans `results/frame_inventory_ocr/cadres_detectes/`.

---

### Méthode 3 : Évaluation & Rapport de Benchmark des Performances
Pour mesurer la précision, le rappel et le taux de couverture global :
```powershell
python scratch/evaluate_geometry_detection.py
```

---

### Méthode 4 : Exécution des Tests Automatisés (Synthétiques & Non-Régression)
Pour valider l'intégrité du code :
```powershell
python -m unittest discover -s tests -p "test_*.py"
```

---

## 🔄 Procédure pour Traiter une Nouvelle Pièce (Vidage du Cache)

Lorsque vous passez à un **autre fichier pièce (`.CATPart`)** ou à un **nouveau jeu de captures**, vous devez vider le cache pour garantir un calcul 100% neuf sans réutiliser les anciennes données.

### Étape 1 : Vider le cache en une seule commande PowerShell
```powershell
python -c "from pathlib import Path; import shutil; root = Path('.'); [c.unlink() for c in root.glob('results/**/cache*.json')]; [c.unlink() for c in root.glob('results/**/frame_inventory_*.json')]; [shutil.rmtree(p) for p in root.glob('**/__pycache__') if p.is_dir()]; print('Cache vidé avec succès !')"
```

### Étape 2 : Mettre à jour les captures
- Placez les nouvelles captures d'écran de votre pièce dans le dossier `captures_annotations/` (au format `.png`).
- *(Optionnel)* : Vous pouvez utiliser la fonction de capture automatique intégrée si CATIA est ouvert :
  ```powershell
  python -c "import sys; sys.path.insert(0, 'module_catia_cotes_fonctionnelles'); import visual_annotation_scanner as s; s.interactive_capture()"
  ```

### Étape 3 : Relancer le traitement
```powershell
python module_catia_cotes_fonctionnelles/catia_functional_tolerances.py
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
└── README.md                            # Documentation complète du projet de A à Z
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


