# Extraction et Analyse Automatique des Tolérances Fonctionnelles CATIA V5 (GD&T / ISO GPS)

Ce projet fournit une solution logicielle industrielle automatisée pour extraire, analyser et réconcilier les tolérances fonctionnelles (GD&T / ISO GPS) depuis **CATIA V5** vers **Microsoft Excel (.xlsx)** en combinant l'**API Automation COM** et la **Vision par Ordinateur (OpenCV, LSD et OCR local)**.

---

## 📖 Sommaire
1. [Pré-requis Système et Outils Indispensables](#1-pré-requis-système-et-outils-indispensables)
2. [Installation Complète de Toutes les Bibliothèques (Pas à Pas)](#2-installation-complète-de-toutes-les-bibliothèques-pas-à-pas)
3. [Déroulement Chronologique d'Exécution du Programme](#3-déroulement-chronologique-dexécution-du-programme)
   - [Étape 1 : Prise des Captures de la Pièce 3D (`visual_annotation_scanner.py`)](#étape-1--prise-des-captures-de-la-pièce-3d)
   - [Étape 2 : Extraction de l'Arbre et Calcul des Tolérances (`catia_functional_tolerances.py`)](#étape-2--extraction-de-larbre-et-calcul-des-tolérances)
4. [Où Trouver Tous les Résultats et Fichiers Générés ?](#4-où-trouver-tous-les-résultats-et-fichiers-générés-)
5. [Procédure pour Traiter une Nouvelle Pièce (Vidage du Cache)](#5-procédure-pour-traiter-une-nouvelle-pièce-vidage-du-cache)
6. [Dépannage et Questions Fréquentes](#6-dépannage-et-questions-fréquentes)

---

## 1. Pré-requis Système et Outils Indispensables

Avant de lancer le projet sur votre poste Windows, assurez-vous que les éléments suivants sont installés :

1. **Système d'exploitation** : Windows 10 ou Windows 11 (64 bits).
2. **Python** : Version **3.10, 3.11 ou 3.12** installée (avec `pip` activé et la case *Add Python to PATH* cochée).
3. **Tesseract OCR pour Windows** (Moteur de reconnaissance optique) :
   - Téléchargez l'installateur officiel : [Tesseract OCR Windows 64-bit Installer](https://github.com/UB-Mannheim/tesseract/wiki).
   - Installez-le dans son répertoire par défaut : `C:\Program Files\Tesseract-OCR\`.
   - *(Vérification)* : Le fichier `C:\Program Files\Tesseract-OCR\tesseract.exe` doit être présent.
4. **CATIA V5** : Installé sur la machine avec une licence active.
5. **Microsoft Excel** : Installé sur la machine (nécessaire pour l'exportation et le formatage automatique via COM).

---

## 2. Installation Complète de Toutes les Bibliothèques (Pas à Pas)

Ouvrez un terminal **PowerShell** ou **Git Bash** et exécutez les étapes suivantes :

### Étape A : Cloner ou Télécharger le Dépôt
```powershell
git clone https://github.com/mouadelhilali1-max/Extract_donnee_Ocr_OpenCv.git
cd Extract_donnee_Ocr_OpenCv
```

### Étape B : Créer et Activer l'Environnement Virtuel Python (`.venv`)
```powershell
# 1. Création de l'environnement virtuel
python -m venv .venv

# 2. Activation sous PowerShell
.\.venv\Scripts\Activate.ps1
```
*(Si PowerShell bloque l'activation avec une erreur de script, lancez une fois : `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned` puis réactivez).*

### Étape C : Installer TOUTES les Bibliothèques Nécessaires

Vous pouvez installer l'ensemble des dépendances en une seule commande :
```powershell
pip install --upgrade pip
pip install -r requirements.txt
```

#### Détail exhaustif de chaque bibliothèque installée :
Si vous souhaitez installer ou vérifier chaque bibliothèque manuellement une par une :
```powershell
pip install opencv-python   # Traitement d'images, détection des contours et transformation affine
pip install numpy           # Calculs matriciels, trigonométrie et projections géométriques
pip install pandas          # Structuration et manipulation des tableaux de données
pip install pytesseract     # Interface Python avec le binaire Tesseract OCR
pip install pywin32         # Connexion API COM avec CATIA V5 et Microsoft Excel
pip install openpyxl        # Génération et styles des classeurs Excel .xlsx
pip install mss             # Capture d'écran ultra-rapide et légère
pip install PyGetWindow     # Détection et positionnement de la fenêtre active CATIA V5
pip install PyAutoGUI       # Automatisation du focus et de la souris
pip install Pillow          # Manipulation et découpages d'images (PNG)
```

---

## 3. Déroulement Chronologique d'Exécution du Programme

Le traitement de votre pièce mécanique s'effectue en **2 étapes séquentielles exactes** :

```
                  FLUX CHRONOLOGIQUE D'EXÉCUTION
                  ═════════════════════════════

  [ CATIA V5 Ouvert avec la Pièce 3D ]
                  │
                  ▼
  ┌────────────────────────────────────────────────────────┐
  │ ÉTAPE 1 : Prise des Captures 3D                        │
  │ (visual_annotation_scanner.py -> captures_annotations/)│
  └───────────────────────┬────────────────────────────────┘
                          │
                          ▼
  ┌────────────────────────────────────────────────────────┐
  │ ÉTAPE 2 : Extraction de l'Arbre & Calcul des Tolérances│
  │ (catia_functional_tolerances.py -> exports/*.xlsx)     │
  └───────────────────────┬────────────────────────────────┘
                          │
                          ▼
  ┌────────────────────────────────────────────────────────┐
  │ RÉSULTATS : Excel structuré + Diagnostics + Debug vert │
  └────────────────────────────────────────────────────────┘
```

---

### Étape 1 : Prise des Captures de la Pièce 3D

Avant d'extraire les données, le système a besoin des captures des cadres visibles dans l'espace 3D de CATIA.

1. Ouvrez votre pièce (`.CATPart`) dans **CATIA V5** et orientez la vue 3D pour faire apparaître clairement un groupe de cadres de tolérances.
2. Lancez l'outil de capture interactive dans votre terminal :
   ```powershell
   python -c "import sys; sys.path.insert(0, 'module_catia_cotes_fonctionnelles'); import visual_annotation_scanner as s; s.interactive_capture()"
   ```
3. Suivez le guide affiché :
   - Basculez sur la fenêtre CATIA V5 et appuyez sur la touche indiquée pour enregistrer la vue actuelle.
   - Tournez la pièce dans CATIA pour cadrer les autres tolérances et prenez 1 ou 2 autres captures pour couvrir toutes les séries.
4. Les images sont enregistrées automatiquement dans le dossier **`captures_annotations/`**.

*(Note : Vous pouvez également placer manuellement vos propres captures d'écran `.png` directement dans le dossier `captures_annotations/`).*

---

### Étape 2 : Extraction de l'Arbre et Calcul des Tolérances

Dès que vos captures sont prêtes dans `captures_annotations/`, lancez le programme principal :

```powershell
python module_catia_cotes_fonctionnelles/catia_functional_tolerances.py
```

**Ce que fait automatiquement le programme :**
1. **Connexion COM** : Se rattache à CATIA V5.
2. **Extraction de l'Arbre** : Lit l'arbre gauche et récupère la liste officielle des séries (`known_series` : `01A01`, `02A01`, `06A01`...).
3. **Moteur de Vision V9.0** :
   - Détecte les cadres OpenCV et redresse les angles (`rotation affine`).
   - Mesure les parois physiques par l'algorithme **LSD** (*Line Segment Detector*).
   - **Isole strictement la Cellule 2 ($IT$)** et vérifie la Référence $A$ en Cellule 3.
   - Extrait les multiplicateurs ($Xn$) dans la bande supérieure.
4. **Consensus et Déduplication** : Fusionne les vues multiples, supprime les doublons et élimine les faux positifs ($FP = 0$).
5. **Livrables Finaux** : Génère le fichier Excel formaté et le rapport de diagnostic JSON.

---

## 4. Où Trouver Tous les Résultats et Fichiers Générés ?

Après l'exécution, tous les fichiers sont stockés dans des dossiers clairs :

### 📊 1. Le Classeur Excel Final (.xlsx)
Formaté avec les colonnes : *Série, Groupe Fonctionnel, Référence, Tolérance IT, Multiplicité, Condition, Confiance, Source* :
- 📁 **Emplacement :** `exports/` (ou `results/excel/`)
- 📄 **Fichier :** `exports/<NomDeLaPiece>_cotes_fonctionnelles_<Date_Heure>.xlsx`
  *(Exemple : `exports/R20_cotes_fonctionnelles_20260825_035825.xlsx`)*

### 🔍 2. Le Rapport de Traçabilité et Diagnostic JSON
- 📁 **Emplacement :** `exports/<NomDeLaPiece>_cotes_fonctionnelles_<Date_Heure>.diagnostic.json`

### 🖼️ 3. Les Images de Debug avec Cadres Verts Détectés
Contient les captures de la pièce avec **un cadre vert tracé autour de chaque tolérance et le texte OCR en vert au-dessus** (ex: `03A03 IT=1.6`, `06A02 IT=1.4`) :
- 📁 **Emplacement :** `results/frame_inventory_ocr/cadres_detectes/`
  *(Exemple : `annotation_view_01_..._v90_lsd_frames.png`)*

### 🗂️ 4. L'Inventaire Consolidé des Cadres
- 📁 **Emplacement :** `results/frame_inventory_ocr/frame_inventory_latest.json`

---

## 5. Procédure pour Traiter une Nouvelle Pièce (Vidage du Cache)

Lorsque vous changez de **fichier pièce (`.CATPart`)** ou de **jeu de captures**, vous devez réinitialiser le cache pour garantir un calcul 100 % neuf :

1. **Exécutez la commande de vidage sous PowerShell** :
   ```powershell
   python -c "from pathlib import Path; import shutil; root = Path('.'); [c.unlink() for c in root.glob('results/**/cache*.json')]; [c.unlink() for c in root.glob('results/**/frame_inventory_*.json')]; [shutil.rmtree(p) for p in root.glob('**/__pycache__') if p.is_dir()]; print('>>> Cache vidé avec succès !')"
   ```

2. **Nettoyez le dossier des captures et avancez la nouvelle pièce** :
   - Placez les nouvelles captures dans `captures_annotations/`.

3. **Relancez le traitement complet** :
   ```powershell
   python module_catia_cotes_fonctionnelles/catia_functional_tolerances.py
   ```

4. **Récupérez votre nouveau fichier Excel dans `exports/`**.

---

## 6. Dépannage et Questions Fréquentes

- **Erreur `TesseractNotFoundError`** :  
  Vérifiez que Tesseract est installé dans `C:\Program Files\Tesseract-OCR\tesseract.exe` (variable `$env:TESSERACT_CMD`).
- **Erreur `CATIA.Application introuvable`** :  
  Vérifiez que CATIA V5 est bien démarré et qu'un fichier `.CATPart` est actif.
- **Tolérance manquante ou non reconnue** :  
  Tournez la pièce 3D dans CATIA pour prendre une capture où la tolérance est bien nette et réexécutez.
