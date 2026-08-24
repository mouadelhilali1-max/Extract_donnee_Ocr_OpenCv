"""
test_anti_overfitting.py
========================
Suite de tests unitaires synthétiques démontrant la généricité totale du moteur OCR.

Objectifs :
1. Prouver que l'extraction d'IT fonctionne fidèlement sur des valeurs arbitraires
   (ex. 0.25, 0.8, 1.4, 2.5, 3.2, 12.5) et ne dévie JAMAIS vers 1.0 ou 1.6.
2. Prouver que la désambiguïsation de séries fonctionne pour n'importe quel code
   (ex. 98X01 vs 98X02, 14F01 vs 14F02) sans aucune règle codée en dur.
3. Prouver que le consensus multi-captures détecte formellement les conflits
   (CONFLIT_IT_MULTI_CAPTURES) lorsque des captures divergentes se contredisent.
"""

import cv2
import numpy as np
from pathlib import Path
import sys
import unittest

project_root = Path(r"C:\Users\pc\Desktop\Projet_Stage\Projet_analyse")
sys.path.insert(0, str(project_root / "module_catia_cotes_fonctionnelles"))
import visual_annotation_scanner as scanner


def _create_synthetic_cell(text: str, width: int = 140, height: int = 65) -> np.ndarray:
    """Génère une cellule synthétique nette simulant le rendu CATIA."""
    img = np.zeros((height, width, 3), dtype=np.uint8)
    # Fond bleu foncé CATIA
    img[:] = (105, 52, 52)
    # Bords de cadre blancs
    cv2.rectangle(img, (0, 0), (width - 1, height - 1), (255, 255, 255), 2)
    # Texte blanc bien centré avec marges nettes
    font = cv2.FONT_HERSHEY_SIMPLEX
    scale = 0.85
    thickness = 2
    (tw, th), baseline = cv2.getTextSize(text, font, scale, thickness)
    tx = max(10, (width - tw) // 2)
    ty = max(th + 10, (height + th) // 2)
    cv2.putText(img, text, (tx, ty), font, scale, (255, 255, 255), thickness, cv2.LINE_AA)
    return img


class TestGenericityAndAntiOverfitting(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        scanner._configure_tesseract(project_root)

    def test_arbitrary_it_values(self):
        """Vérifie l'extraction exacte de valeurs non standard."""
        test_values = [0.25, 0.8, 1.4, 2.5, 3.2, 5.0, 12.5]
        for val in test_values:
            cell = _create_synthetic_cell(f"{val:g}")
            extracted, conf, texts = scanner.extract_exact_it_value(cell)
            self.assertIsNotNone(extracted, f"Échec d'extraction pour la valeur {val}")
            self.assertAlmostEqual(
                extracted, val, places=2,
                msg=f"La valeur {val} a été altérée en {extracted} (vérifier l'absence de mapping fixe)"
            )
            self.assertNotIn(
                extracted, [1.0, 1.6] if val not in (1.0, 1.6) else [],
                msg=f"La valeur {val} a été déviée vers 1.0 ou 1.6 !"
            )

    def test_series_disambiguation_arbitrary_codes(self):
        """Vérifie le scoring de séries arbitraires inconnues de la pièce de test."""
        pool = {"98X01", "98X02", "14F01", "14F02", "55K03"}
        
        # Test 1 : Lecture parfaite de 98X01
        scores = scanner._series_candidate_scores(["98X01"], pool)
        code, score, margin = scanner._choose_series(scores)
        self.assertEqual(code, "98X01")
        self.assertGreater(score, 0.90)

        # Test 2 : Lecture parfaite de 98X02
        scores = scanner._series_candidate_scores(["98X02"], pool)
        code, score, margin = scanner._choose_series(scores)
        self.assertEqual(code, "98X02")
        self.assertGreater(score, 0.90)

        # Test 3 : Confusion de caractère OCR 98X0I -> 98X01
        scores = scanner._series_candidate_scores(["98X0I"], pool)
        code, score, margin = scanner._choose_series(scores)
        self.assertEqual(code, "98X01")

    def test_multi_capture_conflict_detection(self):
        """Vérifie la détection explicite d'un conflit d'IT entre plusieurs captures."""
        dummy_poly = np.array([[0, 0], [10, 0], [10, 10], [0, 10]])
        # Capture 1 donne IT=2.5 avec confiance 0.92
        obs1 = scanner.PhysicalObservation(
            image_path=Path("view1.png"), crop=np.zeros((10, 10, 3), dtype=np.uint8),
            angle=0.0, crop_polygon=dummy_poly, texts=["2.5"],
            candidate_scores={"98X01": 0.95}, series_code="98X01", series_score=0.95,
            tolerance_value=2.5, multiplicity=None, datum_raw="A", datums={"A": True},
            layout="CADRE_REFERENCES", condition_text="", confidence=0.92,
            diagnostic="Observation 1"
        )
        # Capture 2 donne IT=0.8 avec confiance 0.90
        obs2 = scanner.PhysicalObservation(
            image_path=Path("view2.png"), crop=np.zeros((10, 10, 3), dtype=np.uint8),
            angle=0.0, crop_polygon=dummy_poly, texts=["0.8"],
            candidate_scores={"98X01": 0.95}, series_code="98X01", series_score=0.95,
            tolerance_value=0.8, multiplicity=None, datum_raw="A", datums={"A": True},
            layout="CADRE_REFERENCES", condition_text="", confidence=0.90,
            diagnostic="Observation 2"
        )

        results = scanner._consensus([obs1, obs2], {"98X01"})
        self.assertEqual(len(results), 1)
        res = results[0]
        self.assertEqual(res.read_status, "CONFLIT_IT_MULTI_CAPTURES")
        self.assertIn("CONFLIT_IT_MULTI_CAPTURES", res.diagnostic)
        self.assertIn("2.5", res.diagnostic)
        self.assertIn("0.8", res.diagnostic)
        self.assertLessEqual(res.confidence, 0.60)


if __name__ == "__main__":
    unittest.main()
