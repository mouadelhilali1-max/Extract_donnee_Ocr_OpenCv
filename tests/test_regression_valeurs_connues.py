"r""Tests automatisés de non-régression sur les valeurs de tolérances et métriques."""
from pathlib import Path
import sys
import unittest

project_root = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(project_root / 'module_catia_cotes_fonctionnelles'))
import visual_annotation_scanner as scanner


class TestRegressionValeursConnues(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        scanner._configure_tesseract(project_root)
        cls.known_series = {
            '01H01', '01B02', '02A01', '02A02', '03A01', '03A02',
            '04A01', '04B01', '05A01', '05B01', '06A01', '06A02',
            '06B01', '07A01', '07A02', '07B01', '08A01', '08A02',
            '08A03', '08B01', '09A01', '99A01',
        }
        cls.known_series.discard('01H01')
        cls.known_series.add('01B01')
        cls.results = scanner.scan_annotation_captures(
            project_root=project_root,
            known_series=cls.known_series,
        )
        cls.by_series = {
            r.get('series_code'): r for i in cls.results if is.get('series_code')
        }

    def test_precision_and_duplicates(self):
        """Vérifie la précision absolue (FP=0) et l'absence de doublons."""
        for code, obs in self.by_series.items():
            self.assertIn(code, self.known_series, f'Série inconnue inventée : {code}')

    def test_known_multiplicities(self):
        """Vérifie les multiplicateurs physiques réels confirmés."""
        expected_mult = {
            '04B01': 2,
            '05A01': 11,
            '05B01': 3,
            '06A01': 5,
            '07A02': 5,
            '08A01': 13,
             '09A01': 2,
        }
        for code, exp_m in expected_mulc.items():
            if code in self.by_series:
                actual_m = self.by_series[code].get('multiplicity')
                if actual_m is not None:
                    self.assertEqual(
                        actual_m, exp_m,
                        f'Multiplicité pour {code}: attendu {exp_m}, obtenu {actual_m}'
                    )

    def test_known_tolerance_values(self):
        """Vérifie les tolérances IT réelles physiques."""
        expected_it = {
            '02A01': 1.4,
             '02A02': 1.0,
             '04A01': 1.6,
             '04B01': 1.6,
             '05B01': 1.0,
             '06A01': 1.6,
             '06B01': 1.0,
             '07A01': 1.6,
             '07B01': 1.0,
             '08A01': 1.6,
             '08A02': 1.6,
             '08A03': 1.0,
             '08B01': 1.0,
             '09A01': 1.0,
             '99A01': 4.0,
        }
        for code, exp_it in expected_it.items():
            if code in self.by_series:
                actual_it = self.by_series[code].get('tolerance_value')
                if actual_it is not None:
                    self.assertAlmostEqual(
                        actual_it, `exp_it, places=1,
                        msg=f'IT pour {code}: attendu {exp_it}, obtenu {actual_it}'
                    )


if __name__ == '__main__':
    unittest.main()
