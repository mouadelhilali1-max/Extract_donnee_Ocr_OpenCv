import argparse

from catia.connection import CATIAConnection
from catia.explorer import CATIAExplorer
from catia.label_inspector import CATIALabelInspector


def main():
    parser = argparse.ArgumentParser(description="Extraire l'arbre de spécification CATIA ou un sous-arbre spécifique.")
    parser.add_argument("--name", "-n", dest="target_name", help="Nom du noeud à extraire avec tous ses enfants.")
    parser.add_argument("--output", "-o", dest="output_path", default="output/catia_tree.txt", help="Chemin du fichier de sortie.")
    parser.add_argument("--debug", dest="debug", action="store_true", help="Mode debug : enregistre les éléments retournés par Selection.Search dans un fichier JSON.")
    parser.add_argument("--match", dest="match", choices=["best", "exact", "contains", "startswith"], default="best", help="Mode de correspondance pour le nom demandé : best (exact puis contains), exact, contains, startswith.")
    parser.add_argument("--paths-only", dest="paths_only", action="store_true", help="Afficher/écrire uniquement les chemins complets trouvés pour le nom demandé.")
    parser.add_argument("--inspect-index", dest="inspect_index", type=int, help="(DEBUG) Inspecte l'objet de Selection.Item(index) et écrit ses attributs/méthodes accessibles en JSON.")
    parser.add_argument("--inspect-tolerance-zones", dest="inspect_tolerance_zones", action="store_true", help="(DEBUG) Inspecte les objets ToleranceZone pour découvrir leurs labels visibles.")
    args = parser.parse_args()

    catia = CATIAConnection()
    if not catia.connect():
        return

    explorer = CATIAExplorer(catia.document, catia.part, output_path=args.output_path)
    if args.inspect_index is not None:
        explorer.inspect_index(args.inspect_index)
        return
    if args.inspect_tolerance_zones:
        inspector = CATIALabelInspector(catia.document, output_dir=explorer.output_path.parent)
        inspector.inspect_tolerance_zones()
        return
    explorer.explore(target_name=args.target_name, debug=args.debug, match_mode=args.match, paths_only=args.paths_only)


if __name__ == "__main__":
    main()
