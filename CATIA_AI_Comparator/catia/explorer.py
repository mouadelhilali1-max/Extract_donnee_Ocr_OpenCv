"""Extraction de la structure accessible d'un CATPart."""

from pathlib import Path
import re
import unicodedata
import json


class _TreeNode:
    def __init__(self, label):
        self.label = label
        self.children = {}

    def add_path(self, path):
        node = self
        for label in path:
            node = node.children.setdefault(label, _TreeNode(label))

    def add_child(self, label):
        return self.children.setdefault(label, _TreeNode(label))

    def lines(self, level=0):
        for child in self.children.values():
            yield f"{'    ' * level}- {child.label}"
            yield from child.lines(level + 1)

    def subtree_lines(self, level=0):
        yield f"{'    ' * level}- {self.label}"
        for child in self.children.values():
            yield from child.subtree_lines(level + 1)


class CATIAExplorer:
    """Extrait seulement les noeuds visibles de l'arbre de specification."""

    def __init__(self, document, part, output_path="output/catia_tree.txt"):
        self.document = document
        self.part = part
        self.output_path = Path(output_path)
        self.label_map = self._load_label_map()

    def explore(self, target_name=None, debug=False, match_mode='best', paths_only=False):
        # debug: when True, dump every item returned by Selection.Search to a JSON file
        self.debug = bool(debug)
        self.match_mode = match_mode
        self.paths_only = bool(paths_only)
        root_name = self._name_of(self.part)
        tree = _TreeNode(root_name)
        report = [
            "=" * 70,
            "ARBRE DE SPECIFICATION CATIA",
            "=" * 70,
            root_name,
        ]
        self.count = 0
        self._add_origin_elements(tree)
        self._add_hybrid_bodies(self.part, tree)
        self._add_bodies(self.part, tree)
        self._add_collection(self.part, "HybridShapes", tree, "HybridShapes")
        self._add_collection(self.part, "Sketches", tree, "Sketches")
        self._add_annotation_sets(tree)
        self._add_collection(self.part, "Publications", tree, "Publication")
        self._add_fta_objects_found_by_search(tree)
        self._add_annotation_results_from_selection(tree, root_name)
        self._build_tree_from_selection(tree, root_name)
        self._add_annotation_result_root(tree, root_name)

        if target_name:
            # Chercher d'abord dans l'arbre reconstruit en utilisant le mode de match demandé
            matches = list(self._find_subtrees(tree, target_name, match_mode=self.match_mode))

            # Si on a trouvé des noeuds, mais qu'ils sont des feuilles synthétiques
            # (comme 'Captures' -> 'ALL' ajoutés par grouping), reconstruire le
            # sous-arbre complet en parcourant Selection et en ajoutant tous les
            # chemins qui contiennent la cible dans leur chemin minimal.
            if matches:
                # build a fresh subtree containing complete paths for the target(s)
                subtree_root = _TreeNode(target_name)
                paths = list(self._collect_paths_for_target(target_name, root_name))

                if getattr(self, 'paths_only', False):
                    out_paths = []
                    for p in paths:
                        try:
                            norm = self._normalize_string(target_name)
                            idx = next(i for i, lbl in enumerate(p) if self._normalize_string(lbl) == norm)
                        except StopIteration:
                            idx = 0
                        out_paths.append(p[idx:])

                    paths_file = self.output_path.parent / f"{self.output_path.stem}_paths.json"
                    paths_file.parent.mkdir(parents=True, exist_ok=True)
                    try:
                        paths_file.write_text(json.dumps(out_paths, ensure_ascii=False, indent=2), encoding='utf-8')
                        print(f"Paths written: {paths_file.resolve()}")
                    except Exception as e:
                        print(f"Failed writing paths file: {e}")

                    report = [
                        "=" * 70,
                        f"CHEMINS TROUVES POUR : {target_name}",
                        "=" * 70,
                    ]
                    for p in out_paths:
                        report.append("/".join(p))
                    self.count = len(out_paths)

                else:
                    if paths:
                        for p in paths:
                            try:
                                norm = self._normalize_string(target_name)
                                idx = next(i for i, lbl in enumerate(p) if self._normalize_string(lbl) == norm)
                            except StopIteration:
                                rel = p
                                subtree_root.add_path(rel)
                                continue

                            rel = p[idx:]
                            subtree_root.add_path(rel)

                        report = [
                            "=" * 70,
                            f"SOUS-ARBRE POUR : {target_name}",
                            "=" * 70,
                        ]
                        self.count = 0
                        report.extend(subtree_root.subtree_lines())
                        self.count += sum(1 for _ in subtree_root.subtree_lines())

                    else:
                        # Fallback : utiliser le sous-arbre déjà reconstruit à partir du tree
                        report = [
                            "=" * 70,
                            f"SOUS-ARBRE POUR : {target_name}",
                            "=" * 70,
                        ]
                        self.count = 0
                        for subtree in matches:
                            report.extend(subtree.subtree_lines())
                            self.count += sum(1 for _ in subtree.subtree_lines())
            else:
                report = [
                    "=" * 70,
                    f"Aucun noeud trouve pour : {target_name}",
                    "=" * 70,
                ]
                self.count = 0
        else:
            report.extend(tree.lines())

        report.append("=" * 70)
        report.append(f"Noeuds affiches : {self.count}")

        result = "\n".join(report)
        print(result)
        self.output_path.parent.mkdir(parents=True, exist_ok=True)
        self.output_path.write_text(result + "\n", encoding="utf-8")
        print(f"\nFichier cree : {self.output_path.resolve()}")

    def _load_label_map(self):
        map_path = self.output_path.parent / "tolerance_zone_label_map.json"
        if not map_path.exists():
            return {}
        try:
            return json.loads(map_path.read_text(encoding='utf-8'))
        except Exception:
            return {}

    def _lookup_label_map(self, technical_name):
        if not technical_name:
            return None
        if technical_name in self.label_map:
            return self.label_map[technical_name]
        normalized = technical_name.strip().lower()
        for key, value in self.label_map.items():
            if key.strip().lower() == normalized:
                return value
        return None

    def _add_origin_elements(self, tree):
        try:
            origin = self.part.OriginElements
        except Exception:
            return
        for property_name in ("PlaneXY", "PlaneYZ", "PlaneZX"):
            try:
                tree.add_child(self._name_of(getattr(origin, property_name)))
                self.count += 1
            except Exception:
                continue

    def _add_hybrid_bodies(self, owner, parent_node):
        self._add_collection(owner, "HybridBodies", parent_node, on_item=self._fill_hybrid_body)

    def _fill_hybrid_body(self, hybrid_body, node):
        # Ces deux collections correspondent aux enfants affiches sous un
        # ensemble geometrique dans l'arbre CATIA.
        self._add_collection(hybrid_body, "HybridBodies", node, on_item=self._fill_hybrid_body)
        self._add_collection(hybrid_body, "Bodies", node, on_item=self._fill_body)
        self._add_collection(hybrid_body, "HybridShapes", node)

    def _add_bodies(self, owner, parent_node):
        self._add_collection(owner, "Bodies", parent_node, on_item=self._fill_body)

    def _fill_body(self, body, node):
        self._add_collection(body, "HybridBodies", node, on_item=self._fill_hybrid_body)
        self._add_collection(body, "Shapes", node)
        self._add_collection(body, "Sketches", node)

    def _add_annotation_sets(self, tree):
        annotation_sets = self._get_collection(self.part, "AnnotationSets")
        if annotation_sets is None:
            return

        for annotation_set in self._items(annotation_sets):
            set_node = tree.add_child(self._name_of(annotation_set))
            self.count += 1
            # Les noms de dossiers sont ceux visibles dans le panneau CATIA.
            self._add_collection(annotation_set, "Captures", set_node, "Captures")
            self._add_collection(annotation_set, "Views", set_node, "Vues")
            self._add_collection(annotation_set, "AnnotatedViews", set_node, "Vues")
            self._add_collection(annotation_set, "AnnotationPlanes", set_node, "Références")
            self._add_annotations(annotation_set, set_node)

    def _add_annotations(self, annotation_set, parent_node):
        collection = self._get_collection(annotation_set, "Annotations")
        if collection is None:
            return
        for annotation in self._items(collection):
            folder = self._annotation_folder(annotation)
            node = parent_node.add_child(folder)
            node.add_child(self._name_of(annotation))
            self.count += 1

    def _add_fta_objects_found_by_search(self, tree):
        """Secours pour les versions CATIA qui ne remplissent pas Part.AnnotationSets.

        Certains documents FTA affichent l'ensemble d'annotations dans l'arbre
        mais ne le retournent pas dans Part.AnnotationSets. La recherche est
        alors utilisee uniquement pour les types visibles FTA et Publication,
        jamais pour la geometrie interne.
        """
        selection = self.document.Selection
        annotation_sets = []
        captures = []
        views = []
        references = []
        annotations = []
        publications = []

        try:
            selection.Clear()
            selection.Search("Name=*,all")
            for index in range(1, self._selection_count(selection) + 1):
                try:
                    item = self._selection_item(selection, index)
                    item_type = self._type_name(item).lower()
                    name = self._name_of(item)
                    lower_name = self._normalize_string(name)

                    if "annotationset" in item_type or "ensemble d'annotations" in lower_name:
                        annotation_sets.append(name)
                    elif item_type == "capture" or item_type.endswith(".capture") or "capture" in lower_name:
                        captures.append(name)
                    elif item_type == "view" or item_type.endswith(".view") or "vue" in lower_name:
                        views.append(name)
                    elif item_type == "annotationplane" or "annotationplane" in item_type or "reference" in lower_name:
                        references.append(name)
                    elif item_type == "annotation" or item_type.endswith(".annotation") or "annotation" in lower_name:
                        annotations.append((name, self._annotation_folder(item)))
                    elif "publication" in item_type:
                        publications.append(name)
                except Exception:
                    continue
        except Exception:
            return
        finally:
            try:
                selection.Clear()
            except Exception:
                pass

        set_names = list(dict.fromkeys(annotation_sets))
        capture_names = list(dict.fromkeys(captures))
        view_names = list(dict.fromkeys(views))
        reference_names = list(dict.fromkeys(references))
        annotation_items = list(dict.fromkeys(annotations))

        if set_names:
            for set_name in set_names:
                set_node = tree.add_child(set_name)
                if capture_names:
                    capture_node = set_node.add_child("Captures")
                    for capture_name in capture_names:
                        capture_node.add_child(capture_name)
                        self.count += 1
                if view_names:
                    view_node = set_node.add_child("Vues")
                    for view_name in view_names:
                        view_node.add_child(view_name)
                        self.count += 1
                if reference_names:
                    ref_node = set_node.add_child("Références")
                    for reference_name in reference_names:
                        ref_node.add_child(reference_name)
                        self.count += 1
                for annotation_name, folder in annotation_items:
                    set_node.add_child(folder).add_child(annotation_name)
                    self.count += 1

        if publications:
            publication_node = tree.add_child("Publication")
            for publication_name in dict.fromkeys(publications):
                publication_node.add_child(publication_name)
                self.count += 1

    def _add_annotation_results_from_selection(self, tree, root_name):
        """Reconstruit les groupes d'annotation en se basant principalement sur les chemins remontés par Selection.

        Cette version détecte les dossiers visibles (Captures, Vues, Références, Cadres de tolérances,
        Tolérance géométrique, Notes) en regardant les labels des parents plutôt que de se fier uniquement
        aux types COM retournés par l'API.
        """
        groups = {}
        selection = self.document.Selection
        try:
            selection.Clear()
            selection.Search("Name=*,all")
            for index in range(1, self._selection_count(selection) + 1):
                try:
                    item = self._selection_item(selection, index)
                    # Construire le chemin minimal parent tel que vu dans l'arbre
                    path = self._minimal_parent_path(item, root_name)
                    if not path:
                        continue

                    # Normaliser les labels du chemin pour détection
                    norm_path = [self._normalize_string(p) for p in path]

                    # Déterminer le conteneur (AnnotationSet) et le type (capture/view/etc.)
                    container_name = self._annotation_set_name(item, root_name)
                    if container_name is None:
                        container_name = "Résultat d'un ensemble d'annotations"

                    item_name = self._name_of(item)

                    # Heuristiques : chercher des dossiers visibles dans le chemin
                    kind = None
                    if any("capture" in p for p in norm_path):
                        kind = "capture"
                    elif any(p.startswith("vue") or "view" in p for p in norm_path):
                        kind = "view"
                    elif any("référence" in p or "reference" in p for p in norm_path):
                        kind = "reference"
                    elif any("cadre" in p or "cadres" in p for p in norm_path):
                        kind = "cadre"
                    elif any("tolérance" in p or "tolerance" in p or "tolérance géométrique" in p for p in norm_path):
                        kind = "tolerance"
                    elif any("note" in p for p in norm_path):
                        kind = "note"
                    else:
                        # fallback sur la detection par type COM
                        kind = self._item_kind(item)

                    if kind is None:
                        continue

                    group = groups.setdefault(container_name, {})
                    group.setdefault(kind, set()).add(item_name)
                except Exception:
                    continue
        except Exception:
            return
        finally:
            try:
                selection.Clear()
            except Exception:
                pass

        # Transcrire les groupes en chemins dans l'arbre
        for container_name, kinds in groups.items():
            # Publications sont traitées à part
            if container_name == "Publication":
                for pub_name in sorted(kinds.get("publication", [])):
                    tree.add_path([container_name, pub_name])
                    self.count += 1
                continue

            # Pour chaque type détecté, choisir le libellé de dossier visible
            for kind, names in kinds.items():
                if kind == "capture":
                    folder_label = "Captures"
                elif kind == "view":
                    folder_label = "Vues"
                elif kind == "reference":
                    folder_label = "Références"
                elif kind == "cadre":
                    folder_label = "Cadres de tolérances"
                elif kind == "tolerance":
                    folder_label = "Tolérance géométrique"
                elif kind == "note":
                    folder_label = "Notes"
                elif kind == "publication":
                    folder_label = "Publication"
                else:
                    folder_label = None

                if folder_label is None:
                    # Cas inattendu : créer un noeud direct sous le container
                    for item_name in sorted(names):
                        tree.add_path([container_name, item_name])
                        self.count += 1
                    continue

                for item_name in sorted(names):
                    tree.add_path([container_name, folder_label, item_name])
                    self.count += 1

    def _annotation_set_name(self, item, root_name):
        current = item
        for _ in range(100):
            try:
                parent = current.Parent
            except Exception:
                return "Résultat d'un ensemble d'annotations"
            if parent is None:
                return "Résultat d'un ensemble d'annotations"

            parent_name = self._name_of(parent)
            if self._is_annotation_set_label(parent_name):
                return parent_name

            current = parent
            if self._name_of(current) == root_name:
                break

        return "Résultat d'un ensemble d'annotations"

    def _build_tree_from_selection(self, tree, root_name):
        selection = self.document.Selection
        debug_records = []
        try:
            selection.Clear()
            selection.Search("Name=*,all")
            for index in range(1, self._selection_count(selection) + 1):
                try:
                    item = self._selection_item(selection, index)
                    path = self._minimal_parent_path(item, root_name)
                    if not path:
                        continue

                    kind = self._item_kind(item)
                    folder_label = self._folder_label_for_kind(kind)
                    if folder_label:
                        normalized_path = [self._normalize_string(label) for label in path]
                        if self._has_annotation_set_in_path(path):
                            if self._normalize_string(folder_label) not in normalized_path:
                                insert_at = self._annotation_set_index(path)
                                if insert_at is not None:
                                    path.insert(insert_at + 1, folder_label)
                        else:
                            if normalized_path and self._normalize_string(path[0]) != self._normalize_string(folder_label):
                                # Les résultats peuvent être fournis sans le dossier visible parent.
                                # Ajouter le libellé de dossier pour retrouver correctement
                                # les sous-arbres par nom cible.
                                path.insert(0, folder_label)
                            # Si le chemin n'indique pas de set d'annotations parent, ajouter
                            # explicitement le parent visible englobant.
                            if not self._has_annotation_set_in_path(path):
                                path.insert(0, self._annotation_result_root_label())

                    tree.add_path(path)

                    if getattr(self, 'debug', False):
                        debug_records.append({
                            'index': index,
                            'name': self._name_of(item),
                            'type': self._type_name(item),
                            'path': path,
                            'normalized_path': [self._normalize_string(p) for p in path],
                        })
                except Exception:
                    continue
        except Exception:
            return
        finally:
            try:
                selection.Clear()
            except Exception:
                pass
 
        if getattr(self, 'debug', False):
            try:
                debug_path = self.output_path.parent / f"{self.output_path.stem}_selection_debug.json"
                debug_path.parent.mkdir(parents=True, exist_ok=True)
                debug_path.write_text(json.dumps(debug_records, ensure_ascii=False, indent=2), encoding='utf-8')
                print(f"Debug selection written: {debug_path.resolve()}")
            except Exception as e:
                print(f"Failed writing debug file: {e}")

    def _add_annotation_result_root(self, tree, root_name):
        """Ajoute explicitement le grand-parent visible pour les résultats d'annotations.

        Certains objets CATIA sont retournés par Selection.Search sans chemin contenant
        'Résultat d'un ensemble d'annotations'. Ce parent doit être forcé pour que
        la recherche par nom comme 'Vues' ou 'Captures' retourne bien tous ses enfants.
        """
        annotation_root = self._annotation_result_root_label()
        added_paths = set()
        selection = self.document.Selection
        try:
            selection.Clear()
            selection.Search("Name=*,all")
            for index in range(1, self._selection_count(selection) + 1):
                try:
                    item = self._selection_item(selection, index)
                    kind = self._item_kind(item)
                    if not self._is_annotation_related_kind(kind):
                        continue

                    folder_label = self._folder_label_for_kind(kind)
                    if folder_label is None:
                        continue

                    item_name = self._name_of(item)
                    path = [annotation_root, folder_label, item_name]
                    path_key = tuple(path)
                    if path_key in added_paths:
                        continue

                    added_paths.add(path_key)
                    tree.add_path(path)
                    self.count += 1
                except Exception:
                    continue
        except Exception:
            return
        finally:
            try:
                selection.Clear()
            except Exception:
                pass

    def inspect_index(self, index):
        """DEBUG: inspecte l'objet Selection.Item(index) et écrit ses attributs/méthodes accessibles en JSON.

        Cette méthode tente d'accéder prudemment aux attributs courants et d'appeler
        les méthodes sans arguments pour repérer où se trouve le libellé visible.
        """
        selection = self.document.Selection
        try:
            selection.Clear()
            selection.Search("Name=*,all")
            try:
                item = self._selection_item(selection, index)
            except Exception as e:
                print(f"Impossible d'obtenir l'élément de selection à l'index {index}: {e}")
                return

            obj = item
            info = {
                'index': index,
                'type': self._type_name(obj),
                'name': self._name_of(obj),
                'attributes': {},
                'callables': {},
                'probe': {}
            }

            # Tentative prudente d'inspecter des attributs utiles
            probe_props = ('Name', 'Caption', 'Title', 'Label', 'UserName', 'DisplayedName', 'Text', 'GetCaption', 'GetName')
            for prop in probe_props:
                try:
                    val = getattr(obj, prop)
                    if callable(val):
                        try:
                            val = val()
                        except Exception as e:
                            val = f"<call-error: {type(e).__name__}: {e}>"
                    info['probe'][prop] = str(val) if val is not None else None
                except Exception:
                    info['probe'][prop] = None

            # Parent info
            try:
                parent = getattr(obj, 'Parent')
                info['probe']['ParentName'] = self._name_of(parent)
                info['probe']['ParentType'] = self._type_name(parent)
            except Exception:
                info['probe']['ParentName'] = None

            # dir-based inspection (liste d'attributs/méthodes visibles)
            names = []
            try:
                names = [n for n in dir(obj) if not n.startswith('_')]
            except Exception:
                names = []

            for n in names:
                try:
                    attr = getattr(obj, n)
                    if callable(attr):
                        try:
                            # appeler uniquement si la méthode a l'air sans effet (heuristique)
                            val = attr()
                            info['callables'][n] = repr(val)
                        except Exception as e:
                            info['callables'][n] = f"<call-error: {type(e).__name__}: {e}>"
                    else:
                        try:
                            info['attributes'][n] = repr(attr)
                        except Exception as e:
                            info['attributes'][n] = f"<attr-error: {type(e).__name__}: {e}>"
                except Exception as e:
                    info['attributes'][n] = f"<get-error: {type(e).__name__}: {e}>"

            out_path = self.output_path.parent / f"{self.output_path.stem}_inspect_{index}.json"
            out_path.parent.mkdir(parents=True, exist_ok=True)
            out_path.write_text(json.dumps(info, ensure_ascii=False, indent=2), encoding='utf-8')
            print(f"Inspect written: {out_path.resolve()}")

        except Exception as e:
            print(f"Inspection error: {e}")
        finally:
            try:
                selection.Clear()
            except Exception:
                pass

    def _collect_paths_for_target(self, target_name, root_name, match_mode=None):
        """Parcourt Selection et retourne tous les chemins complets (listes de labels)
        dont un segment correspond à target_name selon match_mode (normalisé).
        """
        if match_mode is None:
            match_mode = self.match_mode if hasattr(self, 'match_mode') else 'best'
        target_norm = self._normalize_string(target_name)

        def path_matches(norm_path):
            if match_mode == 'best':
                # exact OR contains
                return any(p == target_norm for p in norm_path) or any(target_norm in p for p in norm_path)
            if match_mode == 'exact':
                return any(p == target_norm for p in norm_path)
            if match_mode == 'contains':
                return any(target_norm in p for p in norm_path)
            if match_mode == 'startswith':
                return any(p.startswith(target_norm) for p in norm_path)
            return False
 
        selection = self.document.Selection
        try:
            selection.Clear()
            selection.Search("Name=*,all")
            for index in range(1, self._selection_count(selection) + 1):
                try:
                    item = self._selection_item(selection, index)
                    path = self._minimal_parent_path(item, root_name)
                    if not path:
                        continue
                    norm_path = [self._normalize_string(p) for p in path]
                    if path_matches(norm_path):
                        yield path
                        continue
 
                    # Si la cible est un dossier d'annotations comme Vues, Captures, Notes, etc.,
                    # certaines plateformes CATIA retournent les enfants sans chemin parent.
                    target_kinds = self._target_kinds_for_name(target_name)
                    item_kind = self._item_kind(item)
                    if target_kinds and item_kind in target_kinds:
                        folder_label = self._folder_label_for_kind(item_kind) or target_name
                        if self._normalize_string(folder_label) != target_norm:
                            folder_label = target_name
                        synthetic_path = [folder_label, self._name_of(item)]
                        if item_kind in {"capture", "view", "reference", "tolerance", "note"}:
                            synthetic_path.insert(0, self._annotation_result_root_label())
                        yield synthetic_path
                except Exception:
                    continue
        except Exception:
            return
        finally:
            try:
                selection.Clear()
            except Exception:
                pass

    def _item_kind(self, item):
        item_type = self._type_name(item).lower()
        name = self._normalize_string(self._name_of(item))
 
        if "annotationset" in item_type or "ensemble d'annotations" in name:
            return "annotation_set"
        if "capture" in item_type or item_type.endswith(".capture") or "capture" in name:
            return "capture"
        if "view" in item_type or item_type.endswith(".view") or "vue" in name:
            return "view"
        if "annotationplane" in item_type or "annotationplane" in name or "reference" in name or "referenceframe" in item_type:
            return "reference"
        if "tolerancezone" in item_type or "tolerance" in item_type or "tolerance" in name or "cadre" in name:
            return "tolerance"
        if "note" in item_type or "note" in name or "noa" in item_type or "flagnote" in name:
            return "note"
        if "annotation" in item_type or item_type.endswith(".annotation") or "annotation" in name:
            return "annotation"
        if "publication" in item_type or "publication" in name:
            return "publication"
        return None

    @staticmethod
    def _folder_label_for_kind(kind):
        return {
            "capture": "Captures",
            "view": "Vues",
            "reference": "Références",
            "tolerance": "Cadres de tolérances",
            "note": "Notes",
            "annotation": "Annotations",
            "publication": "Publication",
        }.get(kind)
 
    def _target_kinds_for_name(self, target_name):
        if not target_name:
            return None
        normalized_target = self._normalize_string(target_name)
        if any(term in normalized_target for term in ("capture", "captures")):
            return {"capture"}
        if any(term in normalized_target for term in ("vue", "views", "vues")):
            return {"view"}
        if any(term in normalized_target for term in ("référence", "reference", "références", "references")):
            return {"reference"}
        if any(term in normalized_target for term in ("cadre", "cadres", "tolérance", "tolerance")):
            return {"tolerance"}
        if any(term in normalized_target for term in ("note", "notes", "noa", "flagnote")):
            return {"note"}
        if "publication" in normalized_target:
            return {"publication"}
        return None
 
    @staticmethod
    def _annotation_result_root_label():
        return "Résultat d'un ensemble d'annotations"
 
    def _is_annotation_related_kind(self, kind):
        return kind in {"capture", "view", "reference", "tolerance", "note", "annotation", "publication"}
 
    def _has_annotation_set_in_path(self, path):
        return any(self._is_annotation_set_label(label) for label in path)

    def _annotation_set_index(self, path):
        for index, label in enumerate(path):
            if self._is_annotation_set_label(label):
                return index
        return None

    def _is_annotation_set_label(self, label):
        normalized = self._normalize_string(label)
        return "ensemble d'annotations" in normalized or "annotationset" in normalized or "resultat" in normalized

    def _add_all_named_objects_without_cycles(self, tree, root_name):
        """Ajoute le parcours global CATIA en supprimant les boucles Parent.

        Cette passe est necessaire pour les objets visibles dans certaines
        licences/workbenches mais absents des collections Part. Un nom deja vu
        dans le chemin coupe immediatement la boucle Capture -> Capture.
        """
        selection = self.document.Selection
        try:
            selection.Clear()
            selection.Search("Name=*,all")
            for index in range(1, self._selection_count(selection) + 1):
                try:
                    path = self._minimal_parent_path(
                        self._selection_item(selection, index), root_name
                    )
                    if path:
                        tree.add_path(path)
                except Exception:
                    continue
        except Exception:
            return
        finally:
            try:
                selection.Clear()
            except Exception:
                pass

    def _minimal_parent_path(self, item, root_name):
        """Retourne un seul chemin, sans parent repete ni cycle COM."""
        path = []
        seen_labels = set()
        current = item

        for _ in range(100):
            label = self._name_of(current)
            if label == "<sans nom>" or label == root_name or label in seen_labels:
                break
            path.append(label)
            seen_labels.add(label)

            try:
                parent = current.Parent
            except Exception:
                break
            current = parent

        path.reverse()
        return path

    def _normalize_string(self, text):
        if text is None:
            return ""
        normalized = unicodedata.normalize("NFKC", str(text))
        normalized = normalized.replace("\u00A0", " ")
        normalized = re.sub(r"\s+", " ", normalized).strip().lower()
        return normalized

    def _find_subtrees(self, node, target_name, match_mode='best'):
        """Parcourt l'arbre d'_TreeNode et renvoie les noeuds dont le label correspond
        selon match_mode : 'best' (exact d'abord, sinon contains), 'exact', 'contains', 'startswith'.
        """
        target_norm = self._normalize_string(target_name)
        node_norm = self._normalize_string(node.label)

        def matches(node_norm, target_norm):
            if match_mode == 'best':
                if node_norm == target_norm:
                    return True
                return target_norm in node_norm
            if match_mode == 'exact':
                return node_norm == target_norm
            if match_mode == 'contains':
                return target_norm in node_norm
            if match_mode == 'startswith':
                return node_norm.startswith(target_norm)
            return False

        if matches(node_norm, target_norm):
            yield node

        for child in node.children.values():
            yield from self._find_subtrees(child, target_name, match_mode=match_mode)

    @staticmethod
    def _annotation_folder(annotation):
        tests = (
            ("IsAReferenceFrame", "Cadres de tolérances"),
            ("IsAToleranceZone", "Tolérance géométrique"),
            ("IsANoa", "Notes"),
            ("IsAFlagNote", "Notes"),
            ("IsADatum", "Références"),
        )
        for method_name, folder in tests:
            try:
                if getattr(annotation, method_name)():
                    return folder
            except Exception:
                continue
        return "Annotations"

    def _add_collection(self, owner, property_name, parent_node, folder_name=None, on_item=None):
        collection = self._get_collection(owner, property_name)
        if collection is None:
            return
        items = list(self._items(collection))
        if not items:
            return
        node = parent_node.add_child(folder_name) if folder_name else parent_node
        for item in items:
            item_node = node.add_child(self._name_of(item))
            self.count += 1
            if on_item is not None:
                on_item(item, item_node)

    @staticmethod
    def _get_collection(owner, property_name):
        try:
            return getattr(owner, property_name)
        except Exception:
            return None

    @staticmethod
    def _selection_count(selection):
        try:
            return selection.Count2
        except Exception:
            return selection.Count

    @staticmethod
    def _selection_item(selection, index):
        try:
            return selection.Item2(index).Value
        except Exception:
            return selection.Item(index).Value

    @staticmethod
    def _items(collection):
        try:
            count = collection.Count
        except Exception:
            return
        for index in range(1, count + 1):
            try:
                yield collection.Item(index)
            except Exception:
                continue

    def _name_of(self, item):
        """Try multiple ways to get the user-visible label for an object.

        CATIA COM objects sometimes expose an internal technical Name (e.g. CATIAToleranceZone0)
        while the user-visible label is stored in other attributes or on a parent object.
        This helper tries Name first, then several common alternatives, then the parent
        name, and finally falls back to a sensible representation.
        """
        candidate = None
        try:
            # Preferred: Name (most objects have it)
            name = getattr(item, 'Name', None)
            if name:
                candidate = str(name)
                # If the Name looks like a technical id (starts with CATIA or CATIA...),
                # don't return it immediately — try other attributes first.
                if not candidate.lower().startswith('catia'):
                    return candidate
        except Exception:
            candidate = None

        # Try common alternative attributes/methods that may hold the visible label
        alt_attrs = ('Caption', 'Title', 'Label', 'UserName', 'DisplayedName', 'Text', 'GetCaption', 'GetName')
        for attr in alt_attrs:
            try:
                if not hasattr(item, attr):
                    continue
                val = getattr(item, attr)
                if callable(val):
                    try:
                        val = val()
                    except Exception:
                        continue
                if val:
                    return str(val)
            except Exception:
                continue

        # As a last resort, try the parent object's Name (sometimes the visible label is there)
        # If the object name is technical, allow a manual label mapping file to override it.
        if candidate:
            mapped = self._lookup_label_map(candidate)
            if mapped:
                return mapped

        # As a last resort, try the parent object's Name (sometimes the visible label is there)
        try:
            parent = getattr(item, 'Parent', None)
            if parent is not None:
                p_name = getattr(parent, 'Name', None)
                if p_name:
                    return str(p_name)
        except Exception:
            pass

        if candidate:
            return candidate

        try:
            return str(item.__class__.__name__)
        except Exception:
            return "<sans nom>"

    @staticmethod
    def _type_name(item):
        try:
            info = item._oleobj_.GetTypeInfo()
            return str(info.GetDocumentation(-1)[0])
        except Exception:
            return item.__class__.__name__
