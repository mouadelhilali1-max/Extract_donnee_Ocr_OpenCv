import json
from pathlib import Path


class CATIALabelInspector:
    """Inspecte les objets CATIA pour découvrir les labels affichés dans l'arbre.

    Cette classe est utile lorsque certains objets retournés par Selection.Search
    n'exposent que des noms techniques (par exemple CATIAToleranceZone0) et que
    l'on souhaite trouver la propriété réelle utilisée par CATIA pour l'affichage.
    """

    def __init__(self, document, output_dir="output"):
        self.document = document
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def inspect_tolerance_zones(self):
        selection = self.document.Selection
        try:
            selection.Clear()
            selection.Search("Name=*,all")
        except Exception as e:
            print("Impossible d'exécuter Selection.Search:", e)
            return

        records = []
        for index in range(1, self._selection_count(selection) + 1):
            try:
                item = self._selection_item(selection, index)
            except Exception:
                continue

            if not self._is_tolerance_zone(item):
                continue

            record = {
                "index": index,
                "type": self._type_name(item),
                "name": self._safe_get_name(item),
                "display_names": self._inspect_display_name_candidates(item),
                "attributes": self._inspect_attributes(item),
                "methods": self._inspect_methods(item),
            }
            records.append(record)

        output_path = self.output_dir / "tolerance_zone_inspection.json"
        try:
            output_path.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"Tolerance zone inspection written: {output_path.resolve()}")
        except Exception as e:
            print(f"Impossible d'écrire le fichier d'inspection : {e}")

    def _inspect_display_name_candidates(self, item):
        candidates = {}
        for attr in ("Name", "Caption", "Label", "Title", "UserName", "DisplayedName", "Text", "GetCaption", "GetName", "Description", "GetDescription", "ShortName"):
            try:
                value = getattr(item, attr)
            except Exception:
                continue
            try:
                if callable(value):
                    value = value()
            except Exception:
                continue
            if value is not None:
                candidates[attr] = self._safe_value(value)
        return candidates

    def _inspect_attributes(self, item):
        attributes = {}
        names = []
        try:
            names = [name for name in dir(item) if not name.startswith("_")]
        except Exception:
            return attributes

        for name in sorted(names):
            if name in ("Name", "Caption", "Label", "Title", "UserName", "DisplayedName", "Text"):
                continue
            try:
                value = getattr(item, name)
            except Exception as error:
                attributes[name] = f"<get-error: {error.__class__.__name__}: {error}>"
                continue

            if self._is_collection(value):
                attributes[name] = self._inspect_collection(value)
                continue

            if callable(value):
                attributes[name] = "<callable>"
            else:
                attributes[name] = self._safe_value(value)
        return attributes

    def _inspect_methods(self, item):
        methods = {}
        try:
            names = [name for name in dir(item) if not name.startswith("_")]
        except Exception:
            return methods

        for name in sorted(names):
            try:
                attr = getattr(item, name)
            except Exception:
                continue
            if not callable(attr):
                continue
            if name.lower() not in ("getcaption", "getname", "getdescription", "getfullname", "gettext"):
                continue
            try:
                value = attr()
                methods[name] = self._safe_value(value)
            except Exception as error:
                methods[name] = f"<call-error: {error.__class__.__name__}: {error}>"
        return methods

    @staticmethod
    def _inspect_collection(collection):
        try:
            count = collection.Count
        except Exception:
            return "<collection-unreadable>"

        result = {"count": count}
        if count > 0:
            items = []
            for idx in range(1, min(count, 20) + 1):
                try:
                    item = collection.Item(idx)
                    items.append(CATIALabelInspector._safe_name(item))
                except Exception:
                    items.append("<item-error>")
            result["items"] = items
        return result

    @staticmethod
    def _is_collection(value):
        if value is None:
            return False
        try:
            count = getattr(value, "Count", None)
            return isinstance(count, int)
        except Exception:
            return False

    @staticmethod
    def _type_name(item):
        try:
            info = item._oleobj_.GetTypeInfo()
            return str(info.GetDocumentation(-1)[0])
        except Exception:
            try:
                return item.__class__.__name__
            except Exception:
                return "<unknown>"

    @staticmethod
    def _safe_get_name(item):
        try:
            name = getattr(item, "Name", None)
            if name:
                return str(name)
        except Exception:
            pass
        return ""

    @staticmethod
    def _safe_value(value):
        if value is None:
            return ""
        try:
            if isinstance(value, str):
                return value
            return str(value)
        except Exception:
            return "<unrepresentable>"

    @staticmethod
    def _safe_name(item):
        try:
            name = getattr(item, "Name", None)
            if name:
                return str(name)
        except Exception:
            pass
        try:
            return str(item)
        except Exception:
            return "<unknown>"

    @staticmethod
    def _selection_count(selection):
        try:
            return selection.Count2
        except Exception:
            try:
                return selection.Count
            except Exception:
                return 0

    @staticmethod
    def _selection_item(selection, index):
        try:
            return selection.Item2(index).Value
        except Exception:
            return selection.Item(index).Value

    @staticmethod
    def _is_tolerance_zone(item):
        try:
            type_name = item._oleobj_.GetTypeInfo().GetDocumentation(-1)[0]
        except Exception:
            try:
                type_name = item.__class__.__name__
            except Exception:
                return False
        type_name = str(type_name).lower()
        if "tolerancezone" in type_name or "tolerance" in type_name:
            return True
        try:
            name = getattr(item, "Name", "")
            if name and "catia tolerance zone" in str(name).lower():
                return True
        except Exception:
            pass
        return False
