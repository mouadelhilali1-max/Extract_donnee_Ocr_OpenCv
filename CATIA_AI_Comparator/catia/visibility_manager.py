class VisibilityManager:

    def __init__(self, document):
        self.document = document
        self.selection = document.Selection

    def set_visibility(self, obj, visible=True):

        self.selection.Clear()

        self.selection.Add(obj)

        vis = self.selection.VisProperties

        if visible:
            vis.SetShow(0)
            print(f"[SHOW] {obj.Name}")
        else:
            vis.SetShow(1)
            print(f"[HIDE] {obj.Name}")

        self.selection.Clear()

    def find_hybrid_body(self, name):

        return self._find_recursive(
            self.document.Part.HybridBodies,
            name
        )

    def _find_recursive(self, collection, name):

        for i in range(1, collection.Count + 1):

            hb = collection.Item(i)

            if hb.Name == name:
                return hb

            try:
                result = self._find_recursive(
                    hb.HybridBodies,
                    name
                )

                if result is not None:
                    return result

            except:
                pass

        return None