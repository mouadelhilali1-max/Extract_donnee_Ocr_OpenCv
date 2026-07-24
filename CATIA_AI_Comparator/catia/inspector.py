class CATIAInspector:

    def __init__(self):
        pass

    def inspect_part(self, part):

        print("\n" + "=" * 80)
        print("PART")
        print("=" * 80)

        self.test_property(part, "Name")
        self.test_property(part, "Bodies")
        self.test_property(part, "HybridBodies")
        self.test_property(part, "Parameters")
        self.test_property(part, "Relations")
        self.test_property(part, "OriginElements")
        self.test_property(part, "ShapeFactory")
        self.test_property(part, "HybridShapeFactory")

    def inspect_hybrid_body(self, hb):

        print("\n" + "=" * 80)
        print(f"HYBRID BODY : {hb.Name}")
        print("=" * 80)

        properties = [

            "Name",
            "Parent",
            "HybridBodies",
            "HybridShapes",
            "Bodies",
            "GeometricElements",
            "Shapes",
            "Selection",
            "VisProperties"

        ]

        for p in properties:

            self.test_property(hb, p)

    def test_property(self, obj, prop):

        try:

            value = getattr(obj, prop)

            if hasattr(value, "Count"):

                print(f"{prop:<25} -> Collection ({value.Count})")

            else:

                print(f"{prop:<25} -> OK")

        except Exception as e:

            print(f"{prop:<25} -> NOT AVAILABLE")