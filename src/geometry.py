class Geometry:

    def __init__(self, shape):
        self.shape = shape

    def analyze(self):

        result = {
            "Plane": 0,
            "Cylinder": 0,
            "Cone": 0,
            "Sphere": 0,
            "Toroid": 0,
            "BSpline": 0,
            "Bezier": 0,
            "Other": 0
        }

        print("\n========== ANALYSE GEOMETRIQUE ==========\n")

        for i, face in enumerate(self.shape.Faces):

            surface = face.Surface
            surface_type = surface.__class__.__name__

            print(f"Face {i+1:3} --> {surface_type}")

            if surface_type in result:
                result[surface_type] += 1
            else:
                result["Other"] += 1

        return result