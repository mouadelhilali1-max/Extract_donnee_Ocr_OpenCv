import win32com.client


class TreeReader:

    def __init__(self):
        pass

    # --------------------------------------------------
    # Affichage
    # --------------------------------------------------

    def print_node(self, name, level=0, symbol="|--"):
        print("    " * level + symbol + " " + str(name))

    # --------------------------------------------------
    # HybridBody
    # --------------------------------------------------

    def read_hybrid_body(self, hb, level=0):

        self.print_node(hb.Name, level)

        # ------------------------
        # Shapes
        # ------------------------

        try:
            shapes = hb.HybridShapes

            for i in range(1, shapes.Count + 1):
                shape = shapes.Item(i)
                self.print_node(shape.Name, level + 1, "+")

        except:
            pass

        # ------------------------
        # Sous HybridBodies
        # ------------------------

        try:
            sub = hb.HybridBodies

            for i in range(1, sub.Count + 1):
                self.read_hybrid_body(sub.Item(i), level + 1)

        except:
            pass

    # --------------------------------------------------
    # Bodies
    # --------------------------------------------------

    def read_bodies(self, part):

        try:

            print("\n===== BODIES =====\n")

            bodies = part.Bodies

            for i in range(1, bodies.Count + 1):

                body = bodies.Item(i)

                self.print_node(body.Name)

        except:
            pass

    # --------------------------------------------------
    # HybridBodies
    # --------------------------------------------------

    def read_hybrid_bodies(self, part):

        try:

            print("\n===== HYBRID BODIES =====\n")

            hbs = part.HybridBodies

            for i in range(1, hbs.Count + 1):

                hb = hbs.Item(i)

                self.read_hybrid_body(hb)

        except:
            pass

    # --------------------------------------------------
    # Publications
    # --------------------------------------------------

    def read_publications(self, part):

        try:

            print("\n===== PUBLICATIONS =====\n")

            pubs = part.Publications

            for i in range(1, pubs.Count + 1):

                pub = pubs.Item(i)

                self.print_node(pub.Name)

        except:
            pass

    # --------------------------------------------------
    # Parameters
    # --------------------------------------------------

    def read_parameters(self, part):

        try:

            print("\n===== PARAMETERS =====")

            print(part.Parameters.Count)

        except:
            pass

    # --------------------------------------------------
    # Relations
    # --------------------------------------------------

    def read_relations(self, part):

        try:

            print("\n===== RELATIONS =====")

            print(part.Relations.Count)

        except:
            pass

    # --------------------------------------------------
    # Point d'entrée
    # --------------------------------------------------

    def read(self, part):

        print("=" * 60)
        print("Nom de la pièce :", part.Name)
        print("=" * 60)

        self.read_bodies(part)

        self.read_hybrid_bodies(part)

        self.read_publications(part)

        self.read_parameters(part)

        self.read_relations(part)