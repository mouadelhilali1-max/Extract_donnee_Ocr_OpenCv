import win32com.client


class CATIAConnection:

    def __init__(self):
        self.catia = None
        self.document = None
        self.part = None

    def connect(self):

        try:
            self.catia = win32com.client.Dispatch("CATIA.Application")

            print("=" * 60)
            print("Connexion réussie")
            print("=" * 60)

            documents = self.catia.Documents
            try:
                count = documents.Count
            except Exception:
                count = 0

            if count == 0:
                print("Aucun document CATIA n'est ouvert. Ouvrez un CATPart dans CATIA avant de lancer le script.")
                return False

            try:
                self.document = self.catia.ActiveDocument
            except Exception:
                try:
                    self.document = documents.Item(1)
                    print("ActiveDocument inaccessible. Utilisation du premier document ouvert :", self.document.Name)
                except Exception as inner_e:
                    print("Impossible d'accéder à un document CATIA ouvert.")
                    print(type(inner_e))
                    print(inner_e)
                    return False

            print("Nom du document :", self.document.Name)

            # Essayer directement d'accéder à la Part
            try:
                self.part = self.document.Part
            except Exception as part_e:
                print("Impossible d'accéder à Part sur le document ouvert. Ce n'est peut-être pas un CATPart.")
                print(type(part_e))
                print(part_e)
                return False

            print("Nom de la pièce :", self.part.Name)
            return True

        except Exception as e:
            print("Erreur :")
            print(type(e))
            print(e)
            return False