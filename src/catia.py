import win32com.client


class Catia:

    def __init__(self):

        try:

            self.catia = win32com.client.GetActiveObject("CATIA.Application")
            print("Connexion au CATIA existant.")

        except:

            self.catia = win32com.client.Dispatch("CATIA.Application")
            self.catia.Visible = True

            print("Nouveau CATIA lancé.")