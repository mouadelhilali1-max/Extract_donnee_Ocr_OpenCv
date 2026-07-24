from pywinauto import Desktop


class UITreeReader:

    def __init__(self):
        self.window = None

    def connect(self):

        windows = Desktop(backend="uia").windows()

        for w in windows:

            try:

                title = w.window_text()

                if "CATIA" in title:

                    self.window = w

                    print("CATIA trouvée :")
                    print(title)

                    return True

            except:
                pass

        return False