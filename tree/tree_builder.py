import pandas as pd


# Tolérance verticale (pixels)
TOLERANCE_Y = 10

# Tolérance horizontale pour les niveaux
TOLERANCE_X = 20


def charger_csv(fichier="output/ocr_resultats.csv"):

    df = pd.read_csv(fichier)

    return df


def fusionner_lignes(df):

    lignes = []

    # Traiter image par image
    for image in df["image"].unique():

        data = df[df["image"] == image].copy()

        # Trier par Y puis X
        data = data.sort_values(["y", "x"])

        groupes = []

        for _, row in data.iterrows():

            trouve = False

            for g in groupes:

                if abs(row["y"] - g["y"]) <= TOLERANCE_Y:

                    g["mots"].append(row)

                    trouve = True

                    break

            if not trouve:

                groupes.append({
                    "y": row["y"],
                    "mots": [row]
                })

        # Fusion des mots
        for g in groupes:

            mots = sorted(g["mots"], key=lambda r: r["x"])

            texte = " ".join([m["texte"] for m in mots])

            lignes.append({

                "image": image,

                "x": mots[0]["x"],

                "y": g["y"],

                "texte": texte

            })

    return pd.DataFrame(lignes)


def calculer_niveaux(df):

    niveaux = []

    positions = []

    for _, row in df.iterrows():

        x = row["x"]

        trouve = False

        for i, p in enumerate(positions):

            if abs(x - p) < TOLERANCE_X:

                niveaux.append(i)

                trouve = True

                break

        if not trouve:

            positions.append(x)

            positions.sort()

            niveau = positions.index(x)

            niveaux.append(niveau)

    df["niveau"] = niveaux

    return df


def sauvegarder(df):

    df = df.sort_values(["image", "y"])

    df.to_csv(
        "output/tree_reconstructed.csv",
        index=False,
        encoding="utf-8-sig"
    )

    print()

    print("===================================")
    print("Arbre reconstruit.")
    print("Fichier : output/tree_reconstructed.csv")
    print("===================================")


def build_tree():

    print("Lecture OCR...")

    df = charger_csv()

    print("Fusion des lignes...")

    lignes = fusionner_lignes(df)

    print("Calcul des niveaux...")

    lignes = calculer_niveaux(lignes)

    sauvegarder(lignes)

    return lignes


if __name__ == "__main__":

    arbre = build_tree()

    print()

    print(arbre.head(30))