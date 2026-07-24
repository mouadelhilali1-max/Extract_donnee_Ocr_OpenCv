import os
import json
import pandas as pd

from ocr.ocr_reader import read_tree

CAPTURE_FOLDER = "captures"

OUTPUT_JSON = "output/all_scans.json"
OUTPUT_EXCEL = "output/all_scans.xlsx"


def process_all_scans():

    os.makedirs("output", exist_ok=True)

    all_lines = []

    images = sorted(os.listdir(CAPTURE_FOLDER))

    capture = 0

    for image in images:

        if not image.lower().endswith((".png", ".jpg", ".jpeg")):
            continue

        print(image)

        path = os.path.join(CAPTURE_FOLDER, image)

        lines = read_tree(path)

        lines = sorted(lines, key=lambda l: l["y"])

        order = 1

        for line in lines:

            all_lines.append({

                "Capture": capture,
                "Image": image,

                "Order": order,

                "Text": line["text"],

                "X": line["x"],
                "Y": line["y"],

                "Score": round(line["score"], 3)

            })

            order += 1

        capture += 1

    with open(OUTPUT_JSON, "w", encoding="utf8") as f:

        json.dump(
            all_lines,
            f,
            indent=4,
            ensure_ascii=False
        )

    df = pd.DataFrame(all_lines)

    df = df.sort_values(
        by=["Capture", "Order"]
    )

    df.to_excel(
        OUTPUT_EXCEL,
        index=False
    )

    print()
    print("Nombre de lignes :", len(df))
    print("JSON créé :", OUTPUT_JSON)
    print("Excel créé :", OUTPUT_EXCEL)


if __name__ == "__main__":
    process_all_scans()