import cv2
import pytesseract
from paddleocr import PaddleOCR

# -----------------------------
# Configuration
# -----------------------------

MIN_SCORE = 0.85

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

ocr = PaddleOCR(
    use_angle_cls=True,
    lang="en",
    use_gpu=False,
    show_log=False,
    enable_mkldnn=False,
    cpu_threads=1
)


# -----------------------------
# Prétraitement
# -----------------------------

def preprocess(img):

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    gray = cv2.resize(
        gray,
        None,
        fx=2,
        fy=2,
        interpolation=cv2.INTER_CUBIC
    )

    gray = cv2.GaussianBlur(gray, (3, 3), 0)

    _, gray = cv2.threshold(
        gray,
        0,
        255,
        cv2.THRESH_BINARY + cv2.THRESH_OTSU
    )

    return gray


# -----------------------------
# PaddleOCR
# -----------------------------

def paddle_read(img):

    results = []

    out = ocr.ocr(img, cls=True)

    if out is None or out[0] is None:
        return results

    for line in out[0]:

        box = line[0]

        text = line[1][0].strip()

        score = float(line[1][1])

        if text == "":
            continue

        if score < MIN_SCORE:
            continue

        results.append({
            "text": text,
            "score": round(score,3),
            "x": int(box[0][0]),
            "y": int(box[0][1]),
            "engine": "Paddle"
        })

    return results


# -----------------------------
# Tesseract
# -----------------------------

def tesseract_read(img):

    data = pytesseract.image_to_data(
        img,
        output_type=pytesseract.Output.DICT,
        config="--oem 3 --psm 6"
    )

    groups = {}

    n = len(data["text"])

    for i in range(n):

        txt = data["text"][i].strip()

        if txt == "":
            continue

        try:
            conf = float(data["conf"][i])/100
        except:
            conf = 0

        if conf < MIN_SCORE:
            continue

        key = (
            data["block_num"][i],
            data["par_num"][i],
            data["line_num"][i]
        )

        if key not in groups:

            groups[key] = {
                "words": [],
                "scores": [],
                "x": data["left"][i],
                "y": data["top"][i]
            }

        groups[key]["words"].append(txt)
        groups[key]["scores"].append(conf)

    results = []

    for g in groups.values():

        score = sum(g["scores"]) / len(g["scores"])

        if score < MIN_SCORE:
            continue

        results.append({
            "text": " ".join(g["words"]),
            "score": round(score,3),
            "x": int(g["x"]),
            "y": int(g["y"]),
            "engine": "Tesseract"
        })

    return results


# -----------------------------
# Fusion Paddle + Tesseract
# -----------------------------

def merge_results(results):

    results.sort(key=lambda r: (r["y"], r["x"]))

    merged = []

    for r in results:

        duplicate = False

        for m in merged:

            if abs(r["y"]-m["y"]) < 10 and abs(r["x"]-m["x"]) < 40:

                duplicate = True

                if r["score"] > m["score"]:
                    m.update(r)

                break

            if r["text"] == m["text"]:

                duplicate = True

                if r["score"] > m["score"]:
                    m.update(r)

                break

        if not duplicate:
            merged.append(r)

    return merged


# -----------------------------
# Lecture complète
# -----------------------------

def read_tree(image_path):

    img = cv2.imread(image_path)

    if img is None:
        return []

    img = preprocess(img)

    results = []

    try:
        results.extend(paddle_read(img))
    except Exception as e:
        print("Paddle :", e)

    try:
        results.extend(tesseract_read(img))
    except Exception as e:
        print("Tesseract :", e)

    results = merge_results(results)

    results.sort(key=lambda r: (r["y"], r["x"]))

    return results