"""Parse 'Ведомость по остаткам ТМЦ на складах' (1C PDF, split column groups).

The PDF is printed column-by-column groups:
  pages 1,4,7,10  -> Склад | Группа | Артикул | Номенклатура | Ед | Начальный остаток
  pages 2,5,8,11  -> Приход | Расход | Конечный остаток | Мин. резерв | Окт25..Апр26
  pages 3,6,9,12  -> Май26..Окт26
Row y-positions continue across the group pages, so rows are joined by order.
"""
import json
import re
import sys
from collections import defaultdict

import pdfplumber

NUM_RE = re.compile(r"^-?\d+(?:[.,]\d+)?$")
MONTHS = ["2025-10", "2025-11", "2025-12", "2026-01", "2026-02", "2026-03",
          "2026-04", "2026-05", "2026-06", "2026-07", "2026-08", "2026-09", "2026-10"]

ITEM_PAGES = [0, 3, 6, 9]
NUM_PAGES_A = [1, 4, 7, 10]   # 11 columns
NUM_PAGES_B = [2, 5, 8, 11]   # 6 columns


def to_num(t):
    return float(t.replace(",", "."))


def group_lines(words, tol=5.0):
    words = sorted(words, key=lambda w: (w["top"], w["x0"]))
    lines = []
    for w in words:
        if lines and abs(w["top"] - lines[-1]["top"]) < tol:
            lines[-1]["words"].append(w)
            lines[-1]["top"] = min(lines[-1]["top"], w["top"])
        else:
            lines.append({"top": w["top"], "words": [w]})
    return lines


def parse_item_pages(pdf):
    rows = []
    for pg in ITEM_PAGES:
        page = pdf.pages[pg]
        words = page.extract_words()
        words = [w for w in words if w["top"] > 125 or pg > 0]
        lines = group_lines(words)
        blocks = []
        for ln in lines:
            is_start = any(w["text"] == "Склад" and w["x0"] < 50 for w in ln["words"])
            if is_start or not blocks:
                blocks.append([])
            blocks[-1].append(ln)
        for block in blocks:
            row = {"склад": "", "группа": "", "артикул": "", "номенклатура": "",
                   "ед": "", "начальный_остаток": None}
            for ln in block:
                for w in sorted(ln["words"], key=lambda w: w["x0"]):
                    x, t = w["x0"], w["text"]
                    if x < 130:
                        row["склад"] += " " + t
                    elif x < 205:
                        row["группа"] += " " + t
                    elif x < 263:
                        row["артикул"] += " " + t
                    elif x < 442:
                        row["номенклатура"] += " " + t
                    elif x < 470:
                        row["ед"] += " " + t
                    elif NUM_RE.match(t):
                        row["начальный_остаток"] = to_num(t)
            for k in ("склад", "группа", "артикул", "номенклатура", "ед"):
                row[k] = re.sub(r"\s+", " ", row[k]).strip()
            rows.append(row)
    return rows


def parse_numeric_page(page):
    words = page.extract_words()
    has_header = any(w["top"] < 126 and w["text"] in ("Ведомость", "Май") for w in words)
    words = [w for w in words if NUM_RE.match(w["text"]) and (not has_header or w["top"] > 125)]
    # cluster right edges
    edges = sorted({round(w["x1"], 0) for w in words})
    clusters = []
    for e in edges:
        if clusters and e - clusters[-1][-1] <= 3:
            clusters[-1].append(e)
        else:
            clusters.append([e])
    top_count = max(len(c) for c in clusters)
    centers = [sum(c) / len(c) for c in clusters if len(c) >= 0.6 * top_count]
    lines = defaultdict(dict)
    for w in words:
        col = min(range(len(centers)), key=lambda i: abs(centers[i] - w["x1"]))
        line = None
        for key in lines:
            if abs(key - w["top"]) < 5:
                line = key
                break
        if line is None:
            line = w["top"]
        lines[line][col] = to_num(w["text"])
    return [lines[k] for k in sorted(lines)], len(centers)


def main(path):
    with pdfplumber.open(path) as pdf:
        items = parse_item_pages(pdf)
        numA, numB = [], []
        for pg in NUM_PAGES_A:
            r, n = parse_numeric_page(pdf.pages[pg])
            assert n == 11, (pg, n)
            numA += r
        for pg in NUM_PAGES_B:
            r, n = parse_numeric_page(pdf.pages[pg])
            assert n == 6, (pg, n)
            numB += r
    assert len(items) == len(numA) == len(numB), (len(items), len(numA), len(numB))
    out = []
    for it, a, b in zip(items, numA, numB):
        row = dict(it)
        row["приход"] = a.get(0, 0.0)
        row["расход"] = a.get(1, 0.0)
        row["конечный_остаток"] = a.get(2, 0.0)
        row["мин_резерв"] = a.get(3, 0.0)
        months = [a.get(4 + i, 0.0) for i in range(7)] + [b.get(i, 0.0) for i in range(6)]
        row["продажи_по_месяцам"] = dict(zip(MONTHS, months))
        out.append(row)
    return out


if __name__ == "__main__":
    data = main(sys.argv[1])
    json.dump(data, open(sys.argv[2], "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("rows:", len(data))
