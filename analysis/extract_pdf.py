"""Извлечение таблицы из PDF «Ведомость по остаткам».

PDF печатается тремя горизонтальными срезами на страницу (колонки 1-6, 7-17, 18-23),
поэтому страницы объединяются блоками по три, а ячейки собираются по координатам слов.
Контроль: начальный остаток + приход - расход = конечный остаток (все строки).

Запуск: pip install pymupdf && python3 analysis/extract_pdf.py <путь-к.pdf>
"""
import pymupdf, json, re, csv, os, sys, unicodedata

PDF = sys.argv[1] if len(sys.argv) > 1 else os.environ.get(
    "VEDOMOST_PDF", "/tmp/arena-coding-uploads/Ведомость мав.pdf")
MONTHS = ["Октябрь 2025","Ноябрь 2025","Декабрь 2025","Январь 2026","Февраль 2026",
          "Март 2026","Апрель 2026","Май 2026","Июнь 2026","Июль 2026","Август 2026",
          "Сентябрь 2026","Октябрь 2026"]
COLS_A = ["sklad","gruppa","artikul","nomenklatura","ed","nach_ostatok"]
COLS_B1 = ["prihod","rashod","konechny_ostatok","min_rezerv"] + MONTHS[:7]
COLS_B2 = MONTHS[7:]

def clean(s):
    if s is None: return ""
    s = unicodedata.normalize("NFC", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def cell_text(words, x0, x1, y0, y1):
    sel = [w for w in words
           if x0-0.6 <= (w[0]+w[2])/2 <= x1+0.6 and y0-0.6 <= (w[1]+w[3])/2 <= y1+0.6]
    sel.sort(key=lambda w: (round(w[1], 1), w[0]))
    lines, cur, ly = [], [], None
    for w in sel:
        if ly is None or abs(w[1]-ly) < 3.0:
            cur.append(w[4]); ly = w[1] if ly is None else ly
        else:
            lines.append(" ".join(cur)); cur = [w[4]]; ly = w[1]
    if cur: lines.append(" ".join(cur))
    return clean(" ".join(lines))

def num(s):
    s = clean(s).replace("\u00a0", "").replace(" ", "")
    if s in ("", "-", "—"): return None
    s = s.replace(",", ".")
    if re.fullmatch(r"-?\d+(\.\d+)?", s):
        return float(s)
    return None

doc = pymupdf.open(PDF)
BLOCKS = [(1,2,3),(4,5,6),(7,8,9),(10,11,12)]

page_meta = {}
for i, p in enumerate(doc):
    t = p.find_tables().tables[0]
    cols = []
    for c in t.rows[0].cells:
        if not cols or abs(c[0]-cols[-1][0]) > 0.5 or abs(c[2]-cols[-1][1]) > 0.5:
            cols.append((c[0], c[2]))
    page_meta[i] = {"table": t, "cols": cols,
                    "bands": [(r.bbox[1], r.bbox[3]) for r in t.rows],
                    "words": p.get_text("words")}

records = []
for blk in BLOCKS:
    pa, pb, pc = [x-1 for x in blk]
    mA, mB, mC = page_meta[pa], page_meta[pb], page_meta[pc]
    bands = mA["bands"]
    # строка заголовка есть только на первой странице документа
    header_row = None
    if cell_text(mA["words"], *mA["cols"][2], *bands[0]) == "Артикул":
        header_row = 0
    for ri, (y0, y1) in enumerate(bands):
        if ri == header_row:
            continue
        rec = {}
        for ci, (cx0, cx1) in enumerate(mA["cols"]):
            rec[COLS_A[ci]] = cell_text(mA["words"], cx0, cx1, y0, y1)
        for meta, names in ((mB, COLS_B1), (mC, COLS_B2)):
            for ci, (cx0, cx1) in enumerate(meta["cols"]):
                if ci < len(names):
                    rec[names[ci]] = cell_text(meta["words"], cx0, cx1, y0, y1)
        rec["page"] = pa+1
        rec["row"] = ri
        records.append(rec)

rows = []
for r in records:
    d = {k: r.get(k, "") for k in COLS_A}
    d["nach_ostatok"] = num(r.get("nach_ostatok"))
    for k in ["prihod","rashod","konechny_ostatok","min_rezerv"]:
        d[k] = num(r.get(k))
    for m in MONTHS:
        d[m] = num(r.get(m))
    d["_months_sum"] = sum((d[m] or 0) for m in MONTHS)
    d["_balance_ok"] = None
    if None not in (d["nach_ostatok"], d["prihod"], d["rashod"], d["konechny_ostatok"]):
        d["_balance_ok"] = abs((d["nach_ostatok"] + d["prihod"] - d["rashod"]) - d["konechny_ostatok"]) < 1e-6
    rows.append(d)

print("всего строк:", len(rows))
print("баланс нач+приход-расход=конечный выполняется:", sum(1 for x in rows if x["_balance_ok"]),
      "не выполняется:", sum(1 for x in rows if x["_balance_ok"] is False),
      "не проверено:", sum(1 for x in rows if x["_balance_ok"] is None))
print("sum(расход) =", sum(x["rashod"] or 0 for x in rows))
print("sum(месячные колонки) =", sum(x["_months_sum"] for x in rows))
print("sum(приход) =", sum(x["prihod"] or 0 for x in rows))
print("совпадений months_sum == расход:", sum(1 for x in rows if abs(x["_months_sum"]-(x["rashod"] or 0))<1e-9))
print("совпадений months_sum == приход:", sum(1 for x in rows if abs(x["_months_sum"]-(x["prihod"] or 0))<1e-9))
print("склады:", sorted({x['sklad'] for x in rows}))
print("группы:", sorted({x['gruppa'] for x in rows}))
print("ед.изм:", sorted({x['ed'] for x in rows}))
print("артикулы непустые:", sum(1 for x in rows if x['artikul'].strip()))

with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "vedomost_ostatki.csv"), "w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=list(rows[0].keys()), delimiter=";")
    w.writeheader(); w.writerows(rows)
json.dump(rows, open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "vedomost_ostatki.json"), "w", encoding="utf-8"),
          ensure_ascii=False, indent=1)
print("\nпервые 6 строк:")
for x in rows[:6]:
    print(f"  {x['nomenklatura'][:44]:46} нач={x['nach_ostatok']:6} прих={x['prihod']:7} расх={x['rashod']:7} кон={x['konechny_ostatok']:6} мин.рез={x['min_rezerv']:4} мес.сумма={x['_months_sum']:6}")
print("\nпоследние 4 строки:")
for x in rows[-4:]:
    print(f"  {x['nomenklatura'][:44]:46} нач={x['nach_ostatok']:6} прих={x['prihod']:7} расх={x['rashod']:7} кон={x['konechny_ostatok']:6} мин.рез={x['min_rezerv']:4} мес.сумма={x['_months_sum']:6}")
