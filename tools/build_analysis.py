"""Build the sales analysis dataset from the parsed 1C stock statement.

Input : JSON produced by tools/parse_vedomost.py
Output: reports/vedomost_2025-10_2026-10.csv
        reports/summary.json (aggregates + item rows for the dashboard)

The statement contains quantities only (no prices), so the analysis works with
physical volume: packages (as in the source) and an estimated litre equivalent.
Litre estimates: explicit "N л" is used as is; "N кг" is converted with an
average paint density of 1.35 kg/l; two names without a size are handled with
an explicit note (see VOLUME_OVERRIDES).

Method notes
------------
* 12 complete months (Oct 2025 - Sep 2026) are used for ABC/XYZ and averages;
  October 2026 is a partial month (1-19 Oct) and is shown separately.
* ABC is built on litres over the 12 full months (80/95 % cumulative).
* XYZ uses the coefficient of variation of monthly litres inside the season
  (Apr-Sep 2026) with thresholds 0.5 / 1.0, because the annual series is
  dominated by seasonality and would put almost every SKU into Z.
"""
from __future__ import annotations

import csv
import json
import re
import statistics
import sys
from collections import Counter, OrderedDict, defaultdict
from pathlib import Path

MONTHS = ["2025-10", "2025-11", "2025-12", "2026-01", "2026-02", "2026-03",
          "2026-04", "2026-05", "2026-06", "2026-07", "2026-08", "2026-09",
          "2026-10"]
MONTH_DAYS = [31, 30, 31, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31]
MONTH_LABELS = ["Окт 25", "Ноя 25", "Дек 25", "Янв 26", "Фев 26", "Мар 26",
                "Апр 26", "Май 26", "Июн 26", "Июл 26", "Авг 26", "Сен 26", "Окт 26*"]
FULL_MONTHS = MONTHS[:12]          # 12 complete months
PARTIAL_MONTH = MONTHS[-1]         # 1-19 Oct 2026
DENSITY_KG_PER_L = 1.35
PERIOD_DAYS = 384                  # 2025-10-01 .. 2026-10-19
PARTIAL_DAYS = 19

VOLUME_OVERRIDES = {
    # name contains no unit; the trailing number is the package size
    "Грунтовка ВД-АК-01 FLAGMAН 5": 5.0,
    # "0,9 (0,8кг)" – the litre value lost the "л" suffix
    "Эмаль ПФ-115 вишневая 0,9 (0,8кг)": 0.9,
}

COLOR_MAP = {
    "белая": "Белая", "белая гл": "Белая", "белая мат": "Белая",
    "св.серая": "Светло-серая", "кр-коричневая": "Красно-коричневая",
    "оранжевая": "Оранжевая", "слон.кость": "Слоновая кость",
    "жел.-корич": "Жёлто-коричневая", "зол.-корич": "Золотисто-коричневая",
    "св.-корич": "Светло-коричневая",
}


def litres(name: str) -> float | None:
    if name in VOLUME_OVERRIDES:
        return VOLUME_OVERRIDES[name]
    m = re.search(r"(\d+(?:[.,]\d+)?)\s*(л|кг)\b", name)
    if not m:
        return None
    value = float(m.group(1).replace(",", "."))
    if m.group(2) == "кг":
        value = value / DENSITY_KG_PER_L
    return round(value, 3)


def product_group(name: str) -> str:
    for prefix, group in (
        ("Эмаль ПФ-115", "Эмаль ПФ-115"),
        ("Эмаль ПФ-266", "Эмаль ПФ-266"),
        ("Краска", "Краска"),
        ("Грунтовка", "Грунтовка"),
        ("Лак", "Лак"),
    ):
        if name.startswith(prefix):
            return group
    return "Прочее"


def colour(name: str, group: str) -> str:
    if group in ("Эмаль ПФ-115", "Эмаль ПФ-266"):
        code = group.replace("Эмаль ", "")
        rest = name.split(code, 1)[1].strip()
        m = re.search(r"\d+(?:[.,]\d+)?\s*(?:л|кг)\b", rest)
        raw = rest[:m.start()] if m else rest
        raw = re.sub(r"\s*\d+(?:[.,]\d+)?\s*\(?$", "", raw)
        raw = raw.replace("ECOL", "").strip(" .-").lower()
        raw = re.sub(r"-1$", "", raw).strip(" .")
        return COLOR_MAP.get(raw, raw[:1].upper() + raw[1:] if raw else "—")
    return "—"


def finish(name: str, group: str) -> str:
    if group != "Эмаль ПФ-115":
        return ""
    if " гл." in name:
        return "Глянцевая"
    if " мат." in name:
        return "Матовая"
    return ""


def brand(name: str) -> str:
    upper = name.upper()
    if "ECOL" in upper:
        return "ECOL"
    if "FLAGMAN" in upper or "FLAGMAН" in upper or "ФЛАГМАН" in upper:
        return "FLAGMAN"
    if "BRAVA" in upper or "БРАВА" in upper:
        return "BRAVA"
    if "BELAKOR" in upper:
        return "BELAKOR"
    if "HAUSE MASTER" in upper:
        return "HAUSE MASTER"
    return "Без явного бренда"


def size_bucket(litres_value: float | None) -> str:
    if litres_value is None:
        return "Не указан"
    if litres_value < 1.5:
        return "до 1 л"
    if litres_value < 2.6:
        return "2–2,4 л"
    if litres_value < 6:
        return "3–5 л"
    return "7–11 л"


def cv(values: list[float]) -> float | None:
    mean = statistics.fmean(values) if values else 0
    if not mean:
        return None
    return statistics.pstdev(values) / mean


def pct(part: float, total: float) -> float:
    return part / total * 100 if total else 0.0


def build(rows: list[dict]) -> tuple[list[dict], dict]:
    items = []
    for source in rows:
        name = source["номенклатура"]
        group = product_group(name)
        monthly = [source["продажи_по_месяцам"][m] for m in MONTHS]
        volume_l = litres(name)
        total_units = sum(monthly)
        item = {
            "name": name,
            "warehouse": source["склад"],
            "group": group,
            "colour": colour(name, group),
            "finish": finish(name, group),
            "brand": brand(name),
            "unit": source["ед"],
            "article": source["артикул"],
            "volume_l": volume_l,
            "size": size_bucket(volume_l),
            "start": source["начальный_остаток"],
            "income": source["приход"],
            "outcome": source["расход"],
            "stock": source["конечный_остаток"],
            "min_reserve": source["мин_резерв"],
            "monthly": monthly,
            "total_units": total_units,
            "total_l": round(total_units * volume_l, 2) if volume_l else None,
            "full12_units": sum(monthly[:12]),
            "full12_l": round(sum(monthly[:12]) * volume_l, 2) if volume_l else None,
            "partial_units": monthly[-1],
            "last3_units": sum(monthly[9:12]),
            "prev3_units": sum(monthly[6:9]),
            "last6_units": sum(monthly[6:12]),
            "prev6_units": sum(monthly[:6]),
            "stock_l": round(source["конечный_остаток"] * volume_l, 2) if volume_l else None,
        }
        avg3 = item["last3_units"] / 3
        item["avg3"] = round(avg3, 2)
        item["cover_months"] = round(source["конечный_остаток"] / avg3, 2) if avg3 > 0 else None
        item["deficit_vs_reserve"] = round(max(0.0, source["мин_резерв"] - source["конечный_остаток"]), 2)
        annual_series = [(monthly[i] * volume_l if volume_l else monthly[i]) for i in range(12)]
        summer_series = [(monthly[i] * volume_l if volume_l else monthly[i]) for i in range(6, 12)]
        item["cv"] = round(cv(annual_series), 3) if sum(annual_series) > 0 else None
        item["cv_summer"] = round(cv(summer_series), 3) if sum(summer_series) > 0 else None
        item["active_months"] = sum(1 for v in monthly[:12] if v > 0)
        items.append(item)

    # --- stock status -----------------------------------------------------
    for item in items:
        avg3 = item["avg3"]
        stock = item["stock"]
        if stock > 0 and item["last6_units"] == 0:
            status = "Мёртвый остаток"
        elif stock == 0 and avg3 > 0:
            status = "Нет в наличии"
        elif stock == 0:
            status = "Нет остатка"
        elif stock < item["min_reserve"]:
            status = "Ниже резерва"
        elif item["cover_months"] is not None and item["cover_months"] > 6:
            status = "Избыток"
        else:
            status = "Норма"
        item["stock_status"] = status

    # --- ABC by litres over 12 full months --------------------------------
    ranked = sorted(items, key=lambda x: x["full12_l"] or 0, reverse=True)
    total_l = sum(x["full12_l"] or 0 for x in items)
    cumulative = 0.0
    for item in ranked:
        cumulative += item["full12_l"] or 0
        share = pct(cumulative, total_l)
        item["abc"] = "A" if share <= 80 else ("B" if share <= 95 else "C")
        item["abc_cum_share"] = round(share, 1)

    # --- XYZ by CV of monthly litres inside the season --------------------
    for item in items:
        c = item["cv_summer"]
        if c is None:
            item["xyz"] = "—"
        elif c <= 0.5:
            item["xyz"] = "X"
        elif c <= 1.0:
            item["xyz"] = "Y"
        else:
            item["xyz"] = "Z"

    # --- aggregates -------------------------------------------------------
    monthly_units = [sum(x["monthly"][i] for x in items) for i in range(13)]
    monthly_l = [round(sum((x["monthly"][i] * x["volume_l"]) if x["volume_l"] else 0
                           for x in items), 1) for i in range(13)]
    active_skus = [sum(1 for x in items if x["monthly"][i] > 0) for i in range(13)]
    total_units = sum(monthly_units)
    total_litres = round(sum(x["total_l"] or 0 for x in items), 1)
    full12_units = sum(monthly_units[:12])
    full12_litres = round(sum(monthly_l[:12]), 1)
    stock_units = sum(x["stock"] for x in items)
    stock_litres = round(sum(x["stock_l"] or 0 for x in items), 1)
    start_litres = round(sum((x["start"] * x["volume_l"]) if x["volume_l"] else 0 for x in items), 1)
    avg_stock = (sum(x["start"] for x in items) + stock_units) / 2
    avg_stock_l = (start_litres + stock_litres) / 2
    outcome_units = sum(x["outcome"] for x in items)

    group_stats = []
    for group in ["Эмаль ПФ-115", "Эмаль ПФ-266", "Краска", "Грунтовка", "Лак", "Прочее"]:
        subset = [x for x in items if x["group"] == group]
        if not subset:
            continue
        group_stats.append({
            "group": group,
            "sku": len(subset),
            "units": sum(x["total_units"] for x in subset),
            "litres": round(sum(x["total_l"] or 0 for x in subset), 1),
            "stock": sum(x["stock"] for x in subset),
            "stock_l": round(sum(x["stock_l"] or 0 for x in subset), 1),
            "unit_share": round(pct(sum(x["total_units"] for x in subset), total_units), 1),
            "litre_share": round(pct(sum(x["total_l"] or 0 for x in subset), total_litres), 1),
        })

    def breakdown(key: str, subset=None):
        subset = subset if subset is not None else items
        acc = defaultdict(lambda: {"units": 0.0, "litres": 0.0, "sku": 0, "stock": 0.0,
                                   "stock_l": 0.0, "last3": 0.0})
        for x in subset:
            k = x[key] if x[key] else "—"
            acc[k]["units"] += x["total_units"]
            acc[k]["litres"] += x["total_l"] or 0
            acc[k]["sku"] += 1
            acc[k]["stock"] += x["stock"]
            acc[k]["stock_l"] += x["stock_l"] or 0
            acc[k]["last3"] += x["last3_units"]
        out = [{"key": k, **{**v, "units": round(v["units"], 1),
                             "litres": round(v["litres"], 1),
                             "stock": round(v["stock"], 1),
                             "stock_l": round(v["stock_l"], 1),
                             "last3": round(v["last3"], 1)}}
               for k, v in acc.items()]
        for row in out:
            row["unit_share"] = round(pct(row["units"], total_units), 1)
            row["litre_share"] = round(pct(row["litres"], total_litres), 1)
            row["cover_months"] = round(row["stock"] / (row["last3"] / 3), 1) if row["last3"] else None
        return sorted(out, key=lambda r: r["litres"], reverse=True)

    status_counts = Counter(x["stock_status"] for x in items)
    status_units = defaultdict(float)
    status_litres = defaultdict(float)
    for x in items:
        status_units[x["stock_status"]] += x["stock"]
        status_litres[x["stock_status"]] += x["stock_l"] or 0

    abc_counts = Counter(x["abc"] for x in items)
    abc_units = defaultdict(float)
    abc_litres = defaultdict(float)
    for x in items:
        abc_units[x["abc"]] += x["total_units"]
        abc_litres[x["abc"]] += x["full12_l"] or 0

    xyz_counts = Counter(x["xyz"] for x in items)

    matrix = OrderedDict()
    for a in ["A", "B", "C"]:
        for z in ["X", "Y", "Z", "—"]:
            subset = [x for x in items if x["abc"] == a and x["xyz"] == z]
            if not subset:
                continue
            matrix[f"{a}{z}"] = {
                "sku": len(subset),
                "litres": round(sum(x["full12_l"] or 0 for x in subset), 1),
                "share": round(pct(sum(x["full12_l"] or 0 for x in subset), total_litres), 1),
            }

    sorted_by_l = sorted(items, key=lambda x: x["full12_l"] or 0, reverse=True)
    pareto = 0
    acc = 0.0
    for x in sorted_by_l:
        acc += x["full12_l"] or 0
        pareto += 1
        if acc >= 0.8 * total_l:
            break

    def share_l(n):
        return round(pct(sum(x["full12_l"] or 0 for x in sorted_by_l[:n]), total_litres), 1)

    top_units = sorted(items, key=lambda x: x["total_units"], reverse=True)[:15]
    top_litres = sorted(items, key=lambda x: x["total_l"] or 0, reverse=True)[:15]
    dead = [x for x in items if x["stock_status"] == "Мёртвый остаток"]
    out_of_stock = [x for x in items if x["stock_status"] == "Нет в наличии"]
    below_reserve = [x for x in items if x["stock"] < x["min_reserve"]]
    overstock = [x for x in items if x["stock_status"] == "Избыток"]
    enamel_colours = breakdown("colour", [x for x in items if x["colour"] != "—"])

    for x in items:
        x["trend_pct"] = (round((x["last3_units"] - x["prev3_units"]) / x["prev3_units"] * 100, 1)
                          if x["prev3_units"] else None)

    growth = sorted([x for x in items if x["prev3_units"] >= 3],
                    key=lambda x: x["last3_units"] - x["prev3_units"], reverse=True)[:10]
    decline = sorted([x for x in items if x["prev3_units"] >= 3],
                     key=lambda x: x["last3_units"] - x["prev3_units"])[:10]

    quarterly = []
    for label, indexes in (("IV кв. 2025", [0, 1, 2]), ("I кв. 2026", [3, 4, 5]),
                           ("II кв. 2026", [6, 7, 8]), ("III кв. 2026", [9, 10, 11]),
                           ("Окт. 2026 (1–19)", [12])):
        u = sum(monthly_units[i] for i in indexes)
        l = round(sum(monthly_l[i] for i in indexes), 1)
        quarterly.append({"label": label, "units": u, "litres": l,
                          "unit_share": round(pct(u, total_units), 1),
                          "litre_share": round(pct(l, total_litres), 1)})

    monthly_rows = []
    for i, m in enumerate(MONTHS):
        prev = monthly_units[i - 1] if i else None
        monthly_rows.append({
            "month": m,
            "label": MONTH_LABELS[i],
            "units": monthly_units[i],
            "litres": monthly_l[i],
            "active_sku": active_skus[i],
            "partial": m == PARTIAL_MONTH,
            "days": PARTIAL_DAYS if m == PARTIAL_MONTH else MONTH_DAYS[i],
            "avg_per_day": round(monthly_units[i] / (PARTIAL_DAYS if m == PARTIAL_MONTH
                                                      else MONTH_DAYS[i]), 2),
            "litres_per_day": round(monthly_l[i] / (PARTIAL_DAYS if m == PARTIAL_MONTH
                                                    else MONTH_DAYS[i]), 2),
            "mom_pct": round((monthly_units[i] - prev) / prev * 100, 1) if prev else None,
        })

    seasonality = [{"month": MONTH_LABELS[i],
                    "index": round(pct(monthly_units[i], full12_units / 12), 0)}
                   for i in range(12)]
    h1 = sum(monthly_units[:6])
    h2 = sum(monthly_units[6:12])
    winter = sum(monthly_units[:5])
    # Mar–Sep 2026 (индексы 5..11); неполный октябрь 2026 в «сезон» не включаем.
    season = sum(monthly_units[5:12])

    summary = {
        "source": "Ведомость по остаткам ТМЦ на складах, 01.10.25–19.10.26",
        "period": {"from": "2025-10-01", "to": "2026-10-19", "days": PERIOD_DAYS,
                   "full_months": 12, "partial_month": PARTIAL_MONTH, "partial_days": PARTIAL_DAYS},
        "months": MONTHS,
        "month_labels": MONTH_LABELS,
        "full_months": FULL_MONTHS,
        "partial_month": PARTIAL_MONTH,
        "sku": len(items),
        "active_sku": sum(1 for x in items if x["total_units"] > 0),
        "dead_sku": sum(1 for x in items if x["total_units"] == 0),
        "warehouses": dict(Counter(x["warehouse"] for x in items)),
        "units": dict(Counter(x["unit"] for x in items)),
        "totals": {
            "units": total_units,
            "litres": total_litres,
            "full12_units": full12_units,
            "full12_litres": full12_litres,
            "start": sum(x["start"] for x in items),
            "start_litres": start_litres,
            "income": sum(x["income"] for x in items),
            "outcome": outcome_units,
            "stock": stock_units,
            "stock_litres": stock_litres,
            "min_reserve": sum(x["min_reserve"] for x in items),
            "avg_stock": avg_stock,
            "avg_stock_l": round(avg_stock_l, 1),
            "partial_units": monthly_units[-1],
            "partial_litres": monthly_l[-1],
        },
        "monthly": monthly_rows,
        "quarterly": quarterly,
        "seasonality": seasonality,
        "halves": {"oct_mar": h1, "apr_sep": h2,
                   "growth_pct": round((h2 - h1) / h1 * 100, 1) if h1 else None,
                   "winter_oct_feb": winter, "season_mar_sep": season,
                   "season_share": round(pct(season, total_units), 1)},
        "groups": group_stats,
        "brands": breakdown("brand"),
        "colours": breakdown("colour"),
        "enamel_colours": enamel_colours,
        "sizes": breakdown("size"),
        "finish": breakdown("finish"),
        "abc": {k: {"sku": abc_counts.get(k, 0), "units": round(abc_units[k], 1),
                    "litres": round(abc_litres[k], 1),
                    "litre_share": round(pct(abc_litres[k], total_litres), 1),
                    "unit_share": round(pct(abc_units[k], total_units), 1)}
                for k in ["A", "B", "C"]},
        "xyz": dict(xyz_counts),
        "matrix": matrix,
        "pareto_sku": pareto,
        "concentration": {"top1": share_l(1), "top5": share_l(5), "top10": share_l(10),
                          "top20": share_l(20)},
        "stock_status": {k: {"sku": v, "units": round(status_units[k], 1),
                             "litres": round(status_litres[k], 1)}
                         for k, v in status_counts.items()},
        "dead": [{"name": x["name"], "stock": x["stock"], "stock_l": x["stock_l"],
                  "volume_l": x["volume_l"]} for x in dead],
        "out_of_stock": [{"name": x["name"], "avg3": x["avg3"], "min_reserve": x["min_reserve"]}
                         for x in out_of_stock],
        "below_reserve": [{"name": x["name"], "stock": x["stock"], "min_reserve": x["min_reserve"],
                           "deficit": x["deficit_vs_reserve"]} for x in below_reserve],
        "overstock": [{"name": x["name"], "stock": x["stock"], "cover_months": x["cover_months"],
                       "stock_l": x["stock_l"]} for x in overstock],
        "lost_demand": {
            "sku": len(out_of_stock),
            "avg3_units": round(sum(x["avg3"] for x in out_of_stock), 2),
            "avg3_litres": round(sum(x["avg3"] * (x["volume_l"] or 0) for x in out_of_stock), 2),
        },
        "deficit_total": round(sum(x["deficit_vs_reserve"] for x in items), 1),
        "reserve_surplus_total": round(sum(max(0.0, x["stock"] - x["min_reserve"]) for x in items), 1),
        "partial_dynamics": {
            "oct26_per_day": round(monthly_units[-1] / PARTIAL_DAYS, 2),
            "sep26_per_day": round(monthly_units[11] / 30, 2),
            "oct25_per_day": round(monthly_units[0] / 31, 2),
            "oct26_vs_sep_pct": round((monthly_units[-1] / PARTIAL_DAYS) /
                                      (monthly_units[11] / 30) * 100 - 100, 1),
            "oct26_vs_oct25_pct": round((monthly_units[-1] / PARTIAL_DAYS) /
                                        (monthly_units[0] / 31) * 100 - 100, 1),
        },
        "top_units": [{"name": x["name"], "units": x["total_units"], "litres": x["total_l"],
                       "stock": x["stock"], "abc": x["abc"], "xyz": x["xyz"]} for x in top_units],
        "top_litres": [{"name": x["name"], "units": x["total_units"], "litres": x["total_l"],
                        "stock": x["stock"], "abc": x["abc"], "xyz": x["xyz"]} for x in top_litres],
        "growth": [{"name": x["name"], "prev3": x["prev3_units"], "last3": x["last3_units"],
                    "delta": x["last3_units"] - x["prev3_units"]} for x in growth],
        "decline": [{"name": x["name"], "prev3": x["prev3_units"], "last3": x["last3_units"],
                     "delta": x["last3_units"] - x["prev3_units"]} for x in decline],
        "assortment": {
            "never_sold": sum(1 for x in items if x["total_units"] == 0),
            "no_sales_6m": sum(1 for x in items if x["last6_units"] == 0),
            "sold_1_2_months": sum(1 for x in items if 0 < x["active_months"] <= 2),
            "active_12_months": sum(1 for x in items if x["active_months"] == 12),
            "median_active_months": statistics.median([x["active_months"] for x in items]),
        },
        "items": sorted(items, key=lambda x: x["total_l"] or 0, reverse=True),
    }

    summary["turnover"] = {
        "turns_per_year": round(full12_units / avg_stock, 2) if avg_stock else None,
        "days": round(365 * avg_stock / full12_units, 0) if full12_units else None,
        "turns_per_year_litres": round(full12_litres / avg_stock_l, 2) if avg_stock_l else None,
        "days_litres": round(365 * avg_stock_l / full12_litres, 0) if full12_litres else None,
        "stock_days_at_last3": round(stock_units / (sum(x["last3_units"] for x in items) / 91), 0),
        "stock_months_at_last3": round(stock_units / (sum(x["last3_units"] for x in items) / 3), 1),
        "stock_days_at_last3_litres": round(stock_litres /
                                            (sum(x["last3_units"] * (x["volume_l"] or 0)
                                                 for x in items) / 91), 0),
    }
    mismatch = [x for x in items if abs(x["total_units"] - x["outcome"]) > 0.01]
    summary["quality"] = {
        "monthly_vs_outcome_diff": round(total_units - outcome_units, 1),
        "monthly_vs_outcome_pct": round(pct(total_units - outcome_units, outcome_units), 1),
        "monthly_eq_outcome_sku": len(items) - len(mismatch),
        "monthly_gt_outcome_sku": sum(1 for x in items if x["total_units"] > x["outcome"] + 0.01),
        "monthly_lt_outcome_sku": sum(1 for x in items if x["total_units"] < x["outcome"] - 0.01),
        "sales_gt_available_sku": sum(1 for x in items
                                      if x["total_units"] > x["start"] + x["income"] + 0.01),
        "with_article": sum(1 for x in items if x["article"]),
        "no_volume_sku": sum(1 for x in items if x["volume_l"] is None),
        "unit_kinds": dict(Counter(x["unit"] for x in items)),
        "balance_ok": all(abs(x["start"] + x["income"] - x["outcome"] - x["stock"]) < 0.01
                          for x in items),
    }
    return items, summary


def main(json_path: str, out_dir: str) -> None:
    rows = json.loads(Path(json_path).read_text(encoding="utf-8"))
    items, summary = build(rows)
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)

    fields = ["name", "warehouse", "group", "colour", "finish", "brand", "unit", "article",
              "volume_l", "size", "start", "income", "outcome", "stock", "min_reserve",
              "total_units", "total_l", "full12_units", "full12_l", "partial_units",
              "last3_units", "prev3_units", "last6_units", "prev6_units", "avg3",
              "cover_months", "stock_status", "abc", "xyz", "cv", "cv_summer", "active_months",
              "deficit_vs_reserve", "stock_l"]
    with (out / "vedomost_2025-10_2026-10.csv").open("w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.writer(fh, delimiter=";")
        writer.writerow(fields + MONTHS)
        for x in items:
            writer.writerow([x.get(k) for k in fields] + x["monthly"])
    (out / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=1),
                                      encoding="utf-8")

    t = summary["totals"]
    print(f"SKU: {summary['sku']} (active {summary['active_sku']}, dead {summary['dead_sku']})")
    print(f"Sales: {t['units']:.0f} units / {t['litres']:.0f} L "
          f"(12 full: {t['full12_units']:.0f} / {t['full12_litres']:.0f} L)")
    print("Monthly units:", [m["units"] for m in summary["monthly"]])
    print("Monthly litres:", [m["litres"] for m in summary["monthly"]])
    print("Active SKU:", [m["active_sku"] for m in summary["monthly"]])
    print("Stock:", t["stock"], "units /", t["stock_litres"], "L; reserve", t["min_reserve"])
    print("Turnover:", summary["turnover"])
    print("ABC:", summary["abc"])
    print("XYZ:", summary["xyz"])
    print("Matrix:", dict(summary["matrix"]))
    print("Pareto:", summary["pareto_sku"], "concentration:", summary["concentration"])
    print("Status:", summary["stock_status"])
    print("Lost demand:", summary["lost_demand"], "deficit total:", summary["deficit_total"])
    print("Assortment:", summary["assortment"])
    print("Quality:", summary["quality"])
    print("Halves:", summary["halves"])
    print("Quarterly:", summary["quarterly"])
    print("Colours (enamel):")
    for c in summary["enamel_colours"][:12]:
        print(f"   {c['key']:22} {c['litres']:8.1f} L {c['litre_share']:5.1f}%  sku={c['sku']}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
