#!/usr/bin/env python3
"""
Parse and clean equipment inventory data from markdown table → JSON + CSV
"""
import json
import csv
import re
import sys


def clean_value(val: str) -> str | None:
    val = val.strip()
    if val in ("", "\\-", "-", "\\\\-", "\\\\\\-"):
        return None
    return val


def parse_price(val: str) -> float | None:
    val = re.sub(r"[\s,]", "", val)
    try:
        return float(val)
    except ValueError:
        return None


def parse_tags(val: str) -> list[str]:
    if not val or not val.strip():
        return []
    return [t.strip() for t in val.split(",") if t.strip()]


def parse_markdown_table(path: str) -> list[dict]:
    with open(path, encoding="utf-8") as f:
        lines = f.readlines()

    headers = None
    items = []

    for line in lines:
        line = line.strip()
        if not line.startswith("|"):
            continue
        # Skip alignment/separator row (entire line is only |, :, -, spaces)
        if re.fullmatch(r"[\|:\- ]+", line):
            continue

        cells = [c.strip() for c in line.split("|")[1:-1]]

        if headers is None:
            headers = cells
            continue

        if len(cells) < 15:
            continue

        item = {
            "fsn":           clean_value(cells[0]),
            "equipment_id":  clean_value(cells[1]),
            "gf_id":         clean_value(cells[2]),
            "fiscal_year":   clean_value(cells[3]),
            "category":      clean_value(cells[4]),
            "tags":          parse_tags(cells[5]),
            "name":          clean_value(cells[6]),
            "location":      clean_value(cells[7]),
            "responsible":   clean_value(cells[8]),
            "brand":         clean_value(cells[9]),
            "model":         clean_value(cells[10]),
            "unit_price":    parse_price(cells[11]),
            "unit":          clean_value(cells[12]),
            "received_date": clean_value(cells[13]),
            "status":        clean_value(cells[14]),
        }
        items.append(item)

    return items


def write_json(items: list[dict], path: str) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(items, f, ensure_ascii=False, indent=2)


def write_csv(items: list[dict], path: str) -> None:
    fields = [
        "fsn", "equipment_id", "gf_id", "fiscal_year", "category", "tags",
        "name", "location", "responsible", "brand", "model",
        "unit_price", "unit", "received_date", "status",
    ]
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        for item in items:
            row = {**item}
            row["tags"] = "|".join(item["tags"])
            writer.writerow(row)


if __name__ == "__main__":
    src = sys.argv[1] if len(sys.argv) > 1 else "data/raw.md"
    items = parse_markdown_table(src)
    write_json(items, "data/equipment.json")
    write_csv(items, "data/equipment.csv")
    print(f"Done: {len(items)} records → data/equipment.json + data/equipment.csv")
