#!/usr/bin/env python3
"""
Parse คลังสื่อ Excel → JSON + CSV
แล้วสร้าง group_map.json เชื่อมกับ equipment.json
"""
import json
import csv
import re
import sys
import openpyxl
from pathlib import Path

EXCEL_PATH = Path(
    "/root/.claude/uploads/d67c8eb6-6ed9-4c01-812e-223d4b206ed0/"
    "bb0f7a59-___________________.xlsx"
)

# ----- Group name normalisation -----
# Maps คลังสื่อ กลุ่มงาน → department_id (short key)
MEDIA_GROUP_MAP = {
    "วัยเรียน":                      "วัยเรียน",
    "ผู้สูงอายุ":                    "ผู้สูงอายุ",
    "สิ่งแวดล้อม":                   "อนามัยสิ่งแวดล้อม",
    "ประเมินผลกระทบต่อสุขภาพ":       "อนามัยสิ่งแวดล้อม",
    "ทันตสาธารณสุข":                 "ทันตสาธารณสุข",
}

# Maps equipment responsible → department_id (short key)
EQUIP_RESPONSIBLE_MAP = {
    "กลุ่มอนามัยวัยเรียน":                      "วัยเรียน",
    "กลุ่มอนามัยแม่และเด็ก":                    "ผู้สูงอายุ",   # closest match
    "กลุ่มอนามัยวัยรุ่นและวัยเจริญพันธ์":       "วัยเรียน",
    "กลุ่มอนามัยสิ่งแวดล้อม":                   "อนามัยสิ่งแวดล้อม",
    "กลุ่มทันตสาธารณสุข":                        "ทันตสาธารณสุข",
    "กลุ่มสื่อสาร ไอที":                          "สื่อสาร ไอที",
    "กลุ่มการพยาบาล":                             "การพยาบาล",
    "กลุ่มเภสัชกรรมและแพทย์แผนไทย":            "เภสัชกรรม",
    "กลุ่มชันสูตร แลป":                           "ชันสูตร",
    "กลุ่มอำนวยการ":                              "อำนวยการ",
    "กลุ่มวิชาการ วิจัยและสนับสนุนศูนย์เขต":     "วิจัย",
    "กลุ่มฝึกอบรม":                               "ฝึกอบรม",
    "ศูนย์พัฒนาการเด็กปฐมวัย(EF)":              "พัฒนาการเด็ก",
    "กายภาพบำบัด":                               "กายภาพบำบัด",
}


def clean_str(v) -> str | None:
    if v is None:
        return None
    s = str(v).strip()
    return s if s else None


def parse_quantity(v) -> int | None:
    if v is None:
        return None
    try:
        return int(float(str(v).replace(",", "").strip()))
    except ValueError:
        return None


def parse_media_items(ws) -> list[dict]:
    items = []
    for row in ws.iter_rows(min_row=3, values_only=True):
        code = clean_str(row[0])
        if not code or not code.startswith("M"):
            continue
        group_raw = clean_str(row[1]) or ""
        items.append({
            "media_id":    code,
            "group_raw":   group_raw,
            "department":  MEDIA_GROUP_MAP.get(group_raw, group_raw),
            "name":        clean_str(row[2]),
            "unit":        clean_str(row[3]),
            "link":        clean_str(row[8]),
        })
    return items


def parse_stock_in(ws) -> dict[str, int]:
    totals: dict[str, int] = {}
    for row in ws.iter_rows(min_row=3, values_only=True):
        code = clean_str(row[3])
        qty  = parse_quantity(row[5])
        status = clean_str(row[7]) or ""
        if code and qty and status != "ยกเลิก":
            totals[code] = totals.get(code, 0) + qty
    return totals


def parse_stock_out(ws) -> dict[str, int]:
    totals: dict[str, int] = {}
    for row in ws.iter_rows(min_row=3, values_only=True):
        code = clean_str(row[4])
        qty  = parse_quantity(row[6])
        status = clean_str(row[8]) or ""
        if code and qty and status != "ยกเลิก":
            totals[code] = totals.get(code, 0) + qty
    return totals


def stock_status(bal: int, received: int) -> str:
    if bal <= 0:
        return "หมด"
    if received > 0 and bal / received < 0.2:
        return "ใกล้หมด"
    return "มีเหลือ"


def main():
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)

    items   = parse_media_items(wb["📋 รายการสื่อ"])
    stock_in  = parse_stock_in(wb["📥 รับเข้า"])
    stock_out = parse_stock_out(wb["📤 จ่ายออก"])

    for item in items:
        mid = item["media_id"]
        received = stock_in.get(mid, 0)
        issued   = stock_out.get(mid, 0)
        balance  = received - issued
        item["received"]  = received
        item["issued"]    = issued
        item["balance"]   = balance
        item["status"]    = stock_status(balance, received)

    # --- Write JSON ---
    out_json = Path("data/media.json")
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(items, f, ensure_ascii=False, indent=2)

    # --- Write CSV ---
    out_csv = Path("data/media.csv")
    fields = ["media_id", "department", "group_raw", "name", "unit",
              "received", "issued", "balance", "status", "link"]
    with open(out_csv, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        w.writerows(items)

    # --- Write group mapping ---
    group_map = {
        "description": "แมปกลุ่มงานระหว่าง คลังสื่อ และ ครุภัณฑ์ เพื่อ JOIN ข้อมูล",
        "media_group_to_department": MEDIA_GROUP_MAP,
        "equipment_responsible_to_department": EQUIP_RESPONSIBLE_MAP,
        "join_key": "department",
        "example_query": (
            "SELECT e.equipment_id, e.name AS equipment_name, m.name AS media_name "
            "FROM equipment e "
            "JOIN media m ON equip_dept_map[e.responsible] = media_dept_map[m.group_raw]"
        ),
    }
    out_map = Path("data/group_map.json")
    with open(out_map, "w", encoding="utf-8") as f:
        json.dump(group_map, f, ensure_ascii=False, indent=2)

    print(f"Done:")
    print(f"  {len(items)} media items  → {out_json}, {out_csv}")
    print(f"  group map               → {out_map}")
    total_bal = sum(i["balance"] for i in items)
    out_items = sum(1 for i in items if i["status"] == "หมด")
    low_items = sum(1 for i in items if i["status"] == "ใกล้หมด")
    print(f"  ยอดรวมคงเหลือ: {total_bal:,} ชิ้น | หมด: {out_items} | ใกล้หมด: {low_items}")


if __name__ == "__main__":
    main()
