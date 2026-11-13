"""QA PSV comparison engine adapted from supplied QApsv_compare_list_v4.py.
Comparison logic and report structure retained; input/output use in-memory streams.
"""
from io import BytesIO
from datetime import datetime
from html import unescape
from pathlib import Path
from typing import Any
import re

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


def clean(value: Any) -> str:
    text = unescape(str(value or "")).replace("\xa0", " ")
    text = re.sub(r"<[^>]*>", "", text)
    return re.sub(r"\\s+", " ", text).strip()


def key(value: Any) -> str:
    return clean(value).casefold()


def normalize_masks(value: Any) -> str:
    text = clean(value)
    if key(text) in {"", "n/a", "na", "-"}:
        return ""
    return ",".join(sorted(x.strip().casefold() for x in text.split(",") if x.strip()))


def normalize_oit_restriction(value: Any) -> str:
    text = key(value)
    if text == "yes":
        return "restricted offering"
    if text in {"no", "", "n/a", "na", "-"}:
        return ""
    return text


def normalize_webdm_restriction(value: Any) -> str:
    text = key(value)
    if text in {"", "n/a", "na", "-", "no"}:
        return ""
    if text in {"yes", "restricted", "restricted offering"}:
        return "restricted offering"
    return text


def maturity_matches(oit_value: Any, webdm_value: Any) -> bool:
    oit, webdm = key(oit_value), key(webdm_value)
    if oit == "yes":
        return webdm == "qualified"
    if oit == "no":
        return webdm in {"development", "frozen"}
    return oit == webdm


def read_oit(file_bytes: bytes, filename: str) -> dict[str, dict[str, Any]]:
    workbook = load_workbook(BytesIO(file_bytes), read_only=True, data_only=True)
    if "OIT_result" not in workbook.sheetnames:
        raise ValueError("OIT1 must contain a sheet named 'OIT_result'.")
    sheet = workbook["OIT_result"]
    headers = {clean(sheet.cell(2, c).value): c for c in range(1, sheet.max_column + 1)}
    required = ["Feature Category", "Feature ID", "Feature Description",
                "Additional Mask List", "Restricted (Yes/No)", "Qualified (Yes/No)"]
    missing = [name for name in required if name not in headers]
    if missing:
        raise ValueError("Missing OIT1 column(s): " + ", ".join(missing))
    records = {}
    for row_num in range(3, sheet.max_row + 1):
        feature_id = clean(sheet.cell(row_num, headers["Feature ID"]).value)
        if not feature_id:
            continue
        records[key(feature_id)] = {
            "id": feature_id, "row": row_num,
            "category": clean(sheet.cell(row_num, headers["Feature Category"]).value),
            "description": clean(sheet.cell(row_num, headers["Feature Description"]).value),
            "masks": clean(sheet.cell(row_num, headers["Additional Mask List"]).value),
            "restrictions": clean(sheet.cell(row_num, headers["Restricted (Yes/No)"]).value),
            "maturity": clean(sheet.cell(row_num, headers["Qualified (Yes/No)"]).value),
        }
    workbook.close()
    return records


def read_webdm(file_bytes: bytes, filename: str) -> dict[str, dict[str, Any]]:
    suffix = Path(filename).suffix.lower()
    records: dict[str, dict[str, Any]] = {}
    if suffix == ".psv":
        text = None
        for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                text = file_bytes.decode(encoding)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            raise ValueError(f"Cannot decode PSV file: {filename}")
        for row_num, raw in enumerate(text.splitlines(), start=1):
            if not raw.startswith("|Device Listing|"):
                continue
            fields = [clean(x) for x in raw.split("|")[1:]]
            if fields and fields[-1] == "":
                fields.pop()
            if len(fields) > 2 and fields[2] == "Device Name":
                headers = {name: index for index, name in enumerate(fields)}
                continue
            if "headers" not in locals():
                continue
            def get_field(name: str) -> str:
                index = headers.get(name)
                return fields[index] if index is not None and index < len(fields) else ""
            device_name = get_field("Device Name")
            if device_name:
                records[key(device_name)] = {
                    "id": device_name, "row": row_num,
                    "category": get_field("Device Class"),
                    "description": get_field("Device Description"),
                    "masks": get_field("Additional Masks"),
                    "restrictions": get_field("Restrictions"),
                    "maturity": get_field("Maturity"),
                }
        if not records:
            raise ValueError("Could not find Device Listing in raw PSV file.")
        return records

    workbook = load_workbook(BytesIO(file_bytes), read_only=True, data_only=True)
    # Format A: PSV-export workbook, pipe-separated rows in column A.
    for sheet in workbook.worksheets:
        headers = None
        for row_num, row in enumerate(sheet.iter_rows(min_col=1, max_col=1, values_only=True), start=1):
            raw = row[0]
            if not isinstance(raw, str) or not raw.startswith("|Device Listing|"):
                continue
            fields = [clean(x) for x in raw.split("|")[1:]]
            if fields and fields[-1] == "":
                fields.pop()
            if len(fields) > 2 and fields[2] == "Device Name":
                headers = {name: index for index, name in enumerate(fields)}
                continue
            if not headers:
                continue
            def get_field(name: str) -> str:
                index = headers.get(name)
                return fields[index] if index is not None and index < len(fields) else ""
            device_name = get_field("Device Name")
            if device_name:
                records[key(device_name)] = {
                    "id": device_name, "row": row_num,
                    "category": get_field("Device Class"),
                    "description": get_field("Device Description"),
                    "masks": get_field("Additional Masks"),
                    "restrictions": get_field("Restrictions"),
                    "maturity": get_field("Maturity"),
                }
        if records:
            workbook.close()
            return records

    required = {"Device Name", "Device Class", "Device Description",
                "Additional Masks", "Restrictions", "Maturity"}
    for sheet in workbook.worksheets:
        for header_row in range(1, min(sheet.max_row, 30) + 1):
            headers = {clean(sheet.cell(header_row, c).value): c
                       for c in range(1, sheet.max_column + 1)
                       if clean(sheet.cell(header_row, c).value)}
            if not required.issubset(headers):
                continue
            for row_num in range(header_row + 1, sheet.max_row + 1):
                device_name = clean(sheet.cell(row_num, headers["Device Name"]).value)
                if not device_name:
                    continue
                records[key(device_name)] = {
                    "id": device_name, "row": row_num,
                    "category": clean(sheet.cell(row_num, headers["Device Class"]).value),
                    "description": clean(sheet.cell(row_num, headers["Device Description"]).value),
                    "masks": clean(sheet.cell(row_num, headers["Additional Masks"]).value),
                    "restrictions": clean(sheet.cell(row_num, headers["Restrictions"]).value),
                    "maturity": clean(sheet.cell(row_num, headers["Maturity"]).value),
                }
            if records:
                workbook.close()
                return records
    workbook.close()
    raise ValueError("Could not find a Device Listing in PSV or normal Excel format.")


def compare_records(oit: dict, webdm: dict):
    all_rows = []
    matched_keys = sorted(set(oit) & set(webdm))
    for record_key in matched_keys:
        left, right = oit[record_key], webdm[record_key]
        checks = [
            ("Feature Category / Device Class", left["category"], right["category"], key(left["category"]) == key(right["category"])),
            ("Device Description", left["description"], right["description"], key(left["description"]) == key(right["description"])),
            ("Additional Masks", left["masks"], right["masks"], normalize_masks(left["masks"]) == normalize_masks(right["masks"])),
            ("Restrictions", left["restrictions"], right["restrictions"], normalize_oit_restriction(left["restrictions"]) == normalize_webdm_restriction(right["restrictions"])),
            ("Maturity", left["maturity"], right["maturity"], maturity_matches(left["maturity"], right["maturity"])),
        ]
        for field, oit_value, webdm_value, same in checks:
            all_rows.append([left["id"], field, oit_value, webdm_value,
                             "Same" if same else "Different", left["row"], right["row"]])
    differences = [row for row in all_rows if row[4] == "Different"]
    oit_only = sorted(set(oit) - set(webdm))
    webdm_only = sorted(set(webdm) - set(oit))
    return matched_keys, all_rows, differences, oit_only, webdm_only


def create_qa_report(oit_bytes: bytes, oit_filename: str, webdm_bytes: bytes, webdm_filename: str):
    oit = read_oit(oit_bytes, oit_filename)
    webdm = read_webdm(webdm_bytes, webdm_filename)
    matched, all_rows, differences, oit_only, webdm_only = compare_records(oit, webdm)

    workbook = Workbook()
    workbook.remove(workbook.active)
    navy, white, red, green = "1F4E78", "FFFFFF", "FCE4D6", "E2F0D9"

    def setup_sheet(sheet, title, headers):
        sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
        title_cell = sheet.cell(1, 1, title)
        title_cell.fill = PatternFill("solid", fgColor=navy)
        title_cell.font = Font(color=white, bold=True, size=14)
        sheet.append(headers)
        for cell in sheet[2]:
            cell.fill = PatternFill("solid", fgColor=navy)
            cell.font = Font(color=white, bold=True)
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        sheet.freeze_panes = "A3"
        sheet.sheet_view.showGridLines = False

    sheet = workbook.create_sheet("Summary")
    setup_sheet(sheet, "OIT vs webDM Device Listing QA", ["Metric", "Count", "Rule / Note"])
    summary_rows = [
        ("OIT IDs", len(oit), oit_filename),
        ("webDM IDs", len(webdm), webdm_filename),
        ("Matched IDs", len(matched), "OIT Feature ID = webDM Device Name"),
        ("Different field records", len(differences), "See Differences tab"),
        ("Devices with differences", len({row[0] for row in differences}), "Unique IDs"),
        ("OIT-only IDs", len(oit_only), "See Unmatched tab"),
        ("webDM-only IDs", len(webdm_only), "See Unmatched tab"),
        ("Maturity mapping", "", "Yes = Qualified; No = Development/Frozen"),
    ]
    for item in summary_rows:
        sheet.append(item)
    for column, width in zip("ABC", [30, 16, 78]):
        sheet.column_dimensions[column].width = width

    detail_headers = ["Feature ID / Device Name", "Field", "OIT Value", "webDM Value", "Result", "OIT Row", "webDM PSV Row"]
    for sheet_name, data in [("Differences", differences), ("All_Comparisons", all_rows)]:
        sheet = workbook.create_sheet(sheet_name)
        setup_sheet(sheet, sheet_name.replace("_", " "), detail_headers)
        for item in data:
            sheet.append(item)
        for row in sheet.iter_rows(min_row=3):
            row[4].fill = PatternFill("solid", fgColor=red if row[4].value == "Different" else green)
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        if sheet.max_row > 2:
            table = Table(displayName=sheet_name.replace("_", "") + "Table", ref=f"A2:G{sheet.max_row}")
            table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False)
            sheet.add_table(table)
        for index, width in enumerate([30, 32, 60, 60, 12, 10, 15], start=1):
            sheet.column_dimensions[get_column_letter(index)].width = width

    sheet = workbook.create_sheet("Unmatched")
    setup_sheet(sheet, "Unmatched IDs", ["Source", "ID", "Source Row", "Comment"])
    for record_key in oit_only:
        sheet.append(["OIT only", oit[record_key]["id"], oit[record_key]["row"], "Not found in webDM"])
    for record_key in webdm_only:
        sheet.append(["webDM only", webdm[record_key]["id"], webdm[record_key]["row"], "Not found in OIT"])
    for column, width in zip("ABCD", [18, 40, 14, 48]):
        sheet.column_dimensions[column].width = width

    sheet = workbook.create_sheet("Read_Me")
    setup_sheet(sheet, "Comparison Rules", ["OIT Field", "webDM Field", "Comparison Rule"])
    rules = [
        ("Feature Category", "Device Class", "HTML/whitespace normalized"),
        ("Feature Description", "Device Description", "HTML/whitespace normalized"),
        ("Additional Mask List", "Additional Masks", "Mask order ignored; blank/n/a equivalent"),
        ("Restricted (Yes/No)", "Restrictions", "Yes = Restricted offering; No/n/a = blank"),
        ("Qualified (Yes/No)", "Maturity", "Yes = Qualified; No = Development/Frozen"),
    ]
    for item in rules:
        sheet.append(item)
    for column, width in zip("ABC", [32, 32, 78]):
        sheet.column_dimensions[column].width = width

    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.calculation.calcMode = "auto"
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    mmdd = datetime.now().strftime("%m%d")
    match = re.search(r"PT\d+", Path(oit_filename).stem, flags=re.IGNORECASE)
    pt = match.group(0).upper() if match else "PTxxxx"
    return output.getvalue(), f"{mmdd}_QApsv_result_{pt}.xlsx", {
        "oit": len(oit), "webdm": len(webdm), "matched": len(matched),
        "differences": len(differences),
        "devices_with_differences": len({row[0] for row in differences}),
        "oit_only": len(oit_only), "webdm_only": len(webdm_only),
    }
