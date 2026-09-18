from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from io import BytesIO
from pathlib import Path
from typing import BinaryIO

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


FileInput = str | Path | BinaryIO


def _text(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\u00a0", " ")).strip()


def _key(value: object) -> str:
    return _text(value).upper()


def _item_id(value: object) -> str:
    text = _text(value)
    return text[:-2] if text.endswith(".0") else text


def _number(value: object) -> float | None:
    if value is None or pd.isna(value) or _text(value) == "":
        return None
    try:
        return float(str(value).replace(",", ""))
    except (TypeError, ValueError):
        return None


def _find_sheet(sheets: dict[str, pd.DataFrame], expected: str) -> tuple[str, pd.DataFrame]:
    target = _key(expected)
    for name, frame in sheets.items():
        if _key(name) == target:
            return name, frame
    raise ValueError(f"Required sheet '{expected}' was not found.")


def _read_workbook(source: FileInput) -> dict[str, pd.DataFrame]:
    return pd.read_excel(source, sheet_name=None, header=None, dtype=object, engine="openpyxl")


@dataclass
class ParsedShift:
    expected_shift: str
    detected_shift: str | None
    detected_date: date | None
    rows: pd.DataFrame
    warnings: list[str] = field(default_factory=list)


@dataclass
class ReportResult:
    report: pd.DataFrame
    errors: list[str]
    warnings: list[str]
    source_stats: dict[str, object]

    @property
    def valid(self) -> bool:
        return not self.errors


def _detect_shift_and_date(sheets: dict[str, pd.DataFrame], expected_shift: str) -> tuple[str | None, date | None]:
    detected_shift = None
    detected_date = None
    for frame in sheets.values():
        values = frame.fillna("").astype(str).values.tolist()
        for row in values[:25]:
            for index, value in enumerate(row):
                text = _text(value)
                shift_match = re.search(r"\bSHIFT\s*:?\s*([AB])\b", text, re.I)
                if shift_match:
                    detected_shift = shift_match.group(1).upper()
                if re.fullmatch(r"DATE\s*:?", text, re.I) and index + 1 < len(row):
                    parsed = pd.to_datetime(row[index + 1], dayfirst=True, errors="coerce")
                    if not pd.isna(parsed):
                        detected_date = parsed.date()
        if detected_shift and detected_date:
            break
    return detected_shift, detected_date


def parse_shift_workbook(source: FileInput, expected_shift: str) -> ParsedShift:
    expected_shift = expected_shift.upper()
    sheets = _read_workbook(source)
    _, frame = _find_sheet(sheets, "PRINTING MATL REQ")
    detected_shift, detected_date = _detect_shift_and_date(sheets, expected_shift)
    warnings: list[str] = []
    if detected_shift and detected_shift != expected_shift:
        warnings.append(f"Uploaded Shift {expected_shift} slot contains a workbook labelled Shift {detected_shift}.")
    if not detected_shift:
        warnings.append(f"Shift label was not detected; the file is being treated as Shift {expected_shift}.")

    contractor = ""
    columns: dict[str, int] = {}
    records: list[dict[str, object]] = []
    aliases = {
        "ITEM ID": "item_id",
        "ITEM GROUP": "item_group",
        "ITEM NAME": "item_name",
        "UOM": "uom",
        "REQUIED QTY": "required_qty",
        "REQUIRED QTY": "required_qty",
        "REQUEST QTY": "request_qty",
        "AVAILABLE FLOOR STOCK": "available_floor_stock",
        "PENDING RECEIVE REQUISITION": "pending_receive_requisition",
    }

    for row_number, row in enumerate(frame.itertuples(index=False, name=None), start=1):
        cells = [_text(value) for value in row]
        joined = " | ".join(value for value in cells if value)
        contractor_match = re.search(r"CONTRACTOR\s+NAME\s*:\s*(.+?)(?:\s*\||$)", joined, re.I)
        if contractor_match:
            contractor = _text(contractor_match.group(1))
            continue

        normalized = [_key(value) for value in row]
        if "ITEM ID" in normalized and "ITEM NAME" in normalized:
            columns = {aliases[name]: index for index, name in enumerate(normalized) if name in aliases}
            continue

        if not columns:
            continue
        item_index = columns.get("item_id")
        name_index = columns.get("item_name")
        request_index = columns.get("request_qty")
        if item_index is None or name_index is None or request_index is None:
            continue
        item = _item_id(row[item_index])
        item_name = _text(row[name_index])
        if not item or not item_name or not re.fullmatch(r"\d+", item):
            continue
        request_qty = _number(row[request_index])
        if request_qty is None:
            warnings.append(f"{expected_shift}: nonnumeric Request Qty at source row {row_number}; treated as zero.")
            request_qty = 0.0
        if not contractor:
            warnings.append(f"{expected_shift}: item {item} at source row {row_number} has no contractor heading.")
        record = {
            "Contractor Name": contractor,
            "Printing Item ID": item,
            "Source Item Name": item_name,
            "Item Group": _text(row[columns["item_group"]]) if "item_group" in columns else "",
            "UOM": _text(row[columns["uom"]]) if "uom" in columns else "",
            "Request PCS": request_qty,
            "Required Qty": _number(row[columns["required_qty"]]) if "required_qty" in columns else None,
            "Available Floor Stock": _number(row[columns["available_floor_stock"]]) if "available_floor_stock" in columns else None,
            "Source Row": row_number,
        }
        records.append(record)

    if not records:
        raise ValueError(f"No printing-material rows were found for Shift {expected_shift}.")
    return ParsedShift(expected_shift, detected_shift, detected_date, pd.DataFrame(records), warnings)


def read_master(source: FileInput) -> tuple[pd.DataFrame, list[str], list[str]]:
    sheets = pd.read_excel(source, sheet_name=None, dtype=object, engine="openpyxl")
    name = next((name for name in sheets if _key(name) == "MASTER FILE"), None)
    if not name:
        raise ValueError("The master workbook must contain a 'Master File' sheet.")
    frame = sheets[name].copy()
    frame.columns = [_text(col) for col in frame.columns]
    required = [
        "Printing Item ID", "Printing Item Name", "Machine Line Code", "Machine Line Name",
        "PCS per KG", "Roll Weight (KG)", "Core/Tare Weight (KG)", "Allowance",
    ]
    missing_columns = [column for column in required if column not in frame.columns]
    if missing_columns:
        raise ValueError("Master columns missing: " + ", ".join(missing_columns))
    frame = frame[required].dropna(how="all").copy()
    frame["Printing Item ID"] = frame["Printing Item ID"].map(_item_id)
    frame["Item Name Key"] = frame["Printing Item Name"].map(_key)
    errors: list[str] = []
    warnings: list[str] = []
    duplicates = frame[frame["Printing Item ID"].duplicated(False)]["Printing Item ID"].unique().tolist()
    if duplicates:
        errors.append("Duplicate Printing Item IDs in master: " + ", ".join(duplicates[:20]))
    duplicate_names = frame[frame["Item Name Key"].duplicated(False)]["Printing Item Name"].unique().tolist()
    if duplicate_names:
        calculation_fields = [
            "Machine Line Code", "Machine Line Name", "PCS per KG",
            "Roll Weight (KG)", "Core/Tare Weight (KG)", "Allowance",
        ]
        conflicts = []
        for item_name, group in frame.groupby("Item Name Key"):
            if len(group) < 2:
                continue
            if any(group[column].dropna().astype(str).nunique() > 1 for column in calculation_fields):
                conflicts.append(item_name)
        if conflicts:
            errors.append("Duplicate Item Names with conflicting master values: " + ", ".join(conflicts[:20]))
        noun = "Item Name has" if len(duplicate_names) == 1 else "Item Names have"
        warnings.append(
            f"{len(duplicate_names)} duplicate {noun} identical calculation values; the master codes were combined."
        )
        aggregations = {column: "first" for column in required if column != "Printing Item ID"}
        aggregations["Printing Item ID"] = lambda values: " / ".join(dict.fromkeys(map(_item_id, values)))
        frame = frame.groupby("Item Name Key", as_index=False).agg(aggregations)
    if (frame["Printing Item ID"] == "").any():
        errors.append("Master contains blank Printing Item IDs.")
    return frame, errors, warnings


def _prepare_shift(parsed: ParsedShift, label: str) -> pd.DataFrame:
    frame = parsed.rows.copy()
    frame["Item Name Key"] = frame["Source Item Name"].map(_key)
    grouped = frame.groupby(
        ["Contractor Name", "Item Name Key"], as_index=False, dropna=False
    ).agg({"Source Item Name": "first", "Printing Item ID": "first", "Request PCS": "sum"})
    return grouped.rename(columns={
        "Source Item Name": f"{label} Source Name",
        "Printing Item ID": f"Shift {label} Item ID",
        "Request PCS": f"Shift {label} PCS",
    })


def build_report(master_source: FileInput, shift_a_source: FileInput, shift_b_source: FileInput,
                 default_allowance_percent: float = 3.0) -> ReportResult:
    errors: list[str] = []
    warnings: list[str] = []
    master, master_errors, master_warnings = read_master(master_source)
    errors.extend(master_errors)
    warnings.extend(master_warnings)
    shift_a = parse_shift_workbook(shift_a_source, "A")
    shift_b = parse_shift_workbook(shift_b_source, "B")
    warnings.extend(shift_a.warnings + shift_b.warnings)

    a = _prepare_shift(shift_a, "A")
    b = _prepare_shift(shift_b, "B")
    report = a.merge(b, on=["Contractor Name", "Item Name Key"], how="outer")
    report["Shift A PCS"] = pd.to_numeric(report["Shift A PCS"], errors="coerce").fillna(0.0)
    report["Shift B PCS"] = pd.to_numeric(report["Shift B PCS"], errors="coerce").fillna(0.0)
    zero_request_count = int(((report["Shift A PCS"] + report["Shift B PCS"]) <= 0).sum())
    if zero_request_count:
        warnings.append(f"{zero_request_count} rows with zero Request Qty in both shifts were excluded from the issue report.")
        report = report[(report["Shift A PCS"] + report["Shift B PCS"]) > 0].copy()
    report = report.merge(master, on="Item Name Key", how="left", validate="many_to_one")

    missing_master = report[report["Printing Item Name"].isna()]["Item Name Key"].unique().tolist()
    if missing_master:
        errors.append("Printing Item Names missing from master: " + ", ".join(missing_master[:30]))
    source_ids = report["Shift A Item ID"].fillna(report["Shift B Item ID"])
    differing_ids = report[
        report["Shift A Item ID"].notna() & report["Shift B Item ID"].notna()
        & (report["Shift A Item ID"] != report["Shift B Item ID"])
    ]
    if not differing_ids.empty:
        warnings.append(f"{len(differing_ids)} rows have different daily Item IDs between shifts; exact Item Name was used.")
    report["Daily Item ID"] = source_ids
    report = report.rename(columns={"Printing Item ID": "Master Item Code"})

    for column in ["Machine Line Code", "Machine Line Name", "PCS per KG"]:
        bad = report[report[column].isna() | (report[column].map(_text) == "")]
        if not bad.empty:
            errors.append(f"{column} missing for: " + ", ".join(bad["Item Name Key"].unique().tolist()[:30]))
    conversion = pd.to_numeric(report["PCS per KG"], errors="coerce")
    invalid_conversion = report[conversion.isna() | (conversion <= 0)]
    if not invalid_conversion.empty:
        errors.append("Invalid PCS per KG for: " + ", ".join(invalid_conversion["Item Name Key"].unique().tolist()[:30]))

    roll = pd.to_numeric(report["Roll Weight (KG)"], errors="coerce")
    core = pd.to_numeric(report["Core/Tare Weight (KG)"], errors="coerce")
    missing_weight = report[roll.isna() | core.isna()]
    if not missing_weight.empty:
        errors.append("Roll/core weight missing for: " + ", ".join(missing_weight["Item Name Key"].unique().tolist()[:30]))
    invalid_pair = report[(roll <= 0) & (core > 0)]
    if not invalid_pair.empty:
        errors.append("Core weight exists but Roll Weight is zero for: " + ", ".join(invalid_pair["Item Name Key"].unique().tolist()[:30]))
    zero_pairs = report[(roll.fillna(0) == 0) & (core.fillna(0) == 0)]
    if not zero_pairs.empty:
        warnings.append(f"{len(zero_pairs)} report rows use explicit zero roll/core weights; no core/tare addition will be applied.")

    allowance = pd.to_numeric(report["Allowance"], errors="coerce")
    allowance = allowance.where(allowance.notna(), default_allowance_percent / 100.0)
    percentage_style = allowance > 1
    if percentage_style.any():
        allowance = allowance.where(~percentage_style, allowance / 100.0)
        warnings.append("Allowance values greater than 1 were interpreted as percentages.")
    invalid_allowance = report[(allowance < 0) | (allowance >= 1)]
    if not invalid_allowance.empty:
        errors.append("Allowance must be between 0% and 100%.")

    report["PCS per KG"] = conversion
    report["Roll Weight (KG)"] = roll
    report["Core/Tare Weight (KG)"] = core
    report["Allowance"] = allowance
    report["Total PCS"] = report["Shift A PCS"] + report["Shift B PCS"]
    for label in ["A", "B"]:
        pcs = report[f"Shift {label} PCS"]
        base = pcs / conversion
        tare = ((base / roll) * core).where(roll > 0, 0.0)
        report[f"Shift {label} KG"] = ((base + tare) * (1 + allowance)).round(0)
    report["Total KG"] = report["Shift A KG"] + report["Shift B KG"]

    output_columns = [
        "Contractor Name", "Daily Item ID", "Master Item Code", "Printing Item Name", "Machine Line Code",
        "Machine Line Name", "Shift A PCS", "Shift B PCS", "Total PCS", "PCS per KG",
        "Roll Weight (KG)", "Core/Tare Weight (KG)", "Allowance",
        "Shift A KG", "Shift B KG", "Total KG",
    ]
    report = report[output_columns].sort_values(
        ["Contractor Name", "Printing Item Name"], kind="stable"
    ).reset_index(drop=True)
    item_prefixes = report["Printing Item Name"].fillna("").astype(str).str.strip().str.upper()
    stats = {
        "Shift A source rows": len(shift_a.rows),
        "Shift B source rows": len(shift_b.rows),
        "Output rows": len(report),
        "Shift A date": shift_a.detected_date,
        "Shift B date": shift_b.detected_date,
        "Total PCS": float(report["Total PCS"].sum()),
        "Shift A KG": float(report["Shift A KG"].sum()),
        "Shift A TAG KG": float(report.loc[item_prefixes.str.startswith("TAG"), "Shift A KG"].sum()),
        "Shift A ENV KG": float(report.loc[item_prefixes.str.startswith("ENV"), "Shift A KG"].sum()),
        "Shift B KG": float(report["Shift B KG"].sum()),
        "Shift B TAG KG": float(report.loc[item_prefixes.str.startswith("TAG"), "Shift B KG"].sum()),
        "Shift B ENV KG": float(report.loc[item_prefixes.str.startswith("ENV"), "Shift B KG"].sum()),
        "Total KG": float(report["Total KG"].sum()),
    }
    return ReportResult(report, list(dict.fromkeys(errors)), list(dict.fromkeys(warnings)), stats)


def create_excel_report(result: ReportResult, report_date: date | None = None) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Contractor Report"
    report_date = report_date or datetime.now().date()
    headers = result.report.columns.tolist()
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    ws.cell(1, 1, f"CONTRACTOR-WISE PRINTING MATERIAL ISSUE REPORT — {report_date:%d-%m-%Y}")
    ws.cell(1, 1).font = Font(bold=True, color="FFFFFF", size=14)
    ws.cell(1, 1).fill = PatternFill("solid", fgColor="17365D")
    ws.cell(1, 1).alignment = Alignment(horizontal="center")
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(3, col, header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    for row_idx, row in enumerate(result.report.itertuples(index=False, name=None), start=4):
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row_idx, col_idx, None if pd.isna(value) else value)
    total_row = ws.max_row + 1
    ws.cell(total_row, 1, "GRAND TOTAL").font = Font(bold=True)
    for col_name in ["Shift A PCS", "Shift B PCS", "Total PCS", "Shift A KG", "Shift B KG", "Total KG"]:
        col = headers.index(col_name) + 1
        ws.cell(total_row, col, f"=SUM({get_column_letter(col)}4:{get_column_letter(col)}{total_row-1})")
        ws.cell(total_row, col).font = Font(bold=True)
    for cell in ws[total_row]:
        cell.fill = PatternFill("solid", fgColor="D9EAD3")
    thin = Side(style="thin", color="D9E2F3")
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.border = Border(bottom=thin)
    widths = [20, 17, 17, 55, 18, 28, 15, 15, 15, 15, 18, 22, 14, 15, 15, 15]
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for row in range(4, ws.max_row + 1):
        for col in [7, 8, 9]:
            ws.cell(row, col).number_format = "#,##0.00"
        for col in [10, 11, 12]:
            ws.cell(row, col).number_format = "#,##0.000"
        for col in [14, 15, 16]:
            ws.cell(row, col).number_format = "#,##0"
        ws.cell(row, 13).number_format = "0.00%"
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:{get_column_letter(len(headers))}{total_row-1}"

    check = wb.create_sheet("Validation")
    check.append(["Status", "PASS" if result.valid else "FAILED"])
    check.append(["Check", "Details"])
    for key, value in result.source_stats.items():
        check.append([key, value])
    for warning in result.warnings:
        check.append(["Warning", warning])
    for error in result.errors:
        check.append(["Error", error])
    for cell in check[1] + check[2]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="17365D")
    check.column_dimensions["A"].width = 24
    check.column_dimensions["B"].width = 110
    check.freeze_panes = "A3"

    output = BytesIO()
    wb.save(output)
    return output.getvalue()


CONTRACTOR_REPORT_COLUMNS = [
    "Contractor Name",
    "Daily Item ID",
    "Printing Item Name",
    "Machine Line Name",
    "Shift A KG",
    "Shift B KG",
    "Total KG",
    "Issued in PCS",
    "Issued in KG",
]


def contractor_report_frame(result: ReportResult, contractor: str) -> pd.DataFrame:
    """Return the concise issue sheet used for one contractor's download."""
    contractor_rows = result.report[result.report["Contractor Name"] == contractor].copy()
    if contractor_rows.empty:
        raise ValueError(f"Contractor '{contractor}' was not found in the report.")

    contractor_rows["Issued in PCS"] = ""
    contractor_rows["Issued in KG"] = ""
    return contractor_rows[CONTRACTOR_REPORT_COLUMNS].reset_index(drop=True)


def create_contractor_excel_report(
    result: ReportResult,
    contractor: str,
    report_date: date | None = None,
) -> bytes:
    """Create a single-sheet, contractor-only workbook with issue-entry columns."""
    report = contractor_report_frame(result, contractor)
    report_date = report_date or datetime.now().date()

    wb = Workbook()
    ws = wb.active
    ws.title = "Contractor Issue Report"
    headers = report.columns.tolist()

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    ws.cell(1, 1, f"{contractor.upper()} — PRINTING MATERIAL ISSUE REPORT — {report_date:%d-%m-%Y}")
    ws.cell(1, 1).font = Font(bold=True, color="FFFFFF", size=14)
    ws.cell(1, 1).fill = PatternFill("solid", fgColor="17365D")
    ws.cell(1, 1).alignment = Alignment(horizontal="center")

    for col, header in enumerate(headers, start=1):
        cell = ws.cell(3, col, header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    for row_idx, row in enumerate(report.itertuples(index=False, name=None), start=4):
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row_idx, col_idx, None if pd.isna(value) else value)

    total_row = ws.max_row + 1
    ws.cell(total_row, 1, "TOTAL").font = Font(bold=True)
    for col_name in ["Shift A KG", "Shift B KG", "Total KG"]:
        col = headers.index(col_name) + 1
        ws.cell(total_row, col, f"=SUM({get_column_letter(col)}4:{get_column_letter(col)}{total_row-1})")
        ws.cell(total_row, col).font = Font(bold=True)
    for cell in ws[total_row]:
        cell.fill = PatternFill("solid", fgColor="D9EAD3")

    thin = Side(style="thin", color="D9E2F3")
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.border = Border(bottom=thin)

    widths = [20, 17, 55, 28, 15, 15, 15, 17, 17]
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for row in range(4, ws.max_row + 1):
        for col in [5, 6, 7, 9]:
            ws.cell(row, col).number_format = "#,##0"
        ws.cell(row, 8).number_format = "#,##0.00"

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:{get_column_letter(len(headers))}{total_row-1}"

    output = BytesIO()
    wb.save(output)
    return output.getvalue()


PRINT_COLUMNS = [
    "Printing Item Name",
    "Machine Line Name",
    "Shift A KG",
    "Shift B KG",
    "Total KG",
    "Issued in PCS",
    "Issued in KG",
]
_PRINT_NUMERIC_COLUMNS = {"Shift A KG", "Shift B KG", "Total KG"}


def _escape_html(value: object) -> str:
    text = "" if value is None or pd.isna(value) else str(value)
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


# A4 pagination budget, all in millimetres. Row/heading heights are picked so the
# CSS font-size + padding used below always renders SHORTER than the budgeted
# value — the budget is intentionally conservative so content never overflows a
# simulated page (which would break the 1 preview page = 1 printed page guarantee).
_PAGE_HEIGHT_MM = 297
_PAGE_WIDTH_MM = 210
_PAGE_PADDING_MM = 14
_CONTENT_HEIGHT_MM = _PAGE_HEIGHT_MM - 2 * _PAGE_PADDING_MM
_CONTENT_WIDTH_MM = _PAGE_WIDTH_MM - 2 * _PAGE_PADDING_MM
_SAFETY_BUFFER_MM = 10
_USABLE_HEIGHT_MM = _CONTENT_HEIGHT_MM - _SAFETY_BUFFER_MM

_TITLE_BLOCK_MM = 12
_RUNNING_HEADER_MM = 8
_CONTRACTOR_HEADER_MM = 8
_COLUMN_HEADER_MM = 7
_DATA_ROW_MM = 6
_TOTAL_ROW_MM = 7
_SPACER_MM = 6

# Column widths adapt to the longest value actually present, so full names
# stay on one line (the row-height budget above assumes single-line rows).
_CHAR_WIDTH_MM = 1.9
_CELL_PADDING_MM = 6.0
_MIN_ITEM_NAME_MM = 45.0
_MAX_ITEM_NAME_MM = 110.0
_MIN_MACHINE_NAME_MM = 28.0
_MAX_MACHINE_NAME_MM = 70.0
_MIN_OTHER_COLUMN_MM = 12.0
_OTHER_COLUMN_COUNT = 5


def _text_column_width(values: list[str], min_mm: float, max_mm: float) -> float:
    longest = max((len(v) for v in values), default=0)
    width = longest * _CHAR_WIDTH_MM + _CELL_PADDING_MM
    return min(max_mm, max(min_mm, width))


def _compute_column_widths(result: ReportResult) -> tuple[float, float, float]:
    item_names = result.report["Printing Item Name"].dropna().astype(str).tolist()
    machine_names = result.report["Machine Line Name"].dropna().astype(str).tolist()

    item_mm = _text_column_width(item_names, _MIN_ITEM_NAME_MM, _MAX_ITEM_NAME_MM)
    machine_mm = _text_column_width(machine_names, _MIN_MACHINE_NAME_MM, _MAX_MACHINE_NAME_MM)

    other_total_floor = _MIN_OTHER_COLUMN_MM * _OTHER_COLUMN_COUNT
    remaining = _CONTENT_WIDTH_MM - item_mm - machine_mm
    if remaining < other_total_floor:
        deficit = other_total_floor - remaining
        flex_total = item_mm + machine_mm
        item_mm -= deficit * (item_mm / flex_total)
        machine_mm -= deficit * (machine_mm / flex_total)
        remaining = other_total_floor

    other_mm = remaining / _OTHER_COLUMN_COUNT
    return item_mm, machine_mm, other_mm


def build_contractor_print_html(result: ReportResult, report_date: date | None = None) -> str:
    """Build a standalone, paginated A4 HTML document listing every contractor's issue sheet.

    Page breaks are computed server-side from fixed row-height budgets (not measured
    in the browser), so the on-screen preview and the physically printed pages are
    guaranteed to split in exactly the same places.
    """
    report_date = report_date or datetime.now().date()
    contractors = sorted(result.report["Contractor Name"].dropna().unique().tolist())

    contractor_data: list[tuple[str, list[tuple], dict[str, float]]] = []
    for contractor in contractors:
        frame = contractor_report_frame(result, contractor)[PRINT_COLUMNS]
        rows = list(frame.itertuples(index=False, name=None))
        totals = {column: float(frame[column].sum()) for column in _PRINT_NUMERIC_COLUMNS}
        contractor_data.append((contractor, rows, totals))

    pages = _paginate_contractors(contractor_data)
    item_mm, machine_mm, other_mm = _compute_column_widths(result)
    pages_html = "".join(
        _render_page(blocks, index, len(pages), report_date, item_mm, machine_mm, other_mm)
        for index, blocks in enumerate(pages)
    )

    return f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  @page {{ size: A4; margin: 0; }}
  * {{ box-sizing: border-box; }}
  html, body {{ margin: 0; padding: 0; background: #e5e7eb; }}
  body {{ font-family: Arial, Helvetica, sans-serif; color: #111827; }}
  .print-bar {{ text-align: center; padding: 16px; }}
  .print-bar button {{
    background: #E8B923; color: #17150c; border: none; border-radius: 6px;
    padding: 10px 22px; font-size: 14px; font-weight: 600; cursor: pointer;
  }}
  .page {{
    width: {_PAGE_WIDTH_MM}mm; height: {_PAGE_HEIGHT_MM}mm; padding: {_PAGE_PADDING_MM}mm;
    margin: 0 auto 16px; background: #ffffff; box-shadow: 0 2px 10px rgba(0,0,0,0.25);
    position: relative; overflow: hidden;
  }}
  .page-title {{
    font-size: 15px; font-weight: bold; text-align: center; color: #17365D;
    margin: 0 0 6mm;
  }}
  .running-header {{
    font-size: 11px; font-weight: bold; text-align: center; color: #17365D;
    margin: 0 0 4mm;
  }}
  table.report-table {{ width: 100%; border-collapse: collapse; table-layout: fixed; }}
  .contractor-header th {{
    background: #D9E2F3; color: #17365D; text-align: left; font-size: 13px;
    font-weight: bold; padding: 5px 8px; border-bottom: 2px solid #17365D;
  }}
  .column-header th {{
    background: #EAF0F8; color: #17365D; font-size: 10.5px; font-weight: bold;
    padding: 4px 6px; text-align: center; border: 1px solid #17365D;
  }}
  td {{
    border: 1px solid #9fb2c8; padding: 3px 6px; font-size: 10.5px; text-align: right;
    color: #111827; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
  }}
  td:nth-child(1), td:nth-child(2) {{ text-align: left; }}
  tr.total-row td {{ font-weight: bold; background: #D9EAD3; color: #14532D; border-color: #4b7a57; }}
  tr.spacer-row td {{ border: none; padding: 0; height: {_SPACER_MM}mm; }}
  .page-number {{
    position: absolute; bottom: 5mm; right: {_PAGE_PADDING_MM}mm;
    font-size: 9px; color: #6b7280;
  }}
  * {{
    -webkit-print-color-adjust: exact; print-color-adjust: exact; color-adjust: exact;
  }}
  @media print {{
    body {{ background: #ffffff; }}
    .print-bar {{ display: none; }}
    .page {{
      margin: 0; box-shadow: none; page-break-after: always;
    }}
    .page:last-child {{ page-break-after: auto; }}
  }}
</style>
</head>
<body>
  <div class="print-bar"><button onclick="window.print()">Print / open print window</button></div>
  {pages_html}
</body>
</html>"""


def _paginate_contractors(
    contractor_data: list[tuple[str, list[tuple], dict[str, float]]],
) -> list[list[tuple]]:
    """Split contractor blocks into pages using fixed mm-height budgets.

    Returns a list of pages; each page is a list of block tuples such as
    ("contractor_header", name, continued), ("column_header",), ("data_row", row),
    ("total_row", totals) or ("spacer",).
    """
    pages: list[list[tuple]] = []
    current: list[tuple] = []
    current_height = 0.0

    def usable_height() -> float:
        header = _TITLE_BLOCK_MM if not pages else _RUNNING_HEADER_MM
        return _USABLE_HEIGHT_MM - header

    def start_new_page() -> None:
        nonlocal current, current_height
        pages.append(current)
        current = []
        current_height = 0.0

    for contractor, rows, totals in contractor_data:
        capacity = usable_height()
        section_head_height = _CONTRACTOR_HEADER_MM + _COLUMN_HEADER_MM
        first_row_height = _DATA_ROW_MM if rows else _TOTAL_ROW_MM
        if current and current_height + section_head_height + first_row_height > capacity:
            start_new_page()
            capacity = usable_height()

        current.append(("contractor_header", contractor, False))
        current.append(("column_header",))
        current_height += section_head_height

        for row in rows:
            if current_height + _DATA_ROW_MM > capacity:
                start_new_page()
                capacity = usable_height()
                current.append(("contractor_header", contractor, True))
                current.append(("column_header",))
                current_height += section_head_height
            current.append(("data_row", row))
            current_height += _DATA_ROW_MM

        if current_height + _TOTAL_ROW_MM > capacity:
            start_new_page()
            capacity = usable_height()
            current.append(("contractor_header", contractor, True))
            current.append(("column_header",))
            current_height += section_head_height
        current.append(("total_row", totals))
        current_height += _TOTAL_ROW_MM

        if current_height + _SPACER_MM <= capacity:
            current.append(("spacer",))
            current_height += _SPACER_MM

    if current:
        pages.append(current)
    return pages or [[]]


def _render_page(
    blocks: list[tuple],
    index: int,
    total_pages: int,
    report_date: date,
    item_mm: float,
    machine_mm: float,
    other_mm: float,
) -> str:
    header_html = (
        f'<div class="page-title">CONTRACTOR-WISE PRINTING MATERIAL ISSUE REPORT &mdash; {report_date:%d-%m-%Y}</div>'
        if index == 0
        else f'<div class="running-header">CONTRACTOR-WISE PRINTING MATERIAL ISSUE REPORT &mdash; {report_date:%d-%m-%Y} (continued)</div>'
    )

    rows_html: list[str] = []
    for block in blocks:
        kind = block[0]
        if kind == "contractor_header":
            _, contractor, continued = block
            label = _escape_html(contractor) + (" (continued)" if continued else "")
            rows_html.append(f"<tr class='contractor-header'><th colspan='7'>{label}</th></tr>")
        elif kind == "column_header":
            rows_html.append(
                "<tr class='column-header'>"
                "<th>Printing Item Name</th><th>Machine Line Name</th><th>Shift A KG</th>"
                "<th>Shift B KG</th><th>Total KG</th><th>Issued in PCS</th><th>Issued in KG</th>"
                "</tr>"
            )
        elif kind == "data_row":
            _, row = block
            cells = []
            for column, value in zip(PRINT_COLUMNS, row):
                if column in _PRINT_NUMERIC_COLUMNS:
                    text = "" if pd.isna(value) else f"{value:,.0f}"
                else:
                    text = _escape_html(value)
                cells.append(f"<td>{text}</td>")
            rows_html.append("<tr>" + "".join(cells) + "</tr>")
        elif kind == "total_row":
            _, totals = block
            rows_html.append(
                "<tr class='total-row'>"
                "<td colspan='2'>TOTAL</td>"
                f"<td>{totals['Shift A KG']:,.0f}</td>"
                f"<td>{totals['Shift B KG']:,.0f}</td>"
                f"<td>{totals['Total KG']:,.0f}</td>"
                "<td></td><td></td>"
                "</tr>"
            )
        elif kind == "spacer":
            rows_html.append("<tr class='spacer-row'><td colspan='7'></td></tr>")

    colgroup = (
        "<colgroup>"
        f"<col style='width:{item_mm:.1f}mm'><col style='width:{machine_mm:.1f}mm'>"
        f"<col style='width:{other_mm:.1f}mm'><col style='width:{other_mm:.1f}mm'>"
        f"<col style='width:{other_mm:.1f}mm'><col style='width:{other_mm:.1f}mm'>"
        f"<col style='width:{other_mm:.1f}mm'>"
        "</colgroup>"
    )

    return (
        "<div class='page'>"
        + header_html
        + "<table class='report-table'>" + colgroup + "".join(rows_html) + "</table>"
        + f"<div class='page-number'>Page {index + 1} of {total_pages}</div>"
        + "</div>"
    )
