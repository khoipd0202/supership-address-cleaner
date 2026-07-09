from __future__ import annotations

import io
import os
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

from clean_address import DATA as DEFAULT_DATA_PATH
from clean_address import parse_address_components
from parse_address import Parser, lookup_labels, strip_diacritics

from .models import CleanResult, CleanStats, OUTPUT_HEADERS


class ColumnNotFoundError(ValueError):
    pass


def _default_data_path() -> str:
    packaged = Path(__file__).resolve().parent / "data" / "vn_units_data.json"
    if packaged.exists():
        return str(packaged)
    return str(DEFAULT_DATA_PATH)


def _norm_header(value: Any) -> str:
    text = strip_diacritics(str(value or "").lower())
    return " ".join(text.replace("/", " ").replace("-", " ").split())


def _find_column(headers: list[Any], *needles: str) -> int | None:
    normalized = [_norm_header(h) for h in headers]
    for idx, header in enumerate(normalized):
        for needle in needles:
            n = _norm_header(needle)
            if n and (n == header or n in header):
                return idx
    return None


def detect_columns(headers: list[Any]) -> dict[str, int]:
    columns = {
        "address": _find_column(headers, "Địa chỉ", "Địa chỉ chi tiết", "Địa chỉ gốc"),
        "ward": _find_column(headers, "Phường/Xã", "Xã/Phường", "Phường", "Xã"),
        "district": _find_column(headers, "Quận/Huyện", "Huyện/Quận", "Quận", "Huyện"),
        "province": _find_column(headers, "Tỉnh/Thành Phố", "Tỉnh/TP", "Tỉnh"),
    }
    missing = [name for name, idx in columns.items() if idx is None]
    if missing:
        available = ", ".join(str(h or "") for h in headers)
        raise ColumnNotFoundError(f"Missing required columns: {', '.join(missing)}. Headers: {available}")
    return {k: int(v) for k, v in columns.items() if v is not None}


class AddressCleaner:
    """Clean SuperShip-style Excel files into a compact six-column workbook."""

    def __init__(self, data_path: str | os.PathLike[str] | None = None):
        self.data_path = str(data_path or _default_data_path())
        self.parser = Parser(self.data_path)

    def clean(
        self,
        raw_address: Any,
        ward: Any = None,
        district: Any = None,
        province: Any = None,
    ) -> CleanResult:
        labels = lookup_labels(self.parser, ward, district, province)
        clean_ward = labels.get("ward") or (str(ward).strip() if ward else "")
        clean_district = labels.get("district") or (str(district).strip() if district else "")
        clean_province = labels.get("province") or (str(province).strip() if province else "")
        ward_new = labels.get("ward_new")
        province_new = labels.get("province_new")

        parsed = parse_address_components(
            raw_address,
            [ward, district, province, clean_ward, clean_district, clean_province, ward_new, province_new],
        )
        return CleanResult(
            poi=parsed.get("poi") or "",
            street=parsed.get("street") or "",
            level4=parsed.get("level4") or "",
            ward=clean_ward or "",
            district=clean_district or "",
            province=clean_province or "",
            confidence=float(parsed.get("confidence") or 0.0),
            flags=tuple(parsed.get("flags") or ()),
        )

    def clean_excel(
        self,
        input_path: str | os.PathLike[str],
        output_path: str | os.PathLike[str],
        *,
        include_empty_rows: bool = False,
        split_components: bool = True,
        sheet_name: str | None = None,
    ) -> CleanStats:
        workbook = openpyxl.load_workbook(input_path, data_only=True)
        source = workbook[sheet_name] if sheet_name else workbook.active
        rows = list(source.iter_rows(values_only=True))
        output, stats = self.clean_rows_to_workbook(
            rows,
            include_empty_rows=include_empty_rows,
            split_components=split_components,
        )
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        output.save(output_path)
        return stats

    def clean_bytes(
        self,
        workbook_bytes: bytes,
        *,
        include_empty_rows: bool = False,
        split_components: bool = True,
        sheet_name: str | None = None,
    ) -> tuple[bytes, CleanStats]:
        workbook = openpyxl.load_workbook(io.BytesIO(workbook_bytes), data_only=True)
        source = workbook[sheet_name] if sheet_name else workbook.active
        rows = list(source.iter_rows(values_only=True))
        output, stats = self.clean_rows_to_workbook(
            rows,
            include_empty_rows=include_empty_rows,
            split_components=split_components,
        )
        buffer = io.BytesIO()
        output.save(buffer)
        return buffer.getvalue(), stats

    def clean_rows_to_workbook(
        self,
        rows: list[tuple[Any, ...]],
        *,
        include_empty_rows: bool = False,
        split_components: bool = True,
    ) -> tuple[openpyxl.Workbook, CleanStats]:
        output = openpyxl.Workbook()
        sheet = output.active
        sheet.title = "Địa chỉ sạch"
        sheet.append(list(OUTPUT_HEADERS))
        self._style_output(sheet)

        stats = CleanStats()
        if not rows:
            return output, stats

        headers = list(rows[0])
        columns = detect_columns(headers)
        for row in rows[1:]:
            stats.input_n += 1
            raw = row[columns["address"]] if columns["address"] < len(row) else ""
            ward = row[columns["ward"]] if columns["ward"] < len(row) else ""
            district = row[columns["district"]] if columns["district"] < len(row) else ""
            province = row[columns["province"]] if columns["province"] < len(row) else ""

            result = self.clean(raw, ward, district, province)
            if not include_empty_rows and not result.has_detail:
                stats.removed += 1
                continue

            output_rows = result.as_component_rows() if split_components else [result.as_output_row()]
            for output_row in output_rows:
                sheet.append(output_row)
            stats.output_n += len(output_rows)
            if result.ward and result.district and result.province:
                stats.full_admin += len(output_rows)

        self._autosize_output(sheet)
        return output, stats

    @staticmethod
    def _style_output(sheet) -> None:
        sheet.freeze_panes = "A2"
        header_fill = PatternFill("solid", fgColor="1F4E79")
        header_font = Font(bold=True, color="FFFFFF")
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

    @staticmethod
    def _autosize_output(sheet) -> None:
        widths = [32, 26, 26, 24, 26, 26]
        for idx, width in enumerate(widths, 1):
            sheet.column_dimensions[openpyxl.utils.get_column_letter(idx)].width = width
        for row in sheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)


def clean_excel(
    input_path: str | os.PathLike[str],
    output_path: str | os.PathLike[str],
    *,
    include_empty_rows: bool = False,
    split_components: bool = True,
    sheet_name: str | None = None,
    data_path: str | os.PathLike[str] | None = None,
) -> CleanStats:
    return AddressCleaner(data_path=data_path).clean_excel(
        input_path,
        output_path,
        include_empty_rows=include_empty_rows,
        split_components=split_components,
        sheet_name=sheet_name,
    )


def clean_workbook_bytes(
    workbook_bytes: bytes,
    *,
    include_empty_rows: bool = False,
    split_components: bool = True,
    sheet_name: str | None = None,
    data_path: str | os.PathLike[str] | None = None,
) -> tuple[bytes, CleanStats]:
    return AddressCleaner(data_path=data_path).clean_bytes(
        workbook_bytes,
        include_empty_rows=include_empty_rows,
        split_components=split_components,
        sheet_name=sheet_name,
    )
