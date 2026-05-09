from __future__ import annotations

import json
import re
from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, Iterable, Optional

from openpyxl import load_workbook
from pypdf import PdfReader

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_MAPPING_FILE = BASE_DIR / "mapping.json"
DEFAULT_TEMPLATE_FILE = BASE_DIR / "GenericBookingTemplate.xlsx"


@dataclass
class ConversionResult:
    excel_bytes: bytes
    parsed: Dict[str, Any]
    output_filename: str


def load_mapping(mapping_file: Path | str = DEFAULT_MAPPING_FILE) -> Dict[str, Any]:
    with open(mapping_file, "r", encoding="utf-8") as f:
        return json.load(f)


def pdf_to_text(pdf_file: str | Path | BytesIO) -> str:
    reader = PdfReader(pdf_file)
    pages = []
    for page in reader.pages:
        pages.append(page.extract_text() or "")
    return "\n".join(pages)


def _clean(value: Optional[str]) -> str:
    return "" if value is None else str(value).strip()


def _to_float(value: Any, default: float = 0.0) -> float:
    if value is None or value == "":
        return default
    value = str(value).strip().replace(",", ".")
    try:
        return float(value)
    except ValueError:
        return default


def _to_int_or_float(value: Any) -> Any:
    num = _to_float(value, 0.0)
    return int(num) if float(num).is_integer() else num


def parse_datetime(value: str | None) -> Optional[datetime]:
    value = _clean(value)
    if not value:
        return None
    for fmt in ("%d/%m/%Y %H:%M", "%d/%m/%Y %H:%M:%S"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    return None


def regex_value(text: str, pattern: str, default: str = "") -> str:
    match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE | re.DOTALL)
    return _clean(match.group(1)) if match else default


def parse_booking_pdf_text(text: str, mapping: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    mapping = mapping or load_mapping()
    lines = [_clean(line) for line in text.splitlines() if _clean(line)]
    normalized = "\n".join(lines)

    driver_name = lines[0] if lines else ""

    phone = regex_value(normalized, r"Born Date\s*0\s*\n([^\n]+)", "")
    if not re.search(r"\d", phone):
        phone = ""

    pickup_section = regex_value(
        normalized,
        r"PickUp Details([\s\S]*?)Drop Off Details",
        ""
    )

    dropoff_section = regex_value(
        normalized,
        r"Drop Off Details([\s\S]*?)Product Code",
        ""
    )

    pickup_dt = regex_value(
        pickup_section,
        r"Date-Time\s*([0-9]{2}/[0-9]{2}/[0-9]{4}\s+[0-9]{2}:[0-9]{2})",
        ""
    )

    pickup_station = regex_value(
        pickup_section,
        r"Station\s*([^\n]+)",
        ""
    )

    dropoff_dt = regex_value(
        dropoff_section,
        r"Date-Time\s*([0-9]{2}/[0-9]{2}/[0-9]{4}\s+[0-9]{2}:[0-9]{2})",
        ""
    )

    dropoff_station = regex_value(
        dropoff_section,
        r"Station\s*([^\n]+)",
        ""
    )

    car_group = regex_value(normalized, r"Car Group\s*\n([^\n\s/]+)")
    product_code = regex_value(normalized, r"Product Code\s*\n([^\n]+)")
    voucher_no = regex_value(normalized, r"Voucher No\s*\n([^\n]+)")
    flight = regex_value(normalized, r"Flight\s*\n([^\n]+)")
    if flight in {"Drop Off Details", "Product Code", "Payment Method Voucher"}:
        flight = ""

    voucher_value = _to_float(regex_value(normalized, r"Voucher Value\s*([0-9]+(?:[\.,][0-9]+)?)"))
    pay_on_arrival = _to_float(regex_value(normalized, r"Pay On Arrival\s*([0-9]+(?:[\.,][0-9]+)?)"))
    total_amount_pdf = _to_float(regex_value(normalized, r"Total Amount\s*([0-9]+(?:[\.,][0-9]+)?)"))
    total_amount = total_amount_pdf or (voucher_value + pay_on_arrival)

    reservation_datetime = parse_datetime(
        regex_value(normalized, r"([0-9]{2}/[0-9]{2}/[0-9]{4}\s+[0-9]{2}:[0-9]{2})\s+Confirmed")
    )
    reservation_number = regex_value(normalized, r"Reservation\s+([0-9]+)")

    station_code_map = mapping.get("station_code_map", {})
    station_location_map = mapping.get("station_location_map", {})

    parsed = {
        "driver_name": driver_name,
        "phone": phone,
        "car_group": car_group,
        "pickup_datetime": parse_datetime(pickup_dt),
        "pickup_station": pickup_station,
        "pickup_station_code": station_code_map.get(pickup_station, ""),
        "pickup_location": station_location_map.get(pickup_station, ""),
        "dropoff_datetime": parse_datetime(dropoff_dt),
        "dropoff_station": dropoff_station,
        "dropoff_station_code": station_code_map.get(dropoff_station, ""),
        "dropoff_location": station_location_map.get(dropoff_station, ""),
        "product_code": product_code,
        "voucher_no": voucher_no,
        "flight": flight,
        "voucher_value": voucher_value,
        "pay_on_arrival": pay_on_arrival,
        "total_amount": total_amount,
        "program": "pp" if voucher_value > 0 else "POA",
        "excess": _to_int_or_float(regex_value(normalized, r"EXCESS\s*([0-9]+(?:[\.,][0-9]+)?)")),
        "deposit": _to_int_or_float(regex_value(normalized, r"Deposit\s*([0-9]+(?:[\.,][0-9]+)?)")),
        "reservation_datetime": reservation_datetime,
        "reservation_number": reservation_number,
        "baby_seats": _to_int_or_float(regex_value(normalized, r"Baby Seats\s*([0-9]+)")),
        "child_seats": _to_int_or_float(regex_value(normalized, r"Child Seats\s*([0-9]+)")),
        "infant_seats": _to_int_or_float(regex_value(normalized, r"Infant Seats\s*([0-9]+)")),
        "add_drivers": _to_int_or_float(regex_value(normalized, r"Add\. Drivers\s*([0-9]+)")),
        "del_fee": _to_float(regex_value(normalized, r"Del Fee\s*([0-9]+(?:[\.,][0-9]+)?)")),
        "col_fee": _to_float(regex_value(normalized, r"Col Fee\s*([0-9]+(?:[\.,][0-9]+)?)")),
        "one_way_fee": _to_float(regex_value(normalized, r"One Way Fee\s*([0-9]+(?:[\.,][0-9]+)?)")),
        "night_fee": _to_float(regex_value(normalized, r"Night Fee\s*([0-9]+(?:[\.,][0-9]+)?)")),
        "raw_text": text,
    }
    return parsed


def build_output_row(headers: Iterable[str], parsed: Dict[str, Any], mapping: Dict[str, Any]) -> Dict[str, Any]:
    fixed = mapping.get("fixed_values", {})
    field_mapping = mapping.get("field_mapping", {})
    row: Dict[str, Any] = {}

    for header in headers:
        if header in fixed:
            row[header] = fixed[header]
        elif header in field_mapping:
            row[header] = parsed.get(field_mapping[header], "")
        else:
            row[header] = ""

    return row


def write_to_template(parsed: Dict[str, Any], template_file: Path | str, mapping: Dict[str, Any]) -> bytes:
    wb = load_workbook(template_file)
    ws = wb["Bookings"] if "Bookings" in wb.sheetnames else wb.active

    headers = [cell.value for cell in ws[1] if cell.value]
    row_data = build_output_row(headers, parsed, mapping)

    for row_idx in range(ws.max_row, 1, -1):
        ws.delete_rows(row_idx)

    output_row_idx = 2
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(output_row_idx, col_idx)
        cell.value = row_data.get(header, "")
        if isinstance(cell.value, datetime):
            cell.number_format = "DD/MM/YYYY HH:MM"

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.read()


def convert_pdf_to_excel(
    pdf_file: str | Path | BytesIO,
    template_file: Path | str = DEFAULT_TEMPLATE_FILE,
    mapping_file: Path | str = DEFAULT_MAPPING_FILE,
) -> ConversionResult:
    mapping = load_mapping(mapping_file)
    text = pdf_to_text(pdf_file)
    parsed = parse_booking_pdf_text(text, mapping)
    excel_bytes = write_to_template(parsed, template_file, mapping)
    res_no = parsed.get("reservation_number") or "booking"
    return ConversionResult(
        excel_bytes=excel_bytes,
        parsed=parsed,
        output_filename=f"GenericBookingOutput_{res_no}.xlsx",
    )
