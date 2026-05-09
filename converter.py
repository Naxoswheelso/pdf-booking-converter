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

VAT_RATE = 1.24


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


def add_vat_24(value: Any) -> float:
    """
    Adds 24% VAT to extra charges.
    Example: 16.13 -> 20.00
    """
    amount = _to_float(value, 0.0)
    return round(amount * VAT_RATE, 2) if amount else 0


def parse_datetime(value: str | None) -> Optional[datetime]:
    value = _clean(value)
    if not value:
        return None
    value = re.sub(r"\s+", " ", value)
    for fmt in ("%d/%m/%Y %H:%M", "%d/%m/%Y %H:%M:%S"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    return None


def regex_value(text: str, pattern: str, default: str = "") -> str:
    match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE | re.DOTALL)
    return _clean(match.group(1)) if match else default


def section_between(text: str, start_label: str, end_label: str) -> str:
    pattern = re.escape(start_label) + r"([\s\S]*?)" + re.escape(end_label)
    return regex_value(text, pattern, "")


def extract_section_datetime(section: str) -> str:
    value = regex_value(
        section,
        r"Date-Time\s*([0-9]{2}/[0-9]{2}/[0-9]{4}\s+[0-9]{2}:[0-9]{2})",
        ""
    )
    if value:
        return value

    match = re.search(
        r"([0-9]{2}/[0-9]{2}/[0-9]{4})\s+([0-9]{2}:[0-9]{2})\s*Date-Time",
        section,
        re.IGNORECASE | re.MULTILINE | re.DOTALL
    )
    if match:
        return f"{match.group(1)} {match.group(2)}"

    return regex_value(
        section,
        r"([0-9]{2}/[0-9]{2}/[0-9]{4}\s+[0-9]{2}:[0-9]{2})",
        ""
    )


def extract_station(section: str) -> str:
    return regex_value(section, r"Station\s*([^\n]+)", "")


def extract_extra_amount(text: str, label: str) -> float:
    """
    Smart extraction for extra charges. Handles all PDF formats:

    Format 1 - "Label qty amount" (zero values):
        "Baby Seats 0 0" -> 0
        "Child Seats 0 0" -> 0

    Format 2 - "Label qty amount" (with values, space preserved):
        "Baby Seats 1 10" -> 10
        "Add. Drivers 1 36.69" -> 36.69

    Format 3 - "Label amount<qty>" (pypdf strips space, qty glued to amount):
        "Add. Drivers 36.691" -> 36.69 (amount=36.69, qty=1)
        "Baby Seats 12.402" -> 12.40 (amount=12.40, qty=2)

    Format 4 - "Label amount" (single value, e.g. P.AI.):
        "P.AI. 16.13" -> 16.13
        "P.AI. 0" -> 0
    """
    # Try Format 1/2 first: two separated numbers after label
    pattern_two = re.escape(label) + r"\s+([0-9]+(?:[\.,][0-9]+)?)\s+([0-9]+(?:[\.,][0-9]+)?)"
    match = re.search(pattern_two, text, re.IGNORECASE)
    if match:
        n2 = _to_float(match.group(2))
        return n2

    # Format 3/4: single number after label (may have qty glued)
    pattern_one = re.escape(label) + r"\s+([0-9]+(?:[\.,][0-9]+)?)"
    match = re.search(pattern_one, text, re.IGNORECASE)
    if not match:
        return 0.0

    num_str = match.group(1).replace(",", ".")
    if "." in num_str:
        int_part, dec_part = num_str.split(".")
        # If decimal part has 3+ digits, the last digit(s) are the quantity glued.
        # Real amounts always have exactly 2 decimal places.
        # e.g. "36.691" = amount 36.69 + qty 1
        # e.g. "12.402" = amount 12.40 + qty 2
        if len(dec_part) >= 3:
            amount_str = f"{int_part}.{dec_part[:2]}"
            return _to_float(amount_str)
        return _to_float(num_str)

    return _to_float(num_str)


def parse_booking_pdf_text(text: str, mapping: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    mapping = mapping or load_mapping()

    lines = [_clean(line) for line in text.splitlines() if _clean(line)]
    normalized = "\n".join(lines)

    driver_name = lines[0] if lines else ""

    phone = regex_value(normalized, r"Born Date\s*0\s*\n([^\n]+)", "")
    if not re.search(r"\d", phone):
        phone = ""

    pickup_section = section_between(normalized, "PickUp Details", "Drop Off Details")
    dropoff_section = section_between(normalized, "Drop Off Details", "Product Code")

    pickup_dt = extract_section_datetime(pickup_section)
    dropoff_dt = extract_section_datetime(dropoff_section)

    pickup_station = extract_station(pickup_section)
    dropoff_station = extract_station(dropoff_section)

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
        regex_value(
            normalized,
            r"([0-9]{2}/[0-9]{2}/[0-9]{4}\s+[0-9]{2}:[0-9]{2})\s+Confirmed",
            ""
        )
    )

    reservation_number = regex_value(
        normalized,
        r"Reservation\s+(?:Last Change)?\s*([0-9]+)",
        ""
    )
    if not reservation_number:
        reservation_number = regex_value(
            normalized,
            r"Reservation\s+Last Change\s*([0-9]+)",
            ""
        )

    station_code_map = mapping.get("station_code_map", {})
    station_location_map = mapping.get("station_location_map", {})

    # Extras: smart extractor handles all PDF formats. Output amounts will get +24% VAT.
    baby_seats_amount = extract_extra_amount(normalized, "Baby Seats")
    # Child seat (CS) catches all "child-sized" seats from PDF: Child, Booster, Todler
    child_seats_amount = (
        extract_extra_amount(normalized, "Child Seats")
        + extract_extra_amount(normalized, "Booster Seats")
        + extract_extra_amount(normalized, "Todler Seats")
    )
    infant_seats_amount = extract_extra_amount(normalized, "Infant Seats")
    add_drivers_amount = extract_extra_amount(normalized, "Add. Drivers")

    # P.AI. uses single-amount format
    pai_amount = extract_extra_amount(normalized, "P.AI.")

    # Convert all extras to gross (with VAT)
    add_drivers_gross = add_vat_24(add_drivers_amount)
    baby_seats_gross = add_vat_24(baby_seats_amount)
    infant_seats_gross = add_vat_24(infant_seats_amount)
    child_seats_gross = add_vat_24(child_seats_amount)
    pai_gross = add_vat_24(pai_amount)

    extras_total = (
        add_drivers_gross
        + baby_seats_gross
        + infant_seats_gross
        + child_seats_gross
        + pai_gross
    )

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

        # Extras with 24% VAT (gross amounts)
        "add_drivers": add_drivers_gross,
        "baby_seats": baby_seats_gross,
        "infant_seats": infant_seats_gross,
        "child_seats": child_seats_gross,
        "pai_amount": pai_gross,

        # Sanity check: SUM of all extras should equal Pay On Arrival
        "extras_total": round(extras_total, 2),
        "extras_match_pay_on_arrival": abs(extras_total - pay_on_arrival) < 0.05,

        "del_fee": add_vat_24(extract_extra_amount(normalized, "Del Fee")),
        "col_fee": add_vat_24(extract_extra_amount(normalized, "Col Fee")),
        "one_way_fee": add_vat_24(extract_extra_amount(normalized, "One Way Fee")),
        "night_fee": add_vat_24(extract_extra_amount(normalized, "Night Fee")),

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


def write_many_to_template(
    parsed_list: list[Dict[str, Any]],
    template_file: Path | str,
    mapping: Dict[str, Any],
) -> bytes:
    """
    Writes multiple parsed bookings into a single Excel file,
    one booking per row, sharing the same template headers.
    """
    wb = load_workbook(template_file)
    ws = wb["Bookings"] if "Bookings" in wb.sheetnames else wb.active

    headers = [cell.value for cell in ws[1] if cell.value]

    # Clear any existing sample rows
    for row_idx in range(ws.max_row, 1, -1):
        ws.delete_rows(row_idx)

    # Write each booking on its own row, starting from row 2
    for i, parsed in enumerate(parsed_list):
        row_data = build_output_row(headers, parsed, mapping)
        row_idx = 2 + i
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row_idx, col_idx)
            cell.value = row_data.get(header, "")
            if isinstance(cell.value, datetime):
                cell.number_format = "DD/MM/YYYY HH:MM"

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.read()


def convert_many_pdfs_to_excel(
    pdf_files: list,
    template_file: Path | str = DEFAULT_TEMPLATE_FILE,
    mapping_file: Path | str = DEFAULT_MAPPING_FILE,
) -> tuple[bytes, list[Dict[str, Any]]]:
    """
    Convert multiple PDFs into a single combined Excel.
    pdf_files: list of (filename, file_bytes_or_path) tuples or BytesIO objects.
    Returns: (excel_bytes, list_of_parsed_dicts)
    """
    mapping = load_mapping(mapping_file)
    parsed_list = []

    for pdf_file in pdf_files:
        text = pdf_to_text(pdf_file)
        parsed = parse_booking_pdf_text(text, mapping)
        parsed_list.append(parsed)

    excel_bytes = write_many_to_template(parsed_list, template_file, mapping)
    return excel_bytes, parsed_list


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
