import json

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from settings_service import (
    APP_NAME,
    APP_VERSION,
    BASE_DIR,
    SETTINGS_PATH,
    DEFAULT_SETTINGS,
    get_output_paths,
    ensure_directories,
    load_or_create_settings,
    save_settings,
)


DEVICE_TYPES = [
    "CPU",
    "ETech",
    "Loose Drive",
    "Mobile Phone",
    "Tablet",
    "USB Drive",
    "SD Card",
    "DVR/NVR Storage",
    "Other"
]


PROCESSING_TYPES = [
    "Mobile Device Extraction",
    "Drive / Media Imaging",
    "Reader Report Generation",
    "Case File Generation",
    "Extraction Export",
    "Other"
]


PROCESSING_STATUSES = [
    "Completed",
    "Completed with Limitations",
    "Partial Extraction / Image",
    "Unable to Process",
    "Cancelled by Requestor",
    "Pending"
]


OUTPUT_TYPES = [
    "Reader Report",
    "Case File",
    "Forensic Image",
    "Extraction Export",
    "Other"
]


STORAGE_UNITS = [
    "GB",
    "TB",
    "MB",
    "Unknown"
]


def safe_filename(value):
    """
    Convert a case number or other text value into a safer filename.
    """
    unsafe_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|']
    safe_value = str(value).strip()

    for char in unsafe_chars:
        safe_value = safe_value.replace(char, "-")

    return safe_value.replace(" ", "_")


def convert_to_gb(size, unit):
    """
    Convert a storage size to GB.

    Returns None when size is unknown or cannot be converted.
    """
    if size is None or unit == "Unknown":
        return None

    if unit == "GB":
        return size

    if unit == "TB":
        return size * 1024

    if unit == "MB":
        return size / 1024

    return None


def format_storage_gb(value):
    """
    Format a GB value for report output.
    """
    if value is None:
        return "Unknown"

    if value >= 1024:
        return f"{value:,.2f} GB ({value / 1024:,.2f} TB)"

    return f"{value:,.2f} GB"


def count_devices_by_type(devices):
    """
    Count total device quantities by device type.
    """
    counts = {}

    for device in devices:
        device_type = device.get("device_type", "Other")
        quantity = device.get("quantity", 0)

        counts[device_type] = counts.get(device_type, 0) + quantity

    return counts


def calculate_total_storage_gb(devices):
    """
    Calculate total known storage across all device entries.

    Unknown storage entries are ignored.
    Returns None if no devices have known storage.
    """
    total = 0
    has_known_storage = False

    for device in devices:
        storage_total_gb = device.get("storage_total_gb")

        if storage_total_gb is not None:
            total += storage_total_gb
            has_known_storage = True

    if not has_known_storage:
        return None

    return total


def add_summary_to_packet(packet):
    """
    Add or refresh the summary section of a packet.

    This is useful because both the terminal app and GUI app
    can build packet dictionaries, then call this function before export.
    """
    devices = packet.get("devices", [])

    packet["summary"] = {
        "device_counts": count_devices_by_type(devices),
        "total_devices": sum(device.get("quantity", 0) for device in devices),
        "total_storage_gb": calculate_total_storage_gb(devices)
    }

    return packet


def build_txt_report(packet):
    """
    Build the plain text acquisition packet report.
    """
    general = packet.get("general_info", {})
    intake = packet.get("intake_info", {})
    subject = packet.get("subject", {})
    processing = packet.get("processing", {})
    output = packet.get("output", {})
    summary = packet.get("summary", {})

    department = packet.get("department", {})
    department_name = department.get("department_name", "")
    unit_name = department.get("unit_name", "")

    lines = []

    lines.append("DIGITAL EVIDENCE ACQUISITION DOCUMENTATION")
    lines.append("=" * 55)

    if department_name:
        lines.append(department_name)

    if unit_name:
        lines.append(unit_name)

    lines.append("")

    lines.append("Administrative Information")
    lines.append("-" * 30)
    lines.append(f"Case Number: {general.get('case_number', '')}")
    lines.append(f"Agency Case Number: {general.get('agency_case_number', '')}")
    lines.append(f"Offense / Incident: {general.get('offense_or_incident', '')}")
    lines.append(f"Subject: {subject.get('last_name', '')}, {subject.get('first_name', '')}")
    lines.append(f"Requesting Officer / Investigator: {general.get('requesting_investigator', '')}")
    lines.append(f"Technician: {general.get('technician', '')}")
    lines.append(f"Date Received: {general.get('date_received', '')}")
    lines.append(f"Time Received: {general.get('time_received', '')}")
    lines.append(f"Date Processed: {general.get('date_processed', '')}")
    lines.append(f"Time Processed: {general.get('time_processed', '')}")
    lines.append("")

    lines.append("Intake Information")
    lines.append("-" * 30)
    lines.append(f"Drop-off Person: {intake.get('dropoff_person', '')}")
    lines.append(f"Received From: {intake.get('received_from', '')}")
    lines.append(f"Evidence Item Number: {intake.get('evidence_item_number', '')}")
    lines.append(f"Checked Out From Evidence: {intake.get('checked_out_from_evidence', '')}")
    lines.append(f"Checked Out Date/Time: {intake.get('checked_out_datetime', '')}")
    lines.append(f"Returned To Evidence: {intake.get('returned_to_evidence', '')}")
    lines.append(f"Returned Date/Time: {intake.get('returned_datetime', '')}")
    lines.append(f"Evidence Location / Locker: {intake.get('evidence_location', '')}")
    lines.append("")

    lines.append("Device / Media Summary")
    lines.append("-" * 30)
    lines.append(f"Total Devices / Media Count: {summary.get('total_devices', 0)}")
    lines.append(f"Total Known Storage: {format_storage_gb(summary.get('total_storage_gb'))}")
    lines.append("")

    device_counts = summary.get("device_counts", {})

    if device_counts:
        for device_type, count in device_counts.items():
            lines.append(f"{device_type}: {count}")
    else:
        lines.append("No device/media entries recorded.")

    lines.append("")
    lines.append("Device / Media Detail")
    lines.append("-" * 30)

    devices = packet.get("devices", [])

    if devices:
        for index, device in enumerate(devices, start=1):
            lines.append(f"Device Entry {index}")
            lines.append(f"  Type: {device.get('device_type', '')}")
            lines.append(f"  Quantity: {device.get('quantity', '')}")
            lines.append(f"  Description: {device.get('description', '')}")
            lines.append(f"  Make: {device.get('make', '')}")
            lines.append(f"  Model: {device.get('model', '')}")
            lines.append(f"  Serial / Identifier: {device.get('serial', '')}")

            size = device.get("capacity_size")
            unit = device.get("capacity_unit")

            if size is None:
                lines.append("  Storage Per Device: Unknown")
            else:
                lines.append(f"  Storage Per Device: {size} {unit}")

            lines.append(f"  Total Known Storage: {format_storage_gb(device.get('storage_total_gb'))}")
            lines.append("")
    else:
        lines.append("No device/media entries recorded.")
        lines.append("")

    lines.append("Tools Used")
    lines.append("-" * 30)

    tools_used = packet.get("tools_used", [])

    if tools_used:
        for tool in tools_used:
            name = tool.get("name", "")
            version = tool.get("version", "")
            lines.append(f"{name} | Version: {version}")
    else:
        lines.append("No tools entered.")

    lines.append("")

    lines.append("Processing Information")
    lines.append("-" * 30)
    lines.append(f"Processing Type: {processing.get('processing_type', '')}")
    lines.append(f"Processing Status: {processing.get('processing_status', '')}")
    lines.append(f"Processing Notes: {processing.get('processing_notes', '')}")
    lines.append("")

    lines.append("Generated Output")
    lines.append("-" * 30)
    lines.append(f"Output Type: {output.get('output_type', '')}")
    lines.append(f"Output Filename / Identifier: {output.get('output_filename', '')}")
    lines.append(f"Output Location: {output.get('output_location', '')}")
    lines.append(f"Reader Report Generated: {output.get('reader_report_generated', '')}")
    lines.append(f"Case File Generated: {output.get('case_file_generated', '')}")
    lines.append("")

    lines.append("Technician Notes")
    lines.append("-" * 30)
    lines.append(packet.get("technician_notes", ""))
    lines.append("")

    lines.append("Limitations and Scope")
    lines.append("-" * 30)
    lines.append(packet.get("scope_statement", ""))
    lines.append("")

    lines.append("Technician Statement")
    lines.append("-" * 30)
    lines.append(
        "I documented the above process based on the actions performed, tool output, "
        "and information available at the time of processing."
    )
    lines.append("")

    lines.append("Report Generated")
    lines.append("-" * 30)
    lines.append(f"Generated On: {packet.get('created_at', '')}")
    lines.append(f"Generated By: {APP_NAME} v{APP_VERSION}")

    return "\n".join(lines)


def save_packet_outputs(packet, settings=None):
    """
    Save the acquisition packet as TXT and JSON.

    Filenames include case number, a case title/offense value, and a timestamp
    so repeated exports do not overwrite earlier packet files.

    Returns:
        tuple: (txt_path, json_path)
    """
    ensure_directories(settings)
    paths = get_output_paths(settings)

    packet = add_summary_to_packet(packet)

    general = packet.get("general_info", {})
    report_info = packet.get("report_info", {})

    case_number = general.get("case_number", "UNKNOWN_CASE")

    case_title = (
        general.get("offense_or_incident", "")
        or report_info.get("case_type", "")
        or general.get("case_type", "")
        or "Untitled_Case"
    )

    timestamp = packet.get("created_at", "")
    if not timestamp:
        timestamp = "unknown_time"

    safe_case = safe_filename(case_number)
    safe_title = safe_filename(case_title)
    safe_timestamp = safe_filename(timestamp)

    base_filename = f"{safe_case}_{safe_title}_{safe_timestamp}_acquisition_packet"

    txt_path = paths["reports_dir"] / f"{base_filename}.txt"
    json_path = paths["saved_packets_dir"] / f"{base_filename}.json"

    report_text = build_txt_report(packet)

    txt_path.write_text(report_text, encoding="utf-8")

    json_path.write_text(
        json.dumps(packet, indent=4),
        encoding="utf-8"
    )

    return txt_path, json_path

def style_header_row(sheet):
    """
    Apply basic styling to the first row of an XLSX sheet.
    """
    fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    font = Font(color="FFFFFF", bold=True)

    for cell in sheet[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center")


def autofit_columns(sheet):
    """
    Set reasonable column widths based on cell contents.
    """
    for column_cells in sheet.columns:
        max_length = 0
        column_letter = get_column_letter(column_cells[0].column)

        for cell in column_cells:
            value = cell.value
            if value is not None:
                max_length = max(max_length, len(str(value)))

        sheet.column_dimensions[column_letter].width = min(max_length + 2, 40)


def get_or_create_workbook(settings=None):
    """
    Open the existing FPR tracking workbook or create a new one.
    """
    paths = get_output_paths(settings)
    tracking_path = paths["tracking_workbook_path"]

    if tracking_path.exists():
        workbook = load_workbook(tracking_path)
    else:
        workbook = Workbook()
        default_sheet = workbook.active
        workbook.remove(default_sheet)

    if "Case Summary" not in workbook.sheetnames:
        sheet = workbook.create_sheet("Case Summary")
        sheet.append([
            "Case Number",
            "Agency Case Number",
            "Subject Last Name",
            "Subject First Name",
            "Offense / Incident",
            "Requesting Investigator",
            "Technician",
            "Date Processed",
            "CPU Count",
            "ETech Count",
            "Loose Drive Count",
            "Mobile Phone Count",
            "Tablet Count",
            "USB Drive Count",
            "SD Card Count",
            "DVR/NVR Storage Count",
            "Other Count",
            "Total Devices",
            "Total Storage GB",
            "Processing Type",
            "Processing Status",
            "Output Type",
            "Report Generated"
        ])
        style_header_row(sheet)

    if "Device Detail" not in workbook.sheetnames:
        sheet = workbook.create_sheet("Device Detail")
        sheet.append([
            "Case Number",
            "Subject Last Name",
            "Subject First Name",
            "Device Type",
            "Quantity",
            "Description",
            "Make",
            "Model",
            "Serial / Identifier",
            "Storage Size Per Device",
            "Storage Unit",
            "Storage GB Per Device",
            "Storage GB Total",
            "Date Processed",
            "Technician"
        ])
        style_header_row(sheet)

    if "FPR Case Info" not in workbook.sheetnames:
        sheet = workbook.create_sheet("FPR Case Info")
        sheet.append([
            "Case Number",
            "Agency Case Number",
            "State / Local Case No.",
            "Subject Last Name",
            "Subject First Name",
            "Case Type",
            "Offense / Incident",
            "City of Offense",
            "State of Offense",
            "Country of Offense",
            "Exam Start Date",
            "Exam End Date",
            "Other Data Analyzed",
            "Case Summary",
            "Requesting Investigator",
            "Technician",
            "Date Processed",
            "Report Generated"
        ])
        style_header_row(sheet)

    return workbook


def append_to_fpr_tracking(packet, settings=None):
    """
    Append acquisition packet summary data to fpr_tracking.xlsx.

    The workbook contains:
    - Case Summary
    - Device Detail

    Returns:
        Path: path to fpr_tracking.xlsx
    """
    packet = add_summary_to_packet(packet)

    ensure_directories(settings)
    paths = get_output_paths(settings)
    workbook = get_or_create_workbook(settings)

    general = packet.get("general_info", {})
    subject = packet.get("subject", {})
    processing = packet.get("processing", {})
    output = packet.get("output", {})
    report_info = packet.get("report_info", {})
    summary = packet.get("summary", {})
    counts = summary.get("device_counts", {})

    case_summary = workbook["Case Summary"]

    known_summary_types = [
        "CPU",
        "ETech",
        "Loose Drive",
        "Mobile Phone",
        "Tablet",
        "USB Drive",
        "SD Card",
        "DVR/NVR Storage"
    ]

    other_count = 0

    for device_type, count in counts.items():
        if device_type not in known_summary_types:
            other_count += count

    case_summary.append([
        general.get("case_number", ""),
        general.get("agency_case_number", ""),
        subject.get("last_name", ""),
        subject.get("first_name", ""),
        general.get("offense_or_incident", ""),
        general.get("requesting_investigator", ""),
        general.get("technician", ""),
        general.get("date_processed", ""),
        counts.get("CPU", 0),
        counts.get("ETech", 0),
        counts.get("Loose Drive", 0),
        counts.get("Mobile Phone", 0),
        counts.get("Tablet", 0),
        counts.get("USB Drive", 0),
        counts.get("SD Card", 0),
        counts.get("DVR/NVR Storage", 0),
        other_count,
        summary.get("total_devices", 0),
        summary.get("total_storage_gb"),
        processing.get("processing_type", ""),
        processing.get("processing_status", ""),
        output.get("output_type", ""),
        packet.get("created_at", "")
    ])

    device_detail = workbook["Device Detail"]

    for device in packet.get("devices", []):
        device_detail.append([
            general.get("case_number", ""),
            subject.get("last_name", ""),
            subject.get("first_name", ""),
            device.get("device_type", ""),
            device.get("quantity", ""),
            device.get("description", ""),
            device.get("make", ""),
            device.get("model", ""),
            device.get("serial", ""),
            device.get("capacity_size"),
            device.get("capacity_unit", ""),
            device.get("storage_each_gb"),
            device.get("storage_total_gb"),
            general.get("date_processed", ""),
            general.get("technician", "")
        ])

    for device in packet.get("devices", []):
        device_detail.append([
            general.get("case_number", ""),
            subject.get("last_name", ""),
            subject.get("first_name", ""),
            device.get("device_type", ""),
            device.get("quantity", ""),
            device.get("description", ""),
            device.get("make", ""),
            device.get("model", ""),
            device.get("serial", ""),
            device.get("capacity_size"),
            device.get("capacity_unit", ""),
            device.get("storage_each_gb"),
            device.get("storage_total_gb"),
            general.get("date_processed", ""),
            general.get("technician", "")
        ])

    for sheet in workbook.worksheets:
        autofit_columns(sheet)

    try:
        workbook.save(paths["tracking_workbook_path"])
    except PermissionError as error:
        raise PermissionError(
            "Unable to update the XLSX tracking workbook.\n\n"
            "The tracking workbook may already be open in Excel or locked by another program.\n\n"
            "Close fpr_tracking.xlsx and try generating the packet again."
        ) from error

    return paths["tracking_workbook_path"]