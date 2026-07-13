from datetime import datetime


DATE_FORMAT = "%Y-%m-%d"


def is_valid_date(value):
    """
    Return True when value is blank or formatted as YYYY-MM-DD.

    Blank values are allowed because not every field is required yet.
    """
    if not value:
        return True

    try:
        datetime.strptime(value, DATE_FORMAT)
        return True
    except ValueError:
        return False


def compare_dates(start_value, end_value):
    """
    Compare two YYYY-MM-DD date strings.

    Returns:
        None if either value is blank or invalid.
        -1 if start < end.
         0 if start == end.
         1 if start > end.
    """
    if not start_value or not end_value:
        return None

    if not is_valid_date(start_value) or not is_valid_date(end_value):
        return None

    start_date = datetime.strptime(start_value, DATE_FORMAT)
    end_date = datetime.strptime(end_value, DATE_FORMAT)

    if start_date < end_date:
        return -1

    if start_date == end_date:
        return 0

    return 1


def validate_packet(packet):
    """
    Validate packet data before review/export.

    Returns:
        tuple: (errors, warnings)

    Errors should block export.
    Warnings should be shown to the user but allow export.
    """
    errors = []
    warnings = []

    general = packet.get("general_info", {})
    subject = packet.get("subject", {})
    processing = packet.get("processing", {})
    output = packet.get("output", {})
    devices = packet.get("devices", [])
    tools = packet.get("tools_used", [])

    case_number = general.get("case_number", "").strip()
    technician = general.get("technician", "").strip()

    if not case_number:
        errors.append("Case Number is required.")

    if not technician:
        errors.append("Technician is recommended before export.")

    if not devices:
        errors.append("At least one device/media entry is required.")

    if not tools:
        errors.append("At least one tool entry is required.")

    date_fields = [
        ("Date Received", general.get("date_received", "")),
        ("Date Processed", general.get("date_processed", "")),
        ("Exam Start Date", processing.get("exam_start_date", "")),
        ("Exam End Date", processing.get("exam_end_date", "")),
    ]

    for label, value in date_fields:
        if value and not is_valid_date(value):
            errors.append(f"{label} must use YYYY-MM-DD format.")

    received_vs_processed = compare_dates(
        general.get("date_received", ""),
        general.get("date_processed", "")
    )

    if received_vs_processed == 1:
        warnings.append("Date Processed is earlier than Date Received.")

    exam_start_vs_end = compare_dates(
        processing.get("exam_start_date", ""),
        processing.get("exam_end_date", "")
    )

    if exam_start_vs_end == 1:
        warnings.append("Exam End Date is earlier than Exam Start Date.")

    if not subject.get("last_name", "").strip() and not subject.get("first_name", "").strip():
        warnings.append("No subject name was entered.")

    if not general.get("offense_or_incident", "").strip():
        warnings.append("Offense / Incident Type is blank.")

    if not output.get("output_filename", "").strip():
        warnings.append("Output Filename / Identifier is blank.")

    for index, device in enumerate(devices, start=1):
        device_label = f"Device Entry {index}"

        if not device.get("device_type", "").strip():
            errors.append(f"{device_label}: Device Type is blank.")

        if device.get("quantity", 0) < 1:
            errors.append(f"{device_label}: Quantity must be greater than zero.")

        if device.get("decrypted", False) and not device.get("encrypted", False):
            warnings.append(f"{device_label}: Decrypted is checked but Encrypted is not checked.")

        if device.get("password_unlocked", False) and not device.get("password_locked", False):
            warnings.append(f"{device_label}: Password Unlocked is checked but Password Locked is not checked.")

        if device.get("tools_used_to_decrypt", "").strip() and not device.get("decrypted", False):
            warnings.append(f"{device_label}: Tools Used to Decrypt is filled in but Decrypted is not checked.")

        if device.get("services_used_to_unlock", "").strip() and not device.get("password_unlocked", False):
            warnings.append(f"{device_label}: Services Used to Unlock is filled in but Password Unlocked is not checked.")

    return errors, warnings