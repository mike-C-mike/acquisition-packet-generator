import json
import sys
from pathlib import Path


APP_NAME = "ByteCase Acquire"
APP_SUBTITLE = "Acquisition Packet Generator"
APP_VERSION = "0.9.1"

SUITE_NAME = "ByteCase"
PUBLISHER_NAME = "Forensics Byte"
PRODUCT_DOMAIN = "byte-case.com"
TOOL_FOLDER_NAME = "acquire"
DEFAULT_ROOT_FOLDER_NAME = "ByteCase"


def get_base_dir():
    """
    Return the correct working directory for local application files.

    When running from source, use the project folder. When running as a
    PyInstaller EXE, use the folder containing the EXE. This keeps local files
    such as settings.json next to the app instead of inside PyInstaller's
    temporary runtime folder.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent

    return Path(__file__).parent


BASE_DIR = get_base_dir()
SETTINGS_PATH = BASE_DIR / "settings.json"


DEFAULT_SETTINGS = {
    "suite_name": SUITE_NAME,
    "publisher": PUBLISHER_NAME,
    "department_name": "Example Police Department",
    "unit_name": "Digital Evidence Unit",
    "default_technician": "Example Technician",

    "appearance": {
        "theme": "system"
    },

    "output_paths": {
        "base_output_dir": "",
        "reports_folder_name": "reports",
        "saved_packets_folder_name": "saved_packets",
        "tracking_folder_name": "tracking",
        "tracking_workbook_name": "fpr_tracking.xlsx"
    },

    "report_branding": {
        "patch_image_path": ""
    },

    "common_technicians": [
        "Example Technician"
    ],

    "common_investigators": [
        "Officer Example",
        "Detective Example"
    ],

    "common_tools": [
        {"name": "Cellebrite UFED", "version": ""},
        {"name": "Cellebrite Physical Analyzer", "version": ""},
        {"name": "Magnet AXIOM", "version": ""},
        {"name": "FTK Imager", "version": ""},
        {"name": "GrayKey", "version": ""}
    ],

    "common_evidence_locations": [
        "Evidence Locker",
        "Property Room",
        "Digital Evidence Storage"
    ],

    "default_scope_statement": (
        "This documentation records the technical acquisition, extraction, imaging, "
        "and/or processing steps performed on the listed device or media. The technician "
        "did not conduct a content examination, investigative review, interpretation of user "
        "activity, or evidentiary analysis unless specifically stated. Any review, "
        "interpretation, or investigative conclusions are the responsibility of the assigned "
        "officer, investigator, or case agent."
    )
}


def merge_settings(defaults, current):
    merged = defaults.copy()

    for key, value in current.items():
        if isinstance(value, dict) and key in merged and isinstance(merged[key], dict):
            merged[key] = merge_settings(merged[key], value)
        else:
            merged[key] = value

    return merged


def normalize_settings(settings):
    settings["suite_name"] = str(settings.get("suite_name", SUITE_NAME)).strip() or SUITE_NAME
    settings["publisher"] = str(settings.get("publisher", PUBLISHER_NAME)).strip() or PUBLISHER_NAME
    settings["department_name"] = str(settings.get("department_name", "")).strip()
    settings["unit_name"] = str(settings.get("unit_name", "")).strip()
    settings["default_technician"] = str(settings.get("default_technician", "")).strip()

    appearance = settings.get("appearance", {})
    if not isinstance(appearance, dict):
        appearance = {}

    theme = str(appearance.get("theme", "system")).strip().lower()
    if theme in {"system default", "default", "os", "os default"}:
        theme = "system"
    if theme not in {"dark", "light", "system"}:
        theme = "system"
    settings["appearance"] = {"theme": theme}

    output_paths = settings.get("output_paths", {})
    if not isinstance(output_paths, dict):
        output_paths = {}

    # Preserve old settings.json values while moving the defaults to ByteCase's
    # case-folder structure.
    settings["output_paths"] = {
        "base_output_dir": str(output_paths.get("base_output_dir", "")).strip(),
        "reports_folder_name": str(output_paths.get("reports_folder_name", "reports")).strip() or "reports",
        "saved_packets_folder_name": str(output_paths.get("saved_packets_folder_name", "saved_packets")).strip() or "saved_packets",
        "tracking_folder_name": str(output_paths.get("tracking_folder_name", "tracking")).strip() or "tracking",
        "tracking_workbook_name": str(output_paths.get("tracking_workbook_name", "fpr_tracking.xlsx")).strip() or "fpr_tracking.xlsx",
    }

    report_branding = settings.get("report_branding", {})
    if not isinstance(report_branding, dict):
        report_branding = {}
    settings["report_branding"] = {
        "patch_image_path": str(report_branding.get("patch_image_path", "")).strip()
    }

    for list_key in ["common_technicians", "common_investigators", "common_evidence_locations"]:
        values = settings.get(list_key, [])
        if not isinstance(values, list):
            values = []
        cleaned = []
        seen = set()
        for value in values:
            text = str(value).strip()
            key = text.lower()
            if text and key not in seen:
                cleaned.append(text)
                seen.add(key)
        settings[list_key] = cleaned

    tools = settings.get("common_tools", [])
    if not isinstance(tools, list):
        tools = []
    cleaned_tools = []
    seen_tools = set()
    for tool in tools:
        if isinstance(tool, dict):
            name = str(tool.get("name", "")).strip()
            version = str(tool.get("version", "")).strip()
        else:
            name = str(tool).strip()
            version = ""
        key = (name.lower(), version.lower())
        if name and key not in seen_tools:
            cleaned_tools.append({"name": name, "version": version})
            seen_tools.add(key)
    settings["common_tools"] = cleaned_tools

    settings["default_scope_statement"] = str(settings.get("default_scope_statement", "")).strip()
    if not settings["default_scope_statement"]:
        settings["default_scope_statement"] = DEFAULT_SETTINGS["default_scope_statement"]

    return settings


def load_or_create_settings():
    if not SETTINGS_PATH.exists():
        settings = normalize_settings(DEFAULT_SETTINGS.copy())
        SETTINGS_PATH.write_text(json.dumps(settings, indent=4), encoding="utf-8")
        return settings

    with SETTINGS_PATH.open("r", encoding="utf-8") as file:
        current_settings = json.load(file)

    merged_settings = normalize_settings(merge_settings(DEFAULT_SETTINGS, current_settings))
    SETTINGS_PATH.write_text(json.dumps(merged_settings, indent=4), encoding="utf-8")
    return merged_settings


def save_settings(settings):
    settings = normalize_settings(settings)
    SETTINGS_PATH.write_text(json.dumps(settings, indent=4), encoding="utf-8")


def safe_case_folder_name(case_number):
    value = str(case_number or "").strip() or "NO_CASE"

    for char in '<>:"/\\|?*':
        value = value.replace(char, "_")

    value = value.replace(" ", "_")
    while "__" in value:
        value = value.replace("__", "_")

    return value.strip("_") or "NO_CASE"


def get_default_output_root():
    return Path.home() / DEFAULT_ROOT_FOLDER_NAME


def get_output_root(settings=None):
    if settings is None:
        settings = load_or_create_settings()

    output_settings = settings.get("output_paths", {})
    base_output_dir = str(output_settings.get("base_output_dir", "")).strip()

    if base_output_dir:
        return Path(base_output_dir)

    return get_default_output_root()


def get_output_paths(settings=None, case_number=None):
    """
    Return ByteCase Acquire output paths.

    Blank output root:
        C:\\Users\\<user>\\ByteCase\\<case_number>\\acquire\\

    Custom output root:
        <custom_root>\\<case_number>\\acquire\\
    """
    if settings is None:
        settings = load_or_create_settings()

    output_settings = settings.get("output_paths", {})
    root_path = get_output_root(settings)

    if case_number:
        base_path = root_path / safe_case_folder_name(case_number) / TOOL_FOLDER_NAME
    else:
        base_path = root_path

    reports_folder_name = output_settings.get("reports_folder_name", "reports")
    saved_packets_folder_name = output_settings.get("saved_packets_folder_name", "saved_packets")
    tracking_folder_name = output_settings.get("tracking_folder_name", "tracking")
    tracking_workbook_name = output_settings.get("tracking_workbook_name", "fpr_tracking.xlsx")

    reports_dir = base_path / reports_folder_name
    saved_packets_dir = base_path / saved_packets_folder_name
    tracking_dir = base_path / tracking_folder_name
    tracking_workbook_path = tracking_dir / tracking_workbook_name

    return {
        "root_path": root_path,
        "base_path": base_path,
        "reports_dir": reports_dir,
        "saved_packets_dir": saved_packets_dir,
        "tracking_dir": tracking_dir,
        "tracking_workbook_path": tracking_workbook_path
    }


def ensure_directories(settings=None, case_number=None):
    paths = get_output_paths(settings, case_number=case_number)

    paths["root_path"].mkdir(parents=True, exist_ok=True)
    paths["base_path"].mkdir(parents=True, exist_ok=True)
    paths["reports_dir"].mkdir(parents=True, exist_ok=True)
    paths["saved_packets_dir"].mkdir(parents=True, exist_ok=True)
    paths["tracking_dir"].mkdir(parents=True, exist_ok=True)

    return paths


def get_settings_path():
    return SETTINGS_PATH


def get_base_path():
    return BASE_DIR