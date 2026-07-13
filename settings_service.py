import json
import sys
from pathlib import Path


APP_NAME = "Acquisition Packet Generator"
APP_VERSION = "0.6"


def get_base_dir():
    """
    Return the correct working directory.

    When running from source, use the project folder.
    When running as a PyInstaller EXE, use the folder containing the EXE.

    This keeps generated local files next to the EXE after packaging,
    instead of writing them into PyInstaller's temporary runtime folder.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent

    return Path(__file__).parent


BASE_DIR = get_base_dir()
SETTINGS_PATH = BASE_DIR / "settings.json"


DEFAULT_SETTINGS = {
    "department_name": "Example Police Department",
    "unit_name": "Digital Evidence Unit",
    "default_technician": "Example Technician",

    "output_paths": {
        "base_output_dir": "",
        "reports_folder_name": "output",
        "saved_packets_folder_name": "saved_packets",
        "tracking_workbook_name": "fpr_tracking.xlsx"
    },

    "common_technicians": [
        "Example Technician"
    ],

    "common_investigators": [
        "Officer Example",
        "Detective Example"
    ],

    "common_tools": [
        {
            "name": "Cellebrite UFED",
            "version": ""
        },
        {
            "name": "Cellebrite Physical Analyzer",
            "version": ""
        },
        {
            "name": "Magnet AXIOM",
            "version": ""
        },
        {
            "name": "FTK Imager",
            "version": ""
        },
        {
            "name": "GrayKey",
            "version": ""
        }
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
    """
    Merge existing settings with default settings.

    This lets the app add new default settings in future versions without
    overwriting a user's existing local settings.

    Example:
        Existing settings.json from v0.3 may not have output_paths.
        DEFAULT_SETTINGS from v0.4 does.
        This function adds output_paths while preserving the user's
        existing department name, technicians, tools, etc.
    """
    merged = defaults.copy()

    for key, value in current.items():
        if (
            isinstance(value, dict)
            and key in merged
            and isinstance(merged[key], dict)
        ):
            merged[key] = merge_settings(merged[key], value)
        else:
            merged[key] = value

    return merged


def load_or_create_settings():
    """
    Load local settings.json.

    If settings.json does not exist, create it from DEFAULT_SETTINGS.

    If settings.json does exist, merge in any new default keys and then
    save the merged settings back to disk.

    settings.json should not be committed to GitHub because it may contain
    local agency names, technician names, investigator names, evidence
    locations, or custom output paths.
    """
    if not SETTINGS_PATH.exists():
        SETTINGS_PATH.write_text(
            json.dumps(DEFAULT_SETTINGS, indent=4),
            encoding="utf-8"
        )
        return DEFAULT_SETTINGS.copy()

    with SETTINGS_PATH.open("r", encoding="utf-8") as file:
        current_settings = json.load(file)

    merged_settings = merge_settings(DEFAULT_SETTINGS, current_settings)

    SETTINGS_PATH.write_text(
        json.dumps(merged_settings, indent=4),
        encoding="utf-8"
    )

    return merged_settings


def save_settings(settings):
    """
    Save updated settings to settings.json.
    """
    SETTINGS_PATH.write_text(
        json.dumps(settings, indent=4),
        encoding="utf-8"
    )


def get_output_paths(settings=None):
    """
    Return output paths based on settings.

    If no custom base output directory is set, use BASE_DIR.

    Returns:
        dict with:
            base_path
            reports_dir
            saved_packets_dir
            tracking_workbook_path
    """
    if settings is None:
        settings = load_or_create_settings()

    output_settings = settings.get("output_paths", {})

    base_output_dir = output_settings.get("base_output_dir", "").strip()

    if base_output_dir:
        base_path = Path(base_output_dir)
    else:
        base_path = BASE_DIR

    reports_folder_name = output_settings.get("reports_folder_name", "output")
    saved_packets_folder_name = output_settings.get(
        "saved_packets_folder_name",
        "saved_packets"
    )
    tracking_workbook_name = output_settings.get(
        "tracking_workbook_name",
        "fpr_tracking.xlsx"
    )

    reports_dir = base_path / reports_folder_name
    saved_packets_dir = base_path / saved_packets_folder_name
    tracking_workbook_path = base_path / tracking_workbook_name

    return {
        "base_path": base_path,
        "reports_dir": reports_dir,
        "saved_packets_dir": saved_packets_dir,
        "tracking_workbook_path": tracking_workbook_path
    }


def ensure_directories(settings=None):
    """
    Create configured output folders if they do not already exist.

    This does not create the XLSX workbook itself. It only ensures the
    configured base folder, reports folder, and saved packets folder exist.
    """
    paths = get_output_paths(settings)

    paths["base_path"].mkdir(parents=True, exist_ok=True)
    paths["reports_dir"].mkdir(parents=True, exist_ok=True)
    paths["saved_packets_dir"].mkdir(parents=True, exist_ok=True)


def get_settings_path():
    """
    Return the path to settings.json.

    Useful for future About/Diagnostics screens.
    """
    return SETTINGS_PATH


def get_base_path():
    """
    Return the application base path.

    When running from source, this is the project folder.
    When running as an EXE, this is the folder containing the EXE.
    """
    return BASE_DIR