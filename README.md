# Acquisition Packet Generator

**Acquisition Packet Generator** is a Python tool for documenting digital evidence intake, mobile extractions, drive/media imaging, generated outputs, and related acquisition packet information.

The goal of this project is to help digital evidence technicians create consistent acquisition documentation while also collecting structured data that can support internal reporting, device tracking, and FPR-style reporting workflows.

This project is currently a GUI-based Windows-focused release candidate with a packaged Windows executable release available through GitHub Releases.

---

## Current Status

**Current version:** v0.9.0  
**Latest packaged release:** v0.9.0  
**Stage:** Release candidate / pre-v1.0 polish  
**Interface:** Tkinter GUI  
**Primary outputs:** TXT acquisition packet, JSON saved packet, DOCX report, XLSX tracking workbook

The project began as a terminal-based proof of concept and has moved into a GUI workflow. The current version includes settings management, configurable output paths, multi-device entries, multi-tool entries, device-level FPR media details, DOCX report generation, optional report branding, review-before-export, validation warnings, build documentation, dependency documentation, and an open output folder option.

Version v0.9.0 is intended to serve as the final release candidate before the first stable public release. The main focus of v0.9.0 is not adding large new features, but hardening the tool, documenting the build process, validating the workflow, and preparing the project for a cleaner v1.0.0 release.

---

## Quick Start for Windows Users

The easiest way to use Acquisition Packet Generator is to download the packaged Windows release.

1. Open the GitHub Releases page.
2. Download the Windows ZIP file for the latest release.
3. Right-click the ZIP file and choose **Extract All**.
4. Open the extracted folder.
5. Double-click:

```text
AcquisitionPacketGenerator.exe
```

6. Use the **Settings** button to configure output folders and local defaults if needed.
7. Complete the intake form.
8. Add one or more devices/media items on the Device tab.
9. Add one or more tools on the Tools tab.
10. Complete the Processing / Output tab.
11. Click **Review Packet**.
12. Review the packet summary and any warnings.
13. Click **Confirm Export** to generate the packet files.
14. Use **Open Output Folder** to view generated files.

The app should be extracted from the ZIP before running. Do not run the executable directly from inside the ZIP preview window.

Windows may show a SmartScreen or antivirus warning because this is an unsigned open-source prototype. The source code is available in the GitHub repository for review.

---

## Basic Operating Workflow

The intended workflow is:

1. Configure Settings if needed.
2. Enter case and subject information.
3. Enter intake and evidence handling information.
4. Add one or more devices or media items.
5. Add one or more tools used during acquisition, extraction, imaging, or output generation.
6. Enter processing and output details.
7. Review the packet before export.
8. Confirm export.
9. Review generated TXT, JSON, DOCX, and XLSX output.

The review step is intentional. It allows the user to confirm the case information, device/media entries, tools used, derived media values, and validation warnings before files are written or the XLSX tracking workbook is updated.

The tool is designed to document what was received, what was processed, what tools were used, what output was generated, and what structured tracking information should be retained. It does not perform forensic analysis or generate investigative conclusions.

---

## Generated Local Files and Folders

When the executable runs, it creates local files and folders next to the executable unless a custom output path is configured in Settings.

Default generated local items:

```text
settings.json
output/
saved_packets/
fpr_tracking.xlsx
```

The `output/` folder stores TXT and DOCX reports.

The `saved_packets/` folder stores structured JSON packet data.

The `fpr_tracking.xlsx` workbook stores structured XLSX tracking rows.

These files are intentionally local. They may contain agency-specific, case-specific, or sensitive data and should not be committed to the public repository.

---

## Intended Use

This tool is designed to document acquisition-related tasks such as:

- Intake of a device or storage media
- Drop-off and evidence handling information
- Requesting officer or investigator information
- Technician information
- Device/media type and storage size
- Device-level media examined information
- Tool(s) used during processing
- Processing status
- Output generated, such as reader reports, case files, exports, or forensic images
- Return of the original device or media to evidence storage
- Structured reporting data for XLSX tracking
- FPR-style case and media examined tracking
- DOCX acquisition report generation

The tool is intended to support documentation consistency and reduce repetitive reporting work while keeping the technician in control of the content.

The application does not attempt to replace forensic tools such as Cellebrite, Magnet AXIOM, FTK Imager, EnCase, GrayKey, or similar acquisition/examination platforms. Instead, it is designed to sit around those tools as a documentation and reporting aid.

The core idea is that a technician can document what was received, what was processed, what tool was used, what output was generated, and what structured tracking information should be retained for later reporting.

---

## Not Intended For

This tool is **not** intended to perform forensic analysis or generate investigative conclusions.

It should not be used to document interpretive findings such as:

- Timeline analysis
- Malware findings
- Documents of interest
- User activity conclusions
- USB history interpretation
- Evidence relevance
- Investigative opinions
- Legal conclusions

The current focus is acquisition documentation only.

Any review, interpretation, or investigative conclusions are the responsibility of the assigned officer, investigator, examiner, or case agent.

This project intentionally separates acquisition documentation from analysis. That separation is important because acquisition work often records technical process, tool output, device handling, and generated files, while analysis requires interpretation, context, judgment, and case-specific conclusions.

---

## Current Features

Version v0.9.0 supports:

- GUI workflow for acquisition packet documentation
- Settings menu
- Custom output/storage location configuration
- Department and unit defaults
- Common technician defaults
- Common investigator defaults
- Common tool defaults
- Common evidence location defaults
- Optional report patch/logo image path for DOCX branding
- Multiple device/media entries through a device table
- Multiple tool entries through a tool table
- Device type tracking
- Device quantity tracking
- Storage size tracking in MB, GB, or TB
- Automatic storage conversion to GB
- Total device count calculation
- Total known storage calculation
- Device-level FPR media detail fields
- Volumes examined tracking
- Volume scale tracking
- Encrypted/decrypted status tracking
- Password locked/unlocked status tracking
- Tools used to decrypt tracking
- Services used to unlock tracking
- Derived FPR media examined values
- Mobile phones counted as ETech for derived FPR credit calculations
- TXT acquisition packet output
- JSON saved packet output
- DOCX report output
- XLSX tracking workbook append
- Separate workbook sheets for case summary, device detail, FPR case info, and FPR media examined
- Review-before-export workflow
- Validation checks before review/export
- Warning display for possible data issues
- Open output folder button
- Packaged Windows `.exe` release

---

## Settings Menu

The v0.9.0 release includes a GUI settings menu.

The settings menu currently includes:

### Output / Storage

- Base output folder
- Reports folder name
- Saved packets folder name
- Tracking workbook name

If no custom base output folder is selected, the app writes output next to the executable or source files, depending on how it is run.

### Report Branding

- Optional patch/logo image path for DOCX reports

Recommended image format:

```text
PNG
```

Supported image formats:

```text
PNG
JPG
JPEG
```

SVG/vector artwork is not currently inserted directly into DOCX reports. Convert vector artwork to PNG first.

If the configured image path is missing, invalid, or unsupported, the DOCX report should still generate without the image.

### Customization / Defaults

- Department name
- Unit name
- Default technician
- Technician list
- Investigator list
- Evidence location list
- Tool list with optional version values

Tool entries may be entered in the following format:

```text
Tool Name | Version
```

Example:

```text
Cellebrite UFED | 7.x
Magnet AXIOM | 8.x
FTK Imager
```

If no version is entered, the tool is still added as a selectable default.

---

## Local Settings

The app creates a local `settings.json` file the first time it runs.

This file stores local defaults such as:

- Department name
- Unit name
- Default technician
- Common technicians
- Common investigators
- Common tools
- Common evidence locations
- Output folder settings
- Tracking workbook name
- Optional DOCX report patch/logo image path
- Default scope statement

`settings.json` is intentionally ignored by Git because it may contain local agency-specific information.

A safe example file is provided:

```text
settings.example.json
```

To manually create your own settings file, copy:

```text
settings.example.json
```

Then rename the copy to:

```text
settings.json
```

Do not commit real agency names, investigator names, technician names, internal evidence locations, report branding paths, or other sensitive local information.

---

## Output and Storage Behavior

The app can write outputs to:

- The folder where the app is running
- A custom base output folder selected by the user

The app creates or uses the following output items:

```text
output/
saved_packets/
fpr_tracking.xlsx
```

These names can be configured in Settings.

### Output Folder

TXT acquisition packet reports and DOCX reports are saved to the configured reports folder.

Example TXT report:

```text
output/TEST-001_Burglary_2026-07-12_14-30-00_acquisition_packet.txt
```

Example DOCX report:

```text
output/TEST-001_Burglary_2026-07-12_14-30-00_acquisition_packet.docx
```

### Saved Packets Folder

Structured JSON packet data is saved to the configured saved packets folder.

Example:

```text
saved_packets/TEST-001_Burglary_2026-07-12_14-30-00_acquisition_packet.json
```

The JSON packet preserves structured data from the form and derived summary fields. This is useful for later review, troubleshooting, future imports, or additional reporting workflows.

### FPR Tracking Workbook

The XLSX tracking workbook is saved as:

```text
fpr_tracking.xlsx
```

The workbook currently contains four sheets:

```text
Case Summary
Device Detail
FPR Case Info
FPR Media Examined
```

The `Case Summary` sheet contains one row per acquisition packet.

The `Device Detail` sheet contains device/media entries associated with each case.

The `FPR Case Info` sheet contains case-level fields used for FPR-style reporting.

The `FPR Media Examined` sheet contains one row per device/media entry with device-level media details such as volumes examined, encryption/decryption status, password lock status, and related unlock/decrypt notes.

---

## Review Before Export

The current workflow uses a review-before-export process.

Clicking **Review Packet** builds the packet in memory and displays a summary before writing files.

The review window includes:

- Administrative summary
- Processing/output summary
- Device/media summary
- Derived FPR media values
- Device/media rows
- Tools used
- Validation warnings, when present

Nothing is written until the user clicks:

```text
Confirm Export
```

This reduces accidental workbook pollution during testing and gives the technician a chance to catch obvious data-entry issues before generating output.

This workflow was added because early versions wrote output immediately after clicking Generate. That worked for a prototype, but it made testing easier to pollute and left less opportunity to catch mistakes before the workbook was updated. Review-before-export makes the workflow safer without adding a database or complex case management layer.

---

## Validation and Warnings

The app includes basic validation before review/export.

Validation may block export for issues such as:

- Missing case number
- Missing technician
- Missing required device/media entries
- Missing required tool entries
- Invalid date format

The app may show warnings for issues such as:

- Exam end date earlier than exam start date
- Date processed earlier than date received
- Missing subject name
- Blank offense/incident type
- Blank output filename/identifier
- Decrypted checked without encrypted checked
- Password unlocked checked without password locked checked
- Decrypt/unlock notes entered without the matching checkbox selected

Warnings do not necessarily block export. They are intended to help the technician review possible inconsistencies before confirming output.

The validation layer is intentionally basic in v0.9.0. It is not intended to replace agency policy, supervisory review, or examiner judgment. Its purpose is to catch common entry problems before output files are generated.

---

## Data Collected

The current release collects information such as:

- Case number
- Agency case number
- State/local case number
- Case type
- Offense or incident type
- City/state/country of offense
- Subject last name
- Subject first name
- Requesting investigator
- Technician
- Date/time received
- Date/time processed
- Exam start date
- Exam end date
- Drop-off person
- Evidence item number
- Evidence check-out/check-in information
- Evidence location
- Device/media type
- Device quantity
- Device make/model/serial, when available
- Device storage size
- Volumes examined
- Volume scale
- Encryption status
- Decryption status
- Tools used to decrypt
- Password locked status
- Password unlocked status
- Services used to unlock
- Tools used
- Tool versions
- Processing type
- Processing status
- Output type
- Output location
- Other data analyzed
- Case summary
- Technician notes
- Derived FPR media examined values

---

## Device and Media Entry

The Device tab supports multiple device/media entries.

Users can enter:

- Device type
- Quantity
- Description
- Make
- Model
- Serial or identifier
- Storage size per device
- Storage unit
- Volumes examined
- Volume scale
- Encrypted/decrypted status
- Tools used to decrypt
- Password locked/unlocked status
- Services used to unlock

After entering device details, click:

```text
Add Device
```

The device appears in the device table and will be included in the generated TXT report, JSON packet, DOCX report, and XLSX tracking workbook.

Users may remove a selected device row before reviewing/exporting the packet.

The Device tab is the source of truth for device-level FPR media details. The app derives FPR media examined values from the device entries instead of requiring duplicate manual credit entry.

---

## Tool Entry

The Tools tab supports multiple tool entries.

Users can select a default tool or choose **Other** and enter a custom tool name.

Each tool entry can include:

- Tool name
- Tool version

After entering tool details, click:

```text
Add Tool
```

The tool appears in the tool table and will be included in the generated acquisition packet.

Users may remove a selected tool row before reviewing/exporting the packet.

---

## Storage Size Handling

For each device or media entry, the app can collect:

- Quantity
- Storage size per device
- Storage unit: MB, GB, TB, or Unknown

Known storage values are converted to GB for tracking purposes.

Example:

```text
2 loose drives
4 TB each
8,192 GB total
```

Unknown storage values are allowed and will not be included in the known storage total.

---

## FPR-Aware Tracking

The app includes FPR-aware tracking fields intended to support structured internal reporting.

Current FPR-related tracking includes:

- FPR case information
- Exam start and end dates
- State/local case number
- Offense location
- Other data analyzed
- Case summary
- Device-level media examined details
- Volumes examined
- Volume scale
- Encrypted/decrypted status
- Password locked/unlocked status
- Unlock/decrypt service notes
- Derived media examined credit values

Derived media examined values are calculated from device/media entries.

Current working credit logic:

```text
Total Media Examined = total quantity of all device/media entries

Hard Drive Credits = CPU + Loose Drive + DVR/NVR Storage

ETech Credits = ETech + Mobile Phone

Media Credits = Tablet + USB Drive + SD Card + Cloud + Storage Media + Other
```

If a custom device type is not part of the known credit mapping, it is currently counted under Media Credits.

The credit logic is centralized in the application code so it can be changed later if agency reporting requirements or FPR interpretation changes. The GUI does not ask the user to manually enter duplicate credit totals because those totals are derived from the device/media table.

---

## DOCX Report Output

The app can generate a DOCX report that follows the acquisition packet structure.

Current DOCX sections include:

- Title and department/unit header
- Administrative Information
- Intake Information
- Examination Information
- Device / Media Summary
- Device Counts
- Device / Media Detail
- Tools Used
- Generated Output
- Other Data Analyzed
- Case Summary
- Technician Notes
- Limitations and Scope
- Technician Statement
- Report Generated

DOCX formatting is functional but not final. The v1.0.0 polish pass will focus on final wording, spacing, and presentation.

The DOCX report can optionally include a patch/logo image at the top of the report when configured in Settings.

---

## Installation for Development

### Requirements

- Python 3.10 or newer recommended
- Windows recommended for current development and packaged executable testing
- `openpyxl` for XLSX workbook output
- `python-docx` for DOCX report output
- `pyinstaller` for Windows executable builds

### Clone the Repository

```bash
git clone https://github.com/mike-C-mike/acquisition-packet-generator.git
cd acquisition-packet-generator
```

### Create a Virtual Environment

```bash
py -m venv .venv
```

Activate it on Windows:

```bash
.venv\Scripts\activate
```

### Install Dependencies

```bash
py -m pip install --upgrade pip
py -m pip install -r requirements.txt
```

To build a Windows executable, also install PyInstaller:

```bash
py -m pip install pyinstaller
```

---

## Running the App from Source

Run the GUI version:

```bash
py gui.py
```

You can also run the GUI through the launcher:

```bash
py main.py
```

The GUI is the primary development target going forward.

Earlier development began as a terminal-based proof of concept, but the current application workflow is GUI-first. The current `main.py` file functions as a launcher for the GUI rather than as a separate terminal workflow.

---

## Building the Windows Executable

The current working PyInstaller build command is:

```powershell
py -m PyInstaller --onefile --windowed --name AcquisitionPacketGenerator --paths . --hidden-import app_core --hidden-import settings_service --hidden-import docx_exporter --hidden-import validators gui.py
```

The executable will be created at:

```text
dist\AcquisitionPacketGenerator.exe
```

Before building a clean release, remove old build artifacts:

```powershell
Remove-Item -Recurse -Force build, dist -ErrorAction SilentlyContinue
Remove-Item -Force AcquisitionPacketGenerator.spec -ErrorAction SilentlyContinue
```

Then rebuild using the PyInstaller command above.

Additional build notes are available in:

```text
BUILD.md
```

---

## Recommended Release ZIP Contents

Do not upload the entire `dist` folder after testing because it may contain generated local files.

A clean release ZIP should include only:

```text
AcquisitionPacketGenerator.exe
README.md
BUILD.md
DEPENDENCIES.md
LICENSE
settings.example.json
```

Do **not** include:

```text
settings.json
output/
saved_packets/
fpr_tracking.xlsx
build/
dist/
*.spec
real case files
test files containing sensitive information
```

Generated files should remain local to the user and should not be included in the public release package.

---

## Dependency and License Inventory

Runtime dependencies currently include:

- Python
- Tkinter / Tcl-Tk
- openpyxl
- python-docx

Packaging uses:

- PyInstaller

For more detail, see:

```text
DEPENDENCIES.md
```

The project is released under the MIT License.

---

## Current Limitations

The current version has several limitations:

- No PDF export yet
- No hash calculation yet
- No hash manifest yet
- No vendor report parsing yet
- No authentication or role-based access
- No database backend
- No case management system integration
- DOCX report formatting is functional but not final
- DOCX branding currently supports PNG/JPG/JPEG images only
- SVG/vector logo support is not currently implemented
- Basic GUI styling only
- Unsigned Windows executable may trigger SmartScreen or antivirus warnings

This version is intended for workflow testing and release-candidate validation before v1.0.0.

---

## Safety and Privacy Warning

Do not use real case data in this project unless your agency, organization, or legal environment has reviewed and approved the tool for that purpose.

Do not commit or upload:

- Real case information
- Real subject names
- Real investigator names
- Real technician names, unless intentionally public
- Real agency settings
- Evidence locations
- Generated reports
- Saved packet JSON files
- FPR tracking workbooks
- Any sensitive or protected data

This project is designed to support documentation workflows, but users are responsible for validating accuracy, policy compliance, retention requirements, and legal defensibility in their own environment.

---

## Recommended Repository Hygiene

The following files should be committed:

```text
app_core.py
gui.py
main.py
docx_exporter.py
validators.py
requirements.txt
README.md
LICENSE
.gitignore
settings.example.json
BUILD.md
DEPENDENCIES.md
```

The following files should not be committed:

```text
settings.json
output/
saved_packets/
fpr_tracking.xlsx
build/
dist/
*.spec
real case files
test files containing sensitive information
```

---

## Suggested Development Workflow

Use VS Code for editing and GitHub Desktop for version control.

Recommended workflow:

1. Create a feature branch in GitHub Desktop.
2. Edit code in VS Code.
3. Run the app locally with fake test data.
4. Review generated output.
5. Confirm no sensitive generated files are staged.
6. Commit source code changes in GitHub Desktop.
7. Push the branch to GitHub.
8. Merge into `main` when tested.

Commit small changes often with clear messages.

Example commit messages:

```text
Add settings menu
Add multi-device GUI table
Add multi-tool GUI table
Add open output folder button
Improve acquisition packet text report
Refactor output path handling
Add device-level FPR media details
Add review before export workflow
Add DOCX export and report branding
Add release candidate validation and build docs
```

---

## Project Direction

Acquisition Packet Generator is intended to be the first tool in a small collection of focused digital evidence documentation utilities.

The goal is not to build one oversized application that attempts to handle every possible forensic workflow. Instead, the long-term direction is to create small, purpose-built tools that share a similar structure, documentation style, and data philosophy.

Potential companion tools may include:

- Acquisition Packet Generator
- Analysis Report Builder
- Case Folder Builder
- Hash Manifest Generator
- Media Sanitization Documentation Tool
- Exhibit / Attachment Index Builder
- Tool Version and Validation Log
- Chain of Custody Supplement Builder

Each tool should remain focused, local-first, and designed around structured data that can support consistent reporting, repeatable documentation, and defensible technical workflows.

The broader direction is to build documentation utilities that help preserve the work performed around digital evidence without trying to replace the examiner, investigator, or forensic platform. The project should remain practical, lightweight, and defensible.

---

## Project Philosophy

This project is intentionally starting simple.

The first priority is a reliable data model and repeatable acquisition workflow. The GUI is being built incrementally so the app can be tested by users who are not comfortable cloning a repository, running Python code, or working from a terminal.

The guiding principle is:

> One intake workflow, multiple defensible outputs.

The app should help reduce repetitive documentation work while keeping acquisition documentation separate from analysis, interpretation, and investigative conclusions.

The project also favors local-first operation. Generated reports, saved packet JSON files, and tracking workbooks are stored locally. The tool does not upload case data, connect to a cloud service, or require external accounts.

---

# Version Evolution

This section explains how the project evolved across each version. The version history is intentionally detailed so users can understand what changed, why it changed, and how each release built on the prior one.

---

## v0.1.0 - Terminal Proof of Concept

The earliest version of Acquisition Packet Generator began as a simple proof of concept. At this stage, the goal was not to build a polished application. The goal was to test whether a repeatable acquisition packet could be created from structured user input.

This first stage focused on the idea that acquisition documentation could be broken into predictable sections:

- Case information
- Subject information
- Device/media information
- Tool information
- Processing information
- Output information
- Technician notes

The application was still terminal-based. A user could enter information, and the project could create early text output and structured packet data. This confirmed the basic concept: one structured intake process could produce repeatable documentation.

### Added in v0.1.0

- Initial terminal-based workflow
- Basic acquisition packet data collection
- Early TXT report output
- Early JSON-style structure
- Initial separation between case information, device information, tools, and output
- Basic project direction established

### Why v0.1.0 mattered

This version proved that the project idea was workable. It did not need a full database, cloud service, or complex interface to begin producing useful acquisition documentation. It also established the core philosophy that later versions continued to follow:

> Structured input should create multiple useful outputs.

---

## v0.2.0 - Structured Data and XLSX Tracking

Version v0.2.0 expanded the project from simple report generation into structured tracking.

At this stage, the project began treating the acquisition packet as more than just a text report. The packet data could also support tracking and internal reporting. This was an important step because digital evidence work often needs both narrative documentation and structured summary data.

The XLSX tracking workbook concept was introduced so that each generated packet could append useful case and device information into a spreadsheet. This made the project more useful for internal reporting, workload tracking, device counts, media counts, and later FPR-style reporting.

### Added in v0.2.0

- Expanded packet structure
- XLSX tracking workbook support
- Case summary style tracking
- Early device detail tracking
- Improved separation between report output and structured data
- Continued terminal/source-based workflow testing

### Why v0.2.0 mattered

This version shifted the project from “generate a report” toward “collect structured acquisition data once and reuse it.” That idea became the foundation for the later TXT, JSON, DOCX, and XLSX outputs.

It also introduced one of the project’s key long-term values: acquisition documentation should support both case-level reporting and administrative tracking.

---

## v0.3.0 - First Packaged GUI Prototype

Version v0.3.0 was the first major usability milestone. The project moved from a terminal/source-code workflow into a basic Tkinter GUI and became available as a packaged Windows executable.

This mattered because many potential users of the tool may not be comfortable cloning a repository, installing dependencies, or running Python scripts from a command line. Packaging the tool as a Windows executable made it easier to test as a practical workflow tool.

At this stage, the GUI was still basic, but it provided a usable form-based workflow for entering case, device, tool, and output information.

### Added in v0.3.0

- First packaged Windows executable release
- Basic Tkinter GUI prototype
- TXT acquisition packet output
- JSON packet save
- XLSX tracking append
- Local settings support
- Initial packaged release workflow

### Why v0.3.0 mattered

This version moved the project from developer-only testing toward user testing. It made the tool more accessible and confirmed that a GUI workflow was the correct direction for the project.

The packaged executable also introduced new considerations that shaped later versions, including:

- Local file creation next to the executable
- Output folder behavior
- Local settings files
- Release ZIP hygiene
- Avoiding accidental inclusion of generated data

---

## v0.4.0 - Settings and Multi-Entry Workflow

Version v0.4.0 focused on making the GUI more practical and configurable.

Earlier versions could collect basic information, but the workflow needed better support for repeated use. Digital evidence technicians often use recurring values such as department names, unit names, technician names, investigator names, evidence locations, and common tools.

The settings menu was added so those recurring values could be configured once and reused. This made the tool more agency-friendly and reduced repetitive typing.

This version also improved the device and tool workflow by supporting multiple device/media entries and multiple tool entries through GUI tables.

### Added in v0.4.0

- Settings menu
- Custom output/storage path configuration
- Editable department and unit defaults
- Editable technician defaults
- Editable investigator defaults
- Editable tool defaults
- Editable evidence location defaults
- Multi-device GUI table
- Multi-tool GUI table
- Open output folder button
- Improved workflow usability

### Why v0.4.0 mattered

This version made the app feel more like a usable workflow tool rather than a simple form. Multi-device and multi-tool support were especially important because real acquisition work often involves more than one phone, drive, SD card, USB device, or generated output.

The settings work also established the pattern that local agency-specific information should live in `settings.json`, while `settings.example.json` should remain safe to commit.

---

## v0.5.0 - FPR Case Fields and Derived Media Summary

Version v0.5.0 expanded the project toward FPR-style reporting support.

The app began collecting more case-level information that could support FPR-related reporting fields. Instead of creating a separate FPR form or a disconnected workflow, the project started adding those fields naturally into the existing acquisition packet workflow.

This version also introduced better output file naming to prevent accidental overwrites. Earlier output names could be reused if the same case number was entered multiple times. The updated naming approach included the case number, case title/offense value, and timestamp.

The project also began centralizing settings and path logic into a separate settings service, improving maintainability and reducing duplicated logic.

### Added in v0.5.0

- Settings service refactor
- New / Clear Packet workflow
- Timestamped output filenames to prevent overwrites
- Natural FPR case information fields
- State/local case number field
- Case type field
- Offense location fields
- Exam start and end date fields
- Other data analyzed field
- Case summary field
- FPR Case Info worksheet
- FPR Media Examined worksheet
- Derived media examined summary values
- Cleaner XLSX permission error handling

### Important v0.5.0 design decision

The project intentionally did not create a standalone “FPR tab” as the primary workflow. Instead, FPR-related fields were added where they naturally belonged in the acquisition packet process.

This kept the acquisition packet as the source of truth and allowed the XLSX workbook to become a structured export of that data.

### Why v0.5.0 mattered

This version moved the tool closer to real administrative usefulness. It began supporting not just packet documentation, but also structured fields that could help with recurring reporting needs.

It also revealed an important correction for later versions: mobile phones should count as ETech for FPR credit logic. That correction was handled in v0.6.0.

---

## v0.6.0 - Device-Level FPR Media Details

Version v0.6.0 was a major FPR-related workflow improvement.

Instead of creating a separate manual media category grid, the app began collecting FPR media details directly at the device row level. This made the Device tab the source of truth for device/media information.

The update added fields for volumes examined, volume scale, encryption/decryption status, password lock/unlock status, tools used to decrypt, and services used to unlock. These fields are captured with each device/media entry rather than being entered separately somewhere else.

This version also corrected the derived credit logic so mobile phones count as ETech.

### Added in v0.6.0

- Device-level FPR media fields
- Volumes examined
- Volume scale
- Encrypted/decrypted status
- Tools used to decrypt
- Password locked/unlocked status
- Services used to unlock
- Mobile phones counted as ETech credits
- Expanded Device Detail worksheet
- FPR Media Examined worksheet changed to one row per device/media entry
- TXT output updated with device-level FPR fields
- JSON output updated with device-level FPR fields

### FPR credit logic updated in v0.6.0

Current working logic:

```text
Total Media Examined = total quantity of all device/media entries

Hard Drive Credits = CPU + Loose Drive + DVR/NVR Storage

ETech Credits = ETech + Mobile Phone

Media Credits = Tablet + USB Drive + SD Card + Cloud + Storage Media + Other
```

### Why v0.6.0 mattered

This version prevented duplicate data entry. Instead of asking the user to manually enter summary counts in one place and device details in another, the app derives summary values from the device table.

That keeps the data model cleaner and reduces the chance that the TXT report, JSON packet, and XLSX workbook will disagree.

---

## v0.7.0 - Review Before Export

Version v0.7.0 changed the export workflow.

Earlier versions wrote output files immediately after the user clicked Generate. That was workable in an early prototype, but it created problems during testing and real use. A user could accidentally write bad data to the tracking workbook before reviewing the packet.

The review-before-export workflow solved that by adding a review window. The app now builds the packet in memory, displays a summary, and waits for the user to confirm before writing files.

### Added in v0.7.0

- Review-before-export workflow
- Review window before writing TXT, JSON, DOCX, or XLSX output
- Cancel option before file creation
- Cleaner export confirmation flow
- Review summary with administrative information
- Review summary with processing/output information
- Review summary with device/media totals
- Review summary with derived FPR media values
- Review summary with devices and tools

### Fixed during v0.7.0 testing

- Review text formatting issue caused by string concatenation behavior
- Repeated review section headers
- Repeated divider formatting

### Why v0.7.0 mattered

This version made the tool safer to use. The tracking workbook should not be updated until the user has had a chance to review the packet.

This was an important step toward v1.0 because it made the tool behave more like a real documentation workflow rather than a quick prototype.

---

## v0.8.0 - DOCX Export and Report Branding

Version v0.8.0 added DOCX report output.

Earlier versions generated TXT, JSON, and XLSX. TXT was useful for plain documentation, JSON preserved structured data, and XLSX supported tracking. DOCX added a more familiar report format that can be opened, reviewed, edited, printed, saved as PDF, or adapted for agency-specific use.

This version also added optional report branding. Users can configure a PNG, JPG, or JPEG image path in Settings, and the image can be placed at the top of DOCX reports.

### Added in v0.8.0

- DOCX report generation
- DOCX report sections matching the acquisition packet structure
- Optional patch/logo image path in Settings
- Report Branding settings tab
- PNG/JPG/JPEG image support for DOCX report header branding
- `python-docx` dependency added
- DOCX output included in export confirmation
- Missing or invalid branding image path handled without blocking report generation

### DOCX report sections added

The DOCX report currently includes:

- Title and department/unit header
- Administrative Information
- Intake Information
- Examination Information
- Device / Media Summary
- Device Counts
- Device / Media Detail
- Tools Used
- Generated Output
- Other Data Analyzed
- Case Summary
- Technician Notes
- Limitations and Scope
- Technician Statement
- Report Generated

### Why v0.8.0 mattered

This version made the app more useful outside of pure testing. DOCX is a practical working format for documentation because it can be reviewed and edited before final use.

The report branding option also made the tool more agency-friendly without requiring hardcoded images or agency-specific files in the repository.

---

## v0.9.0 - Release Candidate Hardening

Version v0.9.0 focused on hardening the app before v1.0.0.

This version added validation, warning display, release build documentation, dependency/license documentation, and updated settings examples. The goal was not to add a large new workflow, but to make the existing workflow safer, clearer, and easier to package.

The validation system checks for missing required information and date formatting issues before the review window is displayed. It also provides warnings for possible inconsistencies without necessarily blocking export.

### Added in v0.9.0

- Packet validation before review/export
- Date format checks
- Date order warnings
- Device-level consistency warnings
- Review-window warning display
- Build documentation
- Dependency and license inventory
- Updated settings example
- Release-candidate hardening before v1.0.0
- Current packaged release version

### Validation added in v0.9.0

Validation may block export for issues such as:

- Missing case number
- Missing technician
- Missing required device/media entries
- Missing required tool entries
- Invalid date format

Warnings may be shown for issues such as:

- Exam end date earlier than exam start date
- Date processed earlier than date received
- Missing subject name
- Blank offense/incident type
- Blank output filename/identifier
- Decrypted checked without encrypted checked
- Password unlocked checked without password locked checked
- Decrypt/unlock notes entered without the matching checkbox selected

### Documentation added in v0.9.0

Version v0.9.0 also added or updated documentation files intended to support a cleaner public release:

- `BUILD.md`
- `DEPENDENCIES.md`
- Updated `settings.example.json`
- Updated README content
- Updated release ZIP guidance
- Updated PyInstaller build command

### Why v0.9.0 mattered

This version prepared the app for a stable public release. The main question shifted from “Can the feature work?” to “Can the tool be packaged, tested, explained, and released cleanly?”

That is the purpose of the v0.9.0 release candidate.

---

# Roadmap

## v1.0.0

The planned v1.0.0 release is the first stable public release.

Focus areas:

- Final README and release documentation
- Final UI wording pass
- Final output validation pass
- Clean packaged Windows ZIP
- Release checksum
- GitHub release notes
- Public release announcement

v1.0.0 is not intended to be a major new-feature release. It should be a polish and packaging release built on top of the v0.9.0 release candidate.

## Future

Future development may include:

- PDF export
- Hash calculation
- SHA-256 and MD5 support
- Optional folder hash manifest generation
- Vendor report parsing
- Further documentation modules
- Expanded companion tools
- Case folder creation
- Exhibit / attachment index generation
- Media sanitization documentation
- Analysis report builder as a separate companion project

Future features should continue following the same project philosophy:

- Keep acquisition documentation separate from analysis
- Keep the tool local-first
- Avoid unnecessary complexity
- Use structured data once and export it in multiple useful formats
- Do not generate investigative conclusions automatically

---

## Contributing

This project is open to future contributions.

Useful contribution areas may include:

- GUI development
- DOCX formatting and template improvements
- XLSX reporting improvements
- Data validation
- Validation rules
- Release testing
- Hashing support
- Code refactoring
- Documentation
- Test data generation
- Packaging and release workflows

Before contributing, avoid including any real case data, agency-specific information, or sensitive operational details.

---

## License

This project is licensed under the MIT License.

See the `LICENSE` file for details.

---

## Disclaimer

This software is provided as a release candidate for documentation workflow development.

It does not perform forensic analysis, validate evidence, guarantee legal admissibility, or replace agency policy, examiner judgment, legal review, or supervisory approval.

Users are responsible for testing, validating, and approving the tool before using it in any official capacity.