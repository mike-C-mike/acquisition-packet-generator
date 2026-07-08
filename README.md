# Acquisition Packet Generator

**Acquisition Packet Generator** is an early-stage Python tool for documenting digital evidence intake, mobile extractions, drive/media imaging, generated outputs, and related acquisition packet information.

The goal of this project is to help digital evidence technicians create consistent acquisition documentation while also collecting structured data that can support internal reporting, device tracking, and FPR-style reporting workflows.

This project is currently an early GUI-based prototype with a packaged Windows executable release available through GitHub Releases.

## Current Status

**Current version:** v0.4.0  
**Latest packaged release:** v0.4.0  
**Stage:** Early prototype  
**Interface:** Tkinter GUI  
**Primary outputs:** TXT acquisition packet, JSON saved packet, XLSX tracking workbook

The project began as a terminal-based proof of concept and has moved into a basic GUI workflow. The current version includes settings management, configurable output paths, multi-device entries, multi-tool entries, and an open output folder option.

## Intended Use

This tool is designed to document acquisition-related tasks such as:

- Intake of a device or storage media
- Drop-off and evidence handling information
- Requesting officer or investigator information
- Technician information
- Device/media type and storage size
- Tool(s) used during processing
- Processing status
- Output generated, such as reader reports, case files, exports, or forensic images
- Return of the original device or media to evidence storage
- Structured reporting data for XLSX tracking

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

## Current Features

Version v0.4.0 supports:

- Basic GUI workflow
- Settings menu
- Custom output/storage location configuration
- Department and unit defaults
- Common technician defaults
- Common investigator defaults
- Common tool defaults
- Common evidence location defaults
- Multiple device/media entries through a device table
- Multiple tool entries through a tool table
- Device type tracking
- Device quantity tracking
- Storage size tracking in MB, GB, or TB
- Automatic storage conversion to GB
- Total device count calculation
- Total known storage calculation
- TXT acquisition packet output
- JSON saved packet output
- XLSX tracking workbook append
- Separate workbook sheets for case summary and device detail
- Open output folder button
- Packaged Windows `.exe` release

## Latest Packaged Release

The latest packaged Windows prototype is available from the GitHub Releases page.

Users who do not want to run the source code directly can download the Windows ZIP release, extract it, and run:

```text
AcquisitionPacketGenerator.exe
```

The app should be extracted from the ZIP before running. Do not run the executable directly from inside the ZIP preview window.

When the executable runs, it creates local files and folders next to the executable unless a custom output path is configured in Settings.

Default generated local items:

```text
settings.json
output/
saved_packets/
fpr_tracking.xlsx
```

## Download and Run Instructions

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
10. Click **Generate Packet**.
11. Use **Open Output Folder** to view generated files.

Windows may show a SmartScreen or antivirus warning because this is an unsigned open-source prototype. The source code is available in the GitHub repository for review.

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

## Project Philosophy

This project is intentionally starting simple.

The first priority is a reliable data model and repeatable acquisition workflow. The GUI is being built incrementally so the app can be tested by users who are not comfortable cloning a repository, running Python code, or working from a terminal.

The guiding principle is:

> One intake workflow, multiple defensible outputs.

The app should help reduce repetitive documentation work while keeping acquisition documentation separate from analysis, interpretation, and investigative conclusions.

## Installation for Development

### Requirements

- Python 3.10 or newer recommended
- Windows recommended for current development and packaged executable testing
- `openpyxl` for XLSX workbook output
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

## Running the App from Source

Run the GUI version:

```bash
py gui.py
```

The older terminal version may still exist for development reference:

```bash
py main.py
```

The GUI is the primary development target going forward.

## Building the Windows Executable

The current working PyInstaller build command is:

```powershell
py -m PyInstaller --onefile --windowed --name AcquisitionPacketGenerator --paths . --hidden-import app_core gui.py
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

## Recommended Release ZIP Contents

Do not upload the entire `dist` folder after testing because it may contain generated local files.

A clean release ZIP should include only:

```text
AcquisitionPacketGenerator.exe
README.txt
LICENSE.txt
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

Do not commit real agency names, investigator names, technician names, internal evidence locations, or other sensitive local information.

## Settings Menu

The v0.4.0 release includes a GUI settings menu.

The settings menu currently includes:

### Output / Storage

- Base output folder
- Reports folder name
- Saved packets folder name
- Tracking workbook name

If no custom base output folder is selected, the app writes output next to the executable or source files, depending on how it is run.

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

TXT acquisition packet reports are saved to the configured reports folder.

Example:

```text
output/TEST-001_acquisition_packet.txt
```

### Saved Packets Folder

Structured JSON packet data is saved to the configured saved packets folder.

Example:

```text
saved_packets/TEST-001_acquisition_packet.json
```

### FPR Tracking Workbook

The XLSX tracking workbook is saved as:

```text
fpr_tracking.xlsx
```

The workbook currently contains two sheets:

```text
Case Summary
Device Detail
```

The `Case Summary` sheet contains one row per acquisition packet.

The `Device Detail` sheet contains device/media entries associated with each case.

## Data Collected

The current prototype collects information such as:

- Case number
- Agency case number
- Offense or incident type
- Subject last name
- Subject first name
- Requesting investigator
- Technician
- Date/time received
- Date/time processed
- Drop-off person
- Evidence item number
- Evidence check-out/check-in information
- Evidence location
- Device/media type
- Device quantity
- Device make/model/serial, when available
- Device storage size
- Tools used
- Tool versions
- Processing type
- Processing status
- Output type
- Output location
- Technician notes

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

After entering device details, click:

```text
Add Device
```

The device appears in the device table and will be included in the generated TXT report, JSON packet, and XLSX tracking workbook.

Users may remove a selected device row before generating the packet.

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

Users may remove a selected tool row before generating the packet.

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

## Current Limitations

The current version has several limitations:

- No DOCX export yet
- No PDF export yet
- No hash calculation yet
- No hash manifest yet
- No vendor report parsing yet
- No authentication or role-based access
- No database backend
- No case management system integration
- Basic GUI styling only
- Unsigned Windows executable may trigger SmartScreen or antivirus warnings

This version is intended for workflow testing and development only.

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

## Recommended Repository Hygiene

The following files should be committed:

```text
app_core.py
gui.py
main.py
requirements.txt
README.md
LICENSE
.gitignore
settings.example.json
BUILD.md
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
```

## Roadmap

### v0.3.0

- First packaged Windows executable release
- Basic GUI prototype
- TXT acquisition packet output
- JSON packet save
- XLSX tracking append
- Local settings support

### v0.4.0

- Settings menu
- Custom output/storage path configuration
- Editable department/unit defaults
- Editable technician defaults
- Editable investigator defaults
- Editable tool defaults
- Editable evidence location defaults
- Multi-device GUI table
- Multi-tool GUI table
- Open output folder button
- Improved workflow usability

### v0.5.0

- DOCX report generation
- Department header support
- Signature block support
- Cleaner report formatting
- Better review-before-export workflow

### v0.6.0

- Hash calculation
- SHA-256 and MD5 support
- File selection for hashing
- Optional folder hash manifest generation
- Hash output included in acquisition packet where applicable

### Future

- Vendor report parsing
- Further documentation modules
- Expanded companion tools
- PDF export
- Case folder creation
- Exhibit / attachment index generation
- Media sanitization documentation
- Analysis report builder as a separate companion project

## Contributing

This project is open to future contributions.

Useful contribution areas may include:

- GUI development
- DOCX template generation
- XLSX reporting improvements
- Data validation
- Hashing support
- Code refactoring
- Documentation
- Test data generation
- Packaging and release workflows

Before contributing, avoid including any real case data, agency-specific information, or sensitive operational details.

## License

This project is licensed under the MIT License.

See the `LICENSE` file for details.

## Disclaimer

This software is provided as an early prototype for documentation workflow development.

It does not perform forensic analysis, validate evidence, guarantee legal admissibility, or replace agency policy, examiner judgment, legal review, or supervisory approval.

Users are responsible for testing, validating, and approving the tool before using it in any official capacity.