# Acquisition Packet Generator

**Acquisition Packet Generator** is an early-stage Python tool for documenting digital evidence intake, mobile extractions, drive/media imaging, generated outputs, and related acquisition packet information.

The goal of this project is to help digital evidence technicians create consistent acquisition documentation while also collecting structured data that can support internal reporting, device tracking, and FPR-style reporting workflows.

This project is currently a **terminal-based prototype**. A GUI and DOCX report generation are planned for a future version.

## Current Status

**Version:** 0.1
**Stage:** Early prototype
**Interface:** Terminal / command-line
**Primary outputs:** TXT acquisition packet, JSON saved packet, XLSX tracking workbook

This version is intended to prove the intake workflow and data structure before adding a graphical interface, DOCX report generation, hashing, vendor report parsing, or packaged executable releases.

## Intended Use

This tool is designed to document acquisition-related tasks such as:

* Intake of a device or storage media
* Drop-off and evidence handling information
* Requesting officer or investigator information
* Technician information
* Device/media type and storage size
* Tool(s) used during processing
* Processing status
* Output generated, such as reader reports, case files, exports, or forensic images
* Return of the original device or media to evidence storage
* Structured reporting data for XLSX tracking

## Not Intended For

This tool is **not** intended to perform forensic analysis or generate investigative conclusions.

It should not be used to document interpretive findings such as:

* Timeline analysis
* Malware findings
* Documents of interest
* User activity conclusions
* USB history interpretation
* Evidence relevance
* Investigative opinions
* Legal conclusions

The current focus is acquisition documentation only.

Any review, interpretation, or investigative conclusions are the responsibility of the assigned officer, investigator, examiner, or case agent.

## Current Features

Version 0.1 currently supports:

* Terminal-based acquisition packet intake
* Local `settings.json` defaults
* Common technician defaults
* Common investigator defaults
* Common tool defaults
* Common evidence location defaults
* Multiple device/media entries
* Device type tracking
* Quantity tracking
* Storage size tracking in MB, GB, or TB
* Automatic storage conversion to GB
* Total device count calculation
* Total known storage calculation
* TXT acquisition packet output
* JSON saved packet output
* XLSX tracking workbook append
* Separate workbook sheets for case summary and device detail

## Project Direction

Acquisition Packet Generator is intended to be the first tool in a small collection of focused digital evidence documentation utilities.

The goal is not to build one oversized application that attempts to handle every possible forensic workflow. Instead, the long-term direction is to create small, purpose-built tools that share a similar structure, documentation style, and data philosophy.

Potential companion tools may include:

* Acquisition Packet Generator
* Analysis Report Builder
* Case Folder Builder
* Hash Manifest Generator
* Media Sanitization Documentation Tool
* Exhibit / Attachment Index Builder
* Tool Version and Validation Log
* Chain of Custody Supplement Builder

Each tool should remain focused, local-first, and designed around structured data that can support consistent reporting, repeatable documentation, and defensible technical workflows.

## Project Philosophy

This project is intentionally starting simple.

The first priority is a reliable data model and repeatable acquisition workflow. The terminal interface is temporary. The long-term goal is to use the same underlying data structure for a future GUI, DOCX reports, XLSX reporting, release builds, and additional acquisition-support utilities.

The guiding principle is:

> One intake workflow, multiple defensible outputs.

## Installation

### Requirements

* Python 3.10 or newer recommended
* Windows, Linux, or macOS should work for the terminal prototype
* `openpyxl` for XLSX workbook output

### Clone the Repository

```bash
git clone https://github.com/YOUR-USERNAME/acquisition-packet-generator.git
cd acquisition-packet-generator
```

Replace `YOUR-USERNAME` with your GitHub username.

### Install Dependencies

```bash
python -m pip install -r requirements.txt
```

On some Windows systems, use:

```bash
py -m pip install -r requirements.txt
```

## Running the App

Run the program from the project folder:

```bash
python main.py
```

Or on some Windows systems:

```bash
py main.py
```

The program opens a terminal menu with options to create a new acquisition packet, view the settings file location, reset settings, or exit.

## Local Settings

The app creates a local `settings.json` file the first time it runs.

This file stores local defaults such as:

* Department name
* Unit name
* Default technician
* Common technicians
* Common investigators
* Common tools
* Common evidence locations
* Default scope statement

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

## Generated Files

The app may generate the following local files and folders:

```text
settings.json
output/
saved_packets/
fpr_tracking.xlsx
```

These files are ignored by Git and should generally remain local.

### Output Folder

TXT acquisition packet reports are saved to:

```text
output/
```

Example:

```text
output/TEST-001_acquisition_packet.txt
```

### Saved Packets Folder

Structured JSON packet data is saved to:

```text
saved_packets/
```

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

* Case number
* Agency case number
* Offense or incident type
* Subject last name
* Subject first name
* Requesting investigator
* Technician
* Date/time received
* Date/time processed
* Drop-off person
* Evidence item number
* Evidence check-out/check-in information
* Evidence location
* Device/media type
* Device quantity
* Device make/model/serial, when available
* Device storage size
* Tools used
* Processing type
* Processing status
* Output type
* Output location
* Technician notes

## Storage Size Handling

For each device or media entry, the app can collect:

* Quantity
* Storage size per device
* Storage unit: MB, GB, TB, or Unknown

Known storage values are converted to GB for tracking purposes.

Example:

```text
2 loose drives
4 TB each
8,192 GB total
```

Unknown storage values are allowed and will not be included in the known storage total.

## Current Limitations

Version 0.1 has several limitations:

* Terminal only
* No GUI yet
* No DOCX export yet
* No PDF export yet
* No packaged `.exe` release yet
* No hash calculation yet
* No hash manifest yet
* No vendor report parsing yet
* No built-in editing after review
* No authentication or role-based access
* No database backend
* No case management system integration

This version is intended for workflow testing and development only.

## Safety and Privacy Warning

Do not use real case data in this project unless your agency, organization, or legal environment has reviewed and approved the tool for that purpose.

Do not commit or upload:

* Real case information
* Real subject names
* Real investigator names
* Real technician names, unless intentionally public
* Real agency settings
* Evidence locations
* Generated reports
* Saved packet JSON files
* FPR tracking workbooks
* Any sensitive or protected data

This project is designed to support documentation workflows, but users are responsible for validating accuracy, policy compliance, retention requirements, and legal defensibility in their own environment.

## Recommended Repository Hygiene

The following files should be committed:

```text
main.py
requirements.txt
README.md
LICENSE
.gitignore
settings.example.json
```

The following files should not be committed:

```text
settings.json
output/
saved_packets/
fpr_tracking.xlsx
real case files
test files containing sensitive information
```

## Suggested Development Workflow

Use VS Code for editing and GitHub Desktop for version control.

Recommended workflow:

1. Edit code in VS Code
2. Run the app locally with fake test data
3. Review generated output
4. Confirm no sensitive generated files are staged
5. Commit source code changes in GitHub Desktop
6. Push to GitHub

Commit small changes often with clear messages.

Example commit messages:

```text
Add device storage tracking
Add XLSX tracking workbook output
Improve acquisition packet text report
Add settings example file
Refactor report builder logic
```

## Roadmap

### v0.1

* Terminal intake workflow
* TXT acquisition packet
* JSON saved packet
* XLSX tracking append
* Local settings defaults
* Multiple device/media entries
* Device quantity and storage tracking

### v0.2

* DOCX report generation
* Basic GUI prototype
* Tabs for general info, intake, devices, tools, output, review, and settings
* Dropdowns and checkbox-based tool selection
* Department header support
* Signature block support

### v0.3

* Packaged Windows `.exe`
* GitHub Releases for downloadable builds
* Basic release notes and checksum support
* Improved setup instructions for non-developer users

### v0.4

* Workflow cleanup
* Data structure improvements
* Better validation for required fields
* Improved review/edit-before-export workflow
* More configurable default options
* Cleaner XLSX column structure
* Improved error handling

### v0.5

* Hash calculation
* SHA-256 and MD5 support
* File selection for hashing
* Optional folder hash manifest generation
* Hash output included in acquisition packet and XLSX tracking where applicable

### Future

* Vendor report parsing
* Further documentation modules
* Expanded companion tools
* PDF export
* Case folder creation
* Exhibit / attachment index generation
* Media sanitization documentation
* Analysis report builder as a separate companion project

## Contributing

This project is open to future contributions.

Useful contribution areas may include:

* GUI development
* DOCX template generation
* XLSX reporting improvements
* Data validation
* Hashing support
* Code refactoring
* Documentation
* Test data generation
* Packaging and release workflows

Before contributing, avoid including any real case data, agency-specific information, or sensitive operational details.

## License

This project is licensed under the MIT License.

See the `LICENSE` file for details.

## Disclaimer

This software is provided as an early prototype for documentation workflow development.

It does not perform forensic analysis, validate evidence, guarantee legal admissibility, or replace agency policy, examiner judgment, legal review, or supervisory approval.

Users are responsible for testing, validating, and approving the tool before using it in any official capacity.
