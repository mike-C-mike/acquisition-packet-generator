# ByteCase Acquire

**Acquisition Packet Generator**

Part of the ByteCase toolset by Forensics Byte.

## Purpose

ByteCase Acquire helps examiners document digital forensic acquisition, extraction, imaging, processing, device/media details, tools used, generated outputs, and FPR-aware tracking data.

ByteCase Acquire does **not** perform forensic acquisition, bypass device security, parse evidence, analyze artifacts, determine evidentiary relevance, or create investigative conclusions.

## v0.9.3 feedback sprint

This sprint incorporates first-user feedback from the early public release line and brings those ideas into the current ByteCase-themed Acquire build.

Added:

- Agency Dropping Item Off field
- Where Device Stored field for each device/media item
- Condition Delivered field for each device/media item
- Optional device photo path for each device/media item
- Device photo copying into each exported packet
- Device photo embedding in DOCX reports when supported
- Larger Processing / Acquisition Narrative text box
- Settings preset tab for common defaults
- Settings lists for storage locations and delivery conditions
- Lab Statistics export by date range

## Output structure

Default output root:

```text
C:\Users\<user>\ByteCase\
```

Acquire packet output:

```text
ByteCase\
  <case_number>\
    acquire\
      reports\
      saved_packets\
      tracking\
      attachments\
        device_photos\
```

Lab statistics output:

```text
ByteCase\
  acquire\
    lab_statistics\
```

## Device photos

Device photos are optional. When provided, supported image files are copied into the packet under:

```text
attachments\device_photos\
```

Supported image extensions:

```text
.png, .jpg, .jpeg, .bmp, .gif, .tif, .tiff
```

DOCX embedding is attempted for supported images. If an image cannot be copied or embedded, export continues and records the issue in the packet.

## Lab statistics

Use **Generate → Export Lab Stats** to export saved packet data across a date range.

The export reads saved ByteCase Acquire packet JSON files and writes:

```text
CSV
XLSX
```

Date filtering uses Date Processed when available. Leave date fields blank to export all saved packets.

## PDF export note

PDF export is not included in v0.9.3. It is parked until a reliable and license-safe PDF implementation path is selected. Current report outputs remain TXT, JSON, DOCX, and XLSX.

## License

MIT License.


## v0.9.3 Notes

This update expands device photo handling to match the Intake-style attachment model. Each device/media row can now include multiple labeled images, such as Front, Back, Top, Bottom, Screen, Ports, or Serial/Identifier. Photos are copied into the acquisition packet under `attachments/device_photos/` during export and are included in TXT/JSON/DOCX outputs.

Settings tabs now use visible scrollbars so smaller screens can reach all fields. The Output / Storage settings tab also exposes the ByteCase output root and folder-name fields for reports, saved packets, tracking, attachments, device photos, and lab statistics.

Default output remains:

```text
C:\Users\<user>\ByteCase\<case_number>\acquire\
```

When a custom output root is selected, Acquire writes case folders directly under that root:

```text
<custom root>\<case_number>\acquire\
```
