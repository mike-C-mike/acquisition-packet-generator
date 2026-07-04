import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

from app_core import (
    APP_NAME,
    APP_VERSION,
    DEVICE_TYPES,
    PROCESSING_TYPES,
    PROCESSING_STATUSES,
    OUTPUT_TYPES,
    STORAGE_UNITS,
    ensure_directories,
    load_or_create_settings,
    convert_to_gb,
    count_devices_by_type,
    calculate_total_storage_gb,
    format_storage_gb,
    save_packet_outputs,
    append_to_fpr_tracking,
)


THEME = {
    "bg": "#0B0B0D",
    "panel": "#151518",
    "accent": "#C9A227",
    "text": "#F2F2F2",
    "muted": "#A8A8A8",
    "input_bg": "#202024",
    "button_text": "#0B0B0D"
}


class AcquisitionPacketGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{APP_NAME} v{APP_VERSION}")
        self.root.geometry("900x750")
        self.root.configure(bg=THEME["bg"])

        ensure_directories()
        self.settings = load_or_create_settings()

        self.create_styles()
        self.create_layout()

    def create_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure(
            "TNotebook",
            background=THEME["bg"],
            borderwidth=0
        )

        style.configure(
            "TNotebook.Tab",
            background=THEME["panel"],
            foreground=THEME["text"],
            padding=[12, 6]
        )

        style.map(
            "TNotebook.Tab",
            background=[("selected", THEME["accent"])],
            foreground=[("selected", THEME["button_text"])]
        )

        style.configure(
            "TCombobox",
            fieldbackground=THEME["input_bg"],
            background=THEME["input_bg"],
            foreground=THEME["text"]
        )

    def create_layout(self):
        header = tk.Frame(self.root, bg=THEME["bg"])
        header.pack(fill="x", padx=20, pady=(15, 5))

        title = tk.Label(
            header,
            text="Acquisition Packet Generator",
            bg=THEME["bg"],
            fg=THEME["accent"],
            font=("Segoe UI", 20, "bold")
        )
        title.pack(anchor="w")

        subtitle = tk.Label(
            header,
            text="Early GUI prototype for acquisition documentation and XLSX tracking",
            bg=THEME["bg"],
            fg=THEME["muted"],
            font=("Segoe UI", 10)
        )
        subtitle.pack(anchor="w")

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=20, pady=10)

        self.general_tab = self.create_tab("General Info")
        self.intake_tab = self.create_tab("Intake")
        self.device_tab = self.create_tab("Device")
        self.tools_tab = self.create_tab("Tools")
        self.output_tab = self.create_tab("Processing / Output")
        self.review_tab = self.create_tab("Generate")

        self.build_general_tab()
        self.build_intake_tab()
        self.build_device_tab()
        self.build_tools_tab()
        self.build_output_tab()
        self.build_generate_tab()

    def create_tab(self, name):
        frame = tk.Frame(self.notebook, bg=THEME["panel"])
        self.notebook.add(frame, text=name)
        return frame

    def label(self, parent, text, row, column=0):
        widget = tk.Label(
            parent,
            text=text,
            bg=THEME["panel"],
            fg=THEME["text"],
            font=("Segoe UI", 10)
        )
        widget.grid(row=row, column=column, sticky="w", padx=10, pady=6)
        return widget

    def entry(self, parent, row, column=1, width=45, default=""):
        widget = tk.Entry(
            parent,
            width=width,
            bg=THEME["input_bg"],
            fg=THEME["text"],
            insertbackground=THEME["text"],
            relief="flat"
        )
        widget.grid(row=row, column=column, sticky="w", padx=10, pady=6)

        if default:
            widget.insert(0, default)

        return widget

    def combo(self, parent, values, row, column=1, width=42, default=None):
        widget = ttk.Combobox(parent, values=values, width=width, state="readonly")
        widget.grid(row=row, column=column, sticky="w", padx=10, pady=6)

        if default:
            widget.set(default)
        elif values:
            widget.set(values[0])

        return widget

    def checkbox(self, parent, text, row, column=0):
        variable = tk.BooleanVar()
        widget = tk.Checkbutton(
            parent,
            text=text,
            variable=variable,
            bg=THEME["panel"],
            fg=THEME["text"],
            selectcolor=THEME["input_bg"],
            activebackground=THEME["panel"],
            activeforeground=THEME["text"]
        )
        widget.grid(row=row, column=column, sticky="w", padx=10, pady=4)
        return variable

    def build_general_tab(self):
        frame = self.general_tab

        self.label(frame, "Case Number", 0)
        self.case_number = self.entry(frame, 0)

        self.label(frame, "Agency Case Number", 1)
        self.agency_case_number = self.entry(frame, 1)

        self.label(frame, "Offense / Incident Type", 2)
        self.offense_or_incident = self.entry(frame, 2)

        self.label(frame, "Subject Last Name", 3)
        self.subject_last = self.entry(frame, 3)

        self.label(frame, "Subject First Name", 4)
        self.subject_first = self.entry(frame, 4)

        investigators = self.settings.get("common_investigators", [])
        self.label(frame, "Requesting Investigator", 5)
        if investigators:
            self.requesting_investigator = self.combo(frame, investigators, 5)
        else:
            self.requesting_investigator = self.entry(frame, 5)

        technicians = self.settings.get("common_technicians", [])
        default_technician = self.settings.get("default_technician", "")
        self.label(frame, "Technician", 6)
        if technicians:
            self.technician = self.combo(frame, technicians, 6)
        else:
            self.technician = self.entry(frame, 6, default=default_technician)

        today = datetime.now().strftime("%Y-%m-%d")

        self.label(frame, "Date Received", 7)
        self.date_received = self.entry(frame, 7, default=today)

        self.label(frame, "Time Received", 8)
        self.time_received = self.entry(frame, 8)

        self.label(frame, "Date Processed", 9)
        self.date_processed = self.entry(frame, 9, default=today)

        self.label(frame, "Time Processed", 10)
        self.time_processed = self.entry(frame, 10)

    def build_intake_tab(self):
        frame = self.intake_tab

        self.label(frame, "Drop-off Person", 0)
        self.dropoff_person = self.entry(frame, 0)

        self.label(frame, "Received From", 1)
        self.received_from = self.entry(frame, 1)

        self.label(frame, "Evidence Item Number", 2)
        self.evidence_item_number = self.entry(frame, 2)

        self.checked_out_from_evidence = self.checkbox(frame, "Checked out from evidence", 3)
        self.checked_out_from_evidence.set(True)

        self.label(frame, "Checked Out Date/Time", 4)
        self.checked_out_datetime = self.entry(frame, 4)

        self.returned_to_evidence = self.checkbox(frame, "Returned to evidence", 5)
        self.returned_to_evidence.set(True)

        self.label(frame, "Returned Date/Time", 6)
        self.returned_datetime = self.entry(frame, 6)

        locations = self.settings.get("common_evidence_locations", [])
        self.label(frame, "Evidence Location", 7)
        if locations:
            self.evidence_location = self.combo(frame, locations, 7)
        else:
            self.evidence_location = self.entry(frame, 7)

    def build_device_tab(self):
        frame = self.device_tab

        note = tk.Label(
            frame,
            text="v0.2 GUI starter supports one device entry. Multi-device GUI comes next.",
            bg=THEME["panel"],
            fg=THEME["muted"],
            font=("Segoe UI", 10, "italic")
        )
        note.grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=10)

        self.label(frame, "Device Type", 1)
        self.device_type = self.combo(frame, DEVICE_TYPES, 1)

        self.label(frame, "Quantity", 2)
        self.device_quantity = self.entry(frame, 2, default="1")

        self.label(frame, "Description", 3)
        self.device_description = self.entry(frame, 3)

        self.label(frame, "Make", 4)
        self.device_make = self.entry(frame, 4)

        self.label(frame, "Model", 5)
        self.device_model = self.entry(frame, 5)

        self.label(frame, "Serial / Identifier", 6)
        self.device_serial = self.entry(frame, 6)

        self.label(frame, "Storage Size Per Device", 7)
        self.device_capacity_size = self.entry(frame, 7)

        self.label(frame, "Storage Unit", 8)
        self.device_capacity_unit = self.combo(frame, STORAGE_UNITS, 8, default="GB")

    def build_tools_tab(self):
        frame = self.tools_tab

        note = tk.Label(
            frame,
            text="Select one common tool or enter a custom tool. Multi-tool GUI comes next.",
            bg=THEME["panel"],
            fg=THEME["muted"],
            font=("Segoe UI", 10, "italic")
        )
        note.grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=10)

        tool_names = [tool.get("name", "") for tool in self.settings.get("common_tools", [])]
        tool_names.append("Other")

        self.label(frame, "Tool Used", 1)
        self.tool_used = self.combo(frame, tool_names, 1)

        self.label(frame, "Tool Version", 2)
        self.tool_version = self.entry(frame, 2)

        self.label(frame, "Custom Tool Name", 3)
        self.custom_tool_name = self.entry(frame, 3)

    def build_output_tab(self):
        frame = self.output_tab

        self.label(frame, "Processing Type", 0)
        self.processing_type = self.combo(frame, PROCESSING_TYPES, 0)

        self.label(frame, "Processing Status", 1)
        self.processing_status = self.combo(frame, PROCESSING_STATUSES, 1)

        self.label(frame, "Processing Notes", 2)
        self.processing_notes = self.entry(frame, 2, width=70)

        self.label(frame, "Output Type", 3)
        self.output_type = self.combo(frame, OUTPUT_TYPES, 3)

        self.label(frame, "Output Filename / Identifier", 4)
        self.output_filename = self.entry(frame, 4, width=70)

        self.label(frame, "Output Location", 5)
        self.output_location = self.entry(frame, 5, width=70)

        self.reader_report_generated = self.checkbox(frame, "Reader report generated", 6)
        self.reader_report_generated.set(True)

        self.case_file_generated = self.checkbox(frame, "Case file generated", 7)

        self.label(frame, "Technician Notes", 8)
        self.technician_notes = self.entry(frame, 8, width=70)

    def build_generate_tab(self):
        frame = self.review_tab

        instructions = tk.Label(
            frame,
            text=(
                "Click Generate Packet to create the TXT acquisition packet, "
                "save JSON packet data, and append to the XLSX tracking workbook."
            ),
            bg=THEME["panel"],
            fg=THEME["text"],
            wraplength=760,
            justify="left",
            font=("Segoe UI", 11)
        )
        instructions.pack(anchor="w", padx=15, pady=15)

        generate_button = tk.Button(
            frame,
            text="Generate Packet",
            command=self.generate_packet,
            bg=THEME["accent"],
            fg=THEME["button_text"],
            activebackground=THEME["accent"],
            activeforeground=THEME["button_text"],
            relief="flat",
            padx=20,
            pady=10,
            font=("Segoe UI", 11, "bold")
        )
        generate_button.pack(anchor="w", padx=15, pady=10)

        self.status_text = tk.Text(
            frame,
            height=18,
            width=95,
            bg=THEME["input_bg"],
            fg=THEME["text"],
            insertbackground=THEME["text"],
            relief="flat"
        )
        self.status_text.pack(fill="both", expand=True, padx=15, pady=15)

    def get_widget_value(self, widget):
        if isinstance(widget, ttk.Combobox):
            return widget.get().strip()
        return widget.get().strip()

    def build_packet_from_form(self):
        case_number = self.get_widget_value(self.case_number)

        if not case_number:
            raise ValueError("Case Number is required.")

        quantity_text = self.get_widget_value(self.device_quantity)

        try:
            quantity = int(quantity_text)
            if quantity < 1:
                raise ValueError
        except ValueError:
            raise ValueError("Device quantity must be a whole number greater than zero.")

        capacity_text = self.get_widget_value(self.device_capacity_size)
        capacity_size = None

        if capacity_text:
            try:
                capacity_size = float(capacity_text)
            except ValueError:
                raise ValueError("Storage size must be a number or left blank.")

        capacity_unit = self.get_widget_value(self.device_capacity_unit)

        if not capacity_text:
            capacity_unit = "Unknown"

        storage_each_gb = convert_to_gb(capacity_size, capacity_unit)
        storage_total_gb = storage_each_gb * quantity if storage_each_gb is not None else None

        device = {
            "device_type": self.get_widget_value(self.device_type),
            "quantity": quantity,
            "description": self.get_widget_value(self.device_description),
            "make": self.get_widget_value(self.device_make),
            "model": self.get_widget_value(self.device_model),
            "serial": self.get_widget_value(self.device_serial),
            "capacity_size": capacity_size,
            "capacity_unit": capacity_unit,
            "storage_each_gb": storage_each_gb,
            "storage_total_gb": storage_total_gb
        }

        selected_tool = self.get_widget_value(self.tool_used)

        if selected_tool == "Other":
            tool_name = self.get_widget_value(self.custom_tool_name)
        else:
            tool_name = selected_tool

        tools_used = []

        if tool_name:
            tools_used.append({
                "name": tool_name,
                "version": self.get_widget_value(self.tool_version)
            })

        packet = {
            "app_name": APP_NAME,
            "app_version": APP_VERSION,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "department": {
                "department_name": self.settings.get("department_name", ""),
                "unit_name": self.settings.get("unit_name", "")
            },
            "general_info": {
                "case_number": case_number,
                "agency_case_number": self.get_widget_value(self.agency_case_number),
                "offense_or_incident": self.get_widget_value(self.offense_or_incident),
                "requesting_investigator": self.get_widget_value(self.requesting_investigator),
                "technician": self.get_widget_value(self.technician),
                "date_received": self.get_widget_value(self.date_received),
                "time_received": self.get_widget_value(self.time_received),
                "date_processed": self.get_widget_value(self.date_processed),
                "time_processed": self.get_widget_value(self.time_processed)
            },
            "intake_info": {
                "dropoff_person": self.get_widget_value(self.dropoff_person),
                "received_from": self.get_widget_value(self.received_from),
                "evidence_item_number": self.get_widget_value(self.evidence_item_number),
                "checked_out_from_evidence": self.checked_out_from_evidence.get(),
                "checked_out_datetime": self.get_widget_value(self.checked_out_datetime),
                "returned_to_evidence": self.returned_to_evidence.get(),
                "returned_datetime": self.get_widget_value(self.returned_datetime),
                "evidence_location": self.get_widget_value(self.evidence_location)
            },
            "subject": {
                "last_name": self.get_widget_value(self.subject_last),
                "first_name": self.get_widget_value(self.subject_first)
            },
            "devices": [device],
            "tools_used": tools_used,
            "processing": {
                "processing_type": self.get_widget_value(self.processing_type),
                "processing_status": self.get_widget_value(self.processing_status),
                "processing_notes": self.get_widget_value(self.processing_notes)
            },
            "output": {
                "output_type": self.get_widget_value(self.output_type),
                "output_filename": self.get_widget_value(self.output_filename),
                "output_location": self.get_widget_value(self.output_location),
                "reader_report_generated": self.reader_report_generated.get(),
                "case_file_generated": self.case_file_generated.get()
            },
            "technician_notes": self.get_widget_value(self.technician_notes),
            "scope_statement": self.settings.get("default_scope_statement", "")
        }

        packet["summary"] = {
            "device_counts": count_devices_by_type(packet["devices"]),
            "total_devices": sum(device["quantity"] for device in packet["devices"]),
            "total_storage_gb": calculate_total_storage_gb(packet["devices"])
        }

        return packet

    def generate_packet(self):
        try:
            packet = self.build_packet_from_form()
            txt_path, json_path = save_packet_outputs(packet)
            xlsx_path = append_to_fpr_tracking(packet)

            message = (
                "Packet generated successfully.\n\n"
                f"TXT report:\n{txt_path}\n\n"
                f"JSON packet:\n{json_path}\n\n"
                f"XLSX tracking workbook:\n{xlsx_path}\n\n"
                f"Total devices: {packet['summary']['total_devices']}\n"
                f"Total known storage: {format_storage_gb(packet['summary']['total_storage_gb'])}"
            )

            self.status_text.delete("1.0", tk.END)
            self.status_text.insert(tk.END, message)
            messagebox.showinfo("Packet Generated", "Acquisition packet generated successfully.")

        except Exception as error:
            messagebox.showerror("Error", str(error))


def main():
    root = tk.Tk()
    app = AcquisitionPacketGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()