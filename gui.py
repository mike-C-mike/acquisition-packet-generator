import os
import subprocess
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
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
    save_settings,
    get_output_paths,
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

        self.settings = load_or_create_settings()
        ensure_directories(self.settings)

        self.devices = []
        self.tools_used = []
        
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

        settings_button = tk.Button(
            header,
            text="Settings",
            command=self.open_settings_window,
            bg=THEME["accent"],
            fg=THEME["button_text"],
            activebackground=THEME["accent"],
            activeforeground=THEME["button_text"],
            relief="flat",
            padx=14,
            pady=6,
            font=("Segoe UI", 10, "bold")
        )
        settings_button.pack(anchor="e", pady=(0, 5))

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
            text="Add one or more devices/media items. Added entries will appear in the table below.",
            bg=THEME["panel"],
            fg=THEME["muted"],
            font=("Segoe UI", 10, "italic")
        )
        note.grid(row=0, column=0, columnspan=4, sticky="w", padx=10, pady=10)

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

        button_frame = tk.Frame(frame, bg=THEME["panel"])
        button_frame.grid(row=9, column=0, columnspan=4, sticky="w", padx=10, pady=10)

        add_button = tk.Button(
            button_frame,
            text="Add Device",
            command=self.add_device_to_table,
            bg=THEME["accent"],
            fg=THEME["button_text"],
            activebackground=THEME["accent"],
            activeforeground=THEME["button_text"],
            relief="flat",
            padx=14,
            pady=6,
            font=("Segoe UI", 10, "bold")
        )
        add_button.pack(side="left", padx=(0, 10))

        remove_button = tk.Button(
            button_frame,
            text="Remove Selected Device",
            command=self.remove_selected_device,
            bg=THEME["input_bg"],
            fg=THEME["text"],
            activebackground=THEME["input_bg"],
            activeforeground=THEME["text"],
            relief="flat",
            padx=14,
            pady=6,
            font=("Segoe UI", 10)
        )
        remove_button.pack(side="left")

        table_frame = tk.Frame(frame, bg=THEME["panel"])
        table_frame.grid(row=10, column=0, columnspan=4, sticky="nsew", padx=10, pady=10)

        columns = (
            "type",
            "quantity",
            "description",
            "make",
            "model",
            "serial",
            "storage",
            "total_gb"
        )

        self.device_tree = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings",
            height=8
        )

        self.device_tree.heading("type", text="Type")
        self.device_tree.heading("quantity", text="Qty")
        self.device_tree.heading("description", text="Description")
        self.device_tree.heading("make", text="Make")
        self.device_tree.heading("model", text="Model")
        self.device_tree.heading("serial", text="Serial / ID")
        self.device_tree.heading("storage", text="Storage Each")
        self.device_tree.heading("total_gb", text="Total GB")

        self.device_tree.column("type", width=120)
        self.device_tree.column("quantity", width=50, anchor="center")
        self.device_tree.column("description", width=160)
        self.device_tree.column("make", width=100)
        self.device_tree.column("model", width=100)
        self.device_tree.column("serial", width=130)
        self.device_tree.column("storage", width=100)
        self.device_tree.column("total_gb", width=100)

        scrollbar = ttk.Scrollbar(
            table_frame,
            orient="vertical",
            command=self.device_tree.yview
        )
        self.device_tree.configure(yscrollcommand=scrollbar.set)

        self.device_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        frame.grid_rowconfigure(10, weight=1)
        frame.grid_columnconfigure(1, weight=1)

    def clear_device_fields(self):
        self.device_quantity.delete(0, tk.END)
        self.device_quantity.insert(0, "1")

        self.device_description.delete(0, tk.END)
        self.device_make.delete(0, tk.END)
        self.device_model.delete(0, tk.END)
        self.device_serial.delete(0, tk.END)
        self.device_capacity_size.delete(0, tk.END)

        if DEVICE_TYPES:
            self.device_type.set(DEVICE_TYPES[0])

        self.device_capacity_unit.set("GB")

    def add_device_to_table(self):
        quantity_text = self.get_widget_value(self.device_quantity)

        try:
            quantity = int(quantity_text)
            if quantity < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror(
                "Invalid Device Quantity",
                "Device quantity must be a whole number greater than zero."
            )
            return

        capacity_text = self.get_widget_value(self.device_capacity_size)
        capacity_size = None

        if capacity_text:
            try:
                capacity_size = float(capacity_text)
            except ValueError:
                messagebox.showerror(
                    "Invalid Storage Size",
                    "Storage size must be a number or left blank."
                )
                return

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

        self.devices.append(device)

        storage_display = "Unknown"
        if capacity_size is not None:
            storage_display = f"{capacity_size:g} {capacity_unit}"

        total_gb_display = "Unknown"
        if storage_total_gb is not None:
            total_gb_display = f"{storage_total_gb:,.2f}"

        self.device_tree.insert(
            "",
            tk.END,
            values=(
                device["device_type"],
                device["quantity"],
                device["description"],
                device["make"],
                device["model"],
                device["serial"],
                storage_display,
                total_gb_display
            )
        )

        self.clear_device_fields()

    def remove_selected_device(self):
        selected_items = self.device_tree.selection()

        if not selected_items:
            messagebox.showwarning(
                "No Device Selected",
                "Select a device row to remove."
            )
            return

        for item in selected_items:
            row_index = self.device_tree.index(item)
            self.device_tree.delete(item)

            if 0 <= row_index < len(self.devices):
                self.devices.pop(row_index)

    def build_tools_tab(self):
        frame = self.tools_tab

        note = tk.Label(
            frame,
            text="Add one or more tools used during acquisition, extraction, imaging, or output generation.",
            bg=THEME["panel"],
            fg=THEME["muted"],
            font=("Segoe UI", 10, "italic")
        )
        note.grid(row=0, column=0, columnspan=4, sticky="w", padx=10, pady=10)

        tool_names = [tool.get("name", "") for tool in self.settings.get("common_tools", [])]
        tool_names.append("Other")

        self.label(frame, "Tool Used", 1)
        self.tool_used = self.combo(frame, tool_names, 1)

        self.label(frame, "Tool Version", 2)
        self.tool_version = self.entry(frame, 2)

        self.label(frame, "Custom Tool Name", 3)
        self.custom_tool_name = self.entry(frame, 3)

        button_frame = tk.Frame(frame, bg=THEME["panel"])
        button_frame.grid(row=4, column=0, columnspan=4, sticky="w", padx=10, pady=10)

        add_button = tk.Button(
            button_frame,
            text="Add Tool",
            command=self.add_tool_to_table,
            bg=THEME["accent"],
            fg=THEME["button_text"],
            activebackground=THEME["accent"],
            activeforeground=THEME["button_text"],
            relief="flat",
            padx=14,
            pady=6,
            font=("Segoe UI", 10, "bold")
        )
        add_button.pack(side="left", padx=(0, 10))

        remove_button = tk.Button(
            button_frame,
            text="Remove Selected Tool",
            command=self.remove_selected_tool,
            bg=THEME["input_bg"],
            fg=THEME["text"],
            activebackground=THEME["input_bg"],
            activeforeground=THEME["text"],
            relief="flat",
            padx=14,
            pady=6,
            font=("Segoe UI", 10)
        )
        remove_button.pack(side="left")

        table_frame = tk.Frame(frame, bg=THEME["panel"])
        table_frame.grid(row=5, column=0, columnspan=4, sticky="nsew", padx=10, pady=10)

        columns = (
            "name",
            "version"
        )

        self.tool_tree = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings",
            height=10
        )

        self.tool_tree.heading("name", text="Tool")
        self.tool_tree.heading("version", text="Version")

        self.tool_tree.column("name", width=300)
        self.tool_tree.column("version", width=160)

        scrollbar = ttk.Scrollbar(
            table_frame,
            orient="vertical",
            command=self.tool_tree.yview
        )
        self.tool_tree.configure(yscrollcommand=scrollbar.set)

        self.tool_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        frame.grid_rowconfigure(5, weight=1)
        frame.grid_columnconfigure(1, weight=1)

    def clear_tool_fields(self):
        if self.tool_used["values"]:
            self.tool_used.set(self.tool_used["values"][0])

        self.tool_version.delete(0, tk.END)
        self.custom_tool_name.delete(0, tk.END)

    def add_tool_to_table(self):
        selected_tool = self.get_widget_value(self.tool_used)

        if selected_tool == "Other":
            tool_name = self.get_widget_value(self.custom_tool_name)
        else:
            tool_name = selected_tool

        tool_version = self.get_widget_value(self.tool_version)

        if not tool_name:
            messagebox.showerror(
                "Missing Tool Name",
                "Select a tool or enter a custom tool name."
            )
            return

        tool = {
            "name": tool_name,
            "version": tool_version
        }

        self.tools_used.append(tool)

        self.tool_tree.insert(
            "",
            tk.END,
            values=(
                tool["name"],
                tool["version"]
            )
        )

        self.clear_tool_fields()

    def remove_selected_tool(self):
        selected_items = self.tool_tree.selection()

        if not selected_items:
            messagebox.showwarning(
                "No Tool Selected",
                "Select a tool row to remove."
            )
            return

        for item in selected_items:
            row_index = self.tool_tree.index(item)
            self.tool_tree.delete(item)

            if 0 <= row_index < len(self.tools_used):
                self.tools_used.pop(row_index)

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

        button_frame = tk.Frame(frame, bg=THEME["panel"])
        button_frame.pack(anchor="w", padx=15, pady=10)
        
        generate_button = tk.Button(
            button_frame,
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
        generate_button.pack(side="left", padx=(0, 10))

        open_folder_button = tk.Button(
            button_frame,
            text="Open Output Folder",
            command=self.open_output_folder,
            bg=THEME["input_bg"],
            fg=THEME["text"],
            activebackground=THEME["input_bg"],
            activeforeground=THEME["text"],
            relief="flat",
            padx=20,
            pady=10,
            font=("Segoe UI", 11)
        )

        clear_button = tk.Button(
            button_frame,
            text="New / Clear Packet",
            command=self.clear_form,
            bg=THEME["input_bg"],
            fg=THEME["text"],
            activebackground=THEME["input_bg"],
            activeforeground=THEME["text"],
            relief="flat",
            padx=20,
            pady=10,
            font=("Segoe UI", 11)
        )
        clear_button.pack(side="left", padx=(10, 0))

        open_folder_button.pack(side="left")

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

    def clear_form(self):
        """
        Clear the current form so a new acquisition packet can be started.

        This does not delete settings, generated reports, saved packets,
        or the tracking workbook.
        """
        confirm = messagebox.askyesno(
            "Clear Current Packet",
            "Clear the current form and start a new packet?\n\n"
            "This will not delete settings or generated output files."
        )

        if not confirm:
            return

        today = datetime.now().strftime("%Y-%m-%d")

        # General Info
        self.case_number.delete(0, tk.END)
        self.agency_case_number.delete(0, tk.END)
        self.offense_or_incident.delete(0, tk.END)
        self.subject_last.delete(0, tk.END)
        self.subject_first.delete(0, tk.END)

        if isinstance(self.requesting_investigator, ttk.Combobox):
            values = self.requesting_investigator["values"]
            if values:
                self.requesting_investigator.set(values[0])
            else:
                self.requesting_investigator.set("")
        else:
            self.requesting_investigator.delete(0, tk.END)

        if isinstance(self.technician, ttk.Combobox):
            default_technician = self.settings.get("default_technician", "")
            values = self.technician["values"]

            if default_technician and default_technician in values:
                self.technician.set(default_technician)
            elif values:
                self.technician.set(values[0])
            else:
                self.technician.set("")
        else:
            self.technician.delete(0, tk.END)
            self.technician.insert(0, self.settings.get("default_technician", ""))

        self.date_received.delete(0, tk.END)
        self.date_received.insert(0, today)

        self.time_received.delete(0, tk.END)

        self.date_processed.delete(0, tk.END)
        self.date_processed.insert(0, today)

        self.time_processed.delete(0, tk.END)

        # Intake
        self.dropoff_person.delete(0, tk.END)
        self.received_from.delete(0, tk.END)
        self.evidence_item_number.delete(0, tk.END)

        self.checked_out_from_evidence.set(True)
        self.checked_out_datetime.delete(0, tk.END)

        self.returned_to_evidence.set(True)
        self.returned_datetime.delete(0, tk.END)

        if isinstance(self.evidence_location, ttk.Combobox):
            values = self.evidence_location["values"]
            if values:
                self.evidence_location.set(values[0])
            else:
                self.evidence_location.set("")
        else:
            self.evidence_location.delete(0, tk.END)

        # Device entry fields and table
        self.clear_device_fields()
        self.devices = []

        for item in self.device_tree.get_children():
            self.device_tree.delete(item)

        # Tool entry fields and table
        self.clear_tool_fields()
        self.tools_used = []

        for item in self.tool_tree.get_children():
            self.tool_tree.delete(item)

        # Processing / Output
        if PROCESSING_TYPES:
            self.processing_type.set(PROCESSING_TYPES[0])

        if PROCESSING_STATUSES:
            self.processing_status.set(PROCESSING_STATUSES[0])

        self.processing_notes.delete(0, tk.END)

        if OUTPUT_TYPES:
            self.output_type.set(OUTPUT_TYPES[0])

        self.output_filename.delete(0, tk.END)
        self.output_location.delete(0, tk.END)

        self.reader_report_generated.set(True)
        self.case_file_generated.set(False)

        self.technician_notes.delete(0, tk.END)

        # Generate/status area
        self.status_text.delete("1.0", tk.END)
        self.status_text.insert(
            tk.END,
            "Ready for a new acquisition packet."
        )

        # Return user to the first tab
        self.notebook.select(self.general_tab)

    def build_packet_from_form(self):
        case_number = self.get_widget_value(self.case_number)

        if not case_number:
            raise ValueError("Case Number is required.")

        if not self.devices:
            raise ValueError("At least one device/media entry must be added before generating a packet.")

        if not self.tools_used:
            raise ValueError("At least one tool must be added before generating a packet.")

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
            "devices": self.devices,
            "tools_used": self.tools_used,
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
    
    def open_settings_window(self):
        settings_window = tk.Toplevel(self.root)
        settings_window.title("Settings")
        settings_window.geometry("850x650")
        settings_window.configure(bg=THEME["bg"])
        settings_window.grab_set()

        notebook = ttk.Notebook(settings_window)
        notebook.pack(fill="both", expand=True, padx=15, pady=15)

        output_tab = tk.Frame(notebook, bg=THEME["panel"])
        customization_tab = tk.Frame(notebook, bg=THEME["panel"])

        notebook.add(output_tab, text="Output / Storage")
        notebook.add(customization_tab, text="Customization / Defaults")

        self.build_output_settings_tab(output_tab)
        self.build_customization_settings_tab(customization_tab)

        button_frame = tk.Frame(settings_window, bg=THEME["bg"])
        button_frame.pack(fill="x", padx=15, pady=(0, 15))

        save_button = tk.Button(
            button_frame,
            text="Save Settings",
            command=lambda: self.save_settings_window(settings_window),
            bg=THEME["accent"],
            fg=THEME["button_text"],
            activebackground=THEME["accent"],
            activeforeground=THEME["button_text"],
            relief="flat",
            padx=18,
            pady=8,
            font=("Segoe UI", 10, "bold")
        )
        save_button.pack(side="right")

    def build_output_settings_tab(self, frame):
        output_paths = self.settings.get("output_paths", {})

        self.settings_base_output_dir = tk.StringVar(
            value=output_paths.get("base_output_dir", "")
        )
        self.settings_reports_folder = tk.StringVar(
            value=output_paths.get("reports_folder_name", "output")
        )
        self.settings_saved_packets_folder = tk.StringVar(
            value=output_paths.get("saved_packets_folder_name", "saved_packets")
        )
        self.settings_tracking_workbook = tk.StringVar(
            value=output_paths.get("tracking_workbook_name", "fpr_tracking.xlsx")
        )

        self.label(frame, "Base Output Folder", 0)
        base_entry = tk.Entry(
            frame,
            textvariable=self.settings_base_output_dir,
            width=65,
            bg=THEME["input_bg"],
            fg=THEME["text"],
            insertbackground=THEME["text"],
            relief="flat"
        )
        base_entry.grid(row=0, column=1, sticky="w", padx=10, pady=6)

        browse_button = tk.Button(
            frame,
            text="Browse",
            command=self.browse_base_output_folder,
            bg=THEME["accent"],
            fg=THEME["button_text"],
            activebackground=THEME["accent"],
            activeforeground=THEME["button_text"],
            relief="flat",
            padx=10
        )
        browse_button.grid(row=0, column=2, sticky="w", padx=5, pady=6)

        self.label(frame, "Reports Folder Name", 1)
        reports_entry = tk.Entry(
            frame,
            textvariable=self.settings_reports_folder,
            width=45,
            bg=THEME["input_bg"],
            fg=THEME["text"],
            insertbackground=THEME["text"],
            relief="flat"
        )
        reports_entry.grid(row=1, column=1, sticky="w", padx=10, pady=6)

        self.label(frame, "Saved Packets Folder Name", 2)
        packets_entry = tk.Entry(
            frame,
            textvariable=self.settings_saved_packets_folder,
            width=45,
            bg=THEME["input_bg"],
            fg=THEME["text"],
            insertbackground=THEME["text"],
            relief="flat"
        )
        packets_entry.grid(row=2, column=1, sticky="w", padx=10, pady=6)

        self.label(frame, "Tracking Workbook Name", 3)
        workbook_entry = tk.Entry(
            frame,
            textvariable=self.settings_tracking_workbook,
            width=45,
            bg=THEME["input_bg"],
            fg=THEME["text"],
            insertbackground=THEME["text"],
            relief="flat"
        )
        workbook_entry.grid(row=3, column=1, sticky="w", padx=10, pady=6)

        paths = get_output_paths(self.settings)

        preview = tk.Label(
            frame,
            text=(
                "Current resolved paths:\n"
                f"Reports: {paths['reports_dir']}\n"
                f"Saved packets: {paths['saved_packets_dir']}\n"
                f"Tracking workbook: {paths['tracking_workbook_path']}"
            ),
            bg=THEME["panel"],
            fg=THEME["muted"],
            justify="left",
            wraplength=760
        )
        preview.grid(row=4, column=0, columnspan=3, sticky="w", padx=10, pady=20)

    def build_customization_settings_tab(self, frame):
        self.settings_department_name = tk.StringVar(
            value=self.settings.get("department_name", "")
        )
        self.settings_unit_name = tk.StringVar(
            value=self.settings.get("unit_name", "")
        )
        self.settings_default_technician = tk.StringVar(
            value=self.settings.get("default_technician", "")
        )

        self.label(frame, "Department Name", 0)
        self.settings_entry_department = tk.Entry(
            frame,
            textvariable=self.settings_department_name,
            width=60,
            bg=THEME["input_bg"],
            fg=THEME["text"],
            insertbackground=THEME["text"],
            relief="flat"
        )
        self.settings_entry_department.grid(row=0, column=1, sticky="w", padx=10, pady=6)

        self.label(frame, "Unit Name", 1)
        self.settings_entry_unit = tk.Entry(
            frame,
            textvariable=self.settings_unit_name,
            width=60,
            bg=THEME["input_bg"],
            fg=THEME["text"],
            insertbackground=THEME["text"],
            relief="flat"
        )
        self.settings_entry_unit.grid(row=1, column=1, sticky="w", padx=10, pady=6)

        self.label(frame, "Default Technician", 2)
        self.settings_entry_default_tech = tk.Entry(
            frame,
            textvariable=self.settings_default_technician,
            width=60,
            bg=THEME["input_bg"],
            fg=THEME["text"],
            insertbackground=THEME["text"],
            relief="flat"
        )
        self.settings_entry_default_tech.grid(row=2, column=1, sticky="w", padx=10, pady=6)

        self.label(frame, "Technicians\none per line", 3)
        self.settings_text_technicians = self.settings_textbox(frame, 3)
        self.settings_text_technicians.insert(
            "1.0",
            "\n".join(self.settings.get("common_technicians", []))
        )

        self.label(frame, "Investigators\none per line", 4)
        self.settings_text_investigators = self.settings_textbox(frame, 4)
        self.settings_text_investigators.insert(
            "1.0",
            "\n".join(self.settings.get("common_investigators", []))
        )

        self.label(frame, "Evidence Locations\none per line", 5)
        self.settings_text_locations = self.settings_textbox(frame, 5)
        self.settings_text_locations.insert(
            "1.0",
            "\n".join(self.settings.get("common_evidence_locations", []))
        )

        self.label(frame, "Tools\nName | Version", 6)
        self.settings_text_tools = self.settings_textbox(frame, 6)

        tool_lines = []
        for tool in self.settings.get("common_tools", []):
            name = tool.get("name", "")
            version = tool.get("version", "")

            if version:
                tool_lines.append(f"{name} | {version}")
            else:
                tool_lines.append(name)

        self.settings_text_tools.insert("1.0", "\n".join(tool_lines))

    def settings_textbox(self, parent, row):
        textbox = tk.Text(
            parent,
            height=5,
            width=60,
            bg=THEME["input_bg"],
            fg=THEME["text"],
            insertbackground=THEME["text"],
            relief="flat"
        )
        textbox.grid(row=row, column=1, sticky="w", padx=10, pady=6)
        return textbox

    def browse_base_output_folder(self):
        folder = filedialog.askdirectory(title="Select Base Output Folder")

        if folder:
            self.settings_base_output_dir.set(folder)

    def get_textbox_lines(self, textbox):
        raw_text = textbox.get("1.0", tk.END)
        lines = []

        for line in raw_text.splitlines():
            clean_line = line.strip()

            if clean_line:
                lines.append(clean_line)

        return lines

    def parse_tools_from_settings_textbox(self):
        lines = self.get_textbox_lines(self.settings_text_tools)
        tools = []

        for line in lines:
            if "|" in line:
                name, version = line.split("|", 1)
                tools.append({
                    "name": name.strip(),
                    "version": version.strip()
                })
            else:
                tools.append({
                    "name": line.strip(),
                    "version": ""
                })

        return tools

    def save_settings_window(self, settings_window):
        self.settings["department_name"] = self.settings_department_name.get().strip()
        self.settings["unit_name"] = self.settings_unit_name.get().strip()
        self.settings["default_technician"] = self.settings_default_technician.get().strip()

        self.settings["output_paths"] = {
            "base_output_dir": self.settings_base_output_dir.get().strip(),
            "reports_folder_name": self.settings_reports_folder.get().strip() or "output",
            "saved_packets_folder_name": self.settings_saved_packets_folder.get().strip() or "saved_packets",
            "tracking_workbook_name": self.settings_tracking_workbook.get().strip() or "fpr_tracking.xlsx"
        }

        self.settings["common_technicians"] = self.get_textbox_lines(self.settings_text_technicians)
        self.settings["common_investigators"] = self.get_textbox_lines(self.settings_text_investigators)
        self.settings["common_evidence_locations"] = self.get_textbox_lines(self.settings_text_locations)
        self.settings["common_tools"] = self.parse_tools_from_settings_textbox()

        save_settings(self.settings)
        ensure_directories(self.settings)

        messagebox.showinfo(
            "Settings Saved",
            "Settings saved successfully. Restart the app to refresh dropdowns."
        )

        settings_window.destroy()

    def open_output_folder(self):
        try:
            paths = get_output_paths(self.settings)
            base_path = paths["base_path"]

            ensure_directories(self.settings)

            if sys.platform.startswith("win"):
                os.startfile(base_path)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(base_path)])
            else:
                subprocess.Popen(["xdg-open", str(base_path)])

        except Exception as error:
            messagebox.showerror(
                "Open Output Folder Error",
                f"Unable to open output folder:\n\n{error}"
            )

    def generate_packet(self):
        try:
            packet = self.build_packet_from_form()
            txt_path, json_path = save_packet_outputs(packet, self.settings)
            xlsx_path = append_to_fpr_tracking(packet, self.settings)

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