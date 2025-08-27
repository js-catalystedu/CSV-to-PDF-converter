import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import inch
import os
import sys

class CSVtoPDFConverter(TkinterDnD.Tk):
    """
    A GUI application for converting CSV files to PDF format.
    Supports drag-and-drop and file browsing.
    """
    def __init__(self):
        super().__init__()
        self.title("CSV to PDF Converter")
        self.geometry("600x450")
        self.configure(bg="#2C3E50")

        # Set an icon for the window
        # To make this work after packaging with PyInstaller, the icon file needs to be bundled.
        # We'll use a try-except block to handle cases where the icon isn't found.
        try:
            # When running as a script
            icon_path = 'icon.ico'
            if not os.path.exists(icon_path):
                 # When running as a bundled exe
                icon_path = os.path.join(sys._MEIPASS, 'icon.ico')
            self.iconbitmap(icon_path)
        except Exception:
            # If the icon is not found, the app will run without it.
            print("Icon not found. Skipping.")
            pass


        self.file_path = None

        # --- UI Elements ---
        self._create_widgets()

    def _create_widgets(self):
        """Creates and places the widgets in the main window."""
        main_frame = tk.Frame(self, bg="#2C3E50")
        main_frame.pack(expand=True, fill="both", padx=20, pady=20)

        # --- Drop Zone ---
        self.drop_zone = tk.Label(
            main_frame,
            text="Drag & Drop CSV File Here",
            font=("Helvetica", 16, "bold"),
            bg="#34495E",
            fg="#ECF0F1",
            relief="solid", # CORRECTED: "dashed" is not a valid relief style in all Tk versions.
            bd=2,
            pady=60
        )
        self.drop_zone.pack(fill="x", pady=(0, 20))

        # Enable dropping files onto the drop_zone
        self.drop_zone.drop_target_register(DND_FILES)
        self.drop_zone.dnd_bind('<<Drop>>', self._on_drop)

        # --- "Or" Separator ---
        separator_label = tk.Label(
            main_frame,
            text="OR",
            font=("Helvetica", 12),
            bg="#2C3E50",
            fg="#BDC3C7"
        )
        separator_label.pack(pady=5)

        # --- Browse Button ---
        self.browse_button = tk.Button(
            main_frame,
            text="Browse for CSV File",
            font=("Helvetica", 12, "bold"),
            bg="#3498DB",
            fg="#ECF0F1",
            activebackground="#2980B9",
            activeforeground="#ECF0F1",
            relief="flat",
            padx=20,
            pady=10,
            command=self._browse_file
        )
        self.browse_button.pack(pady=(10, 20))

        # --- File Info Label ---
        self.file_info_label = tk.Label(
            main_frame,
            text="No file selected",
            font=("Helvetica", 10),
            bg="#2C3E50",
            fg="#ECF0F1",
            wraplength=550
        )
        self.file_info_label.pack(pady=(0, 20))

        # --- Convert Button ---
        self.convert_button = tk.Button(
            main_frame,
            text="Convert to PDF",
            font=("Helvetica", 14, "bold"),
            bg="#27AE60",
            fg="#ECF0F1",
            activebackground="#229954",
            activeforeground="#ECF0F1",
            relief="flat",
            padx=30,
            pady=15,
            state="disabled",
            command=self._convert_to_pdf
        )
        self.convert_button.pack(pady=10)

        # --- Status Bar ---
        self.status_bar = tk.Label(
            self,
            text="Ready",
            font=("Helvetica", 10),
            bg="#1C2833",
            fg="#ECF0F1",
            relief="sunken",
            anchor="w",
            padx=10
        )
        self.status_bar.pack(side="bottom", fill="x")

    def _on_drop(self, event):
        """Handles the file drop event."""
        # The event.data contains a string of file paths, sometimes in curly braces
        file_path = event.data.strip('{}')
        if file_path.lower().endswith('.csv'):
            self._update_file_path(file_path)
        else:
            messagebox.showerror("Invalid File", "Please drop a .csv file.")
            self._update_status("Error: Invalid file type.", "red")

    def _browse_file(self):
        """Opens a file dialog to select a CSV file."""
        file_path = filedialog.askopenfilename(
            filetypes=[("CSV Files", "*.csv"), ("All files", "*.*")]
        )
        if file_path:
            self._update_file_path(file_path)

    def _update_file_path(self, path):
        """Updates the UI with the selected file path."""
        self.file_path = path
        self.file_info_label.config(text=f"Selected: {os.path.basename(path)}")
        self.convert_button.config(state="normal")
        self._update_status(f"File '{os.path.basename(path)}' loaded.", "white")

    def _update_status(self, message, color="white"):
        """Updates the status bar message and color."""
        self.status_bar.config(text=message, fg=color)
        self.update_idletasks()

    def _convert_to_pdf(self):
        """Performs the CSV to PDF conversion."""
        if not self.file_path:
            messagebox.showerror("Error", "No file selected.")
            return

        try:
            self._update_status("Reading CSV file...", "#F1C40F")
            # Read CSV data using pandas
            df = pd.read_csv(self.file_path)
            # Replace NaN values with empty strings for cleaner PDF output
            df.fillna('', inplace=True)

            # Prepare data for the ReportLab table (headers + data)
            data = [df.columns.tolist()] + df.values.tolist()

            # Define output PDF path
            output_pdf_path = os.path.splitext(self.file_path)[0] + ".pdf"
            
            self._update_status("Generating PDF...", "#F1C40F")

            # Create the PDF document
            doc = SimpleDocTemplate(output_pdf_path, pagesize=(20*inch, 15*inch)) # Larger page size for wide tables
            elements = []
            styles = getSampleStyleSheet()

            # Add a title
            title = Paragraph(f"Report for {os.path.basename(self.file_path)}", styles['h1'])
            elements.append(title)
            elements.append(Spacer(1, 0.2*inch))

            # Create the table
            table = Table(data, repeatRows=1) # Repeat headers on each page

            # Define table style
            style = TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#4A6B8A")), # Header background
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'), # Header font
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor("#D0D3D4")), # Body background
                ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ])
            table.setStyle(style)

            elements.append(table)
            doc.build(elements)

            self._update_status(f"Successfully converted to {os.path.basename(output_pdf_path)}", "#2ECC71")
            messagebox.showinfo("Success", f"PDF created successfully at:\n{output_pdf_path}")

        except Exception as e:
            self._update_status(f"An error occurred: {e}", "red")
            messagebox.showerror("Conversion Error", f"An error occurred during conversion:\n{e}")

if __name__ == "__main__":
    app = CSVtoPDFConverter()
    app.mainloop()
