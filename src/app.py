import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from etl import transform_excel
import pandas as pd
import os

class ExcelETLApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel ETL Tool")
        self.geometry("600x600")

        # Destination headers loaded from template file (initially empty)
        self.new_headers = []

        # Variables
        self.old_file_path = ""
        self.template_file_path = ""
        self.mapping_widgets = {}

        # UI Elements
        self.main_frame = tk.Frame(self)
        self.main_frame.pack(padx=10, pady=10, fill='x')

        # --- File Selection ---
        file_frame = tk.Frame(self.main_frame)
        file_frame.pack(fill='x', pady=5)
        
        self.old_file_button = tk.Button(file_frame, text="Upload Source Excel File", command=self.upload_old_file)
        self.old_file_button.pack(side='left', expand=True, fill='x')

        self.template_file_button = tk.Button(file_frame, text="Upload Template (new) File", command=self.upload_template_file)
        self.template_file_button.pack(side='left', expand=True, fill='x', padx=(5,0))
        
        self.file_label = tk.Label(self.main_frame, text="No source file selected", relief=tk.SUNKEN)
        self.file_label.pack(fill='x', pady=5)

        self.template_label = tk.Label(self.main_frame, text="No template file selected", relief=tk.SUNKEN)
        self.template_label.pack(fill='x', pady=5)

        # --- New Filename ---
        filename_frame = tk.Frame(self.main_frame)
        filename_frame.pack(fill='x', pady=5)
        
        tk.Label(filename_frame, text="New Filename:").pack(side='left')
        self.new_filename_entry = tk.Entry(filename_frame)
        self.new_filename_entry.pack(side='left', expand=True, fill='x', padx=5)
        tk.Label(filename_frame, text=".xlsx").pack(side='left')

        # --- Mapping UI ---
        self.canvas = tk.Canvas(self)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.mapping_frame = ttk.Frame(self.canvas)

        self.mapping_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.canvas.create_window((0, 0), window=self.mapping_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        self.scrollbar.pack(side="right", fill="y")

        # --- Transform Button ---
        self.transform_button = tk.Button(self.main_frame, text="Transform and Save", command=self.transform)
        self.transform_button.pack(pady=10, fill='x')

    def upload_old_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            self.old_file_path = file_path
            self.file_label.config(text=f"Source: {os.path.basename(file_path)}")
            self.create_mapping_ui()

    def upload_template_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            self.template_file_path = file_path
            self.template_label.config(text=f"Template: {os.path.basename(file_path)}")
            # Load headers from template (use first sheet only)
            try:
                df = pd.read_excel(self.template_file_path, nrows=0)
                self.new_headers = df.columns.tolist()
                # If mapping UI already present, refresh dropdowns
                self.create_mapping_ui()
            except Exception as e:
                messagebox.showerror("Error", f"Could not read template headers:\n{e}")

    def create_mapping_ui(self):
        for widget in self.mapping_frame.winfo_children():
            widget.destroy()

        if not self.old_file_path:
            return

        try:
            xls = pd.ExcelFile(self.old_file_path)
            old_headers = set()
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name, nrows=0)
                old_headers.update(df.columns.tolist())
            
            sorted_headers = sorted(list(old_headers))

            self.mapping_widgets = {}

            # If template headers are not loaded yet, show a message
            if not self.new_headers:
                tk.Label(self.mapping_frame, text="Upload a template (new) file to load destination headers.", foreground='red').pack(pady=5)

            for old_col in sorted_headers:
                frame = tk.Frame(self.mapping_frame)
                frame.pack(fill='x', pady=2, padx=5)
                
                label = tk.Label(frame, text=f"{old_col}:", width=25, anchor='w')
                label.pack(side='left')

                variable = tk.StringVar(self)
                
                # Dropdown will contain the new headers loaded from template
                dropdown_values = [""] + (self.new_headers if self.new_headers else [])
                dropdown = ttk.Combobox(frame, textvariable=variable, values=dropdown_values, state="readonly")
                dropdown.set("") # Default to empty
                dropdown.pack(side='left', expand=True, fill='x')
                
                self.mapping_widgets[old_col] = variable

        except Exception as e:
            messagebox.showerror("Error Reading File", f"Could not read headers from the file.\n\n{e}")

    def transform(self):
        if not self.old_file_path:
            messagebox.showwarning("Warning", "Please upload a source file.")
            return

        if not self.template_file_path:
            messagebox.showwarning("Warning", "Please upload a template (new) file first.")
            return

        new_filename = self.new_filename_entry.get().strip()
        if not new_filename:
            messagebox.showwarning("Warning", "Please enter a filename for the new file.")
            return

        # The mapping is {old_column: new_column}
        mapping = {old_col: var.get() for old_col, var in self.mapping_widgets.items() if var.get()}
        
        if not mapping:
            if not messagebox.askyesno("Confirmation", "No columns are mapped. This will create an empty file with only headers. Continue?"):
                return

        try:
            output_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'transformed'))
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            output_path = os.path.join(output_dir, f"{new_filename}.xlsx")

            transform_excel(self.old_file_path, output_path, mapping, self.new_headers)
            messagebox.showinfo("Success", f"Transformation Complete!\nFile saved to:\n{output_path}")
        except Exception as e:
            messagebox.showerror("Error During Transformation", str(e))

if __name__ == "__main__":
    app = ExcelETLApp()
    app.mainloop()