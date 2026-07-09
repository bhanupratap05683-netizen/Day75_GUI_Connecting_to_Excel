"""
Day 75 - GUI: Connecting to Excel
Phase 6: GUI & File Management
Author: Bhanu Pratap Singh

A desktop tool: Upload Excel -> Python processes -> Download result.
Covers: tkinter window, filedialog (open/save), messagebox,
        pandas read/write Excel, status updates, error handling.
"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd


# -------------------------------------------------------------------
# Helper: create a sample practice Excel file (so you can test today)
# -------------------------------------------------------------------
def create_sample_file(path="sales_input.xlsx"):
    """Generate the practice sheet if it doesn't already exist."""
    if not os.path.exists(path):
        data = {
            "Product":    ["Laptop", "Mouse", "Keyboard", "Monitor", "Headphone"],
            "Units_Sold": [12, 90, 45, 20, 60],
            "Unit_Price": [55000, 450, 1200, 9500, 1800],
        }
        pd.DataFrame(data).to_excel(path, index=False)
        print(f"Sample file created: {path}")


# -------------------------------------------------------------------
# Main Application Class
# -------------------------------------------------------------------
class ExcelProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Day 75 - Excel Processor Tool")
        self.root.geometry("480x300")
        self.root.resizable(False, False)

        # Holds the path of the uploaded file (state between clicks)
        self.input_path = None

        # ---- Title label ----
        tk.Label(
            root, text="📊 Excel Sales Processor",
            font=("Arial", 16, "bold")
        ).pack(pady=15)

        # ---- Upload button ----
        tk.Button(
            root, text="1. Upload Excel File",
            width=25, command=self.upload_file, bg="#d6eaf8"
        ).pack(pady=5)

        # ---- Processing option dropdown ----
        tk.Label(root, text="Choose processing option:").pack(pady=(10, 0))
        self.option = ttk.Combobox(
            root, values=["Add Revenue Column", "Only Summary"],
            state="readonly", width=25
        )
        self.option.current(0)   # default selection
        self.option.pack(pady=5)

        # ---- Process button ----
        tk.Button(
            root, text="2. Process & Save Result",
            width=25, command=self.process_file, bg="#d5f5e3"
        ).pack(pady=10)

        # ---- Status label (updates dynamically) ----
        self.status = tk.StringVar()
        self.status.set("Status: Waiting for file...")
        tk.Label(
            root, textvariable=self.status,
            fg="blue", font=("Arial", 10)
        ).pack(pady=10)

    # ---------------------------------------------------------------
    # Step 1: Let user pick an Excel file
    # ---------------------------------------------------------------
    def upload_file(self):
        path = filedialog.askopenfilename(
            title="Select an Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if path:
            self.input_path = path
            filename = os.path.basename(path)
            self.status.set(f"Status: Loaded '{filename}' ✅")
        else:
            self.status.set("Status: No file selected.")

    # ---------------------------------------------------------------
    # Step 2: Read -> process -> save
    # ---------------------------------------------------------------
    def process_file(self):
        # Guard: make sure a file was uploaded
        if not self.input_path:
            messagebox.showerror("Error", "Please upload an Excel file first!")
            return

        try:
            # Read the uploaded Excel into a DataFrame
            df = pd.read_excel(self.input_path)

            choice = self.option.get()

            if choice == "Add Revenue Column":
                # Add calculated column
                df["Total_Revenue"] = df["Units_Sold"] * df["Unit_Price"]
                # Add a summary row
                summary = pd.DataFrame({
                    "Product": ["TOTAL"],
                    "Units_Sold": [df["Units_Sold"].sum()],
                    "Unit_Price": [""],
                    "Total_Revenue": [df["Total_Revenue"].sum()],
                })
                result = pd.concat([df, summary], ignore_index=True)

            else:  # "Only Summary"
                total_units = df["Units_Sold"].sum()
                total_rev = (df["Units_Sold"] * df["Unit_Price"]).sum()
                result = pd.DataFrame({
                    "Metric": ["Total Units Sold", "Total Revenue"],
                    "Value": [total_units, total_rev],
                })

            # Ask user where to save the output
            save_path = filedialog.asksaveasfilename(
                title="Save processed file as...",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )

            if not save_path:
                self.status.set("Status: Save cancelled.")
                return

            # Write result to Excel
            result.to_excel(save_path, index=False)

            self.status.set("Status: Done ✅ File saved!")
            messagebox.showinfo(
                "Success",
                f"Processed file saved:\n{os.path.basename(save_path)}"
            )

        except Exception as e:
            # Graceful error handling (Day 19 concept)
            messagebox.showerror("Processing Error", str(e))
            self.status.set("Status: Error occurred ❌")


# -------------------------------------------------------------------
# Run the app
# -------------------------------------------------------------------
if __name__ == "__main__":
    create_sample_file()          # make practice file for testing
    root = tk.Tk()
    app = ExcelProcessorApp(root)
    root.mainloop()