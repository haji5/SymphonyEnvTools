import os
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import sys

import pandas as pd
import pythoncom
import win32com.client


def get_executable_dir():
    """
    Get the directory where the executable is located.
    Works for both PyInstaller bundles and regular Python scripts.
    """
    if getattr(sys, 'frozen', False):
        # Running as a PyInstaller bundle
        return os.path.dirname(sys.executable)
    else:
        # Running as a regular Python script
        return os.path.dirname(os.path.abspath(__file__))


class PRIZMReportGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PRIZM Report Generator")
        self.root.geometry("900x650")

        # Configure high DPI awareness and scaling
        try:
            self.root.tk.call('tk', 'scaling', 1.5)
        except:
            pass

        # Configure style for better appearance
        style = ttk.Style()
        style.theme_use('clam')

        # PRIZM Set Name to Identifier Code Mapping (1-67)
        self.prizm_set_to_code = {
            "The A-List": 1,
            "Wealthy & Wise": 2,
            "Asian Sophisticates": 3,
            "Turbo Burbs": 4,
            "First-Class Families": 5,
            "Downtown Verve": 6,
            "Mature & Secure": 7,
            "Multiculture-ish": 8,
            "Boomer Bliss": 9,
            "Asian Achievement": 10,
            "Modern Suburbia": 11,
            "Eat, Play, Love": 12,
            "Vie de Rêve": 13,
            "Kick-Back Country": 14,
            "South Asian Enterprise": 15,
            "Savvy Seniors": 16,
            "Asian Avenues": 17,
            "Multicultural Corners": 18,
            "Family Mode": 19,
            "New Asian Heights": 20,
            "Scenic Retirement": 21,
            "Indieville": 22,
            "Mid-City Mellow": 23,
            "All-Terrain Families": 24,
            "Suburban Sports": 25,
            "Country Traditions": 26,
            "Diversité Nouvelle": 27,
            "Latte Life": 28,
            "C'est Tiguidou": 29,
            "South Asian Society": 30,
            "Metro Melting Pot": 31,
            "Diverse & Determined": 32,
            "New Country": 33,
            "Familles Typiques": 34,
            "Vie Dynamique": 35,
            "Middle-Class Mosaic": 36,
            "Keep on Trucking": 37,
            "Stressed in Suburbia": 38,
            "Évolution Urbaine": 39,
            "Les Énerjeunes": 40,
            "Down to Earth": 41,
            "Banlieues Tranquilles": 42,
            "Happy Medium": 43,
            "Un Grand Cru": 44,
            "Slow-Lane Suburbs": 45,
            "Patrimoine Rustique": 46,
            "Social Networkers": 47,
            "Agri-Biz": 48,
            "Backcountry Boomers": 49,
            "Country & Western": 50,
            "On Their Own Again": 51,
            "Friends & Roomies": 52,
            "Silver Flats": 53,
            "Vie au Village": 54,
            "Enclaves Multiethniques": 55,
            "Jeunes Biculturels": 56,
            "Juggling Acts": 57,
            "Old Town Roads": 58,
            "La Vie Simple": 59,
            "Value Villagers": 60,
            "Came From Away": 61,
            "Suburban Recliners": 62,
            "Amants de la Nature": 63,
            "Midtown Movers": 64,
            "Âgés & Traditionnels": 65,
            "Indigenous Families": 66,
            "Just Getting By": 67,
        }

        self.prizm_file_path = tk.StringVar()
        self.subdivision_file_path = tk.StringVar()
        self.template_file_path = tk.StringVar()
        self.output_folder_path = tk.StringVar()

        # Set default paths using executable directory
        exe_dir = get_executable_dir()
        self.prizm_file_path.set(os.path.join(exe_dir, "PRIZM_Merged_Data.xlsx"))
        self.subdivision_file_path.set(os.path.join(exe_dir, "Subdivisions_Merged_Data.xlsx"))
        self.template_file_path.set(os.path.join(exe_dir, "PRIZM Report 2.xlsx"))
        self.output_folder_path.set(os.path.join(exe_dir, "Output_PDFs"))

        self.is_running = False
        self.setup_ui()

    def setup_ui(self):
        # Main frame with better padding
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky="nsew")

        # Title with improved font
        title_label = ttk.Label(main_frame, text="PRIZM Report Generator",
                               font=("Segoe UI", 18, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 25))

        # PRIZM Data file selection
        ttk.Label(main_frame, text="PRIZM Master Data File:",
                 font=("Segoe UI", 10, "bold")).grid(row=1, column=0, sticky=tk.W, pady=8)
        prizm_entry = ttk.Entry(main_frame, textvariable=self.prizm_file_path,
                               width=70, font=("Segoe UI", 9))
        prizm_entry.grid(row=1, column=1, padx=8, pady=8, sticky="ew")
        ttk.Button(main_frame, text="Browse",
                  command=lambda: self.browse_file(self.prizm_file_path, "PRIZM Master Data"),
                  style="Browse.TButton").grid(row=1, column=2, pady=8, padx=(5, 0))

        # Subdivision Data file selection
        ttk.Label(main_frame, text="Subdivision Master Data File:",
                 font=("Segoe UI", 10, "bold")).grid(row=2, column=0, sticky=tk.W, pady=8)
        subdiv_entry = ttk.Entry(main_frame, textvariable=self.subdivision_file_path,
                                width=70, font=("Segoe UI", 9))
        subdiv_entry.grid(row=2, column=1, padx=8, pady=8, sticky="ew")
        ttk.Button(main_frame, text="Browse",
                  command=lambda: self.browse_file(self.subdivision_file_path, "Subdivision Master Data"),
                  style="Browse.TButton").grid(row=2, column=2, pady=8, padx=(5, 0))

        # Template file selection
        ttk.Label(main_frame, text="Report Template File:",
                 font=("Segoe UI", 10, "bold")).grid(row=3, column=0, sticky=tk.W, pady=8)
        template_entry = ttk.Entry(main_frame, textvariable=self.template_file_path,
                                  width=70, font=("Segoe UI", 9))
        template_entry.grid(row=3, column=1, padx=8, pady=8, sticky="ew")
        ttk.Button(main_frame, text="Browse",
                  command=lambda: self.browse_file(self.template_file_path, "Report Template"),
                  style="Browse.TButton").grid(row=3, column=2, pady=8, padx=(5, 0))

        # Output folder selection
        ttk.Label(main_frame, text="Output Folder:",
                 font=("Segoe UI", 10, "bold")).grid(row=4, column=0, sticky=tk.W, pady=8)
        output_entry = ttk.Entry(main_frame, textvariable=self.output_folder_path,
                                width=70, font=("Segoe UI", 9))
        output_entry.grid(row=4, column=1, padx=8, pady=8, sticky="ew")
        ttk.Button(main_frame, text="Browse",
                  command=lambda: self.browse_folder(self.output_folder_path, "Output Folder"),
                  style="Browse.TButton").grid(row=4, column=2, pady=8, padx=(5, 0))

        # Configure button styles
        style = ttk.Style()
        style.configure("Browse.TButton", font=("Segoe UI", 9))
        style.configure("Generate.TButton", font=("Segoe UI", 11, "bold"), padding=(10, 8))
        style.configure("Stop.TButton", font=("Segoe UI", 11, "bold"), padding=(10, 8))

        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=3, pady=25)

        # Generate button
        self.generate_button = ttk.Button(button_frame, text="Generate Reports",
                                         command=self.start_report_generation, style="Generate.TButton")
        self.generate_button.grid(row=0, column=0, padx=(0, 10))

        # Stop button (initially disabled)
        self.stop_button = ttk.Button(button_frame, text="Stop Generation",
                                     command=self.stop_generation, style="Stop.TButton",
                                     state="disabled")
        self.stop_button.grid(row=0, column=1, padx=(10, 0))

        # Progress bar
        self.progress_var = tk.StringVar()
        self.progress_var.set("Ready to generate reports")
        ttk.Label(main_frame, textvariable=self.progress_var,
                 font=("Segoe UI", 9)).grid(row=6, column=0, columnspan=3, pady=(10, 5))

        self.progress_bar = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress_bar.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(0, 15))

        # Processing log area
        ttk.Label(main_frame, text="Processing Log:",
                 font=("Segoe UI", 10, "bold")).grid(row=8, column=0, sticky=tk.W, pady=(15, 8))

        # Text widget with scrollbar
        text_frame = ttk.Frame(main_frame)
        text_frame.grid(row=9, column=0, columnspan=3, sticky="nsew", pady=8)

        self.log_text = tk.Text(text_frame, height=16, width=100,
                               font=("Consolas", 9),
                               bg="#f8f8f8", fg="#333333",
                               selectbackground="#0078d4",
                               wrap="word", padx=8, pady=8)
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)

        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        # Configure grid weights for responsive design
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(9, weight=1)
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # Initial log message
        self.log_message("PRIZM Report Generator Ready")
        self.log_message("Please verify file paths and click 'Generate Reports' to start")

    def browse_file(self, file_var, file_type):
        if "Template" in file_type:
            file_path = filedialog.askopenfilename(
                title=f"Select {file_type}",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
        else:
            file_path = filedialog.askopenfilename(
                title=f"Select {file_type}",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )

        if file_path:
            file_var.set(file_path)

    def browse_folder(self, folder_var, folder_type):
        folder_path = filedialog.askdirectory(title=f"Select {folder_type}")
        if folder_path:
            folder_var.set(folder_path)

    def log_message(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def update_progress(self, message):
        self.progress_var.set(message)
        self.root.update_idletasks()

    def start_report_generation(self):
        # Validate inputs
        prizm_file = self.prizm_file_path.get()
        subdivision_file = self.subdivision_file_path.get()
        template_file = self.template_file_path.get()
        output_folder = self.output_folder_path.get()

        if not all([prizm_file, subdivision_file, template_file, output_folder]):
            messagebox.showerror("Error", "Please select all required files and output folder.")
            return

        if not os.path.exists(prizm_file):
            messagebox.showerror("Error", f"PRIZM data file not found: {prizm_file}")
            return

        if not os.path.exists(subdivision_file):
            messagebox.showerror("Error", f"Subdivision data file not found: {subdivision_file}")
            return

        if not os.path.exists(template_file):
            messagebox.showerror("Error", f"Template file not found: {template_file}")
            return

        # Clear log
        self.log_text.delete(1.0, tk.END)

        # Update UI state
        self.is_running = True
        self.generate_button.config(state="disabled")
        self.stop_button.config(state="normal")
        self.progress_bar.start(10)

        # Start generation in separate thread
        self.generation_thread = threading.Thread(target=self.run_report_generation)
        self.generation_thread.daemon = True
        self.generation_thread.start()

    def stop_generation(self):
        self.is_running = False
        self.update_progress("Stopping generation...")
        self.log_message("\n🛑 Stop requested by user")

    def run_report_generation(self):
        try:
            # Run the report generation
            self.automate_reports()
        except Exception as e:
            self.log_message(f"\n❌ Fatal error: {str(e)}")
        finally:
            # Reset UI state
            self.root.after(0, self.reset_ui_state)

    def reset_ui_state(self):
        self.is_running = False
        self.generate_button.config(state="normal")
        self.stop_button.config(state="disabled")
        self.progress_bar.stop()
        self.update_progress("Generation completed")

    def automate_reports(self):
        """
        Main function to automate the generation of PDF reports based on unique values in column B.
        """
        # Get file paths from GUI
        master_file_path = self.prizm_file_path.get()
        subdivisions_file_path = self.subdivision_file_path.get()
        template_path = self.template_file_path.get()
        output_folder = self.output_folder_path.get()

        # Create output folder if it doesn't exist
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
            self.log_message(f"Created output folder: {output_folder}")

        # --- 2. LOAD DATA ---
        self.update_progress("Loading master data...")
        self.log_message("Loading master data...")
        try:
            master_data_df = pd.read_excel(master_file_path)
            self.log_message(f"Loaded {len(master_data_df)} rows from master data file")
        except Exception as e:
            self.log_message(f"Error loading master data: {e}")
            return

        if not self.is_running:
            return

        self.update_progress("Loading subdivision data...")
        self.log_message("Loading subdivision data...")
        try:
            subdivisions_data_df = pd.read_excel(subdivisions_file_path)
            self.log_message(f"Loaded {len(subdivisions_data_df)} rows from subdivision data file")

            # Pre-process subdivision data into a dictionary for faster filtering
            self.log_message("Pre-processing subdivision data for faster filtering...")
            subdivisions_by_target = {}
            for target_group in subdivisions_data_df['TargetGroupName'].unique():
                if pd.notna(target_group):
                    subdivisions_by_target[target_group] = subdivisions_data_df[
                        subdivisions_data_df['TargetGroupName'] == target_group
                    ].copy()
            self.log_message(f"Pre-processed {len(subdivisions_by_target)} target groups from subdivision data")

        except Exception as e:
            self.log_message(f"Error loading subdivision data: {e}")
            return

        if not self.is_running:
            return

        # --- 3. GET UNIQUE VALUES FROM COLUMN B ---
        if 'B' not in master_data_df.columns and len(master_data_df.columns) > 1:
            # If column B doesn't exist by name, use the second column (index 1)
            column_b_name = master_data_df.columns[1]
            self.log_message(f"Using column '{column_b_name}' as column B")
        else:
            column_b_name = 'B'

        unique_values = master_data_df[column_b_name].unique()
        unique_values = [val for val in unique_values if pd.notna(val)]  # Remove NaN values
        self.log_message(f"Found {len(unique_values)} unique values in column B: {unique_values[:5]}...")

        # --- 4. INITIALIZE EXCEL APPLICATION ---
        excel_app = None
        # Initialize counters before try block to avoid "referenced before assignment" errors
        total_reports = 0
        successful_reports = 0
        failed_reports = 0

        try:
            self.update_progress("Initializing Excel application...")
            self.log_message("Initializing Excel application...")

            # Kill any existing Excel processes first
            try:
                import subprocess
                subprocess.run(["taskkill", "/f", "/im", "excel.exe"], capture_output=True, check=False)
                time.sleep(2)
            except:
                pass

            pythoncom.CoInitialize()  # Initialize COM

            excel_app = win32com.client.Dispatch("Excel.Application")
            excel_app.Visible = False  # Set to True for debugging
            excel_app.DisplayAlerts = False
            excel_app.ScreenUpdating = False  # Improve performance

            self.log_message("Excel application initialized successfully")

            # --- 5. PROCESS EACH UNIQUE VALUE SEQUENTIALLY ---
            total_reports = len(unique_values)

            for i, unique_value in enumerate(unique_values, 1):
                if not self.is_running:
                    self.log_message("\n🛑 Generation stopped by user")
                    break

                self.update_progress(f"Processing {i}/{total_reports}: {unique_value}")
                self.log_message(f"\nProcessing {i}/{total_reports}: {unique_value}")

                # Filter data for this unique value
                filtered_data = master_data_df[master_data_df[column_b_name] == unique_value]
                self.log_message(f"  Found {len(filtered_data)} rows for {unique_value}")

                # Process this unique value with retry logic
                success = self.process_unique_value_with_retry(
                    excel_app, template_path, filtered_data, unique_value, output_folder, subdivisions_by_target
                )

                if success:
                    successful_reports += 1
                    self.log_message(f"  ✓ Completed: {unique_value}")
                else:
                    failed_reports += 1
                    self.log_message(f"  ✗ Failed: {unique_value}")

        except Exception as e:
            self.log_message(f"Fatal error: {e}")
            failed_reports = total_reports - successful_reports

        finally:
            # --- 6. CLEANUP ---
            self.log_message("\nCleaning up...")
            if excel_app:
                try:
                    excel_app.Quit()
                except:
                    pass
            try:
                pythoncom.CoUninitialize()
            except:
                pass

        # Final results
        self.log_message(f"\n📊 Final Results:")
        self.log_message(f"✓ Successful reports: {successful_reports}")
        self.log_message(f"✗ Failed reports: {failed_reports}")

        if self.is_running:
            self.log_message(f"📁 Check the Output_PDFs folder for PDF files.")
            self.update_progress(f"All reports completed! Success: {successful_reports}, Failed: {failed_reports}")
        else:
            self.update_progress("Generation stopped by user")

    def process_unique_value_with_retry(self, excel_app, template_path, filtered_data, unique_value, output_folder, subdivisions_by_target, max_retries=3):
        """
        Process a single unique value with retry logic for COM errors.
        """
        for attempt in range(max_retries):
            if not self.is_running:
                return False

            try:
                return self.process_unique_value(excel_app, template_path, filtered_data, unique_value, output_folder, subdivisions_by_target)
            except Exception as e:
                self.log_message(f"  Attempt {attempt + 1} failed: {e}")
                if attempt < max_retries - 1:
                    self.log_message(f"  Retrying in 3 seconds...")
                    time.sleep(3)
                else:
                    self.log_message(f"  All retries failed for {unique_value}")
                    return False
        return False

    def process_unique_value(self, excel_app, template_path, filtered_data, unique_value, output_folder, subdivisions_by_target):
        """
        Process a single unique value - update Excel sheets and generate PDF.
        """
        workbook = None
        try:
            if not self.is_running:
                return False

            # Open template file
            self.log_message(f"  Opening template file...")
            workbook = excel_app.Workbooks.Open(template_path)
            time.sleep(1)  # Give Excel time to fully open the file

            # --- UPDATE DATA SHEET ---
            self.log_message(f"  Updating Data sheet...")
            data_sheet = workbook.Worksheets("Data")

            # Clear ALL existing data (preserve headers in row 1)
            self.log_message(f"    Clearing all existing data from Data sheet...")
            try:
                # Clear everything from row 2 onwards - this ensures complete data removal
                data_sheet.Range("A2:XFD1048576").ClearContents()
                self.log_message(f"    ✓ Cleared all data from row 2 onwards")
            except Exception as e:
                self.log_message(f"    Fallback: clearing large range due to error: {e}")
                # Fallback: clear a very large range
                data_sheet.Range("A2:ZZZ10000").ClearContents()

            # Write filtered data to Data sheet
            if not filtered_data.empty:
                self.log_message(f"  Writing {len(filtered_data)} rows to Data sheet...")
                self.write_data_to_sheet(data_sheet, filtered_data)

            if not self.is_running:
                return False

            # --- UPDATE SUBDIVISIONS SHEET ---
            self.log_message(f"  Updating Subdivisions sheet...")
            try:
                # Wait a bit before accessing the subdivisions sheet
                time.sleep(0.5)
                subdivisions_sheet = workbook.Worksheets("Subdivisions")

                # Clear existing subdivisions data with retry logic
                clear_attempts = 0
                max_clear_attempts = 3
                while clear_attempts < max_clear_attempts:
                    try:
                        used_range_sub = subdivisions_sheet.UsedRange
                        if used_range_sub.Rows.Count > 1:
                            clear_range_sub = subdivisions_sheet.Range(f"A2:{self.get_column_letter(used_range_sub.Columns.Count)}{used_range_sub.Rows.Count}")
                            clear_range_sub.ClearContents()
                        break
                    except Exception as clear_error:
                        clear_attempts += 1
                        self.log_message(f"    Clear attempt {clear_attempts} failed: {clear_error}")
                        if clear_attempts < max_clear_attempts:
                            time.sleep(1)
                        else:
                            # Fallback clearing method
                            subdivisions_sheet.Range("A2:ZZ10000").ClearContents()

                # Get subdivision data with fast dictionary lookup
                filtered_subdivision_data = subdivisions_by_target.get(unique_value, pd.DataFrame())
                self.log_message(f"    Found {len(filtered_subdivision_data)} subdivision rows for {unique_value}")

                if not filtered_subdivision_data.empty:
                    self.log_message(f"    Writing {len(filtered_subdivision_data)} subdivision rows to Subdivisions sheet...")
                    # Add a small delay before writing data
                    time.sleep(0.5)
                    self.write_data_to_sheet_with_retry(subdivisions_sheet, filtered_subdivision_data)
                else:
                    self.log_message(f"    No subdivision data found for {unique_value}")

            except Exception as e:
                self.log_message(f"  Warning: Could not update Subdivisions sheet: {e}")
                # Try to continue with the rest of the process

            if not self.is_running:
                return False

            # --- UPDATE VARIABLES AND AGGREGATES SHEET ---
            self.log_message(f"  Updating Variables and Aggregates sheet...")
            try:
                variables_sheet = workbook.Worksheets("Variables & Aggregates")
                # Update cell A5 with the unique name
                variables_sheet.Cells(5, 1).Value = str(unique_value)
                self.log_message(f"    ✓ Set A5 to: {unique_value}")
                time.sleep(3)
                # Update cell A6 with the corresponding identifier code
                identifier_code = self.prizm_set_to_code.get(str(unique_value))
                if identifier_code is not None:
                    variables_sheet.Cells(6, 1).Value = identifier_code
                    self.log_message(f"    ✓ Set A6 to identifier code: {identifier_code}")
                else:
                    self.log_message(f"    ⚠ Warning: No identifier code found for '{unique_value}' in mapping")
                    # Optionally set a default value or leave empty
                    # variables_sheet.Cells(6, 1).Value = ""

            except Exception as e:
                self.log_message(f"  Warning: Could not update Variables and Aggregates sheet: {e}")

            # --- RECALCULATE WORKBOOK ---
            self.log_message(f"  Recalculating formulas...")
            workbook.Application.CalculateUntilAsyncQueriesDone()
            time.sleep(3)  # Give Excel more time to process calculations

            if not self.is_running:
                return False

            # --- EXPORT REPORT SHEET AS PDF ---
            self.log_message(f"  Exporting to PDF...")
            report_sheet = workbook.Worksheets("Report")

            # Create safe filename
            safe_filename = "".join(c for c in str(unique_value) if c.isalnum() or c in (' ', '-', '_')).rstrip()
            pdf_filename = f"PRIZM_Report_{safe_filename}.pdf"
            pdf_path = os.path.join(output_folder, pdf_filename)

            # Export pages 1-17 to PDF
            try:
                report_sheet.ExportAsFixedFormat(
                    Type=0,  # xlTypePDF
                    Filename=pdf_path,
                    Quality=0,  # xlQualityStandard
                    IncludeDocProperties=True,  # Fixed parameter name
                    IgnorePrintAreas=False,
                    From=1,  # Start page
                    To=17,   # End page
                    OpenAfterPublish=False
                )
                self.log_message(f"  ✓ PDF saved: {pdf_filename}")
                return True

            except Exception as e:
                self.log_message(f"  ✗ Error exporting PDF: {e}")
                return False

        finally:
            # Close workbook without saving
            if workbook:
                try:
                    workbook.Close(SaveChanges=False)
                except:
                    pass

    def write_data_to_sheet_with_retry(self, sheet, dataframe, max_retries=5, batch_size=250):
        """
        Write DataFrame to Excel sheet with improved retry logic for COM errors.
        Uses aggressive chunking for large datasets to prevent Excel overwhelm.
        """
        num_rows = len(dataframe)

        # For large datasets (>500 rows), immediately use smaller batches to prevent Excel overwhelm
        if num_rows > 500:
            self.log_message(f"    Large dataset detected ({num_rows} rows). Using smaller batch approach...")
            return self.write_large_dataset_in_chunks(sheet, dataframe, chunk_size=100)

        # For smaller datasets, try the original approach with retries
        for attempt in range(max_retries):
            try:
                # Force COM to process any pending messages before attempting write
                pythoncom.PumpWaitingMessages()

                self.write_data_to_sheet(sheet, dataframe, batch_size)
                return True
            except Exception as e:
                error_code = getattr(e, 'hresult', None) if hasattr(e, 'hresult') else None

                # Check for specific COM errors that indicate Excel is busy
                is_com_busy_error = (
                    error_code == -2147418111 or  # Call was rejected by callee
                    error_code == -2147352567 or  # Exception occurred
                    "rejected by callee" in str(e).lower() or
                    "busy" in str(e).lower()
                )

                self.log_message(f"    Write attempt {attempt + 1} failed: {e}")

                if attempt < max_retries - 1:
                    if is_com_busy_error:
                        # For COM busy errors, wait longer and try to release COM resources
                        wait_time = (attempt + 1) * 3  # Exponential backoff: 3, 6, 9, 12 seconds
                        self.log_message(f"    COM busy error detected. Waiting {wait_time} seconds and releasing COM resources...")

                        # Try to release any pending COM operations
                        try:
                            pythoncom.PumpWaitingMessages()
                            sheet.Application.Calculate()  # Force Excel to finish calculations
                        except:
                            pass

                        time.sleep(wait_time)
                    else:
                        # For other errors, shorter wait
                        self.log_message(f"    Retrying in 2 seconds...")
                        time.sleep(2)
                else:
                    # If all retries failed, fall back to chunked approach
                    self.log_message(f"    All write attempts failed after {max_retries} retries. Falling back to chunked approach...")
                    return self.write_large_dataset_in_chunks(sheet, dataframe, chunk_size=50)
        return False

    def write_large_dataset_in_chunks(self, sheet, dataframe, chunk_size=100):
        """
        Write large datasets in very small chunks to prevent Excel from being overwhelmed.
        This is the most reliable method for large datasets.
        """
        try:
            self.log_message(f"    Writing {len(dataframe)} rows in chunks of {chunk_size}...")

            # Convert to values once
            data_values = dataframe.values
            num_rows, num_cols = data_values.shape

            # Clear the target range first
            if num_rows > 0:
                end_col = self.get_column_letter(num_cols)
                try:
                    clear_range = sheet.Range(f"A2:{end_col}{num_rows + 1}")
                    clear_range.ClearContents()
                    self.log_message(f"    Cleared target range A2:{end_col}{num_rows + 1}")
                except:
                    # Fallback clearing
                    sheet.Range("A2:ZZ1000").ClearContents()

            # Write data in small chunks
            successful_chunks = 0
            total_chunks = (num_rows + chunk_size - 1) // chunk_size

            for chunk_idx in range(0, num_rows, chunk_size):
                if not self.is_running:
                    self.log_message(f"    Operation stopped by user during chunk processing")
                    return False

                chunk_end = min(chunk_idx + chunk_size, num_rows)
                chunk_data = data_values[chunk_idx:chunk_end].tolist()

                chunk_num = (chunk_idx // chunk_size) + 1
                self.log_message(f"    Writing chunk {chunk_num}/{total_chunks} (rows {chunk_idx + 1}-{chunk_end})...")

                # Calculate Excel range for this chunk
                start_row = chunk_idx + 2  # +2 for 1-based indexing and header row
                end_row = start_row + len(chunk_data) - 1
                end_col = self.get_column_letter(num_cols)

                # Try to write this chunk with retries
                chunk_success = False
                for retry in range(3):
                    try:
                        # Small delay between chunks to let Excel breathe
                        if chunk_idx > 0:
                            time.sleep(0.2)

                        pythoncom.PumpWaitingMessages()

                        chunk_range = sheet.Range(f"A{start_row}:{end_col}{end_row}")
                        chunk_range.Value = chunk_data

                        chunk_success = True
                        break

                    except Exception as chunk_error:
                        if retry < 2:
                            self.log_message(f"      Chunk {chunk_num} attempt {retry + 1} failed, retrying...")
                            time.sleep(1)
                        else:
                            self.log_message(f"      Chunk {chunk_num} failed after all retries: {chunk_error}")
                            # Try cell-by-cell as last resort for this chunk
                            try:
                                for i, row_data in enumerate(chunk_data):
                                    row_idx = start_row + i
                                    for col_idx, value in enumerate(row_data, start=1):
                                        if pd.notna(value):
                                            sheet.Cells(row_idx, col_idx).Value = value
                                chunk_success = True
                                self.log_message(f"      Chunk {chunk_num} written cell-by-cell as fallback")
                            except:
                                self.log_message(f"      Chunk {chunk_num} completely failed")

                if chunk_success:
                    successful_chunks += 1

            success_rate = (successful_chunks / total_chunks) * 100
            self.log_message(f"    ✓ Chunked write completed: {successful_chunks}/{total_chunks} chunks successful ({success_rate:.1f}%)")

            return successful_chunks > 0  # Return True if at least some data was written

        except Exception as e:
            self.log_message(f"    ✗ Chunked write method failed: {e}")
            return False

    def write_data_to_sheet(self, sheet, dataframe, batch_size=1000):
        """
        Write DataFrame to Excel sheet using bulk operations for better performance.
        """
        self.log_message(f"    Converting data to array format...")
        data_values = dataframe.values
        num_rows, num_cols = data_values.shape

        # Clear the existing data range first
        if num_rows > 0:
            end_col = self.get_column_letter(num_cols)
            clear_range = sheet.Range(f"A2:{end_col}{num_rows + 1}")
            clear_range.ClearContents()
            self.log_message(f"    Cleared existing data range A2:{end_col}{num_rows + 1}")

        # Write data in bulk using Excel ranges
        if num_rows > 0 and num_cols > 0:
            self.log_message(f"    Writing {num_rows} rows x {num_cols} columns in bulk...")

            # Convert to format Excel expects (list of lists)
            excel_data = data_values.tolist()

            # Define the target range starting from A2 (assuming row 1 has headers)
            end_col = self.get_column_letter(num_cols)
            target_range = sheet.Range(f"A2:{end_col}{num_rows + 1}")

            try:
                # Write all data at once - much faster than cell by cell
                target_range.Value = excel_data
                self.log_message(f"    ✓ Successfully wrote all data to range A2:{end_col}{num_rows + 1}")
            except Exception as e:
                self.log_message(f"    Bulk write failed, falling back to batch method: {e}")
                # Fallback to smaller batches if bulk write fails
                self.write_data_in_batches(sheet, excel_data, batch_size)

    def write_data_in_batches(self, sheet, excel_data, batch_size):
        """
        Fallback method to write data in smaller batches.
        """
        num_rows = len(excel_data)
        num_cols = len(excel_data[0]) if num_rows > 0 else 0

        for batch_start in range(0, num_rows, batch_size):
            if not self.is_running:
                break

            batch_end = min(batch_start + batch_size, num_rows)
            batch_data = excel_data[batch_start:batch_end]

            self.log_message(f"    Writing batch {batch_start + 1}-{batch_end} of {num_rows}...")

            # Calculate range for this batch
            start_row = batch_start + 2  # +2 for 1-based indexing and header row
            end_row = batch_end + 1
            end_col = self.get_column_letter(num_cols)

            batch_range = sheet.Range(f"A{start_row}:{end_col}{end_row}")

            try:
                batch_range.Value = batch_data
                self.log_message(f"    ✓ Batch {batch_start + 1}-{batch_end} written successfully")
            except Exception as e:
                self.log_message(f"    ✗ Batch {batch_start + 1}-{batch_end} failed: {e}")
                # If batch fails, fall back to individual cells for this batch only
                for i, row_data in enumerate(batch_data):
                    row_idx = start_row + i
                    for col_idx, value in enumerate(row_data, start=1):
                        if pd.notna(value):
                            try:
                                sheet.Cells(row_idx, col_idx).Value = value
                            except:
                                continue

            # Brief pause between batches
            if batch_end < num_rows:
                time.sleep(0.1)

    def get_column_letter(self, col_num):
        """
        Convert column number to Excel column letter (e.g., 1 -> A, 27 -> AA).
        """
        string = ""
        while col_num > 0:
            col_num, remainder = divmod(col_num - 1, 26)
            string = chr(65 + remainder) + string
        return string
def main():
    root = tk.Tk()
    app = PRIZMReportGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
