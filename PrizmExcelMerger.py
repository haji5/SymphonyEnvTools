import pandas as pd
import os
import glob
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import sys

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

def find_excel_files_in_subfolders(main_folder):
    """
    Find all Excel files within subfolders of the main folder.
    Each subfolder should contain one Excel file.
    """
    excel_files = []

    # Get all subdirectories
    subfolders = [f for f in os.listdir(main_folder) if os.path.isdir(os.path.join(main_folder, f))]

    for subfolder in subfolders:
        subfolder_path = os.path.join(main_folder, subfolder)
        # Find Excel files in this subfolder
        excel_files_in_subfolder = glob.glob(os.path.join(subfolder_path, "*.xlsx"))
        excel_files.extend(excel_files_in_subfolder)

    return excel_files

def merge_excel_files_from_folder(source_folder, output_filename, data_type):
    """
    Merge all Excel files from subfolders within the source folder into a single Excel file.
    Appends all data and removes duplicate rows at the end.
    """
    print(f"\n=== Processing {data_type} folder: {os.path.basename(source_folder)} ===")

    # Get all Excel files from subfolders
    excel_files = find_excel_files_in_subfolders(source_folder)

    print(f"Found {len(excel_files)} Excel files to merge...")

    if not excel_files:
        print("No Excel files found in subfolders!")
        return None

    # List to store all dataframes
    all_dataframes = []

    # Process each Excel file
    for i, file_path in enumerate(excel_files, 1):
        try:
            # Get the filename and subfolder for progress tracking
            filename = os.path.basename(file_path)
            subfolder = os.path.basename(os.path.dirname(file_path))
            print(f"Processing file {i}/{len(excel_files)}: {subfolder}/{filename}")

            # Read the Excel file
            df = pd.read_excel(file_path)

            # Append to the list of dataframes
            all_dataframes.append(df)

        except Exception as e:
            filename = os.path.basename(file_path) if 'file_path' in locals() else "Unknown file"
            print(f"Error processing {filename}: {str(e)}")
            continue

    if not all_dataframes:
        print("No valid Excel files could be processed!")
        return None

    # Combine all dataframes
    print("\nCombining all dataframes...")
    merged_df = pd.concat(all_dataframes, ignore_index=True)

    print(f"Combined data shape before duplicate removal: {merged_df.shape}")

    print("Removing duplicate rows...")
    initial_count = len(merged_df)

    # Keep the first occurrence of each duplicate
    merged_df_deduplicated = merged_df.drop_duplicates(keep='first')

    final_count = len(merged_df_deduplicated)
    duplicates_removed = initial_count - final_count

    print(f"Removed {duplicates_removed} duplicate rows")
    print(f"Final data shape: {merged_df_deduplicated.shape}")

    # Save the merged and deduplicated data to Excel
    print(f"\nSaving merged data to {output_filename}...")
    merged_df_deduplicated.to_excel(output_filename, index=False)

    print(f"Successfully merged {len(excel_files)} files into {output_filename}")
    print(f"Total rows in final file: {len(merged_df_deduplicated)}")

    # Display some statistics
    print(f"\n=== {data_type.upper()} MERGE STATISTICS ===")
    print(f"Files processed: {len(excel_files)}")
    print(f"Total rows before deduplication: {initial_count:,}")
    print(f"Duplicate rows removed: {duplicates_removed:,}")
    print(f"Final rows: {final_count:,}")
    print(f"Deduplication rate: {(duplicates_removed/initial_count)*100:.2f}%")

    # Show unique target groups if the column exists
    if 'TargetGroupName' in merged_df_deduplicated.columns:
        unique_groups = merged_df_deduplicated['TargetGroupName'].nunique()
        print(f"Unique target groups: {unique_groups}")

    return merged_df_deduplicated

class DataMergerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PRIZM & Subdivision Data Merger")
        self.root.geometry("650x450")

        # Configure high DPI awareness and scaling
        try:
            self.root.tk.call('tk', 'scaling', 1.5)  # Increase scaling for better quality
        except:
            pass

        # Configure style for better appearance
        style = ttk.Style()
        style.theme_use('clam')  # Use a modern theme

        self.prizm_folder_path = tk.StringVar()
        self.subdivision_folder_path = tk.StringVar()

        self.setup_ui()

    def setup_ui(self):
        # Main frame with better padding
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky="nsew")

        # Title with improved font
        title_label = ttk.Label(main_frame, text="PRIZM & Subdivision Data Merger",
                               font=("Segoe UI", 18, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 25))

        # PRIZM Data folder selection with improved fonts and spacing
        ttk.Label(main_frame, text="PRIZM Data Folder:",
                 font=("Segoe UI", 10, "bold")).grid(row=1, column=0, sticky=tk.W, pady=8)
        prizm_entry = ttk.Entry(main_frame, textvariable=self.prizm_folder_path,
                               width=50, font=("Segoe UI", 9))
        prizm_entry.grid(row=1, column=1, padx=8, pady=8, sticky="ew")
        ttk.Button(main_frame, text="Browse",
                  command=lambda: self.browse_folder(self.prizm_folder_path, "PRIZM"),
                  style="Browse.TButton").grid(row=1, column=2, pady=8, padx=(5, 0))

        # Subdivision Data folder selection with improved fonts and spacing
        ttk.Label(main_frame, text="Subdivision Data Folder:",
                 font=("Segoe UI", 10, "bold")).grid(row=2, column=0, sticky=tk.W, pady=8)
        subdiv_entry = ttk.Entry(main_frame, textvariable=self.subdivision_folder_path,
                                width=50, font=("Segoe UI", 9))
        subdiv_entry.grid(row=2, column=1, padx=8, pady=8, sticky="ew")
        ttk.Button(main_frame, text="Browse",
                  command=lambda: self.browse_folder(self.subdivision_folder_path, "Subdivision"),
                  style="Browse.TButton").grid(row=2, column=2, pady=8, padx=(5, 0))

        # Configure button styles
        style = ttk.Style()
        style.configure("Browse.TButton", font=("Segoe UI", 9))
        style.configure("Merge.TButton", font=("Segoe UI", 11, "bold"), padding=(10, 8))

        # Merge button with better styling
        merge_button = ttk.Button(main_frame, text="Merge Excel Files",
                                 command=self.merge_files, style="Merge.TButton")
        merge_button.grid(row=3, column=0, columnspan=3, pady=25)

        # Progress and log area with improved label
        ttk.Label(main_frame, text="Processing Log:",
                 font=("Segoe UI", 10, "bold")).grid(row=4, column=0, sticky=tk.W, pady=(15, 8))

        # Text widget with scrollbar and better fonts
        text_frame = ttk.Frame(main_frame)
        text_frame.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=8)

        self.log_text = tk.Text(text_frame, height=15, width=70,
                               font=("Consolas", 9),
                               bg="#f8f8f8", fg="#333333",
                               selectbackground="#0078d4",
                               wrap=tk.WORD, padx=8, pady=8)
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)

        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        # Configure grid weights for responsive design
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(5, weight=1)
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

    def browse_folder(self, folder_var, data_type):
        folder_path = filedialog.askdirectory(title=f"Select {data_type} folder containing subfolders with Excel files")
        if folder_path:
            folder_var.set(folder_path)

    def log_message(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update()

    def merge_files(self):
        prizm_folder = self.prizm_folder_path.get()
        subdivision_folder = self.subdivision_folder_path.get()
        
        if not prizm_folder or not subdivision_folder:
            messagebox.showerror("Error", "Please select both PRIZM and Subdivision folders before merging.")
            return
        
        if not os.path.exists(prizm_folder) or not os.path.exists(subdivision_folder):
            messagebox.showerror("Error", "One or both selected folders do not exist.")
            return
        
        self.log_text.delete(1.0, tk.END)
        
        try:
            # Get the executable directory for output files
            exe_dir = get_executable_dir()

            # Process PRIZM folder
            self.log_message("Starting merge process...")
            self.log_message(f"Output files will be saved to: {exe_dir}")
            self.log_message(f"Processing PRIZM Data Folder: {os.path.basename(prizm_folder)}")
            
            output_file_prizm = os.path.join(exe_dir, "PRIZM_Merged_Data.xlsx")

            # Redirect print output to GUI
            import sys
            from io import StringIO
            
            # Capture print output for PRIZM processing
            old_stdout = sys.stdout
            sys.stdout = captured_output = StringIO()
            
            try:
                result_prizm = merge_excel_files_from_folder(prizm_folder, output_file_prizm, "PRIZM")
            finally:
                sys.stdout = old_stdout
                
            # Display captured output
            for line in captured_output.getvalue().split('\n'):
                if line.strip():
                    self.log_message(line)
            
            if result_prizm is not None:
                self.log_message(f"✅ PRIZM data processed successfully: {os.path.basename(output_file_prizm)}")
            else:
                self.log_message("❌ Failed to process PRIZM data")
            
            # Process Subdivision folder
            self.log_message(f"\nProcessing Subdivision Data Folder: {os.path.basename(subdivision_folder)}")
            
            output_file_subdivision = os.path.join(exe_dir, "Subdivisions_Merged_Data.xlsx")

            # Capture print output for Subdivision processing
            sys.stdout = captured_output = StringIO()
            
            try:
                result_subdivision = merge_excel_files_from_folder(subdivision_folder, output_file_subdivision, "Subdivision")
            finally:
                sys.stdout = old_stdout
                
            # Display captured output
            for line in captured_output.getvalue().split('\n'):
                if line.strip():
                    self.log_message(line)
            
            if result_subdivision is not None:
                self.log_message(f"✅ Subdivision data processed successfully: {os.path.basename(output_file_subdivision)}")
            else:
                self.log_message("❌ Failed to process Subdivision data")
            
            # Summary
            if result_prizm is not None and result_subdivision is not None:
                self.log_message(f"\n🎉 Both PRIZM and Subdivision data processed successfully!")
                self.log_message(f"📁 Output files saved to: {exe_dir}")
                messagebox.showinfo("Success", f"Both PRIZM and Subdivision folders have been processed successfully!\n\nOutput files saved to:\n{exe_dir}")
            elif result_prizm is not None or result_subdivision is not None:
                self.log_message("\n⚠️ One data type processed successfully, one failed.")
                messagebox.showwarning("Partial Success", "One data type was processed successfully, but the other failed. Check the log for details.")
            else:
                self.log_message("\n❌ Both data types failed to process.")
                messagebox.showerror("Error", "Both data types failed to process. Check the log for details.")
                
        except Exception as e:
            error_msg = f"Error during merge process: {str(e)}"
            self.log_message(f"\n❌ {error_msg}")
            messagebox.showerror("Error", error_msg)

def main():
    root = tk.Tk()
    app = DataMergerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
