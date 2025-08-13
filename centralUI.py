import tkinter as tk
from tkinter import ttk, messagebox
import threading
import sys
import os

# Import the modules that will be launched
try:
    from PRIZM_Report_GUI import PRIZMReportGUI
except ImportError:
    PRIZMReportGUI = None

try:
    from PrizmExcelMerger import main as excel_merger_main
except ImportError:
    excel_merger_main = None

# Import the other modules with proper names
transformation_available = True  # We'll check this when trying to import
prizm_grabber_available = True
subdivision_grabber_available = True

# Don't import any of these modules immediately - they all have GUI code that runs on import
# We'll import them only when their respective buttons are clicked

class CentralUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Environics Data Tools - Central Hub")
        self.root.geometry("650x700")  # Increased height to accommodate new buttons
        self.root.resizable(True, True)

        # Configure high DPI awareness
        try:
            self.root.tk.call('tk', 'scaling', 1.2)
        except:
            pass

        # Configure style
        style = ttk.Style()
        style.theme_use('clam')

        # Get the current directory where this script is located
        if getattr(sys, 'frozen', False):
            # If running as PyInstaller bundle
            self.script_dir = os.path.dirname(sys.executable)
        else:
            # If running as regular Python script
            self.script_dir = os.path.dirname(os.path.abspath(__file__))

        self.create_widgets()

    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)

        # Title
        title_label = ttk.Label(main_frame, text="Environics Data Tools",
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, pady=(0, 20))

        subtitle_label = ttk.Label(main_frame, text="Select a tool to launch",
                                  font=('Arial', 10))
        subtitle_label.grid(row=1, column=0, pady=(0, 30))

        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, sticky=(tk.W, tk.E))
        button_frame.columnconfigure(0, weight=1)

        # Button 1: PRIZM Report GUI
        btn1 = ttk.Button(button_frame, text="PRIZM Report Generator",
                         command=self.launch_prizm_report,
                         style='Large.TButton')
        btn1.grid(row=0, column=0, pady=10, padx=20, sticky=(tk.W, tk.E))

        # Description for button 1
        desc1 = ttk.Label(button_frame, text="Generate PRIZM reports from data files",
                         font=('Arial', 9), foreground='gray')
        desc1.grid(row=1, column=0, pady=(0, 15))

        # Button 2: PRIZM Excel Merger
        btn2 = ttk.Button(button_frame, text="PRIZM Excel Merger",
                         command=self.launch_excel_merger,
                         style='Large.TButton')
        btn2.grid(row=2, column=0, pady=10, padx=20, sticky=(tk.W, tk.E))

        # Description for button 2
        desc2 = ttk.Label(button_frame, text="Merge multiple Excel files from subfolders",
                         font=('Arial', 9), foreground='gray')
        desc2.grid(row=3, column=0, pady=(0, 15))

        # Button 3: Transformation Tool
        btn3 = ttk.Button(button_frame, text="Data Transformation Tool",
                         command=self.launch_transformation,
                         style='Large.TButton')
        btn3.grid(row=4, column=0, pady=10, padx=20, sticky=(tk.W, tk.E))

        # Description for button 3
        desc3 = ttk.Label(button_frame, text="Transform and process data using web automation",
                         font=('Arial', 9), foreground='gray')
        desc3.grid(row=5, column=0, pady=(0, 15))

        # Button 4: PRIZM Data Grabber GUI
        btn4 = ttk.Button(button_frame, text="PRIZM Data Grabber",
                         command=self.launch_prizm_grabber,
                         style='Large.TButton')
        btn4.grid(row=6, column=0, pady=10, padx=20, sticky=(tk.W, tk.E))

        # Description for button 4
        desc4 = ttk.Label(button_frame, text="Automated PRIZM dashboard creation (204 dashboards for 51 target sets)",
                         font=('Arial', 9), foreground='gray')
        desc4.grid(row=7, column=0, pady=(0, 15))

        # Button 5: Subdivision Data Grabber GUI
        btn5 = ttk.Button(button_frame, text="Subdivision Data Grabber",
                         command=self.launch_subdivision_grabber,
                         style='Large.TButton')
        btn5.grid(row=8, column=0, pady=10, padx=20, sticky=(tk.W, tk.E))

        # Description for button 5
        desc5 = ttk.Label(button_frame, text="Automated subdivision dashboard creation for British Columbia data",
                         font=('Arial', 9), foreground='gray')
        desc5.grid(row=9, column=0, pady=(0, 20))

        # Configure button style
        style = ttk.Style()
        style.configure('Large.TButton', font=('Arial', 11), padding=(10, 10))

        # Status frame
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(20, 0))
        status_frame.columnconfigure(0, weight=1)

        self.status_label = ttk.Label(status_frame, text="Ready",
                                     font=('Arial', 9), foreground='green')
        self.status_label.grid(row=0, column=0)

        # Exit button
        exit_btn = ttk.Button(status_frame, text="Exit",
                             command=self.root.quit)
        exit_btn.grid(row=1, column=0, pady=(10, 0))

    def launch_prizm_report(self):
        """Launch the PRIZM Report GUI"""
        try:
            self.status_label.config(text="Launching PRIZM Report Generator...", foreground='blue')
            self.root.update()

            if PRIZMReportGUI is not None:
                # Create new window in the main thread (no separate thread needed for Tkinter)
                new_root = tk.Toplevel(self.root)

                # Configure window to stay on top of central UI
                self.configure_child_window(new_root)

                app = PRIZMReportGUI(new_root)
                # Don't call mainloop() on Toplevel windows - they're managed by the main root
                self.status_label.config(text="PRIZM Report Generator launched successfully", foreground='green')
            else:
                raise ImportError("PRIZM Report GUI module not available")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch PRIZM Report Generator:\n{str(e)}")
            self.status_label.config(text="Error launching PRIZM Report Generator", foreground='red')

    def launch_excel_merger(self):
        """Launch the PRIZM Excel Merger"""
        try:
            self.status_label.config(text="Launching PRIZM Excel Merger...", foreground='blue')
            self.root.update()

            if excel_merger_main is not None:
                # Import the PrizmExcelMerger module to get the GUI class
                import PrizmExcelMerger

                # Create the Excel Merger GUI in the main thread, not in a separate thread
                def run_merger():
                    try:
                        # Create a new Tkinter window for the Excel Merger
                        new_root = tk.Toplevel(self.root)

                        # Configure window to stay on top of central UI
                        self.configure_child_window(new_root)

                        app = PrizmExcelMerger.DataMergerGUI(new_root)
                        # Don't call mainloop() - it's managed by the main root
                    except Exception as e:
                        print(f"Error running PRIZM Excel Merger: {e}")

                # Run in the main thread, not a separate thread
                run_merger()
                self.status_label.config(text="PRIZM Excel Merger launched successfully", foreground='green')
            else:
                raise ImportError("PRIZM Excel Merger module not available")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch PRIZM Excel Merger:\n{str(e)}")
            self.status_label.config(text="Error launching PRIZM Excel Merger", foreground='red')

    def launch_transformation(self):
        """Launch the Data Transformation Tool"""
        try:
            self.status_label.config(text="Launching Data Transformation Tool...", foreground='blue')
            self.root.update()

            if transformation_available:
                # Create a simple launcher GUI directly in the central UI instead of using the module's startup window
                def run_transformation():
                    try:
                        # Import the module fresh each time
                        import importlib
                        import sys

                        # Clean up any existing module to avoid conflicts
                        if 'Transformation_3' in sys.modules:
                            del sys.modules['Transformation_3']

                        # Import fresh
                        import Transformation_3

                        # Instead of using the module's startup window, call main_program directly
                        # but first create a simple dialog - but only if the main window still exists
                        if not self.root.winfo_exists():
                            return

                        dialog = tk.Toplevel(self.root)
                        dialog.title("Transformation 3.0")
                        dialog.geometry("350x200")
                        dialog.resizable(False, False)
                        dialog.configure(bg='white')

                        # Center the dialog relative to the main window
                        dialog.transient(self.root)
                        dialog.grab_set()

                        # Create dialog content
                        main_frame = tk.Frame(dialog, bg='white')
                        main_frame.pack(fill='both', expand=True, padx=30, pady=30)

                        title_label = tk.Label(main_frame, text="Transformation 3.0",
                                             font=('Segoe UI', 18, 'normal'), bg='white', fg='#2c3e50')
                        title_label.pack(pady=(0, 10))

                        subtitle_label = tk.Label(main_frame, text="Data Processing Tool",
                                                font=('Segoe UI', 10, 'normal'), bg='white', fg='#7f8c8d')
                        subtitle_label.pack(pady=(0, 25))

                        def start_main_program():
                            try:
                                dialog.destroy()
                                # Run main_program in a separate thread to avoid blocking
                                # Restore stdout to prevent conflicts with other GUI components
                                sys.stdout = sys.__stdout__
                                threading.Thread(target=Transformation_3.main_program, daemon=True).start()
                            except Exception as e:
                                print(f"Error starting transformation main program: {e}")

                        start_button = tk.Button(main_frame, text="Start Processing", command=start_main_program,
                                               bg='#3498db', fg='white', font=('Segoe UI', 11, 'normal'),
                                               width=20, height=2, relief='flat', bd=0,
                                               activebackground='#2980b9', activeforeground='white',
                                               cursor='hand2')
                        start_button.pack(pady=10)

                        # Handle dialog close
                        def on_dialog_close():
                            try:
                                dialog.destroy()
                            except:
                                pass

                        dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)

                    except tk.TclError as e:
                        if "application has been destroyed" in str(e) or "can't invoke" in str(e):
                            # Application is shutting down, just exit gracefully
                            return
                        else:
                            print(f"Tkinter error in transformation tool: {e}")
                    except Exception as e:
                        print(f"Error running transformation tool: {e}")

                # Run in the main thread to avoid threading issues
                try:
                    run_transformation()

                    # Check if the central UI still exists before updating status
                    if self.root.winfo_exists():
                        self.status_label.config(text="Data Transformation Tool launched successfully", foreground='green')
                except tk.TclError:
                    # Application might be shutting down, ignore silently
                    pass
                except Exception as e:
                    print(f"Error in transformation launch: {e}")
            else:
                raise ImportError("Transformation module not available")

        except tk.TclError as e:
            if "application has been destroyed" in str(e) or "can't invoke" in str(e):
                # Application is shutting down, just exit gracefully
                return
            else:
                print(f"Tkinter error: {e}")
        except Exception as e:
            # Check if the central UI still exists before showing error
            try:
                if self.root.winfo_exists():
                    messagebox.showerror("Error", f"Failed to launch Data Transformation Tool:\n{str(e)}")
                    self.status_label.config(text="Error launching Data Transformation Tool", foreground='red')
            except:
                # If error display fails, just print to console
                print(f"Failed to launch Data Transformation Tool: {e}")

    def launch_prizm_grabber(self):
        """Launch the PRIZM Data Grabber GUI"""
        try:
            self.status_label.config(text="Launching PRIZM Data Grabber...", foreground='blue')
            self.root.update()

            if prizm_grabber_available:
                # Import PRIZM_data_grabber here to avoid immediate execution of GUI code
                import PRIZM_data_grabber

                # Create the PRIZM data grabber GUI in the main thread, not in a separate thread
                def run_grabber():
                    try:
                        # Create a new Tkinter window for the PRIZM data grabber
                        new_root = tk.Toplevel(self.root)

                        # Configure window to stay on top of central UI
                        self.configure_child_window(new_root)

                        app = PRIZM_data_grabber.PrizmDataGrabberGUI(new_root)
                        # Don't call mainloop() - it's managed by the main root
                    except Exception as e:
                        print(f"Error running PRIZM data grabber: {e}")

                # Run in the main thread, not a separate thread
                run_grabber()
                self.status_label.config(text="PRIZM Data Grabber launched successfully", foreground='green')
            else:
                raise ImportError("PRIZM Data Grabber module not available")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch PRIZM Data Grabber:\n{str(e)}")
            self.status_label.config(text="Error launching PRIZM Data Grabber", foreground='red')

    def launch_subdivision_grabber(self):
        """Launch the Subdivision Data Grabber GUI"""
        try:
            self.status_label.config(text="Launching Subdivision Data Grabber...", foreground='blue')
            self.root.update()

            if subdivision_grabber_available:
                # Import Subdivision_data_grabber here to avoid immediate execution of GUI code
                import Subdivision_data_grabber

                # Create the Subdivision data grabber GUI in the main thread, not in a separate thread
                def run_grabber():
                    try:
                        # Create a new Tkinter window for the Subdivision data grabber
                        new_root = tk.Toplevel(self.root)

                        # Configure window to stay on top of central UI
                        self.configure_child_window(new_root)

                        app = Subdivision_data_grabber.SubdivisionDataGrabberGUI(new_root)
                        # Don't call mainloop() - it's managed by the main root
                    except Exception as e:
                        print(f"Error running subdivision data grabber: {e}")

                # Run in the main thread, not a separate thread
                run_grabber()
                self.status_label.config(text="Subdivision Data Grabber launched successfully", foreground='green')
            else:
                raise ImportError("Subdivision Data Grabber module not available")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch Subdivision Data Grabber:\n{str(e)}")
            self.status_label.config(text="Error launching Subdivision Data Grabber", foreground='red')

    def center_window_relative_to_parent(self, child_window):
        """Center a child window relative to the parent window"""
        # Update the window to get accurate dimensions
        child_window.update_idletasks()

        # Get parent window position and size
        parent_x = self.root.winfo_x()
        parent_y = self.root.winfo_y()
        parent_width = self.root.winfo_width()
        parent_height = self.root.winfo_height()

        # Get child window size
        child_width = child_window.winfo_reqwidth()
        child_height = child_window.winfo_reqheight()

        # Calculate center position
        x = parent_x + (parent_width // 2) - (child_width // 2)
        y = parent_y + (parent_height // 2) - (child_height // 2)

        # Ensure the window doesn't go off screen
        x = max(0, x)
        y = max(0, y)

        child_window.geometry(f"+{x}+{y}")

    def configure_child_window(self, child_window):
        """Configure a child window to stay on top of the central UI"""
        child_window.transient(self.root)  # Make it a transient window of the central UI
        child_window.lift()  # Bring to front
        child_window.focus_force()  # Force focus
        self.center_window_relative_to_parent(child_window)

        # Bind event to ensure the window stays on top when activated
        def on_window_focus(event):
            if child_window.winfo_exists():
                child_window.lift()

        child_window.bind('<FocusIn>', on_window_focus)

def main():
    root = tk.Tk()
    app = CentralUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
