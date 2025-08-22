from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
import time
import tkinter as tk
from tkinter import ttk, scrolledtext
import threading
import sys

def setup_driver():
    """Set up Chrome driver with options"""
    chrome_options = Options()
    # Uncomment the line below if you want to run in headless mode
    # chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")

    driver = webdriver.Chrome(options=chrome_options)
    return driver

def login_to_environics(driver):
    """Login to Environics Analytics website"""
    try:
        # Navigate to the login page
        print("Navigating to Environics Analytics...")
        driver.get("https://en.environicsanalytics.com/Envision/Home")

        # Wait for page to load
        wait = WebDriverWait(driver, 10)
        wait.until(EC.presence_of_element_located((By.ID, "Username")))

        # Set credentials
        username = ""
        password = ""

        # Find and fill username and password fields, then click button
        print("Entering username and password...")
        driver.find_element(By.ID,"Username").send_keys(username)
        driver.find_element(By.ID,"Password").send_keys(password)
        driver.find_element(By.ID,"button").click()

        # Wait for login to complete
        time.sleep(3)
        print("Login completed!")

    except Exception as e:
        print(f"Error during login: {e}")
        return False

    return True

def navigate_to_subdivision_dashboard(driver):
    """Navigate to the subdivision dashboard URL"""
    try:
        print("Navigating to subdivision dashboard...")
        dashboard_url = "https://en.environicsanalytics.com/Envision/ExecDashboard/Build?workpageId=1765&type=All&dbType=System&dbId=5460&wpOwnerType=System"
        driver.get(dashboard_url)

        # Wait for page to load
        time.sleep(5)
        print("Subdivision dashboard page loaded!")

    except Exception as e:
        print(f"Error navigating to subdivision dashboard: {e}")
        return False

    return True

def click_next_button(driver):
    """Click the Next button on the dashboard"""
    try:
        wait = WebDriverWait(driver, 10)

        print("Looking for Next button...")

        # Use the correct selector for the Next button
        next_button = wait.until(EC.element_to_be_clickable((By.ID, "page-bottom")))

        if next_button:
            print("Clicking Next button...")
            next_button.click()
            time.sleep(2)
            print("Next button clicked!")
            return True
        else:
            print("Next button not found.")
            return False

    except Exception as e:
        print(f"Error clicking Next button: {e}")
        return False

def select_british_columbia_geography(driver):
    """Select British Columbia from the geography table"""
    try:
        wait = WebDriverWait(driver, 10)

        # Wait for the geography table to load
        geography_table = wait.until(EC.presence_of_element_located((By.ID, "tbl-25f8e13ebbff4825adce9b3b742c75bf")))

        print("Selecting British Columbia from geography table...")

        # First check if BC is already selected
        try:
            selected_bc = driver.find_element(By.XPATH, "//tr[contains(.//span, 'British Columbia')]//i[@class='selector glyphicons selected circle_ok']")
            if selected_bc:
                print("British Columbia is already selected!")
                return True
        except:
            pass  # Not selected, continue with selection

        # Try to find British Columbia by text content
        try:
            bc_selector = wait.until(EC.element_to_be_clickable((By.XPATH, "//tr[contains(.//span, 'British Columbia')]//i[@class='selector glyphicons ok']")))
            bc_selector.click()
            print("British Columbia geography selected!")
            time.sleep(2)
            return True
        except Exception as e:
            print(f"Text search method failed: {e}")

        # Alternative method: Search through all rows for BC
        try:
            rows = driver.find_elements(By.CSS_SELECTOR, "#tbl-25f8e13ebbff4825adce9b3b742c75bf tbody tr")

            for i, row in enumerate(rows):
                row_text = row.text
                if "British Columbia" in row_text or "Colombie-Britannique" in row_text:
                    print(f"Found British Columbia in row {i+1}: {row_text}")
                    selector = row.find_element(By.CSS_SELECTOR, "i.selector.glyphicons.ok")
                    selector.click()
                    print("British Columbia geography selected using row search!")
                    time.sleep(2)
                    return True

            print("Could not find British Columbia in any row")

        except Exception as e:
            print(f"Row search method failed: {e}")

        return False

    except Exception as e:
        print(f"Error selecting British Columbia geography: {e}")
        return False

def select_target_set(driver, target_name):
    """Select a target set by searching for it"""
    try:
        wait = WebDriverWait(driver, 10)

        print(f"Selecting target set: {target_name}")

        # Find the search input for target sets
        search_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder='Filter your target sets']")))
        search_input.clear()
        search_input.send_keys(target_name)

        # Wait for search results to filter
        time.sleep(2)

        # Find and click the selector button for this target set
        target_table = wait.until(EC.presence_of_element_located((By.ID, "tbl-36771431bfc84f91b4ebc2ef26a4d264")))
        selector_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#tbl-36771431bfc84f91b4ebc2ef26a4d264 tbody tr:first-child td:first-child i.selector.glyphicons.ok")))
        selector_button.click()

        print(f"Selected target set: {target_name}")
        time.sleep(1)
        return True

    except Exception as e:
        print(f"Error selecting target set {target_name}: {e}")
        return False

def select_census_subdivision(driver):
    """Select Census Subdivision from the geography level dropdown"""
    try:
        wait = WebDriverWait(driver, 10)

        print("Selecting Census Subdivision geography level...")

        # Find the geography level dropdown
        geography_dropdown = wait.until(EC.presence_of_element_located((By.ID, "lst-140b72bff88247e2930074e35d3a775e")))

        # Use Select class to select by value
        select = Select(geography_dropdown)
        select.select_by_value("102")  # Census Subdivision value

        print("Census Subdivision selected!")
        time.sleep(2)
        return True

    except Exception as e:
        print(f"Error selecting Census Subdivision: {e}")
        return False

def select_prizm_data_source(driver):
    """Select PRIZM data source from the dropdown"""
    try:
        wait = WebDriverWait(driver, 10)

        print("Selecting PRIZM data source...")

        # Find the data source dropdown
        data_source_dropdown = wait.until(EC.presence_of_element_located((By.ID, "lst-d8ca3e52-52e7-4a21-8eed-68ebcc4d558c")))

        # Use Select class to select PRIZM
        select = Select(data_source_dropdown)
        select.select_by_value("sa|1053")  # PRIZM value

        print("PRIZM data source selected!")
        time.sleep(3)  # Wait for the tree to load
        return True

    except Exception as e:
        print(f"Error selecting PRIZM data source: {e}")
        return False

def select_total_population_variable(driver):
    """Select Total Population from the Basics category"""
    try:
        wait = WebDriverWait(driver, 10)

        print("Selecting Total Population variable...")

        # Wait for the PRIZM tree to be available
        time.sleep(2)

        # First expand the PRIZM category by clicking on its toggle
        try:
            print("Expanding PRIZM category...")
            # Target the specific toggle for the PRIZM category
            prizm_toggle = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[@id='sa|1053']/i[@class='jstree-icon jstree-ocl e-backstagechild']")))
            prizm_toggle.click()
            time.sleep(3)  # Wait for the PRIZM category to expand
            print("Expanded PRIZM category")
        except Exception as e:
            print(f"Could not find or click PRIZM toggle: {e}")
            # Try alternative method using CSS selector with escaped characters
            try:
                prizm_toggle = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#sa\\|1053 > i.jstree-icon.jstree-ocl.e-backstagechild")))
                prizm_toggle.click()
                time.sleep(3)
                print("Expanded PRIZM category using alternative CSS method")
            except Exception as e2:
                print(f"Alternative PRIZM toggle method also failed: {e2}")
                return False

        # Now expand the Basics category by clicking on its toggle
        try:
            print("Expanding Basics category...")
            # Target the specific toggle for the Basics category
            basics_toggle = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[@id='sa|1053|Basics']/i[@class='jstree-icon jstree-ocl e-backstagechild']")))
            basics_toggle.click()
            time.sleep(3)  # Wait for the Basics category to expand
            print("Expanded Basics category")
        except Exception as e:
            print(f"Could not find or click Basics toggle: {e}")
            # Try alternative method using CSS selector with escaped characters
            try:
                basics_toggle = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#sa\\|1053\\|Basics > i.jstree-icon.jstree-ocl.e-backstagechild")))
                basics_toggle.click()
                time.sleep(3)
                print("Expanded Basics category using alternative CSS method")
            except Exception as e2:
                print(f"Alternative Basics toggle method also failed: {e2}")
                # Try clicking the Basics anchor directly
                try:
                    basics_anchor = wait.until(EC.element_to_be_clickable((By.ID, "sa|1053|Basics_anchor")))
                    basics_anchor.click()
                    time.sleep(3)
                    print("Expanded Basics category by clicking anchor")
                except Exception as e3:
                    print(f"Basics anchor click method also failed: {e3}")
                    return False

        # Now find and click the Total Population checkbox
        try:
            print("Looking for Total Population checkbox...")
            total_pop_checkbox = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#sa\\|1053\\|Basics\\|ECYBASPOP_anchor > i.jstree-icon.jstree-checkbox")))
            total_pop_checkbox.click()
            print("Total Population variable selected!")
            time.sleep(1)
            return True
        except Exception as e:
            print(f"Direct Total Population selection failed: {e}")

            # Alternative method using XPath
            try:
                total_pop_checkbox = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@id='sa|1053|Basics|ECYBASPOP_anchor']/i[@class='jstree-icon jstree-checkbox']")))
                total_pop_checkbox.click()
                print("Total Population variable selected using alternative method!")
                time.sleep(1)
                return True
            except Exception as e2:
                print(f"Alternative Total Population method also failed: {e2}")

                # Try to find it by text content
                try:
                    total_pop_element = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Total Population')]/i[@class='jstree-icon jstree-checkbox']")))
                    total_pop_element.click()
                    print("Total Population variable selected using text search!")
                    time.sleep(1)
                    return True
                except Exception as e3:
                    print(f"Text search method also failed: {e3}")
                    # Final attempt using JavaScript
                    try:
                        driver.execute_script("""
                            // First expand PRIZM if not already expanded
                            var prizmToggle = document.querySelector('#sa\\|1053 > i.jstree-icon.jstree-ocl');
                            if (prizmToggle) {
                                prizmToggle.click();
                                setTimeout(function() {
                                    // Then expand Basics
                                    var basicsToggle = document.querySelector('#sa\\|1053\\|Basics > i.jstree-icon.jstree-ocl');
                                    if (basicsToggle) {
                                        basicsToggle.click();
                                        setTimeout(function() {
                                            // Finally select Total Population
                                            var totalPopCheckbox = document.querySelector('#sa\\|1053\\|Basics\\|ECYBASPOP_anchor > i.jstree-icon.jstree-checkbox');
                                            if (totalPopCheckbox) {
                                                totalPopCheckbox.click();
                                            }
                                        }, 2000);
                                    }
                                }, 2000);
                            }
                        """)
                        time.sleep(5)  # Wait longer for the JavaScript execution
                        print("Total Population variable selected using JavaScript!")
                        return True
                    except Exception as e4:
                        print(f"JavaScript method also failed: {e4}")
                        return False

    except Exception as e:
        print(f"Error selecting Total Population variable: {e}")
        return False

def create_subdivision_dashboard(driver, dashboard_name):
    """Enter dashboard name and click create dashboard button"""
    try:
        wait = WebDriverWait(driver, 10)

        print(f"Creating subdivision dashboard with name: {dashboard_name}")

        # Find dashboard name input field
        dashboard_name_input = wait.until(EC.presence_of_element_located((By.ID, "NewDashboardName")))

        if dashboard_name_input:
            dashboard_name_input.clear()
            dashboard_name_input.send_keys(dashboard_name)
            print(f"Entered dashboard name: {dashboard_name}")
        else:
            print("Could not find dashboard name input field")
            return False

        # Find and click create dashboard button
        try:
            create_button = wait.until(EC.element_to_be_clickable((By.ID, "btn-create")))
            create_button.click()
            print("Clicked Create Dashboard button")
            time.sleep(3)  # Wait for dashboard creation
            return True
        except Exception as e:
            print(f"Could not find or click Create Dashboard button: {e}")
            return False

    except Exception as e:
        print(f"Error creating subdivision dashboard: {e}")
        return False

def process_single_target_set(driver, target_name, iteration_index):
    """Process a single target set to create one subdivision dashboard"""
    try:
        dashboard_url = "https://en.environicsanalytics.com/Envision/ExecDashboard/Build?workpageId=1765&type=All&dbType=System&dbId=5460&wpOwnerType=System"

        print(f"\n--- Creating subdivision dashboard for {target_name} ({iteration_index + 1}/51) ---")

        # Navigate back to dashboard URL if not the first iteration
        if iteration_index > 0:
            print("Navigating back to subdivision dashboard URL...")
            driver.get(dashboard_url)
            time.sleep(5)

            if not click_next_button(driver):
                print(f"Could not click Next button for target: {target_name}")
                return False

        # Select British Columbia geography
        if not select_british_columbia_geography(driver):
            print(f"Could not select British Columbia geography")
            return False

        # Select the target set
        if not select_target_set(driver, target_name):
            print(f"Could not select target set: {target_name}")
            return False

        # Select Census Subdivision geography level
        if not select_census_subdivision(driver):
            print(f"Could not select Census Subdivision geography level")
            return False

        # Select PRIZM data source
        if not select_prizm_data_source(driver):
            print(f"Could not select PRIZM data source")
            return False

        # Select Total Population variable
        if not select_total_population_variable(driver):
            print(f"Could not select Total Population variable")
            return False

        # Create dashboard with descriptive name
        dashboard_name = f"Subdivision_{target_name}_{iteration_index + 1}"
        if not create_subdivision_dashboard(driver, dashboard_name):
            print(f"Could not create subdivision dashboard: {dashboard_name}")
            return False

        print(f"Successfully created subdivision dashboard: {dashboard_name}")
        return True

    except Exception as e:
        print(f"Error processing target set {target_name}: {e}")
        return False

def process_all_target_sets(driver, target_set_names):
    """Process all target sets, creating one subdivision dashboard for each"""
    try:
        print(f"Processing {len(target_set_names)} target sets for subdivision data...")

        for i, target_name in enumerate(target_set_names):
            print(f"\n=== Processing target set {i+1}/{len(target_set_names)}: {target_name} ===")

            if not process_single_target_set(driver, target_name, i):
                print(f"Failed to process target set: {target_name}")
                continue

            print(f"Successfully completed subdivision dashboard for target set: {target_name}")

        print("\nAll target sets processed for subdivision data!")
        return True

    except Exception as e:
        print(f"Error processing target sets: {e}")
        return False

class TextRedirector:
    """Class to redirect stdout to the GUI text widget"""
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, string):
        self.text_widget.insert(tk.END, string)
        self.text_widget.see(tk.END)
        self.text_widget.update()

    def flush(self):
        pass

class SubdivisionDataGrabberGUI:
    """Simple GUI for the Subdivision Data Grabber"""

    def __init__(self, root):
        self.root = root
        self.root.title("Subdivision Data Grabber")
        self.root.geometry("800x600")
        self.window_closed = False  # Add flag to track if window is closed

        # Main frame
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")

        # Configure grid weights
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)

        # Title
        title_label = ttk.Label(main_frame, text="Subdivision Data Grabber",
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, pady=(0, 20))

        # Start button
        self.start_button = ttk.Button(main_frame, text="Start Data Grabbing",
                                      command=self.start_process)
        self.start_button.grid(row=1, column=0, pady=10, sticky="n")

        # Log frame
        log_frame = ttk.LabelFrame(main_frame, text="Process Log", padding="10")
        log_frame.grid(row=2, column=0, sticky="nsew", pady=(20, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        # Log text area with scrollbar
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, width=80, height=20)
        self.log_text.grid(row=0, column=0, sticky="nsew")

        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=3, column=0, sticky="ew", pady=(10, 0))

        # Status label
        self.status_label = ttk.Label(main_frame, text="Ready to start", font=("Arial", 10))
        self.status_label.grid(row=4, column=0, pady=(5, 0))

        # Redirect stdout to log text
        self.text_redirector = TextRedirector(self.log_text)

        # Thread for running the main process
        self.process_thread = None
        self.is_running = False

        # Set up proper cleanup when window is closed
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        """Handle window closing event"""
        self.window_closed = True
        if self.process_thread and self.process_thread.is_alive():
            print("Window closing - process will continue in background...")
        self.root.destroy()

    def start_process(self):
        """Start the data grabbing process in a separate thread"""
        if self.is_running:
            return

        self.is_running = True
        self.start_button.config(text="Running...", state="disabled")
        self.progress.start()
        self.status_label.config(text="Starting subdivision data grabbing...")

        # Clear the log
        self.log_text.delete(1.0, tk.END)

        # Redirect stdout to the GUI
        sys.stdout = self.text_redirector

        # Start the process in a separate thread
        self.process_thread = threading.Thread(target=self.run_main_process)
        self.process_thread.daemon = True
        self.process_thread.start()

    def run_main_process(self):
        """Run the main data grabbing process"""
        try:
            main()
        except Exception as e:
            print(f"An error occurred in the main process: {e}")
        finally:
            # Only try to update GUI if window is still open
            if not self.window_closed:
                try:
                    self.root.after(0, self.process_finished)
                except:
                    pass  # Window might have been closed

    def process_finished(self):
        """Called when the process is finished"""
        # Check if window is still open before updating GUI elements
        if self.window_closed:
            return

        try:
            self.is_running = False
            self.start_button.config(text="Start Data Grabbing", state="normal")
            self.progress.stop()
            self.status_label.config(text="Process completed")

            # Restore stdout
            sys.stdout = sys.__stdout__
        except Exception as e:
            # GUI elements might have been destroyed
            print(f"Could not update GUI after process completion: {e}")

def main():
    """Main function to run the subdivision data grabbing process"""
    driver = None

    try:
        # Set up the driver
        driver = setup_driver()

        # Login to the website
        if not login_to_environics(driver):
            print("Login failed!")
            return

        # Navigate to subdivision dashboard
        if not navigate_to_subdivision_dashboard(driver):
            print("Navigation to subdivision dashboard failed!")
            return

        # Click Next button for the first time
        if not click_next_button(driver):
            print("Could not find or click Next button. Please check the page manually.")
            return

        # Process target sets - each will create 1 subdivision dashboard
        target_sets = ["The A-List", "Wealthy & Wise", "Asian Sophisticates", "Turbo Burbs", "First-Class Families", "Downtown Verve", "Mature & Secure", "Multiculture-ish", "Boomer Bliss", "Asian Achievement", "Modern Suburbia", "Eat, Play, Love", "Kick-Back Country", "South Asian Enterprise", "Savvy Seniors", "Asian Avenues", "Multicultural Corners", "Family Mode", "New Asian Heights", "Scenic Retirement", "Indieville", "Mid-City Mellow", "All-Terrain Families", "Suburban Sports", "Country Traditions", "Latte Life", "South Asian Society", "Metro Melting Pot", "Diverse & Determined", "New Country", "Middle-Class Mosaic", "Keep on Trucking", "Stressed in Suburbia", "Down to Earth", "Happy Medium", "Slow-Lane Suburbs", "Social Networkers", "Agri-Biz", "Backcountry Boomers", "Country & Western", "On Their Own Again", "Friends & Roomies", "Silver Flats", "Juggling Acts", "Old Town Roads", "Value Villagers", "Came From Away", "Suburban Recliners", "Midtown Movers", "Indigenous Families", "Just Getting By"]

        if not process_all_target_sets(driver, target_sets):
            print("Could not process target sets. Please check the page manually.")

        print("All subdivision dashboards created successfully! Browser will remain open for further instructions.")
        print("Press Enter to close the browser...")
        input()

    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        if driver:
            driver.quit()

def create_gui():
    """Create and run the GUI"""
    root = tk.Tk()
    app = SubdivisionDataGrabberGUI(root)
    root.mainloop()

if __name__ == "__main__":
    create_gui()
