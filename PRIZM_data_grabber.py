from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
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
        username = "charnley@symphonytourism.ca"
        password = "Symphony&TOTA24!"

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

def navigate_to_dashboard(driver):
    """Navigate to the specified dashboard URL"""
    try:
        print("Navigating to dashboard...")
        dashboard_url = "https://en.environicsanalytics.com/Envision/ExecDashboard/Build?workpageId=289&dbType=System&dbId=7633&type=All&wpOwnerType=System"
        driver.get(dashboard_url)

        # Wait for page to load
        time.sleep(5)
        print("Dashboard page loaded!")

    except Exception as e:
        print(f"Error navigating to dashboard: {e}")
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

def select_data_source(driver, data_source_value):
    """Select a data source from the dropdown"""
    try:
        wait = WebDriverWait(driver, 10)

        print(f"Selecting data source: {data_source_value}")

        # Try to interact with the actual select element using JavaScript
        try:
            dropdown_element = wait.until(EC.presence_of_element_located((By.ID, "lst-1cc5241b-e22e-41bd-8d68-04f4772db1b2")))

            # Use JavaScript to make the select visible and change its value
            driver.execute_script("""
                var select = arguments[0];
                var value = arguments[1];
                select.style.display = 'block';
                select.style.visibility = 'visible';
                select.value = value;
                
                // Trigger change event
                var event = new Event('change', { bubbles: true });
                select.dispatchEvent(event);
                
                // Also trigger jQuery change event if jQuery is available
                if (typeof $ !== 'undefined') {
                    $(select).trigger('change');
                }
            """, dropdown_element, data_source_value)

            print(f"Set dropdown value to {data_source_value} using JavaScript")
            time.sleep(3)  # Wait longer for the change to be processed
            return True

        except Exception as e:
            print(f"Error selecting data source: {e}")
            return False

    except Exception as e:
        print(f"Error selecting data source: {e}")
        return False

def select_variables_opticks_numeris_set1(driver):
    """Select Basics, Restaurants, Loyalty Programs, Media - TV - Usage, Media - Internet - Usage variables"""
    try:
        wait = WebDriverWait(driver, 10)

        print("Selecting Opticks Numeris variables - Set 1...")

        # First expand the Opticks Numeris tree
        opticks_numeris_toggle = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#sa\\|1063 > i.jstree-icon.jstree-ocl")))
        opticks_numeris_toggle.click()
        time.sleep(2)

        # Variables to select with their expected IDs
        variables = [
            ("Basics", "sa|1063|Basics_anchor"),
            ("Restaurants", "sa|1063|Restaurants_anchor"),
            ("Loyalty Programs", "sa|1063|Loyalty Programs_anchor"),
            ("Media - TV - Usage", "sa|1063|Media - TV - Usage_anchor"),
            ("Media - Internet - Usage", "sa|1063|Media - Internet - Usage_anchor")
        ]

        for variable_name, variable_id in variables:
            try:
                # Find and click the checkbox for each variable using the specific ID
                variable_checkbox = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f"#{variable_id.replace('|', '\\|').replace(' ', '\\ ').replace('-', '\\-')} > i.jstree-icon.jstree-checkbox")))
                variable_checkbox.click()
                print(f"Selected variable: {variable_name}")
                time.sleep(1)
            except Exception as e:
                print(f"Could not select variable '{variable_name}': {e}")
                # Try alternative selector
                try:
                    variable_element = wait.until(EC.element_to_be_clickable((By.XPATH, f"//a[@id='{variable_id}']/i[@class='jstree-icon jstree-checkbox']")))
                    variable_element.click()
                    print(f"Selected variable using alternative method: {variable_name}")
                    time.sleep(1)
                except Exception as e2:
                    print(f"Alternative method also failed for '{variable_name}': {e2}")

        return True

    except Exception as e:
        print(f"Error selecting Opticks Numeris variables - Set 1: {e}")
        return False

def select_variables_opticks_numeris_set2(driver):
    """Select Media - Direct & Outdoor - Usage, Sports & Leisure, Travel - Personal & Business variables"""
    try:
        wait = WebDriverWait(driver, 10)

        print("Selecting Opticks Numeris variables - Set 2...")

        # First expand the Opticks Numeris tree
        opticks_numeris_toggle = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#sa\\|1063 > i.jstree-icon.jstree-ocl")))
        opticks_numeris_toggle.click()
        time.sleep(2)

        # Variables to select with their expected IDs
        variables = [
            ("Media - Direct & Outdoor - Usage", "sa|1063|Media - Direct & Outdoor - Usage_anchor"),
            ("Sports & Leisure", "sa|1063|Sports & Leisure_anchor"),
            ("Travel - Personal & Business", "sa|1063|Travel - Personal & Business_anchor")
        ]

        for variable_name, variable_id in variables:
            try:
                # Find and click the checkbox for each variable using the specific ID
                variable_checkbox = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f"#{variable_id.replace('|', '\\|').replace(' ', '\\ ').replace('-', '\\-').replace('&', '\\&')} > i.jstree-icon.jstree-checkbox")))
                variable_checkbox.click()
                print(f"Selected variable: {variable_name}")
                time.sleep(1)
            except Exception as e:
                print(f"Could not select variable '{variable_name}': {e}")
                # Try alternative selector
                try:
                    variable_element = wait.until(EC.element_to_be_clickable((By.XPATH, f"//a[@id='{variable_id}']/i[@class='jstree-icon jstree-checkbox']")))
                    variable_element.click()
                    print(f"Selected variable using alternative method: {variable_name}")
                    time.sleep(1)
                except Exception as e2:
                    print(f"Alternative method also failed for '{variable_name}': {e2}")

        return True

    except Exception as e:
        print(f"Error selecting Opticks Numeris variables - Set 2: {e}")
        return False

def select_my_variables(driver):
    """Select the My Variables checkbox"""
    try:
        wait = WebDriverWait(driver, 10)

        print("Selecting My Variables...")

        # Click the checkbox for My Variables
        my_variables_checkbox = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#udv\\:sasl\\|0_anchor > i.jstree-icon.jstree-checkbox")))
        my_variables_checkbox.click()

        print("My Variables selected!")
        time.sleep(1)
        return True

    except Exception as e:
        print(f"Error selecting My Variables: {e}")
        return False

def select_opticks_social(driver):
    """Select the main Opticks Social checkbox"""
    try:
        wait = WebDriverWait(driver, 10)

        print("Selecting Opticks Social...")

        # Click the checkbox for Opticks Social
        opticks_social_checkbox = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#sa\\|1061_anchor > i.jstree-icon.jstree-checkbox")))
        opticks_social_checkbox.click()

        print("Opticks Social selected!")
        time.sleep(1)
        return True

    except Exception as e:
        print(f"Error selecting Opticks Social: {e}")
        return False

def process_target_set_with_variables(driver, target_name, iteration_index):
    """Process a single target set with all 4 variable configurations"""
    try:
        wait = WebDriverWait(driver, 10)
        dashboard_url = "https://en.environicsanalytics.com/Envision/ExecDashboard/Build?workpageId=289&dbType=System&dbId=7633&type=All&wpOwnerType=System"

        # Configuration for the 4 different dashboards
        dashboard_configs = [
            {
                "name_suffix": "Numeris_Set1",
                "data_source": "sa|1063",  # Opticks Numeris
                "variable_function": select_variables_opticks_numeris_set1
            },
            {
                "name_suffix": "Numeris_Set2",
                "data_source": "sa|1063",  # Opticks Numeris
                "variable_function": select_variables_opticks_numeris_set2
            },
            {
                "name_suffix": "MyVariables",
                "data_source": "sa|1063",  # Opticks Numeris (to access My Variables)
                "variable_function": select_my_variables
            },
            {
                "name_suffix": "Social",
                "data_source": "sa|1061",  # Opticks Social
                "variable_function": select_opticks_social
            }
        ]

        for config_index, config in enumerate(dashboard_configs):
            try:
                print(f"\n--- Creating dashboard {config_index + 1}/4 for {target_name}: {config['name_suffix']} ---")

                # Navigate back to dashboard URL if not the first dashboard
                if config_index > 0:
                    print("Navigating back to dashboard URL...")
                    driver.get(dashboard_url)
                    time.sleep(5)

                    if not click_next_button(driver):
                        print(f"Could not click Next button for dashboard: {config['name_suffix']}")
                        continue

                # Select the target set
                if not select_single_target_set(driver, target_name):
                    print(f"Could not select target set: {target_name}")
                    continue

                # Select British Columbia benchmark
                if not select_british_columbia_benchmark(driver):
                    print(f"Could not select British Columbia benchmark")
                    continue

                # Select data source
                if not select_data_source(driver, config["data_source"]):
                    print(f"Could not select data source: {config['data_source']}")
                    continue

                # Select variables using the appropriate function
                if not config["variable_function"](driver):
                    print(f"Could not select variables for: {config['name_suffix']}")
                    continue

                # Create dashboard with descriptive name
                dashboard_name = f"{target_name}_{config['name_suffix']}_{iteration_index + 1}"
                if not create_dashboard(driver, dashboard_name):
                    print(f"Could not create dashboard: {dashboard_name}")
                    continue

                print(f"Successfully created dashboard: {dashboard_name}")

            except Exception as e:
                print(f"Error creating dashboard {config['name_suffix']} for {target_name}: {e}")
                continue

        return True

    except Exception as e:
        print(f"Error processing target set {target_name}: {e}")
        return False

def select_single_target_set(driver, target_name):
    """Select a single target set by name"""
    try:
        wait = WebDriverWait(driver, 10)

        # Wait for the target sets table to load
        target_table = wait.until(EC.presence_of_element_located((By.ID, "tbl-5d7536c144f9416c9e60fcbcf581a82a")))

        # Find and use the search input for target sets
        search_input = driver.find_element(By.CSS_SELECTOR, "#tbl-5d7536c144f9416c9e60fcbcf581a82a_filter input")
        search_input.clear()
        search_input.send_keys(target_name)

        # Wait for search results to filter
        time.sleep(2)

        # Find and click the selector button for this target set
        selector_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#tbl-5d7536c144f9416c9e60fcbcf581a82a tbody tr:first-child td:first-child i.selector.glyphicons.ok")))
        selector_button.click()

        print(f"Selected target set: {target_name}")
        time.sleep(1)
        return True

    except Exception as e:
        print(f"Error selecting target set {target_name}: {e}")
        return False

def select_target_sets(driver, target_set_names):
    """Select target sets one by one, creating 4 dashboards for each"""
    try:
        print(f"Processing target sets: {target_set_names}")

        for i, target_name in enumerate(target_set_names):
            print(f"\n=== Processing target set {i+1}/{len(target_set_names)}: {target_name} ===")

            if not process_target_set_with_variables(driver, target_name, i):
                print(f"Failed to process target set: {target_name}")
                continue

            print(f"Successfully completed all dashboards for target set: {target_name}")

        print("\nAll target sets processed!")
        return True

    except Exception as e:
        print(f"Error processing target sets: {e}")
        return False

def select_british_columbia_benchmark(driver):
    """Select British Columbia as the benchmark"""
    try:
        wait = WebDriverWait(driver, 10)

        # Wait for the benchmark table to load
        benchmark_table = wait.until(EC.presence_of_element_located((By.ID, "tbl-64dec1191446488db16dc83e40c62a7f")))

        print("Selecting British Columbia benchmark...")

        # First check if BC is already selected
        try:
            selected_bc = driver.find_element(By.XPATH, "//tr[contains(.//span, 'British Columbia')]//i[@class='selector glyphicons selected circle_ok']")
            if selected_bc:
                print("British Columbia is already selected!")
                return True
        except:
            pass  # Not selected, continue with selection

        # Method 1: Try to find British Columbia by text content
        try:
            # Look for the row containing "British Columbia" text with unselected selector
            bc_selector = wait.until(EC.element_to_be_clickable((By.XPATH, "//tr[contains(.//span, 'British Columbia')]//i[@class='selector glyphicons ok']")))
            bc_selector.click()
            print("British Columbia benchmark selected using text search!")
            time.sleep(2)
            return True
        except Exception as e:
            print(f"Text search method failed: {e}")

        # Method 2: Search through all rows for BC
        try:
            rows = driver.find_elements(By.CSS_SELECTOR, "#tbl-64dec1191446488db16dc83e40c62a7f tbody tr")

            for i, row in enumerate(rows):
                row_text = row.text
                if "British Columbia" in row_text or "Colombie-Britannique" in row_text:
                    print(f"Found British Columbia in row {i+1}: {row_text}")
                    # Find the unselected selector in this row
                    selector = row.find_element(By.CSS_SELECTOR, "i.selector.glyphicons.ok")
                    selector.click()
                    print("British Columbia benchmark selected using row search!")
                    time.sleep(2)
                    return True

            print("Could not find British Columbia in any row")

        except Exception as e:
            print(f"Row search method failed: {e}")

        # Method 3: JavaScript approach (last resort)
        try:
            result = driver.execute_script("""
                var table = document.getElementById('tbl-64dec1191446488db16dc83e40c62a7f');
                var rows = table.getElementsByTagName('tr');
                
                for (var i = 0; i < rows.length; i++) {
                    var rowText = rows[i].textContent || rows[i].innerText;
                    if (rowText.includes('British Columbia') || rowText.includes('Colombie-Britannique')) {
                        var selector = rows[i].querySelector('i.selector.glyphicons.ok');
                        if (selector) {
                            selector.click();
                            return true;
                        }
                    }
                }
                return false;
            """)

            if result:
                print("British Columbia benchmark selected using JavaScript!")
                time.sleep(2)
                return True
            else:
                print("JavaScript method could not find unselected BC selector")
        except Exception as e:
            print(f"JavaScript method failed: {e}")

        print("All methods failed to select British Columbia benchmark")
        return False

    except Exception as e:
        print(f"Error selecting British Columbia benchmark: {e}")
        return False

def create_dashboard(driver, dashboard_name):
    """Enter dashboard name and click create dashboard button"""
    try:
        wait = WebDriverWait(driver, 10)

        print(f"Creating dashboard with name: {dashboard_name}")

        # Find dashboard name input field using the exact selector
        dashboard_name_input = wait.until(EC.presence_of_element_located((By.ID, "NewDashboardName")))

        if dashboard_name_input:
            dashboard_name_input.clear()
            dashboard_name_input.send_keys(dashboard_name)
            print(f"Entered dashboard name: {dashboard_name}")
        else:
            print("Could not find dashboard name input field")
            return False

        # Find and click create dashboard button using the correct selector
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
        print(f"Error creating dashboard: {e}")
        return False

class TextRedirector:
    """Class to redirect stdout to the GUI text widget"""
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag

    def write(self, str):
        self.widget.configure(state='normal')
        self.widget.insert("end", str, (self.tag,))
        self.widget.configure(state='disabled')
        self.widget.see("end")

    def flush(self):
        pass

class PrizmDataGrabberGUI:
    """GUI for PRIZM Data Grabber"""

    def __init__(self, root):
        self.root = root
        self.root.title("PRIZM Data Grabber")
        self.root.geometry("800x600")

        # Variables
        self.is_running = False
        self.automation_thread = None
        self.window_closed = False  # Add flag to track if window is closed

        self.create_widgets()
        self.setup_logging()

        # Set up proper cleanup when window is closed
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        """Create the GUI widgets"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)

        # Title
        title_label = ttk.Label(main_frame, text="PRIZM Data Grabber", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))

        # Start button
        self.start_button = ttk.Button(main_frame, text="Start Data Grabbing", command=self.start_automation)
        self.start_button.grid(row=0, column=2, padx=(10, 0))

        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=0, column=3, padx=(10, 0), sticky=(tk.W, tk.E))
        main_frame.columnconfigure(3, weight=1)

        # Log area label
        log_label = ttk.Label(main_frame, text="Console Output:")
        log_label.grid(row=1, column=0, sticky=tk.W, pady=(10, 5))

        # Log text area with scrollbar
        self.log_text = scrolledtext.ScrolledText(
            main_frame,
            height=25,
            width=80,
            state='disabled',
            font=("Consolas", 9)
        )
        self.log_text.grid(row=2, column=0, columnspan=4, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))

        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready to start")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=3, column=0, columnspan=4, sticky=(tk.W, tk.E))

    def setup_logging(self):
        """Setup logging redirection to GUI"""
        # Redirect stdout to the text widget
        sys.stdout = TextRedirector(self.log_text, "stdout")

    def start_automation(self):
        """Start the automation in a separate thread"""
        if not self.is_running:
            self.is_running = True
            self.start_button.configure(text="Running...", state='disabled')
            self.progress.start()
            self.status_var.set("Starting automation...")

            # Clear the log
            self.log_text.configure(state='normal')
            self.log_text.delete(1.0, tk.END)
            self.log_text.configure(state='disabled')

            # Start automation in separate thread
            self.automation_thread = threading.Thread(target=self.run_automation, daemon=True)
            self.automation_thread.start()

    def on_closing(self):
        """Handle window closing event"""
        self.window_closed = True
        if self.automation_thread and self.automation_thread.is_alive():
            print("Window closing - automation will continue in background...")
        self.root.destroy()

    def run_automation(self):
        """Run the main automation function"""
        try:
            main()
        except Exception as e:
            print(f"Error in automation: {e}")
        finally:
            # Only try to update GUI if window is still open
            if not self.window_closed:
                try:
                    self.root.after(0, self.automation_complete)
                except:
                    pass  # Window might have been closed

    def automation_complete(self):
        """Called when automation completes"""
        # Check if window is still open before updating GUI elements
        if self.window_closed:
            return

        try:
            self.is_running = False
            self.start_button.configure(text="Start Data Grabbing", state='normal')
            self.progress.stop()
            self.status_var.set("Automation completed")
        except Exception as e:
            # GUI elements might have been destroyed
            print(f"Could not update GUI after automation completion: {e}")

def create_gui():
    """Create and run the GUI"""
    root = tk.Tk()
    app = PrizmDataGrabberGUI(root)
    root.mainloop()

def main():
    """Main function to run the automation"""
    driver = None

    try:
        # Set up the driver
        driver = setup_driver()

        # Login to the website
        if not login_to_environics(driver):
            print("Login failed!")
            return

        # Navigate to dashboard
        if not navigate_to_dashboard(driver):
            print("Navigation to dashboard failed!")
            return

        # Click Next button for the first time
        if not click_next_button(driver):
            print("Could not find or click Next button. Please check the page manually.")
            return

        # Process target sets - each will create 4 dashboards with different variables
        target_sets = ["The A-List", "Wealthy & Wise", "Asian Sophisticates", "Turbo Burbs", "First-Class Families", "Downtown Verve", "Mature & Secure", "Multiculture-ish", "Boomer Bliss", "Asian Achievement", "Modern Suburbia", "Eat, Play, Love", "Kick-Back Country", "South Asian Enterprise", "Savvy Seniors", "Asian Avenues", "Multicultural Corners", "Family Mode", "New Asian Heights", "Scenic Retirement", "Indieville", "Mid-City Mellow", "All-Terrain Families", "Suburban Sports", "Country Traditions", "Latte Life", "South Asian Society", "Metro Melting Pot", "Diverse & Determined", "New Country", "Middle-Class Mosaic", "Keep on Trucking", "Stressed in Suburbia", "Down to Earth", "Happy Medium", "Slow-Lane Suburbs", "Social Networkers", "Agri-Biz", "Backcountry Boomers", "Country & Western", "On Their Own Again", "Friends & Roomies", "Silver Flats", "Juggling Acts", "Old Town Roads", "Value Villagers", "Came From Away", "Suburban Recliners", "Midtown Movers", "Indigenous Families", "Just Getting By"]
        if not select_target_sets(driver, target_sets):
            print("Could not process target sets. Please check the page manually.")

        print("All dashboards created successfully! Browser will remain open for further instructions.")
        print("Press Enter to close the browser...")
        input()

    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        if driver:
            driver.quit()

if __name__ == "__main__":
    create_gui()
