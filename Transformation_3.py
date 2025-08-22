import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import math
from tkinter import *

# Files location - Updated to be more generic
Directory = 'C:\\Users\\Downloads\\VisitorView Extract\\Prince George'


# Digital credentials to log in to envronics

username = ""
password = ""

def main_program():
    """Main program functionality"""
    # Go to the web page and use credentials to log in
    try:
        driver = webdriver.Chrome()

        driver.get("https://en.environicsanalytics.com/Envision/Import-Data/Customers/Geocode")
        driver.find_element(By.ID,"Username").send_keys(username)
        driver.find_element(By.ID,"Password").send_keys(password)
        driver.find_element(By.ID,"button").click()

        wait = WebDriverWait(driver, 60)

        results='/html/body/div[1]/div[1]/div[2]/ul[4]/li[2]/a/div'

        wait.until(lambda driver: driver.find_element(By.XPATH,results)).click()

        t=2

    except Exception as e:
        # Remove messagebox error window, just print to console
        print(f"Error initializing browser or logging in: {e}")
        return

    root= Tk()

    root.title("Transformation to Excel")
    root.geometry("200x100")

    Label(root, text="Files").grid(row=2,column=10)
    Label(root, text="Position").grid(row=3,column=10)
    Label(root, text="Page").grid(row=4,column=10)

    fileTextbox = Entry(root)
    fileTextbox.grid(row=2, column=11)
    positionTextbox = Entry(root)
    positionTextbox.grid(row=3, column=11)
    pageTextbox = Entry(root)
    pageTextbox.grid(row=4, column=11)

    def parameters():
        try:
            NUMBER = int(fileTextbox.get())
            Position = int(positionTextbox.get())  # This will be used for navigation
            page = int(pageTextbox.get())

            Type = '//*[@id="optLong"]/a/div[1]'  # Second position

            CLICKS = int(math.ceil(NUMBER/10))
            rest = NUMBER % 10

            # Fix: Handle case when rest is 0
            if rest == 0:
                rest = 10

            # Navigate to the starting page first
            if page > 0:
                for nav_page in range(page):
                    # Use the correct selector for the Next button
                    next_button = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#tableview_next a")))
                    webdriver.ActionChains(driver).move_to_element(next_button).click(next_button).perform()
                    time.sleep(t)

            for k in range(page, CLICKS):
                if k == CLICKS - 1:
                    cicle = rest
                else:
                    cicle = 10

                for i in range(cicle):
                    # Removed the nested navigation loop that was causing issues

                    exp = "/html/body/div[1]/div[2]/div[3]/div/div[2]/div/table/tbody/tr["+str(i+1)+"]/td[2]/span/img[1]"

                    time.sleep(t)

                    wait.until(lambda driver: driver.find_element(By.XPATH, exp)).click()

                    time.sleep(t)

                    Desplegar = '/html/body/div[1]/div[2]/div[3]/div/div[1]/a[2]/i[1]'
                    element1 = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, Desplegar)))
                    webdriver.ActionChains(driver).move_to_element(element1).click(element1).perform()

                    time.sleep(t)

                    DL = '/html/body/div[1]/div[2]/div[3]/div/div[1]/div[1]/div[1]/div[2]/div[1]/h4'

                    element1 = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, DL)))
                    webdriver.ActionChains(driver).move_to_element(element1).click(element1).perform()

                    time.sleep(t)

                    element1 = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, Type)))
                    webdriver.ActionChains(driver).move_to_element(element1).click(element1).perform()

                    time.sleep(t)

                    excel = '/html/body/div[1]/div[2]/div[3]/div/div[1]/div[4]/div[2]/div/div[2]/div/div/div[1]'

                    element1 = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, excel)))
                    webdriver.ActionChains(driver).move_to_element(element1).click(element1).perform()

                    time.sleep(t)

                    create = '/html/body/div[1]/div[2]/div[3]/div/div[1]/div[4]/div[2]/div/div[3]/form/div[2]/button[1]'

                    element1 = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, create)))
                    webdriver.ActionChains(driver).move_to_element(element1).click(element1).perform()

                    time.sleep(t)

                    # Go back to results page
                    element1 = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, results)))
                    webdriver.ActionChains(driver).move_to_element(element1).click(element1).perform()

                    time.sleep(t)

                    # Navigate back to the correct page if not on page 0
                    if k > 0:
                        current_page = k + 1  # Convert to 1-based page number
                        print(f"Navigating back to page {current_page}")

                        # Try to click directly on the page number if it's visible
                        try:
                            page_link = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.XPATH, f"//a[text()='{current_page}']"))
                            )
                            webdriver.ActionChains(driver).move_to_element(page_link).click(page_link).perform()
                            time.sleep(t)
                        except:
                            # If direct page link is not available, navigate using Next button
                            print(f"Direct page link not found, using Next button to navigate to page {current_page}")
                            for nav_step in range(current_page - 1):
                                try:
                                    next_button = WebDriverWait(driver, 10).until(
                                        EC.element_to_be_clickable((By.CSS_SELECTOR, "#tableview_next:not(.disabled) a"))
                                    )
                                    webdriver.ActionChains(driver).move_to_element(next_button).click(next_button).perform()
                                    time.sleep(t)
                                except:
                                    print(f"Could not navigate to page {current_page}")
                                    break

                    print(f"Processed item {i+1} on page {k+1}")

                # Navigate to next page after processing all items on current page
                if k < CLICKS - 1:
                    # Check if Next button is not disabled before clicking
                    try:
                        next_button = WebDriverWait(driver, 60).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, "#tableview_next:not(.disabled) a"))
                        )
                        webdriver.ActionChains(driver).move_to_element(next_button).click(next_button).perform()
                        time.sleep(t)
                        print(f"Navigated to page {k+2}")
                    except Exception as e:
                        print(f"Could not navigate to next page: {e}")
                        break

                print(f"Successfully completed page {k+1}")

        except ValueError as e:
            print(f"Error: Please enter valid numbers in all fields. {e}")
        except Exception as e:
            print(f"An error occurred: {e}")

    Button(root,text="ok", command=parameters, bg='brown', fg='white', font=('helvetica', 9, 'bold')).grid(row=5,column=11)

    def on_closing():
        """Handle application cleanup when window is closed"""
        try:
            driver.quit()
        except:
            pass
        root.destroy()

    # Set up proper cleanup when window is closed
    root.protocol("WM_DELETE_WINDOW", on_closing)

    root.mainloop()

def start_application():
    """Start the main application and close startup window"""
    startup_window.destroy()
    main_program()

# Create startup GUI
startup_window = Tk()
startup_window.title("Transformation 3.0")
startup_window.geometry("350x200")
startup_window.resizable(False, False)
startup_window.configure(bg='white')

# Center the window
startup_window.eval('tk::PlaceWindow . center')

# Create main frame
main_frame = Frame(startup_window, bg='white', relief='flat', bd=0)
main_frame.pack(fill=BOTH, expand=True, padx=30, pady=30)

# Title label
title_label = Label(main_frame, text="Transformation 3.0",
                   font=('Segoe UI', 18, 'normal'), bg='white', fg='#2c3e50')
title_label.pack(pady=(0, 10))

# Subtitle label
subtitle_label = Label(main_frame, text="Data Processing Tool",
                      font=('Segoe UI', 10, 'normal'), bg='white', fg='#7f8c8d')
subtitle_label.pack(pady=(0, 25))

# Start button with professional styling
start_button = Button(main_frame, text="Start Processing", command=start_application,
                     bg='#3498db', fg='white', font=('Segoe UI', 11, 'normal'),
                     width=20, height=2, relief='flat', bd=0,
                     activebackground='#2980b9', activeforeground='white',
                     cursor='hand2')
start_button.pack(pady=10)

# Add a subtle border frame around the button
button_frame = Frame(main_frame, bg='#34495e', height=2)
button_frame.pack(fill=X, pady=(5, 0))

# Run startup GUI
startup_window.mainloop()
