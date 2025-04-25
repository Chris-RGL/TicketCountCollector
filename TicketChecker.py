#Credits: Christopher Gallinger-Long
import os
import App
import time
import pygame
import keyboard
from threading import Thread
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
from datetime import datetime, timedelta



def get_login():
    login_info = App.login_info

    return login_info

def login_and_verify(driver, login):
    driver.get('https://service.uoregon.edu/TDNext/Home/Desktop/Desktop.aspx')

    try:
        # Detect username field by HTML ID and enter login info
        username_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'username'))
        )
        username_field.send_keys(login[0])

        # Detect password field by HTML ID and enter login info
        password_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'password'))
        )
        password_field.send_keys(login[1])

        # Detect login button and press it
        login_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//input[@type="submit"]'))
        )
        login_button.click()
        
        # Wait for the page to load and DUO verification to be completed
        verification_code = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CLASS_NAME, 'verification-code'))
        )
        verification_code_element = driver.find_element(By.CLASS_NAME, 'verification-code')
        verification_code = verification_code_element.text

        print(f'DUO Verification Code: {verification_code}')

        try:
            # Wait for the specific element in the content you want to scrape
            time.sleep(10)
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.ID, 'trust-browser-button'))
            )
            yes_device_button = driver.find_element(By.ID, 'trust-browser-button')
            yes_device_button.click()
        except TimeoutException:
            print("Device check not found, proceeding without it.")
        
    except Exception as e:
        print(f"An error occurred during login: {e}")
    


def check_for_ticket(driver):
    try:
        # Check if content is within an iframe switch if true
        iframes = driver.find_elements(By.TAG_NAME, 'iframe')
        if iframes:
            driver.switch_to.frame(iframes[0])

        time.sleep(5)

        try:
            # Wait for the specific element in the content you want to scrape
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//td[text()="New"]'))
            )
            print("New ticket found!")

        except TimeoutException:
            print("No new ticket found.")

    except NoSuchElementException as e:
        print(f"The table element was not found: {e}")
    except TimeoutException as e:
        print(f"Timed out waiting for element: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Ensure we switch back to the default content
        driver.switch_to.default_content()

def check_ticket_counts(driver):
    try:
        # Step 1: Switch to iframe if it exists
        iframes = driver.find_elements(By.TAG_NAME, 'iframe')
        if iframes:
            driver.switch_to.frame(iframes[0])

        # Step 2: Navigate directly to the report page using window.location.href
        driver.execute_script("""
            const reportUrl = "/TDNext/Apps/430/Reporting/ReportViewer?ReportID=102376";
            window.location.href = reportUrl;
        """)

        # Step 3: Wait for navigation to complete
        time.sleep(10)  # Use WebDriverWait for production use

        # Step 4: Extract data after the new page has loaded
        ticket_data = driver.execute_script("""
            const rows = document.querySelectorAll('tr');
            const ticketData = [];

            rows.forEach(row => {
                const cells = row.querySelectorAll('td');
                if (cells.length >= 2) {
                    const countCell = cells[0].textContent.trim();
                    const nameCell = cells[1].textContent.trim();
                    if (!isNaN(countCell) && nameCell.length > 0) {
                        ticketData.push({count: countCell, name: nameCell});
                    }
                }
            });

            return ticketData;
        """)
        """
        # Step 5: Display results in Python
        if ticket_data:
            print("Extracted ticket data:")
            for item in ticket_data:
                print(f"Name: {item['name']}, Ticket Count: {item['count']}")
        else:
            print("No ticket data extracted.")
        """
        # Inside your check_ticket_counts function, replace the "Step 5" section with:
        if ticket_data:
            export_to_excel(ticket_data)
        else:
            print("No ticket data to export.")

    except NoSuchElementException as e:
        print(f"The table element was not found: {e}")
    except TimeoutException as e:
        print(f"Timed out waiting for element: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Always return to the top-level browsing context
        driver.switch_to.default_content()

def export_to_excel(ticket_data, excel_file="ticket_counts.xlsx", append=True):
    """
    Export ticket data to an Excel file.
    
    Parameters:
    - ticket_data: List of dictionaries with 'name' and 'count' keys
    - excel_file: Filename for the Excel file (default: ticket_counts.xlsx)
    - append: If True, append to existing file; if False, create new file (default: True)
    
    Returns:
    - True if export was successful, False otherwise
    """
    try:
        
        # Convert to pandas DataFrame
        data = [{"Name": item['name'], "Ticket Count": item['count']} for item in ticket_data]
        df = pd.DataFrame(data)
        
        # Add timestamp column
        df['Timestamp'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Check if file exists and append or create new
        if append:
            try:
                # Try to read existing file
                existing_df = pd.read_excel(excel_file)
                combined_df = pd.concat([existing_df, df], ignore_index=True)
                combined_df.to_excel(excel_file, index=False)
            except FileNotFoundError:
                # If file doesn't exist, create new
                df.to_excel(excel_file, index=False)
        else:
            # Always create new file
            df.to_excel(excel_file, index=False)
        
        print(f"Ticket data exported to {excel_file}")
        return True
    except Exception as e:
        print(f"Error exporting to Excel: {e}")
        return False

def main():
    options = webdriver.ChromeOptions()
    options.add_argument('--no-sandbox')
    #options.add_argument('--headless')
    options.add_argument('--disable-gpu')

    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

    while not App.login_info:
        time.sleep(1)
    login = get_login()

    try:
        login_and_verify(driver, login)
        time.sleep(10)
        print("Press ESC to exit")
        
        while True:
            if keyboard.is_pressed('esc'):
                print("ESC key pressed. Stopping the ticket checker.")
                break
                
            # Get current time
            now = datetime.now()
            
            # Run the check immediately the first time
            check_ticket_counts(driver)
            print(f"Checked tickets at {now.strftime('%Y-%m-%d %H:%M:%S')}")
            
            # Calculate time until next hour
            next_hour = now.replace(minute=0, second=0, microsecond=0) + timedelta(hours=1)
            wait_seconds = (next_hour - datetime.now()).total_seconds()
            
            # Don't wait if it's negative (can happen if calculations took time)
            if wait_seconds < 0:
                wait_seconds = 0
                
            print(f"Waiting {wait_seconds:.1f} seconds until next check at {next_hour.strftime('%H:00:00')}")
            
            # Check every 5 seconds if ESC is pressed while waiting
            wait_start = time.time()
            while time.time() - wait_start < wait_seconds:
                if keyboard.is_pressed('esc'):
                    print("ESC key pressed. Stopping the ticket checker.")
                    return
                time.sleep(5)  # Check every 5 seconds
                
    except KeyboardInterrupt:
        print("Stopping the ticket checker.")
    finally:
        driver.quit()


if __name__ == '__main__':
    # Start the Flask app in a separate thread
    flask_thread = Thread(target=App.run_flask)
    flask_thread.start()

    # Give Flask a moment to start up
    time.sleep(1)
    # Open the web browser
    App.open_browser()

    # Run the main function
    main()