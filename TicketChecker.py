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
    """
    Retrieve login credentials from the popup application.
    
    Returns:
        list: A list containing the username and password.
    """
    login_info = App.login_info
    return login_info


def login_and_verify(driver, login):
    """
    Log in to TDX service portal and complete DUO verification.
    
    Args:
        driver: Selenium WebDriver instance.
        login: List containing username and password credentials.
    """
    driver.get('https://service.uoregon.edu/TDNext/Home/Desktop/Desktop.aspx')

    try:
        # Enter username
        username_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'username'))
        )
        username_field.send_keys(login[0])

        # Enter password
        password_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'password'))
        )
        password_field.send_keys(login[1])

        # Click login button
        login_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//input[@type="submit"]'))
        )
        login_button.click()
        
        # Wait for DUO verification code and display it
        verification_code = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CLASS_NAME, 'verification-code'))
        )
        verification_code_element = driver.find_element(By.CLASS_NAME, 'verification-code')
        verification_code = verification_code_element.text

        print(f'DUO Verification Code: {verification_code}')

        try:
            # Handle "trust this browser" option if it appears
            time.sleep(10)
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.ID, 'trust-browser-button'))
            )
            yes_device_button = driver.find_element(By.ID, 'trust-browser-button')
            yes_device_button.click()
        except TimeoutException:
            print("Device trust prompt not found, proceeding without it.")
        
    except Exception as e:
        print(f"An error occurred during login: {e}")


def check_for_ticket(driver):
    """
    Check if there are any new tickets in the system.
    
    Args:
        driver: Selenium WebDriver instance.
    """
    try:
        # Switch to iframe if present
        iframes = driver.find_elements(By.TAG_NAME, 'iframe')
        if iframes:
            driver.switch_to.frame(iframes[0])

        time.sleep(5)

        try:
            # Look for "New" status in ticket table
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
        # Return to main content
        driver.switch_to.default_content()


def check_ticket_counts(driver):
    """
    Retrieve ticket counts for each person from TDX report.
    
    Args:
        driver: Selenium WebDriver instance.
    """
    try:
        # Step 1: Switch to iframe if it exists
        iframes = driver.find_elements(By.TAG_NAME, 'iframe')
        if iframes:
            driver.switch_to.frame(iframes[0])

        # Step 2: Navigate directly to the ticket count report
        driver.execute_script("""
            const reportUrl = "/TDNext/Apps/430/Reporting/ReportViewer?ReportID=102376";
            window.location.href = reportUrl;
        """)

        # Step 3: Wait for navigation to complete
        time.sleep(10)

        # Step 4: Extract ticket count data
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

        # Export data if available
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
        # Return to main content
        driver.switch_to.default_content()


def export_to_excel(ticket_data, excel_file="ticket_counts.xlsx", append=True):
    """
    Export ticket count data to an Excel file.
    
    Args:
        ticket_data: List of dictionaries with 'name' and 'count' keys.
        excel_file: Filename for the Excel output file (default: ticket_counts.xlsx).
        append: If True, append to existing file; if False, create new file (default: True).
        
    Returns:
        bool: True if export was successful, False otherwise.
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
    """
    Main function that initializes the WebDriver, handles login,
    and runs the ticket checking process on an hourly schedule.
    """
    # Configure Chrome WebDriver
    options = webdriver.ChromeOptions()
    options.add_argument('--no-sandbox')
    #options.add_argument('--headless')
    options.add_argument('--disable-gpu')

    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

    # Wait for login information from the app
    while not App.login_info:
        time.sleep(1)
    login = get_login()

    try:
        # Perform login and verification
        login_and_verify(driver, login)
        time.sleep(10)
        print("Press ESC to exit")
        
        while True:
            if keyboard.is_pressed('esc'):
                print("ESC key pressed. Stopping the ticket checker.")
                break
                
            # Get current time
            now = datetime.now()
            
            # Run the check immediately
            check_ticket_counts(driver)
            print(f"Checked tickets at {now.strftime('%Y-%m-%d %H:%M:%S')}")
            
            # Calculate time until next hour
            next_hour = now.replace(minute=0, second=0, microsecond=0) + timedelta(hours=1)
            wait_seconds = (next_hour - datetime.now()).total_seconds()
            
            # Ensure wait time is not negative
            if wait_seconds < 0:
                wait_seconds = 0
                
            print(f"Waiting {wait_seconds:.1f} seconds until next check at {next_hour.strftime('%H:00:00')}")
            
            # Check periodically for ESC keypress while waiting
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
    
    # Open the web browser for login dialog
    App.open_browser()

    # Run the main function
    main()