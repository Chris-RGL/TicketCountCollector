# Ticket Count Collector  

A Python automation tool for monitoring and tracking tickets in TeamDynamix (TDX) Service Management platform.
Overview
Ticket Count Collector an automated solution designed for IT support teams using the TeamDynamix service management platform. This tool logs into the TDX portal, monitors ticket counts for team members, and exports the data to Excel for analysis and reporting.
Features

Automated Login: Securely logs into the TDX service portal with credentials provided through a user-friendly interface
DUO Authentication Support: Handles two-factor authentication by displaying DUO verification codes
Hourly Ticket Monitoring: Automatically checks ticket counts on an hourly basis
Excel Reporting: Exports ticket count data to Excel with timestamps for trend analysis
Easy Exit: Simple ESC key functionality to stop the monitoring process

## Requirements

Python 3.6+
Chrome browser
Pip packages (see installation)

## Installation

Clone the repository:
git clone https://github.com/yourusername/TicketCountCollector.git
cd TicketCountCollector

Install the required packages:
pip install selenium pandas openpyxl pygame keyboard webdriver-manager

### Set up Chrome WebDriver:

The program will automatically download and install the appropriate Chrome WebDriver using webdriver-manager



Usage

Run the script:
python main.py

Enter your TDX credentials in the popup dialog
Complete the DUO verification when prompted
The program will check ticket counts hourly and save data to ticket_counts.xlsx
Press ESC at any time to exit the program

# How It Works

## Authentication Flow:

Launches a Flask web app to securely collect login credentials
Automates browser login to TDX
Handles DUO two-factor authentication


## Ticket Monitoring:

Navigates to the TDX reporting page
Extracts ticket count data for each team member
Appends data to Excel with timestamps


## Scheduling:

Runs initial check immediately upon startup
Schedules subsequent checks at the top of each hour
Continuously monitors for ESC keypress to exit



## File Structure

main.py - Main script containing the automation logic
App.py - Flask application for login form and browser integration
ticket_counts.xlsx - Output file containing ticket count data

## Troubleshooting

Login Issues: Ensure your TDX credentials are correct
DUO Timeout: Make sure to approve DUO notifications promptly
Chrome Driver Errors: Try upgrading Chrome to the latest version
Excel Access Denied: Make sure the Excel file is not open in another program

For University of Oregon Users
This script is configured for the UO TDX instance by default. If you're using a different TDX instance, update the URL in the login_and_verify function:
pythondriver.get('https://service.uoregon.edu/TDNext/Home/Desktop/Desktop.aspx')
# License
MIT License
# Credits
Created by Christopher Gallinger-Long
