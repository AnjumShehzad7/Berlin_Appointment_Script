from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import openpyxl
from selenium.webdriver.chrome.service import Service

# Configuration
url = 'https://service.berlin.de/terminvereinbarung/termin/day/1717106400/'
date_to_check = '17'  # Enter the Date for appointment
month_to_check = 'May 2024'   # Enter the Month for appointment
check_interval = 5  # Time interval in seconds (30 seconds)
chrome_driver_path = 'C:\\Users\\anjum\\PycharmProjects\\berlinAppointment\\chromedriver-win64\\chromedriver.exe'  # Path to your ChromeDriver executable

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')

# Excel file configuration
excel_file = 'available_dates.xlsx'


# Initialize Excel file
def initialize_excel():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Available Dates"
    ws.append(["Date", "Time Checked"])
    wb.save(excel_file)


# Log available dates in Excel
def log_available_date(date):
    wb = openpyxl.load_workbook(excel_file)
    ws = wb.active
    ws.append([date, time.strftime('%Y-%m-%d %H:%M:%S')])
    wb.save(excel_file)


def check_appointment(driver):
    # Open the webpage
    driver.get(url)

    # Wait for the calendar to load (adjust the wait time if necessary)
    wait = WebDriverWait(driver, 10)
    try:
        calendar_loaded = wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'calendar')))

        # Check for the correct month
        month_displayed = driver.find_element(By.CLASS_NAME, 'calendar').find_element(By.CLASS_NAME, 'month').text
        if month_to_check in month_displayed:
            # Find all available dates
            available_dates = driver.find_elements(By.CLASS_NAME, 'available')

            # Check if the desired date is available
            date_found = False
            for date in available_dates:
                available_date = date.text
                if available_date == date_to_check:
                    date_found = True
                    message = f"Appointment available on {date_to_check} {month_to_check}"
                    print(message)
                    # Log available date
                    log_available_date(f"{available_date} {month_to_check}")
                    break
                else:
                    log_available_date(f"{available_date} {month_to_check}")

            if not date_found:
                print(f"No appointment available on {date_to_check} {month_to_check}")
        else:
            print(f"The calendar is not displaying {month_to_check}. Please navigate to the correct month.")
    except Exception as e:
        print(f"Error loading calendar: {e}. Retrying...")


# Initialize the Excel file
initialize_excel()

# Initialize WebDriver
service = Service(executable_path=chrome_driver_path)
driver = webdriver.Chrome(service=service, options=chrome_options)

# Repeatedly check for the appointment
while True:
    check_appointment(driver)
    print(f"Checked appointment at {time.strftime('%Y-%m-%d %H:%M:%S')}. Waiting for the next check...")
    for i in range(check_interval, 0, -1):
        print(f"Next check in {i} seconds", end='\r')
        time.sleep(1)
    driver.refresh()
