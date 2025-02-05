import asyncio
import os
import time
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By

from xl_util import get_row_count, read_data, write_data

# Get the current directory
current_directory = os.path.dirname(os.path.abspath(__file__)) + "\\"
print(f"The current directory is: {current_directory}")

# Set up the Chrome WebDriver
driver = webdriver.Chrome()
time.sleep(2)
driver.maximize_window()

# Declare fluent wait
wait = WebDriverWait(driver, 10, poll_frequency=1, ignored_exceptions=[TimeoutException])
# Navigate to website
driver.get("https://www.saucedemo.com/")


# Get the file path
data_file = "data.xlsx"
file_path = current_directory + data_file
print(f"File path is: {file_path}")

# Get the number of rows
rows = get_row_count(file_path, "Sheet1")


# Function to check if element exists
def check_element_exists(by, value):
    try:
        element = driver.find_element(by, value)
        return True
    except NoSuchElementException:
        return


# Loop through the rows
for r in range(2, rows+1):
    print(f"Row number: {r}")
    source_username = read_data(file_path, "Sheet1", r, 1)
    source_password = read_data(file_path, "Sheet1", r, 2)

    print(f"Username: {source_username}")
    print(f"Password: {source_password}")

    # Find the username field with wait
    username = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@name='user-name']")))

    # Enter credentials
    username.send_keys(source_username)
    driver.find_element(By.XPATH, "//input[@name='password']").send_keys(source_password)
    time.sleep(2)
    driver.find_element(By.XPATH, "//input[@name='login-button']").click()
    time.sleep(2)
    
    
    # Check if an element exists
    element_exists = check_element_exists(By.XPATH, "//button[text()='Open Menu']")
    print(f"Element exists: {element_exists}")
    
    
    if element_exists == True:
        print("PASS")
        write_data(file_path, "Sheet1", r, 3, "PASS")  
        driver.find_element(By.XPATH, "//button[text()='Open Menu']").click()
        time.sleep(2)
        driver.find_element(By.XPATH, "//a[text()='Logout']").click()
        time.sleep(2)
    else:
        print("FAIL")
        write_data(file_path, "Sheet1", r, 3, "FAIL")
        driver.refresh()
    


# Close the browser
driver.quit()