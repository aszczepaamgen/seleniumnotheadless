from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
import chromedriver_autoinstaller
from pyvirtualdisplay import Display
display = Display(visible=0, size=(800, 800))  
display.start()

chromedriver_autoinstaller.install()  # Check if the current version of chromedriver exists
                                      # and if it doesn't exist, download it automatically,
                                      # then add chromedriver to path

chrome_options = webdriver.ChromeOptions()    
# Add your options as needed    
options = [
  # Define window size here
   "--window-size=1200,1200",
    "--ignore-certificate-errors"
 
    "--headless",
    #"--disable-gpu",
    #"--window-size=1920,1200",
    #"--ignore-certificate-errors",
    #"--disable-extensions",
    #"--no-sandbox",
    #"--disable-dev-shm-usage",
    #'--remote-debugging-port=9222'
]

for option in options:
    chrome_options.add_argument(option)

    
# Import necessary libraries
import smartsheet
from selenium.webdriver import Chrome
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.common.exceptions import StaleElementReferenceException
import time

# Retry mechanism to handle stale element reference
def retry_find_element(driver, by, value, retries=3):
    for _ in range(retries):
        try:
            return WebDriverWait(driver, 60).until(EC.presence_of_element_located((by, value)))
        except StaleElementReferenceException:
            continue
        except Exception as e:
            print("An error occurred:", e)
            time.sleep(5)  # Add delay before retrying
    raise Exception("Failed to find element after retries")

# Initialize Smartsheet client with API access token
smartsheet_client = smartsheet.Smartsheet('jCRr1bxrdgYULNJTRsyvsKHq3E3Qn2W5Qy9D5')

# Specify source and target sheet IDs, and column ID
source_sheet_id = 6874500020785028
column_id = 2767095891709828
target_sheet_id = 6086409492320132

# Define website URLs and CSS selectors
base_url = 'https://www.ssllc.com'
search_bar = 'input[type=text]'
items = '.ais-Hits'

# Configure ChromeOptions
options = ChromeOptions()
options.add_argument('-headless')  # Optional: Run Chrome in headless mode for faster execution

try:
    # Delete all existing rows in the target sheet
    sheet = smartsheet_client.Sheets.get_sheet(target_sheet_id)
    rows_to_delete = [row.id for row in sheet.rows]
    if rows_to_delete:
        response = smartsheet_client.Sheets.delete_rows(target_sheet_id, rows_to_delete)
        print("Deleted rows:", rows_to_delete)
    else:
        print("No rows to delete.")
    
    # Load the source sheet
    source_sheet = smartsheet_client.Sheets.get_sheet(source_sheet_id)

    # Retrieve search queries from specified column
    search_queries = []
    for row in source_sheet.rows:
        for cell in row.cells:
            if cell.column_id == column_id and cell.value:
                search_queries.append(cell.value)

    # Initialize Chrome WebDriver outside the loop
    with Chrome(options=options) as driver:
        # Loop through each search query
        for search_query in search_queries:
            # Load the website
            driver.get(base_url)

            # Find and interact with the search input field using retry mechanism
            search_input = retry_find_element(driver, By.CSS_SELECTOR, search_bar)
            search_input.send_keys(search_query)
            search_input.send_keys(Keys.RETURN)
            search_input.clear()

            # Wait for search results to load
            time.sleep(5)  # Add delay to allow time for results to load
            WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.CSS_SELECTOR, items)))

            # Extract and format search results
            search_results = driver.find_elements(By.CSS_SELECTOR, items)

            # Initialize a new row object
            row = smartsheet.models.Row()
            row.to_top = True
            row.cells.append({
                'column_id': 2141929578909572,  # Modify if necessary
                'value': search_query
            })

            if search_results:
                for result in search_results:
                    search_result = result.text.strip()
                    if search_result:
                        row.cells.append({
                            'column_id': 6645529206280068,  # Modify if necessary
                            'value': search_result
                        })
                    else:
                        row.cells.append({
                            'column_id': 6645529206280068,  # Modify if necessary
                            'value': 'No result'
                        })
            else:
                row.cells.append({
                    'column_id': 6645529206280068,  # Modify if necessary
                    'value': 'No result'
                })

            # Add the row to the target sheet
            response = smartsheet_client.Sheets.add_rows(target_sheet_id, [row])
            print("Added row:", response.result)

except smartsheet.exceptions.ApiError as e:
    print("Smartsheet API error:", e)
except Exception as e:
    print("An error occurred:", e)



