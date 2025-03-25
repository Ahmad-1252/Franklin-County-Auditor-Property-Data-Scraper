from datetime import datetime, timedelta
from matplotlib.dates import relativedelta
import numpy as np
import os 
from urllib.parse import urlencode
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException
from functools import wraps
from selenium.common.exceptions import StaleElementReferenceException


def retries(max_retries=3, delay=2, exceptions=(Exception,)):
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            attempts = 0
            while attempts < max_retries:
                try:
                    return func(*args, **kwargs)
                except exceptions as e:
                    attempts += 1
                    print(f"Function '{func.__name__}' crashed on attempt {attempts}/{max_retries}: {e}")
                    if attempts < max_retries:
                        print(f"Retrying function '{func.__name__}'...")
                        time.sleep(delay)
                    else:
                        print(f"Function '{func.__name__}' failed after {max_retries} retries.")
                        raise
        return wrapper
    return decorator

def get_url(start_date, end_date):
    base_url = "https://franklin.oh.publicsearch.us/results"
    params = {
        "department": "RP",
        "limit": 250,
        "offset": 50,
        "recordedDateRange": f"{start_date},{end_date}",
        "searchOcrText": "false",
        "searchType": "quickSearch",
        "searchValue": "ENVIRONMENTAL DIVISION"
    }

    # Construct the URL
    dynamic_url = f"{base_url}?{urlencode(params)}"
    return dynamic_url

@retries(max_retries=5, delay=1, exceptions=(ElementClickInterceptedException, TimeoutException))
def get_chromedriver(headless=False):
    current_dir = os.getcwd()  # Get current working directory for downloads
    chrome_options = Options()
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": current_dir,  # Set the download folder
        "download.prompt_for_download": False,  # Don't prompt for download
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    chrome_options.add_argument("--disable-logging")
    chrome_options.add_argument("--start-maximized")
    if headless:
        chrome_options.add_argument("--headless")

    driver = webdriver.Chrome(options=chrome_options)
    pid = driver.service.process.pid
    print(f"Chrome WebDriver Process ID: {pid}")
    return driver, pid

def wait_for_element(driver, Xpath, timeout=10):
    try:
        element = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, Xpath)))
        return element
    except TimeoutException:
        print(f"Element with ID '{Xpath}' not found after {timeout} seconds.")
        return None


def get_table_data(driver):
    """Fetch table data with improved structure and error handling."""
    try:
        # Wait for table rows to be present
        table_rows = WebDriverWait(driver, 300).until(
            EC.presence_of_all_elements_located((By.XPATH, '//div[@data-tourid="searchResults"]//table/tbody/tr'))
        )
        print("table loaded") 
    except Exception as e:
        print(f"Error: Table not loaded - {e}")
        return []

    table_data = []

    for index in range(len(table_rows)):
        try:
            # Wait and click on the specific row
            row_element = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable(
                    (By.XPATH, f'(//div[@data-tourid="searchResults"]//table/tbody/tr)[{index+1}]')
                )
            )
            row_element.click()
            print(f"Clicked row {index + 1}")
            time.sleep(3)

            # Fetch pin elements from the new table
            try:
                pins = WebDriverWait(driver, 30).until(
                    EC.presence_of_all_elements_located((By.XPATH, '//table[@class="css-1uz5dol"]/tbody/tr/td[7]'))
                )
                if pins is not None:
                    pin_texts = [pins_element.text for pins_element in pins if pins_element]
                    print(pin_texts)
                    table_data.append(pin_texts)
                    print(f"Fetched pins for row {index + 1}")
            except Exception as e:
                print(f"Error: Failed to fetch pin elements for row {index + 1}")
                

            # Click the back button
            try:
                back_btn = WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, '//button[@class="css-1ihxvt8"]'))
                )
                back_btn.click()
                print(f"Clicked back button after row {index + 1}")
                time.sleep(1)
            except Exception as e:
                print(f"Failed to click back button ")
                break

        except StaleElementReferenceException as e:
            print(f"StaleElementReferenceException: element not found ")
            time.sleep(2)
        except Exception as e:
            print(f" Pin Id not found for row {index + 1}")
            continue

    return table_data


@retries(max_retries=5, delay=1, exceptions=(ElementClickInterceptedException, TimeoutException))
def search_and_get_case_data(driver,ParcelId):
    try:
                # Initialize case data dictionary
        case_data = {}

        if ParcelId == '' or ParcelId == 'N/A' or ParcelId is None:
            return {}
        print("Opened the browser")

        # Fill out search form
        def fill_input(xpath, value, field_name):
            try:
                input_field = WebDriverWait(driver, 60).until(
                    EC.presence_of_element_located((By.XPATH, xpath))
                )
                input_field.send_keys(value)
                print(f"{field_name} input sent: {value}")
            except TimeoutException:
                print(f"{field_name} input field not found.")

        fill_input('//input[@id="inpParid"]', ParcelId, "Parcel Id")
        # Click the search button
        try:
            search_btn = WebDriverWait(driver, 60).until(
                EC.element_to_be_clickable((By.XPATH, '//button[@id="btSearch"]'))
            )
            search_btn.click()
            print("Clicked the Search Button")
        except TimeoutException:
            print("Search button not found.")
            return 

        time.sleep(3)

        # Check for "No Records Found" error
        try:
            WebDriverWait(driver, 11).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//large[contains(text(), "Your search did not find any records")]')
                )
            )
            print("No records found for the search.")
            case_data['parcel_id'] = ParcelId
            return case_data
        except TimeoutException:
            print("Records found. Continuing...")

        try:
            record_table_btn = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, '(//table[@id="searchResults"]/tbody/tr)[1]'))
            )
            record_table_btn.click()
            print("Clicked the First Row")
        except TimeoutException:
            print("Search Results timed out")

        # Helper function to extract data with XPath
        def extract_data(xpath, key, description=None):
            try:
                element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath)))
                if key == 'parcel_id':
                    case_data[key] = element.text.split(':')[1].strip()
                else:
                    case_data[key] = element.text.strip()
                if description:
                    print(f"{description}: {element.text.strip()}")
            except TimeoutException:
                case_data[key] = ""
                if description:
                    print(f"{description} not found.")

        # Extract main case data
        case_data['parcel_id'] = ParcelId
        extract_data('//tr[td[contains(text(), "Site (Property) Address")]]/td[@class="DataletData"]',
                     'property_address', "Property Address")
        extract_data('//tr[td[contains(text(), "Zip Code")]]/td[@class="DataletData"]',
                     'property_zip_code', "property_zip_code")
        
        try:
            # Wait until the "Legal Description" row is present
            legal_description_row = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//tr[td[contains(text(), "Legal Description")]]')
                )
            )

            # Extract the next two siblings after "Legal Description"
            next_rows = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located(
                    (By.XPATH, '//tr[td[contains(text(), "Legal Description")]]/following-sibling::tr[position() <= 2]')
                )
            )

            # Collect the text of the second td of each row (including "Legal Description")
            concatenated_text = legal_description_row.find_element(By.XPATH, './td[@class="DataletData"]').text.strip() + "\n"
            for row in next_rows:
                td_elements = row.find_elements(By.TAG_NAME, "td")
                if len(td_elements) > 1:  # Ensure there are at least two td elements
                    concatenated_text += td_elements[1].text.strip() + "\n"

            # Add the concatenated description to case_data
            case_data['description'] = concatenated_text.strip()  # Remove trailing newline
        except Exception as e:
            print("Description not found:")
            case_data['description'] = ""

        try:
            # Locate all `<a>` elements containing owner names
            owner_elements = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, '//tr[td[contains(text(), "Owner")]]/td[@class="DataletData"]/a'))
            )

            # Handle cases where there are multiple owners
            if len(owner_elements) == 0:
                case_data['owner_names'] = []
                case_data['owner_names_string'] = ''

            # Extract the text content from each `<a>` element
            owner_names = [owner.text.strip() for owner in owner_elements if owner.text.strip()]

            # Store the owner names as a list and a comma-separated string
            case_data['owner_names'] = owner_names
            case_data['owner_names_string'] = ', '.join(owner_names)

            print(f"Owner Names: {owner_names}")

        except TimeoutException:
            # Handle missing owners by initializing an empty list and string
            case_data['owner_names'] = []
            case_data['owner_names_string'] = ''
            print("Owner Names not found.")


        # Append address info
        case_data.update({
            'property_city': 'columbus',
            'property_state': 'OH',
        })

        # Extract additional fields
        extract_data('//tr[td[contains(text(), "Owner Mailing /")]]/td[@class="DataletData"]',
                     'mailing_address', "Mailing Address")
        extract_data('//tr[td[contains(text(), "Contact Address")]]/td[@class="DataletData"]',
                     'contact_address', "Contact Address")

        # Dwelling data
        dwelling_data = {
            '(//table[@id="Dwelling Data"]//td)[10]': 'bedrooms',
            '(//table[@id="Dwelling Data"]//td)[11]': 'bathrooms',
            '(//table[@id="Dwelling Data"]//td)[8]': 'Tot Fin Area',
            '(//table[@id="Dwelling Data"]//td)[7]': 'Year built'
        }
        for xpath, key in dwelling_data.items():
            extract_data(xpath, key, key)

        # Transfer details
        extract_data('//tr[td[contains(text(), "Transfer Date")]]/td[@class="DataletData"]',
                     'Transfer Date', "Transfer Date")
        extract_data('//tr[td[contains(text(), "Transfer Price")]]/td[@class="DataletData"]',
                     'Transfer Price', "Transfer Price")
        extract_data('//tr[td[contains(text(), "Property Class")]]/td[@class="DataletData"]',
                     'Property Class', "Property Class")

        # Click "Rental Contact" button
        try:
            rental_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//a[span[contains(text(), "Rental Contact")]]'))
            )
            rental_btn.click()
            print("Navigating to Rental Contact page...")
            time.sleep(3)
        except TimeoutException:
            print("Rental Contact button not found.")

        # Extract rental contact details
        rental_headers = [
            "Owner Name:", "Owner Business:", "Title:", "Address1:", "Address2:",
            "City:", "State:", "Zip Code:", "Phone Number:", "E-Mail Address:"
        ]
        for header in rental_headers:
            header_key = header.replace(":", "").replace(" ", "_").lower()
            if header in ["City:", "State:"]:
                header_key = f"rental_{header_key}"
            extract_data(f'//tr[td[contains(text(), "{header}")]]/td[@class="DataletData"]',
                        header_key, header)

        return case_data

    except Exception as e:
        print(f"An error occurred while searching for the address: {e}")
        raise

def process_owner_data(all_data, split_full_name, processed_data):
    for item in all_data:
        if item is None:
            continue
        
        # Get the list of owner names
        owner_names = item.get('owner_names', [])
        
        for i in range(len(owner_names)):
            # Ensure we process only valid owner names
            if owner_names[i].strip():
                full_owner_name = owner_names[i].strip()
                print(f'Processing owner: {full_owner_name}')
                
                # Split the full owner name into first and last names
                name_parts = split_full_name(owner_names[i])
                first_name = name_parts['first_name']
                last_name = name_parts['last_name']

                # Split mailing_address into mailing_city, mailing_state, and mailing_zip
                contact_address = item.get('contact_address', '')
                mailing_parts = contact_address.split(" ")
                mailing_city = mailing_parts[0] if len(mailing_parts) > 0 else ""
                mailing_state = mailing_parts[1] if len(mailing_parts) > 1 else ""
                mailing_zip = mailing_parts[2] if len(mailing_parts) > 2 else ""

                # Append processed data for this owner
                processed_data.append({
                    "EVH No": item.get('EVH No', ''),
                    "parcel": item.get('parcel_id', ''),
                    "full_name": owner_names[i],
                    "first_name": first_name,
                    "last_name": last_name,
                    "property_address": item.get('property_address', ''),
                    "property_city": item.get('property_city', ''),
                    "property_state": item.get('property_state', ''),
                    "property_zip_code": item.get('property_zip_code', ''),
                    "description": item.get('description', ''),
                    "mailing_address": item.get('mailing_address', ''),
                    "mailing_city": mailing_city,
                    "mailing_state": mailing_state,
                    "mailing_zip": mailing_zip,
                    "owner_name": item.get('owner_name', ''),
                    "owner_business": item.get('owner_business', ''),
                    "title": item.get('title', ''),
                    "address_1": item.get('address1', ''),
                    "address_2": item.get('address2', ''),
                    "rental_city": item.get('rental_city', ''),
                    "rental_state": item.get('rental_state', ''),
                    "rental_zipcode": item.get('zip_code', ''),
                    "phone": item.get('phone_number', ''),
                    "email": item.get('e-mail_address', ''),
                    "bedroom": item.get('bedrooms', ''),
                    "bathroom": item.get('bathrooms', ''),
                    "Tot Fin Area": item.get('Tot Fin Area', ''),
                    "year built": item.get('Year built', ''),
                    "Property Class": item.get('Property Class', ''),  # Assuming no Property Class in data
                    "Transfer Date": item.get('Transfer Date', ''),
                    "Transfer Price": item.get('Transfer Price', '')
                })

        if owner_names == []:
            # Handle case where owner name is missing or blank
            print(f"No valid owner name found in this entry, processing other data.")
            # Append data with missing owner information
            processed_data.append({
                "EVH No": item.get('EVH #', ''),
                "parcel": item.get('parcel_id', ''),
                "full_name": '',
                "first_name": '',
                "last_name": '',
                "property_address": item.get('property_address', ''),
                "property_city": item.get('property_city', ''),
                "property_state": item.get('property_state', ''),
                "property_zip_code": item.get('property_zip_code', ''),
                "description": item.get('description', ''),
                "mailing_address": item.get('mailing_address', ''),
                "mailing_city": '',
                "mailing_state": '',
                "mailing_zip": '',
                "owner_name": item.get('owner_name', ''),
                "owner_business": item.get('owner_business', ''),
                "title": item.get('title', ''),
                "address_1": item.get('address1', ''),
                "address_2": item.get('address2', ''),
                "rental_city": item.get('rental_city', ''),
                "rental_state": item.get('rental_state', ''),
                "rental_zipcode": item.get('zip_code', ''),
                "phone": item.get('phone_number', ''),
                "email": item.get('e-mail_address', ''),
                "bedroom": item.get('bedrooms', ''),
                "bathroom": item.get('bathrooms', ''),
                "Tot Fin Area": item.get('Tot Fin Area', ''),
                "year built": item.get('Year built', ''),
                "Property Class": item.get('Property Class', ''),
                "Transfer Date": item.get('Transfer Date', ''),
                "Transfer Price": item.get('Transfer Price', '')
            })


def split_full_name(full_name):
    # Check if the name is likely an organization or business
    if any(keyword in full_name.upper() for keyword in ["LLC", "INC", "CORP", "COMPANY", "INVESTMENTS", "ENTERPRISES"]):
        # Treat the entire name as the last name (organization name)
        return {"first_name": "", "last_name": full_name.strip()}
    
    # Common name prefixes and suffixes to exclude from splitting
    prefixes = ["Dr.", "Mr.", "Ms.", "Mrs.", "Miss", "Prof."]
    suffixes = ["Jr.", "Sr.", "II", "III", "IV", "Ph.D.", "M.D.", "Esq."]

    # Clean and normalize the name
    full_name = full_name.strip()

    # Remove prefixes and suffixes using regex
    for prefix in prefixes:
        if full_name.startswith(prefix):
            full_name = full_name[len(prefix):].strip()

    for suffix in suffixes:
        if full_name.endswith(suffix):
            full_name = full_name[: -len(suffix)].strip()

    # Split name into parts
    name_parts = full_name.split()

    # Handle single-word names
    if len(name_parts) == 1:
        return {"first_name": name_parts[0], "last_name": ""}

    # Handle multi-word names (assign everything but the first part to last name)
    first_name = name_parts[0]
    last_name = " ".join(name_parts[1:])

    return {"first_name": first_name, "last_name": last_name}


def extract_all_pin_ids(driver):
    try:
        all_pin_Ids = []
        # Wait until the table loads and iterate through pages
        while True:
            try:
                try:
                    WebDriverWait(driver, 30).until(
                        EC.presence_of_element_located((By.XPATH, '//h3[text() =" No Results Found "]'))
                    )
                    print('Record not found')
                    break
                except TimeoutException:
                    print("Records found")

                try:
                    total_text = WebDriverWait(driver, 300).until(
                        EC.presence_of_element_located((By.XPATH, '//span[@aria-label="Search Result Totals"]'))
                    ).text.strip()
                    print(f"Total Records Text: {total_text}")

                    # Extract the total number of records from the text
                    total_records = total_text.split("of")[-1].split()[0]
                    total_records = int(total_records.replace(',' , ""))
                    print(f"Total Records: {total_records}")
                except TimeoutException:
                    print("Total records element not found. Exiting.")

                # Extract data from the current page
                page_data = get_table_data(driver)
                all_pin_Ids.extend(pin for sublist in page_data for pin in sublist)
                try:
                    # Locate the 'Next' button
                    NextBtn = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located(
                            (By.XPATH, '//button[@aria-label="next page"]')
                        )
                    )

                    print("The next page btn found")
                    
                    # Check if the 'Next' button is disabled
                    is_disabled = (
                        NextBtn.get_attribute("disabled") is not None or 
                        NextBtn.get_attribute("aria-disabled") == "true"
                    )
                    if is_disabled:
                        print("Next button is disabled. Breaking the loop.")
                        break
                    
                    # Click the 'Next' button if it's enabled
                    print("Clicking the 'Next' button.")
                    NextBtn.click()
                    time.sleep(1)  # To ensure the page loads after the click
                except: 
                    print(f"An error occurred while clicking the 'Next' button: {e}")
                    break
            except TimeoutException:
                print("Next button not found or timeout occurred. Exiting the loop.")
                break
            except Exception as e:
                print(f"An error occurred while processing the page: {e}")
                break

        return all_pin_Ids
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return []
    

# Helper function to generate a list of months between start_date and end_date
def generate_month_ranges(start_date, end_date):
    start = datetime.strptime(start_date, "%Y%m%d")
    end = datetime.strptime(end_date, "%Y%m%d")
    month_ranges = []

    while start <= end:
        month_start = start
        month_end = (start + relativedelta(months=1)) - timedelta(days=1)
        if month_end > end:
            month_end = end
        month_ranges.append((month_start.strftime("%Y%m%d"), month_end.strftime("%Y%m%d")))
        start += relativedelta(months=1)
    
    return month_ranges


if __name__ == "__main__":
    start_date = input("Enter the start date YYYYMMDD : \t")  # January 1, 2023
    end_date = input("Enter the End date YYYYMMDD : \t")  # January 1, 2023

    # Generate month ranges
    month_ranges = generate_month_ranges(start_date, end_date)
    print(f"Generated Month Ranges: {month_ranges}")

    # Initialize a list to store all data
    all_data = []

    # Loop through each month range
    for month_start, month_end in month_ranges:
        print(f"Processing data from {month_start} to {month_end}")
        try:
            # Reinitialize WebDriver for each range
            driver, pid = get_chromedriver(headless=True)
            print(f"Chrome WebDriver Process ID: {pid}")

            # Get the dynamic URL for the date range
            url = get_url(month_start, month_end)
            print(f"URL: {url}")
            
            # Fetch data
            driver.get(url)
            print("Getting the table")
            all_pin_ids = extract_all_pin_ids(driver)
            
            # Store the data for this month
            all_data.extend(all_pin_ids)
        except Exception as e:
            # Log the error and continue
            print(f"Error processing data from {month_start} to {month_end}: {e}")
        finally:
            # Ensure the WebDriver is closed after each range
            try:
                driver.quit()
                print("WebDriver closed")
            except Exception as close_error:
                print(f"Error closing WebDriver: {close_error}")

    # Save all data to a single file
    df = pd.DataFrame(all_data, columns=["Pin IDs"])
    df.drop_duplicates(inplace=True)
    df = df[df['Pin IDs'] != 'N/A']

    output_file = "ParcelIDFile_Complete.csv"
    df.to_csv(output_file, index=False)
    print(f"All data saved to {output_file}")

    data = pd.read_csv('ParcelIDFile_Complete.csv')
    data.fillna('N/A' , inplace=True)
    data.drop_duplicates(inplace=True)
    data = data[data['Pin IDs'] != 'N/A']

    # getting the chrome driver
    driver , pid = get_chromedriver(headless=True)
    
    all_data = []
    for index, row in data.iterrows():
        print(f"Processing record: {row['Pin IDs']}")
        driver.get("https://property.franklincountyauditor.com/_web/search/commonsearch.aspx?mode=parid")
        case_data = search_and_get_case_data(driver, row['Pin IDs'] )
        print('case_data , ', case_data)
        all_data.append(case_data)
    
    driver.quit()

    columns = [
        "evh #", "parcel" , "full name", "first_name", "last_name", "property_address",
        "property_city", "property_state", "property_zip_code", "description",
        "mailing_address", "mailing_city", "mailing_state", "mailing_zip",
        "owner_name", "owner_business", "title", "address_1", "address_2",
        "rental_city", "rental_state", "rental_zipcode", "phone", "email",
        "bedroom", "bathroom", "Tot Fin Area", "year built", "Property Class",
        "Transfer Date", "Transfer Price"
    ]

    processed_data = []  # List to hold processed data
    process_owner_data(all_data, split_full_name, processed_data)

    # # Create a DataFrame
    df = pd.DataFrame(processed_data, columns=columns)
    output_file = "Output.xlsx"

    renamed_file = 'Previous_output.xlsx'
    # Check if the renamed file exists
    if os.path.exists(renamed_file):
        # Delete the renamed file
        os.remove(renamed_file)
        print(f"File {renamed_file} has been deleted")

    # Save to Excel
    if os.path.exists(output_file):
        # Rename the file
        os.rename(output_file, renamed_file)
        if os.path.exists('ProcessedIDs.csv'):
            os.remove('ProcessedIDs.csv')
            print(f"File ProcessedIDs.csv has been deleted")
        os.rename('ParcelIDFile_Complete.csv', 'ProcessedIDs.csv')
        print(f"File renamed to {renamed_file}")

    df.to_excel(output_file, index=False)
    print(f"Data saved to {output_file}")
