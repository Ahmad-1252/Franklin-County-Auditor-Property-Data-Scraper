# Franklin County Auditor Property Data Scraper

This Python script scrapes property data from the Franklin County Auditor's website (Ohio) for parcels associated with the "ENVIRONMENTAL DIVISION." It extracts parcel IDs within a specified date range and then fetches detailed property information for each parcel.

## Features

- Scrapes parcel IDs from search results filtered by date range
- Extracts detailed property information including:
  - Owner names and contact details
  - Property characteristics (bedrooms, bathrooms, square footage)
  - Transfer history
  - Rental information
- Handles pagination automatically
- Includes robust error handling and retry mechanisms
- Outputs data to CSV and Excel formats

## Requirements

- Python 3.7+
- Chrome browser installed
- Required Python packages (install via `pip install -r requirements.txt`):
selenium
pandas
numpy
matplotlib
urllib3

Copy

## Installation

1. Clone this repository:
 git clone https://github.com/yourusername/franklin-county-scraper.git
 cd franklin-county-scraper

Install dependencies:
pip install -r requirements.txt
Download the appropriate ChromeDriver version for your Chrome browser from https://chromedriver.chromium.org/ and place it in your PATH or in the project directory.

## Usage
Run the script:
python scraper.py
When prompted, enter:

Start date in YYYYMMDD format (e.g., 20230101)
End date in YYYYMMDD format (e.g., 20231231)

### The script will:
First collect all parcel IDs within the date range
Then fetch detailed information for each parcel
Save results to:
ParcelIDFile_Complete.csv - All parcel IDs found
Output.xlsx - Detailed property information
Previous outputs are automatically renamed to Previous_output.xlsx

## Functions Overview
### Main Functions
get_url(): Constructs the search URL with date range parameters

get_chromedriver(): Initializes and configures Chrome WebDriver

extract_all_pin_ids(): Extracts parcel IDs from search results with pagination

search_and_get_case_data(): Fetches detailed property information for a given parcel ID

process_owner_data(): Processes and structures owner information

### Helper Functions
retries(): Decorator for automatic retry of failed functions

wait_for_element(): Waits for an element to be present

split_full_name(): Splits owner names into first/last names

generate_month_ranges(): Breaks date range into monthly chunks

## Output Fields
The final output Excel file includes these fields:

Property identification (parcel ID, address)
Owner information (names, contact details)
Property characteristics (size, rooms, year built)
Transfer history (date, price)
Rental information (if available)

## Error Handling

The script includes multiple error handling mechanisms:
Automatic retries for failed operations
Timeout handling for slow-loading elements
Graceful handling of missing data
Process isolation between date ranges to prevent complete failure

## Limitations
Website structure changes may break the script

Rate limiting or IP blocking may occur with large queries

Requires Chrome browser and matching ChromeDriver version

