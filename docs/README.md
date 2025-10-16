# Google Maps Business Scraper

A tool for scraping business information from Google Maps, including phone numbers, addresses, websites, and more.

## Features

- **Accurate Phone Number Extraction**: Uses multiple methods to find phone numbers, with validation to prevent incorrect values
- **Smart Company Name Matching**: Automatically tries similar variations of company names when exact matches fail
- **Multiple Interfaces**: GUI application, command-line batch processing, and Excel import/export
- **Robust Error Handling**: Automatically retries failed searches and recovers from errors

## Key Improvements

- **Phone Number Validation**: Now validates phone numbers to prevent incorrect extraction (like "2000000000")
- **Company Name Matching**: When an exact match fails, tries variations like:
  - Adding "PT" prefix if not present
  - Adding "Indonesia" or "Services" suffix
  - Trying with just the first few words of long names
  - Removing location keywords
- **Name Mapping Tracking**: Keeps track of both original and mapped company names
- **Detailed Logging**: Shows the search process and results

## Usage

### GUI Application

```
python gui_scraper.py
```

### Command Line

For a single company:
```
python maps_scraper.py
```

For batch processing:
```
python batch_scraper.py
```

### Excel Import/Export

```
python import_excel.py
```

Or use the GUI version:
```
python import_excel_gui.py
```

## Requirements

- Python 3.6+
- Chrome browser installed
- Required packages: `selenium`, `pandas`, `openpyxl`, `webdriver-manager`

Install dependencies:
```
pip install -r requirements.txt
```

## Troubleshooting

If you experience issues with the Chrome WebDriver:

1. Run `fix_selenium.bat` (Windows) or manually install the appropriate ChromeDriver for your Chrome version
2. Ensure your Chrome browser is up-to-date
3. Check the `error_screenshot_*.png` images to see what went wrong during scraping

## Understanding Results

When a company name is mapped to a different name in Google Maps:

1. Both original and mapped names are displayed
2. The "Original Company" is the name from your input
3. The "Matched Company" is what was found on Google Maps
4. The "Search Variation" shows which variation of the name was successful

## Output Files

- Regular data output: CSV with all scraped information
- Mappings output: Separate CSV file showing which company names were mapped

## License

MIT 