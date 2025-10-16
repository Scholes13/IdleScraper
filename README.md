# Google Maps Business Scraper - Idle Scrape

A tool for scraping business information from Google Maps, including phone numbers, addresses, websites, and more.

## Features

- **Accurate Phone Number Extraction**: Uses multiple methods to find phone numbers, with validation to prevent incorrect values
- **Smart Company Name Matching**: Automatically tries similar variations of company names when exact matches fail
- **Multiple Interfaces**: GUI application, command-line batch processing, and Excel import/export
- **Robust Error Handling**: Automatically retries failed searches and recovers from errors
- **Website Scraping**: Scrapes company websites to find contact information when not available on Google Maps
- **Enhanced Phone Number Detection**: Intelligently categorizes Indonesian phone numbers by type (mobile/office) and carrier/location
- **Phone Number Validation**: Uses Google's phonenumbers library for accurate validation, formatting, and carrier detection
- **Email Validation and Normalization**: Validates email format, detects disposable domains, and normalizes to consistent format
- **Anti-Detection System**: Implements rotating user-agents, adaptive delays, and proxy support to avoid blocking
- **Intelligent Caching**: Stores and reuses scraped data to improve performance and reduce requests

## Key Improvements

- **Phone Number Validation**: Now validates phone numbers to prevent incorrect extraction (like "2000000000")
- **Company Name Matching**: When an exact match fails, tries variations like:
  - Adding "PT" prefix if not present
  - Adding "Indonesia" or "Services" suffix
  - Trying with just the first few words of long names
  - Removing location keywords
- **Name Mapping Tracking**: Keeps track of both original and mapped company names
- **Detailed Logging**: Shows the search process and results
- **Website Contact Extraction**: When a phone number isn't found on Google Maps but a website is available:
  - Visits the company website
  - Extracts phone numbers and email addresses
  - Prioritizes contact and about pages
  - Validates extracted data to minimize false positives
- **Advanced Phone Detection**:
  - Identifies mobile vs. office/landline numbers
  - Detects carriers for mobile numbers (Telkomsel, Indosat, XL, etc.)
  - Identifies locations for office numbers (Jakarta, Bandung, etc.)
  - Works with various formats (local and international)
- **Phone Number Normalization**:
  - Formats all phone numbers to international standard (e.g., +62 8xx-xxx-xxx)
  - Validates phone numbers against Google's phonenumbers library
  - Ensures consistency across all extracted numbers
- **Email Validation**:
  - RFC-compliant format validation
  - Disposable email domain detection
  - Consistent normalized format
  - Optional deliverability checking (DNS/MX records)
- **Scraping Resilience System**:
  - User-Agent Rotation: Automatically rotates between modern browser user-agents
  - Adaptive Delay: Intelligently adjusts delay times based on response patterns
  - Proxy Support: Optional rotation of free proxies to distribute requests
  - Result Caching: Stores scraped data to reduce duplicate requests

## Anti-Detection Features

The scraper now includes advanced anti-detection features that help prevent being blocked by Google:

- **User-Agent Rotation**:
  - Automatically uses different browser and device identities
  - Includes mobile and desktop user-agents
  - Helps avoid fingerprinting and detection

- **Adaptive Delay System**:
  - Intelligent delay that adapts to throttling patterns
  - Exponential backoff when errors are detected
  - Gradual reduction after successful requests
  - Mimics human browsing behavior with random timing

- **Proxy Rotation (Experimental)**:
  - Automatically fetches and tests free proxies
  - Rotates connections through working proxies
  - Distributes requests across multiple IP addresses
  - Helps overcome IP-based rate limits

- **Result Caching**:
  - Stores previously scraped information
  - Avoids unnecessary duplicate requests
  - Configurable retention period (default: 30 days)

See [Anti-Detection Documentation](docs/anti_detection_features.md) for details on configuration and best practices.

## Project Structure

The project has been reorganized for better maintainability:

```
├── src/                      # Source code
│   ├── core/                 # Core scraping functionality 
│   ├── utils/                # Utility scripts and helpers
│   └── ui/                   # User interface components
├── assets/                   # Static assets
│   ├── icons/                # Application icons
│   └── images/               # Images used by the application
└── docs/                     # Documentation
```

For a detailed project structure, see [PROJECT_INDEX.md](docs/PROJECT_INDEX.md).

## Usage

### Easy Execution with Main Script

```
python main.py             # Launches the GUI by default
python main.py --gui       # Launches the GUI explicitly
python main.py --excel     # Launches the Excel import/export GUI
python main.py --batch     # Runs in batch processing mode
```

### Legacy Usage

#### GUI Application

```
python src/ui/gui_scraper.py
```

#### Command Line

For a single company:
```
python src/core/maps_scraper.py
```

For batch processing:
```
python src/core/batch_scraper.py
```

#### Excel Import/Export

```
python src/core/import_excel.py
```

Or use the GUI version:
```
python src/ui/import_excel_gui.py
```

## Requirements

- Python 3.6+
- Chrome browser installed
- Required packages: `selenium`, `pandas`, `openpyxl`, `webdriver-manager`, `requests`, `beautifulsoup4`, `phonenumbers`, `email-validator`

Install dependencies:
```
pip install -r requirements.txt
```

## Phone Number Detection

The scraper now includes advanced phone number detection for Indonesian numbers:

- **Mobile Numbers**: Identified by prefixes (08xx, +628xx)
  - Uses Google's phonenumbers library for carrier detection
  - Example: "0812345678" → "Mobile (Telkomsel)"

- **Office/Landline Numbers**: Identified by area codes
  - Uses geocoder information for accurate location identification
  - Example: "02112345678" → "Office (Jabodetabek)"

- **Number Validation and Formatting**:
  - Validates all phone numbers against international standards
  - Formats numbers consistently in international format
  - Example: "0811234567" → "+62 811-234-567"

This feature helps better organize and prioritize scraped contact information.
See [Phone Detection Documentation](docs/phone_detection.md) for details.

## Email Validation

The scraper now includes advanced email validation capabilities:

- **Format Validation**: Ensures emails adhere to RFC standards
  - Detects malformed emails before they enter your data
  - Reports detailed errors for invalid formats

- **Domain Analysis**:
  - Detects disposable/temporary email domains
  - Optional deliverability checking

- **Normalization**:
  - Ensures consistent email format
  - Preserves specialized formats like gmail's plus addressing

See [Email Validation Documentation](docs/email_validation.md) for details.

## Troubleshooting

If you experience issues with the Chrome WebDriver:

1. Run `src/utils/fix_selenium.bat` (Windows) or manually install the appropriate ChromeDriver for your Chrome version
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