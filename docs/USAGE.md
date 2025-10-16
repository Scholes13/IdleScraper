# Google Maps Scraper Usage Guide

## Installation

1. Ensure you have Python 3.6+ installed
2. Install required dependencies:
   ```
   pip install -r requirements.txt
   ```
3. Make sure Chrome browser is installed on your system

## Running the Application

### Using the Main Script

The most straightforward way to use the application is through the main script:

```
python main.py             # Launches the GUI by default
python main.py --gui       # Launches the GUI explicitly
python main.py --excel     # Launches the Excel import/export GUI
python main.py --batch     # Runs in batch processing mode
```

### Using Individual Components

#### GUI Application

The GUI application provides a user-friendly interface for scraping:

```
python src/ui/gui_scraper.py
```

Features:
- Search for individual companies
- View results directly in the interface
- Save results to CSV files
- Configure search parameters
- Enable/disable smart search and website scraping

#### Batch Scraper

For processing multiple companies at once:

```
python src/core/batch_scraper.py
```

Features:
- Process a list of companies from a CSV file
- Automatically handle errors and retries
- Save results to a consolidated CSV file

#### Excel Import/Export

For working with Excel files:

```
python src/core/import_excel.py
```

Or use the GUI version:

```
python src/ui/import_excel_gui.py
```

Features:
- Import company names from Excel files
- Export results back to Excel format
- Maintain formatting and additional data

## Advanced Features

### Smart Search

The Smart Search feature tries different variations of company names when exact matches fail:

1. Enable Smart Search in the GUI
2. Adjust the similarity threshold (default: 0.7)
3. Higher threshold values (closer to 1.0) require closer name matches
4. Lower values (closer to 0.0) are more lenient but may return incorrect matches

### Website Scraping

The Website Scraping feature extracts contact information from company websites:

1. Enable Website Scraping in the GUI
2. When Google Maps doesn't provide phone numbers but does have a website link
3. The scraper will:
   - Visit the company website
   - Look for phone numbers and email addresses on the homepage
   - Search for and examine contact/about pages
   - Validate extracted information to minimize false positives
4. If found, phone numbers and emails will be added to the results

## Input File Format

### CSV Input

The CSV input file should have a column named "Company Name" or similar containing the company names to search for.

Example:
```
Company Name
PT Telkom Indonesia
Bank Mandiri
Gojek
```

### Excel Input

For Excel files, the same structure applies - a column with company names is required.

## Output Files

The scraper generates the following output files:

1. **Results CSV/Excel**: Contains all scraped information including:
   - Company name
   - Address
   - Phone number
   - Website
   - Email
   - Rating
   - Number of reviews
   - Category
   - Geographic coordinates
   - Source of contact information (Google Maps or website)

2. **Mapping File**: Shows which company names were mapped to different names on Google Maps

## Troubleshooting

### WebDriver Issues

If you experience issues with Chrome WebDriver:

1. Run the fix script:
   ```
   src/utils/fix_selenium.bat
   ```

2. Ensure Chrome is up-to-date

3. Check error screenshots in the assets/images directory

### Search Problems

If the scraper cannot find certain companies:

1. Try different name variations manually
2. Add location information (city, country) to the company name
3. Increase the retry count and similarity threshold in the settings

### Website Scraping Issues

If the website scraper doesn't find contact information:

1. Check if the website uses non-standard formats for displaying contact info
2. Some websites use images or JavaScript to display contact information, which can't be extracted
3. Try increasing the max_pages parameter in the WebsiteScraper class for deeper website exploration 