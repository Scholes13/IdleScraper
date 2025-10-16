# Google Maps Scraper Project Index

## Project Structure

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

## Core Components

### Core Scraping (`src/core`)
- Core scraping functionality
- Batch processing
- Data extraction logic
- Website scraping for contact information

#### Key Files:
- `maps_scraper.py`: Main Google Maps scraping functionality
- `batch_scraper.py`: Batch processing of multiple companies
- `website_scraper.py`: Extracts contact information from company websites

### Utilities (`src/utils`)
- WebDriver management
- Installation scripts
- Build utilities
- Helper functions

### User Interface (`src/ui`)
- GUI applications
- Excel import/export interfaces

## Assets

### Icons (`assets/icons`)
- Application icons (.ico files)

### Images (`assets/images`)
- Screenshots
- Logo images
- Example images

## Other Important Files

- `requirements.txt`: Python dependencies
- `README.md`: Project documentation

## Portable Application

The project includes a portable application version for easy distribution:
- `Werkudara - Google Maps Scraper Portable/`
- `build_portable_app.py`: Script to build the portable application 