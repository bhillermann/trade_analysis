# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a trade analysis tool for Victorian Native Vegetation Credit Register (NVCR) data. It analyzes General Habitat Units (GHU), Large Tree (LT), and Species Habitat Units (SHU) trading data, combining it with supply data scraped from the NVCR website to generate comprehensive Excel reports with per-CMA (Catchment Management Authority) analysis.

## Development Environment

This project uses Nix flakes for dependency management. The development environment includes:
- Python 3 with packages: numpy, pandas, openpyxl, beautifulsoup4, selenium, thefuzz, requests, lxml, xlsxwriter
- Firefox and geckodriver for web scraping

### Commands

**Enter development shell:**
```bash
nix develop
```

**Build the package:**
```bash
nix build
```

**Run the main analysis (from dev shell):**
```bash
python trade_analysis.py
```

The main script accepts several arguments:
- `-s/--supply`: Path to existing supply Excel file (otherwise scrapes new data)
- `-i/--input`: Path to existing NVCR trade prices file (otherwise downloads new file)
- `-o/--output`: Output filename (default: Trade-Analysis.xlsx)
- `-b/--start`: Start date for analysis (default: 12 months ago from end date)
- `-e/--end`: End date for analysis in YYYY-MM-DD format (default: last day of previous month)
- `--download-nvcr`: Download NVCR trade data and save to specified path, then exit without running analysis

**Run individual scripts:**
```bash
# Scrape supply data only
python ghu_search.py --output /mnt/c/Users/BrendonHillermann/OneDrive\ -\ VL/Documents/Trade\ Analysis/Supply_output.xlsx

# Download NVCR trade data only (without running analysis)
python trade_analysis.py --download-nvcr --output /mnt/c/Users/BrendonHillermann/OneDrive\ -\ VL/Documents/Trade\ Analysis/output_file.xlsx

# Clean and export trade data to CSV
python clean_traded_credits.py --output /mnt/c/Users/BrendonHillermann/OneDrive\ -\ VL/Documents/Trade\ Analysis/Supply_output.xlsx
```

## Architecture

### Core Module Dependencies

```
trade_analysis.py (main)
├── ghu_search.py (get_supply - returns Dict[str, DataFrame])
└── Contains inlined: get_trade_data(), save_nvcr_file(), wait_for_download()
```

### Data Flow

1. **Data Collection** (optional, can use cached files):
   - `get_trade_data()` in `trade_analysis.py`: Selenium-based function that downloads "Traded credits information" Excel file to temporary directory and returns pd.ExcelFile object
   - `get_supply()` in `ghu_search.py`: Selenium-based function that searches NVCR for supply data across all CMAs and returns Dict[CMA_name → DataFrame]
   - `save_nvcr_file()` in `trade_analysis.py`: Downloads NVCR file and saves to specified location without running analysis (used by --download-nvcr)
   - **Automatic cleanup**: Downloaded files use temporary directories that are automatically cleaned up after loading into memory

2. **Data Processing** (`trade_analysis.py`):
   - Parses "Trade Prices by HU" sheet from NVCR Excel file
   - Separates GHU and SHU trades
   - Filters data by date range
   - Uses fuzzy matching (thefuzz) to normalize CMA names
   - Merges Port Phillip and Westernport into Melbourne Water
   - Calculates per-CMA metrics including:
     - Trade volumes and values
     - Average/median/floor prices with and without LTs
     - Theoretical LT values
     - Supply vs demand (years of supply)
     - Water Authority holdings

3. **Report Generation**:
   - Creates multi-sheet Excel workbook with xlsxwriter
   - Applies post-processing formatting with openpyxl (Rubik Light font, currency formats)
   - Sheets include: HU Data, SHU Data, HU Summary, and one sheet per CMA

### Key Data Structures

**CMAs (Catchment Management Authorities):**
- Corangamite, Melbourne Water, Wimmera, Glenelg Hopkins, Goulburn Broken, West Gippsland, East Gippsland, Mallee, North Central, North East
- Port Phillip and Westernport is merged into Melbourne Water

**Water Authority Property IDs:**
- Hardcoded in `wa` dict in trade_analysis.py:88-72
- Used to identify and exclude WA holdings from supply calculations

**Trade Data Columns:**
- date, cma, sbv, ghu, lt, sbu, ghu_price, shu_price, species, price_in_gst, price_ex_gst

### Selenium Web Scraping

Both scrapers use Firefox with Selenium in headless mode:
- `get_supply()`: Iterates through CMAs, fills search form, extracts supply table (the 5th table, index 4), returns data structures in memory
- `get_trade_data()`: Navigates to NVCR page, finds "Traded credits information" link using BeautifulSoup, downloads to temp directory with automatic cleanup, returns pd.ExcelFile
- **Temporary file handling**: Downloads use `tempfile.TemporaryDirectory()` with context managers for automatic cleanup
- **Firefox download preferences**: Configured to save files to temp directory without user prompts

## Utility Scripts

- `clean_traded_credits.py`: Standalone script to parse NVCR Excel and export clean HU/SHU CSV files
- `format.py`: Example script showing openpyxl formatting (not used in main workflow)

## Important Implementation Notes

- **Data passing**: Functions return pandas DataFrames/ExcelFile objects instead of writing intermediate files to disk
- **Temporary files**: All downloads use `tempfile.TemporaryDirectory()` for automatic cleanup - no intermediate files remain in working directory
- **Memory efficiency**: Files are loaded into pandas objects before temp directories are cleaned up
- Date filtering uses Python datetime with date range from last 12 months (default)
- SHU analysis includes both 1-year and 3-year summaries
- Currency values are in AUD, ex-GST for calculations
- Final Excel formatting requires openpyxl because xlsxwriter cannot read/modify existing files
- LT (Large Tree) value is calculated theoretically: (Total value - (GHUs without LTs × avg price without LTs)) / Total LTs
