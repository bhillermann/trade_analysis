# Copilot Instructions for Trade Analysis Repository

## Project Overview
This repository contains a trade analysis tool for the Victorian Native Vegetation Credit Register (NVCR). It processes General Habitat Units (GHU), Large Tree (LT), and Species Habitat Units (SHU) trading data, combining it with supply data scraped from the NVCR website to generate detailed Excel reports with per-CMA (Catchment Management Authority) analysis.

## Key Components

### Core Scripts
- **`trade_analysis.py`**: Main script for data analysis and report generation.
  - Functions:
    - `get_trade_data()`: Downloads NVCR trade data using Selenium.
    - `save_nvcr_file()`: Saves NVCR trade data to a specified location.
    - `wait_for_download()`: Waits for file downloads to complete.
  - Handles data parsing, filtering, and report generation.
- **`ghu_search.py`**: Scrapes supply data for all CMAs.
- **`clean_traded_credits.py`**: Cleans and exports trade data to CSV.

### Data Flow
1. **Data Collection**:
   - `get_trade_data()`: Downloads trade data as an Excel file.
   - `get_supply()`: Scrapes supply data for CMAs.
2. **Data Processing**:
   - Filters and normalizes data.
   - Calculates metrics like trade volumes, values, and theoretical LT values.
3. **Report Generation**:
   - Creates multi-sheet Excel reports using `xlsxwriter`.
   - Applies final formatting with `openpyxl`.

### External Dependencies
- **Python Libraries**: `numpy`, `pandas`, `openpyxl`, `beautifulsoup4`, `selenium`, `thefuzz`, `lxml`, `xlsxwriter`
- **Browser Tools**: Firefox and geckodriver for Selenium-based scraping.

## Development Workflow

### Setting Up the Environment
This project uses Nix flakes for dependency management.
- Enter the development shell:
  ```bash
  nix develop
  ```
- Build the package:
  ```bash
  nix build
  ```

### Running the Main Script
Run `trade_analysis.py` with optional arguments:
```bash
python trade_analysis.py -s <supply_file> -i <input_file> -o <output_file>
```
- `-s/--supply`: Path to supply Excel file (default: scrape new data).
- `-i/--input`: Path to NVCR trade prices file (default: download new file).
- `-o/--output`: Output filename (default: `Trade-Analysis.xlsx`).
- `-b/--start` and `-e/--end`: Start and end dates for analysis.
- `--download-nvcr`: Download NVCR trade data and exit.

### Utility Scripts
- Scrape supply data:
  ```bash
  python ghu_search.py --output <output_path>
  ```
- Download NVCR trade data:
  ```bash
  python trade_analysis.py --download-nvcr --output <output_path>
  ```
- Clean and export trade data:
  ```bash
  python clean_traded_credits.py --output <output_path>
  ```

## Project-Specific Conventions

### Data Structures
- **CMAs**: Catchment Management Authorities (e.g., Corangamite, Melbourne Water).
- **Trade Data Columns**: `date`, `cma`, `sbv`, `ghu`, `lt`, `sbu`, `ghu_price`, `shu_price`, `species`, `price_in_gst`, `price_ex_gst`.
- **Water Authority Property IDs**: Defined in `wa` dictionary in `trade_analysis.py`.

### Patterns and Practices
- **Temporary Files**: Use `tempfile.TemporaryDirectory()` for downloads.
- **Data Passing**: Functions return `pandas` objects instead of writing intermediate files.
- **Date Filtering**: Default range is the last 12 months.
- **Excel Formatting**: Use `openpyxl` for post-processing (e.g., fonts, currency formats).

### Selenium Configuration
- Headless Firefox with custom download preferences.
- Use `BeautifulSoup` for parsing HTML and locating download links.

## Notes for AI Agents
- Follow the data flow and modular structure when adding new features.
- Maintain the use of temporary directories for file handling.
- Ensure compatibility with Nix-based environments and Python 3.
- Adhere to existing conventions for CMA normalization and data filtering.
- Adhere to DRY principles by reusing functions for data retrieval and processing.
- Use Type Hints (Annotations) for all variables and function signatures.
- Make regular commits with clear messages when modifying code.
- Create a new branch for significant changes or features.

For further details, refer to `CLAUDE.md` or the source code.