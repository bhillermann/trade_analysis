# NVCR Trade Analysis

Trade analysis tool for the Victorian Native Vegetation Credit Register (NVCR).
It combines NVCR trade price data with scraped supply data and writes a
multi-sheet Excel report with per-CMA analysis of General Habitat Units (GHU),
Large Trees (LT), and Species Habitat Units (SHU).

## Requirements

The project uses a Nix flake. It provides Python 3 with pandas, openpyxl,
xlsxwriter, beautifulsoup4, selenium, thefuzz, requests and lxml, plus Firefox
and geckodriver for scraping.

```bash
nix develop     # enter the development shell
nix build       # build the package
```

## Usage

```bash
python trade_analysis.py [options]
```

With the package installed from the flake, the entry point is `trade-analysis`.

### Options

| Option | Description |
| --- | --- |
| `-s`, `--supply PATH` | Read supply data from an existing Excel file. Default is to scrape new data. |
| `-i`, `--input PATH` | Read NVCR trade prices from an existing file. Default is to download a new file. |
| `-o`, `--output PATH` | Output file. Default `Trade-Analysis.xlsx`. |
| `-b`, `--start YYYY-MM-DD` | Start of the analysis period. Default is 12 months before the end date. |
| `-e`, `--end YYYY-MM-DD` | End of the analysis period. Default is the last day of the previous month. |
| `--download-nvcr PATH` | Download the NVCR trade price file to PATH, then exit. |
| `--download-supply [PATH]` | Scrape supply data to PATH, then exit. Default `Supply_{timestamp}.xlsx`. |

`--download-supply` writes one sheet per CMA. A literal `{}` in PATH is replaced
with a timestamp.

### Examples

Run the analysis with cached supply data:

```bash
python trade_analysis.py -s ~/NVCR-Data/supply.xlsx -o Trade_Analysis.xlsx
```

Save the supply scrape for later runs:

```bash
python trade_analysis.py --download-supply ~/NVCR-Data/supply.xlsx
```

Save the NVCR trade price file without running the analysis:

```bash
python trade_analysis.py --download-nvcr ~/NVCR-Data/NVCR_Trade-prices.xlsx
```

## Other scripts

| Script | Purpose |
| --- | --- |
| `ghu_search.py -o PATH` | Scrape supply data only. Same output as `--download-supply`. |
| `clean_traded_credits.py -i PATH` | Parse an NVCR file and export clean HU/SHU CSV files. |
| `format.py` | Example openpyxl formatting. Not used in the main workflow. |

## Catchment Management Authorities

Corangamite, Melbourne Water, Wimmera, Glenelg Hopkins, Goulburn Broken,
West Gippsland, East Gippsland, Mallee, North Central, North East.

Port Phillip and Westernport is merged into Melbourne Water.
