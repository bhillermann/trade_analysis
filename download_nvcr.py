#!/usr/bin/env python3
from __future__ import annotations
import logging
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from datetime import datetime
from pathlib import Path
from typing import Optional
import sys

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

NVCR_URL: str = (
    "https://www.environment.vic.gov.au/"
    "native-vegetation/native-vegetation-removal-regulations"
)

def get_trade_data(filename: Path) -> None:
    """Download traded credits information file from NVCR website using Selenium."""
    logging.info("Starting get_trade_data()")

    # Configure Firefox in headless mode
    options = Options()
    # options.add_argument("--headless")

    # Initialize WebDriver
    driver = webdriver.Firefox(options=options)
    try:
        driver.get(NVCR_URL)
        logging.info("Page loaded successfully.")

        # Get page source after rendering
        html = driver.page_source
    except Exception as e:
        logging.error("Failed to load the NVCR page: {e}")
        sys.exit(1)

    soup = BeautifulSoup(html, "lxml")
    download_link: Optional[str] = None

    for link in soup.find_all("a", href=True):
        if "Traded credits information" in link.get_text(strip=True):
            download_link = link["href"]
            break

    if not download_link:
        logging.error("Download link for traded credits information not found.")
        raise ValueError("Download link not found.")

    logging.info(f"Download link found: {download_link}")

    # Download the file using requests
    try:
        driver.get(download_link)
        logging.info("File downloaded successfully.")
    except Exception as e:
        logging.error(f"Failed to download file: {e}")
        sys.exit(1)

#     filename.parent.mkdir(parents=True, exist_ok=True)
#    with open(filename, "wb") as file:
#        file.write(response)
#        logging.info(f"File saved to {filename}")

if __name__ == "__main__":
    output_dir = Path("/home/bhillermann/Documents/Trade Analysis")
    filename = output_dir / f"NVCR_Trade-prices-{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    get_trade_data(filename)
