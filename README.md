# ec-call-popularity-scraper

A headless Selenium tool for scraping popularity and visibility signals from EC opportunity pages.

## Overview

This script extracts popularity indicators (such as call titles or visibility metrics) from hyperlinks listed in an Excel file. It automates browser access to each URL using Selenium and collects data for further analysis.

## Requirements

Install dependencies:

    pip install selenium pandas openpyxl

Also required:
- An Excel file named `xxx.xlsx`, with hyperlinks in column A of the "Future Calls" worksheet
- A browser driver (e.g. `geckodriver` for Firefox, or `chromedriver` for Chrome)

## Configuration

Adjust parameters in `parse.py` as needed:

```python
agent = 'Firefox'      # or 'Chrome'
headless = True        # run without opening a browser window
retries = 2            # retry attempts per URL
title_filter = '-'     # filter condition for Excel titles
