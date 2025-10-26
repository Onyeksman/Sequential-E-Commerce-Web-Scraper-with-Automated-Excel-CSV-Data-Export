# 🕸️ Sequential E-Commerce Data Scraper (Scrapeme.live)

## Overview
This project is a Python Scrapy-based web scraper designed to extract and structure all product listings from [scrapeme.live/shop](https://scrapeme.live/shop/).  
It was engineered for accuracy, readability, and presentation — providing clean datasets in both `.xlsx` and `.csv` formats.

## Key Features
- **Sequential Crawling:** Scrapes each page in visible order, preserving the exact layout of the site.  
- **Clean Data Output:** Automatically removes duplicates, fills missing fields with `N/A`, and normalizes text encoding.  
- **Smart Price Handling:** Extracts both numeric and formatted prices with `$` currency formatting.  
- **Professional Excel Design:**  
  - Dark blue headers with white bold text  
  - Light grey alternate rows  
  - Medium borders around all cells  
  - Auto-fitted columns, filters, and frozen headers  
  - Timestamped metadata note for traceability  

## Tech Stack
- **Language:** Python  
- **Libraries:** Scrapy, OpenPyXL, FTFY, CSV, Datetime, Regex  
- **Outputs:** `pokemon_YYYY-MM-DD.xlsx`, `pokemon_YYYY-MM-DD.csv`

## Impact
The scraper provides a streamlined, reusable foundation for professional data collection and reporting. It delivers reliable, presentation-ready outputs for use in analytics, e-commerce monitoring, or client reporting.

## Author
**Onyekachi Ejimofor**  
Data Scraping & Automation Specialist 

