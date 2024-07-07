# Product Scraper

## Description

`Product Scraper` is a Python script for searching products on Google Shopping. The script automates the search process and collects data on products that match specified criteria.

## What the Code Does

**This script performs the following tasks**:
    1. **Automated Google Shopping Search**: Uses Selenium to automate the process of searching for products on Google Shopping based on the names listed in the `products.xlsx` file.
    2. **Data Collection**: Scrapes information about products, including price, shipping cost, and link to the product.
    3. **Data Filtering**: Filters the products to find the best match based on similarity to the expected product name and within the specified maximum purchase price.
    4. **Data Storage**: Saves the collected data and filtered results into separate Excel files.
    5. **Periodic Browser Restart**: Restarts the browser periodically to avoid detection by anti-bot measures such as CAPTCHA.


## Installation

### Requirements

- Python 3.x
- pip (Python package manager)
- Chrome WebDriver

### Installation Steps

1. **Clone the repository**:

   ```bash
   git clone https://github.com/Abdualiev01/product-scraper.git

2. **Navigate to the project directory**:
    cd product-scraper

3. **Install dependencies**:
    pip install -r requirements.txt


