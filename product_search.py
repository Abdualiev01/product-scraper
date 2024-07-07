import pandas as pd
import time
import random
import re
import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException, ElementClickInterceptedException, TimeoutException
import difflib
import sys

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def start_browser():
    options = webdriver.ChromeOptions()
    options.add_argument('--incognito')
    options.add_argument('--disable-extensions')
    options.add_argument('--start-maximized')
    driver = webdriver.Chrome(options=options)
    driver.get("https://www.google.com/shopping")
    accept_cookies(driver)
    return driver

def accept_cookies(driver):
    try:
        wait = WebDriverWait(driver, 10)
        accept_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="yDmH0d"]/c-wiz/div/div/div/div[2]/div[1]/div[3]/div[1]/div[1]/form[2]/div/div/button')))
        accept_button.click()
    except (NoSuchElementException, ElementNotInteractableException, ElementClickInterceptedException, TimeoutException) as e:
        print("Error while accepting cookies:", e)

def search_product(driver, product_name):
    search_box = driver.find_element(By.NAME, 'q')
    search_box.clear()
    time.sleep(random.uniform(1, 3))
    search_box.send_keys(product_name)
    time.sleep(random.uniform(1, 3))
    search_box.send_keys(Keys.RETURN)
    time.sleep(random.uniform(3, 6))

def is_similar(title1, title2):
    return difflib.SequenceMatcher(None, title1.lower(), title2.lower()).ratio()

def contains_all_keywords(title, keywords):
    return all(keyword.lower() in title.lower() for keyword in keywords)

def get_best_match_title(product_elements, expected_title, purchase_price_max):
    best_match = None
    highest_similarity = 0
    expected_keywords = expected_title.split()

    all_products = []
    for product_element in product_elements:
        try:
            title_element = None
            try:
                title_element = product_element.find_element(By.XPATH, ".//h3[@class='tAxDx']")
            except NoSuchElementException:
                try:
                    title_element = product_element.find_element(By.XPATH, ".//div[@class='sh-np__product-title translate-content']")
                except NoSuchElementException:
                    pass
            if not title_element:
                continue

            title_text = title_element.text
            price_element = product_element.find_element(By.XPATH, ".//span[contains(@class, 'a8Pemb')]")
            price_text = price_element.text
            price = parse_price(price_text)
            shipping_element = product_element.find_element(By.XPATH, ".//div[contains(@class, 'vEjMR')]")
            shipping_text = shipping_element.text
            shipping_cost = parse_shipping_cost(shipping_text)
            total_price = price + (shipping_cost if shipping_cost is not None else 0)
            link_element = product_element.find_element(By.XPATH, ".//a[@class='xCpuod']")
            link = link_element.get_attribute("href")
            
            similarity = is_similar(expected_title, title_text)
            if contains_all_keywords(title_text, expected_keywords) or similarity >= 0.8:
                all_products.append((total_price, link, title_text))
                if similarity > highest_similarity:
                    highest_similarity = similarity
                    best_match = (total_price, link, title_text)
        except NoSuchElementException as e:
            continue
    return best_match, highest_similarity, all_products

def parse_price(price_text):
    try:
        price_text = price_text.replace('$', '').replace(',', '').replace('PLN', '').replace(' ', '').replace(',', '.')
        price = float(price_text)
        return price
    except ValueError:
        return 0.0

def parse_shipping_cost(shipping_text):
    if 'Free' in shipping_text:
        return 0.0
    try:
        match = re.search(r'\d+(\.\d+)?', shipping_text)
        if match:
            return float(match.group())
    except ValueError:
        return 0.0
    return None

def get_prices_from_page(driver, expected_title, purchase_price_max):
    products = []
    try:
        product_elements = driver.find_elements(By.XPATH, "//div[contains(@class, 'sh-dgr__grid-result')]")
        best_match, similarity, all_products = get_best_match_title(product_elements, expected_title, purchase_price_max)
    except NoSuchElementException:
        pass
    return best_match, all_products

def get_all_prices(driver, product_name, purchase_price_max):
    search_product(driver, product_name)
    all_products = []
    best_match = None
    while True:
        match, products = get_prices_from_page(driver, product_name, purchase_price_max)
        if products:
            all_products.extend(products)
        if not best_match and match:
            best_match = match
        try:
            next_button = driver.find_element(By.ID, 'pnnext')
            next_button.click()
            time.sleep(random.uniform(3, 6))
        except NoSuchElementException:
            break
    return best_match, all_products

def sanitize_filename(filename):
    filename = re.sub(r'[<>:"/\\|?*]', '', filename)
    return filename[:255]

input_file = os.path.join(os.getcwd(), 'data', 'products.xlsx')
temp_file = os.path.join(os.getcwd(), 'data', 'products_temp.xlsx')

if not os.path.exists(input_file):
    print(f"Error: The file {input_file} does not exist.")
    sys.exit(1)

df = pd.read_excel(input_file)

if not os.path.exists('data'):
    os.makedirs('data')

driver = start_browser()

for index, row in df.iterrows():
    if index > 0 and index % 4 == 0:
        driver.quit()
        print("Processed 4 products. Restarting browser to avoid reCAPTCHA.")
        time.sleep(random.uniform(240, 360))
        driver = start_browser()
        
    product_name = row['TITLE']
    purchase_price_max = row['PURCHASE PRICE MAX']
    
    print(f"Processing product: {product_name}")
    best_match, products = get_all_prices(driver, product_name, purchase_price_max)
    
    if products:
        filtered_products = [product for product in products if is_similar(product_name, product[2]) >= 0.8 or contains_all_keywords(product[2], product_name.split())]
        if filtered_products:
            product_df = pd.DataFrame(filtered_products, columns=['Total Price', 'Link', 'Title'])
            try:
                sanitized_filename = sanitize_filename(product_name)
                filepath = os.path.join('data', f'{sanitized_filename}.xlsx')
                product_df.to_excel(filepath, index=False)
                print(f"Filtered suitable products for {product_name} saved to {filepath}.")
            except PermissionError:
                sanitized_filename = sanitize_filename(product_name + '_temp')
                filepath = os.path.join('data', f'{sanitized_filename}.xlsx')
                product_df.to_excel(filepath, index=False)
                print(f"Could not save to original file. Saved to {filepath} instead.")
    
    if best_match and best_match[0] <= purchase_price_max:
        df.at[index, 'SOURCE LINK'] = best_match[1]
        df.at[index, 'source price \n(include all fees)'] = best_match[0]
    else:
        comment = f"No product found within the price {purchase_price_max}"
        df.at[index, 'SOURCE LINK'] = str(comment)
        df.at[index, 'source price \n(include all fees)'] = str(comment)

    try:
        df.to_excel(input_file, index=False)
        print("Data successfully saved to Excel file.")
    except PermissionError:
        print("Error: The file is open or locked. Please close the file and try again.")
        try:
            df.to_excel(temp_file, index=False)
            print("Data successfully saved to temporary Excel file.")
        except Exception as e:
            print(f"Error saving data to temporary Excel file: {e}")

driver.quit()
