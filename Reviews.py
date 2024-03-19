import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from openpyxl import Workbook
from openpyxl import load_workbook
import signal

def save_reviews_and_exit(driver, workbook, file_path, total_reviews_collected):
    workbook.save(filename=file_path)
    print(f"\nTotal number of collected reviews: {total_reviews_collected}")
    print(f"The Excel file has been saved at: {os.path.abspath(file_path)}")
    driver.quit()
    exit(0)

def signal_handler(sig, frame, driver, workbook, file_path, total_reviews_collected):
    print('\nYou pressed Ctrl+C!')
    save_reviews_and_exit(driver, workbook, file_path, total_reviews_collected)

def get_user_input():
    user_input_item = input("Enter the item you want to scrape: ")
    max_pages = input("Enter the number of pages to collect reviews up to (leave blank to scrape all pages): ")
    max_pages = int(max_pages) if max_pages else float('inf')
    return user_input_item, max_pages

def initialize_driver():
    os.environ["PATH"] += os.pathsep + 'C:\\chromedriver.exe'
    chrome_options = Options()
    chrome_options.add_argument("--incognito")
    driver = webdriver.Chrome(options=chrome_options)
    return driver

def get_base_url(user_input_item):
    base_url = f"https://www.daraz.com.bd/catalog/?q={user_input_item}&_keyori=ss&from=input&spm=a2a0e.home.search.go.285012f7rLpOTH"
    return base_url

def initialize_workbook(file_path):
    try:
        workbook = load_workbook(filename=file_path)
        worksheet = workbook.active
    except FileNotFoundError:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = 'Reviews'
        worksheet.append(['Item', 'Product Name', 'Total Ratings', 'Answered Questions', 'Price After Discount', 'Actual Price', 'Discount Percentage', 'Review'])
    return workbook, worksheet

def scroll_to_reviews(driver):
    current_height = 0
    while True:
        driver.execute_script(f"window.scrollTo(0, {current_height});")
        time.sleep(2)
        reviews = driver.find_elements(By.CLASS_NAME, "review-content-sl")
        if reviews:
            break
        current_height += 500
        if current_height >= driver.execute_script("return document.body.scrollHeight;"):
            return None
    return reviews

def collect_product_info(driver):
    product_name_element = driver.find_element(By.CSS_SELECTOR, ".pdp-mod-product-badge-title")
    product_name = product_name_element.text if product_name_element else ""

    total_ratings_element = driver.find_element(By.CSS_SELECTOR, ".pdp-review-summary__link")
    total_ratings = total_ratings_element.text.split()[0] if total_ratings_element else ""

    answered_questions_element = driver.find_element(By.XPATH, "//a[contains(text(), 'Answered Questions')]")
    answered_questions = answered_questions_element.text.split()[0] if answered_questions_element else ""

    price_after_discount_element = driver.find_element(By.CSS_SELECTOR, ".pdp-price_type_normal")
    price_after_discount = price_after_discount_element.text if price_after_discount_element else ""

    actual_price_element = driver.find_element(By.CSS_SELECTOR, ".pdp-price_type_deleted")
    actual_price = actual_price_element.text if actual_price_element else ""

    discount_percentage_element = driver.find_element(By.CSS_SELECTOR, ".pdp-product-price__discount")
    discount_percentage = discount_percentage_element.text if discount_percentage_element else ""

    return product_name, total_ratings, answered_questions, price_after_discount, actual_price, discount_percentage

def collect_reviews(driver, user_input_item, worksheet, max_pages):
    current_page = 1
    total_reviews_collected = 0
    while current_page <= max_pages:
        reviews = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "review-content-sl"))
        )
        product_name, total_ratings, answered_questions, price_after_discount, actual_price, discount_percentage = collect_product_info(driver)
        for review in reviews:
            review_text = review.text
            worksheet.append([user_input_item, product_name, total_ratings, answered_questions, price_after_discount, actual_price, discount_percentage, review_text])
            total_reviews_collected += 1
        print(f"\rCollecting reviews on page {current_page} of the item's review. {total_reviews_collected} reviews collected.", end="", flush=True)
        pagination_items = driver.find_elements(By.CSS_SELECTOR, ".ant-pagination-item")
        next_page_link = None
        for item in pagination_items:
            if item.get_attribute("title") == str(current_page + 1):
                next_page_link = item.find_element(By.TAG_NAME, "a")
                break
        if next_page_link:
            next_page_link.click()
            current_page += 1
            time.sleep(3)
        else:
            break
    return total_reviews_collected

def main():
    user_input_item, max_pages = get_user_input()
    driver = initialize_driver()
    base_url = get_base_url(user_input_item)
    current_dir = os.path.dirname(os.path.abspath(__file__))
    file_name = "daraz_reviews.xlsx"
    file_path = os.path.join(current_dir, file_name)
    workbook, worksheet = initialize_workbook(file_path)
    next_row = worksheet.max_row + 1
    product_index = 1
    total_reviews_collected = 0
    signal.signal(signal.SIGINT, lambda sig, frame: signal_handler(sig, frame, driver, workbook, file_path, total_reviews_collected))
    while True:
        driver.get(base_url)
        wait = WebDriverWait(driver, 10)
        products = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".product-card--vHfY9")))
        if product_index > len(products):
            break
        products[product_index - 1].click()
        time.sleep(3)
        reviews = scroll_to_reviews(driver)
        if not reviews:
            print(f"\rSkipping item {product_index} due to missing reviews.", end="", flush=True)
            product_index += 1
            continue
        total_reviews_collected += collect_reviews(driver, user_input_item, worksheet, max_pages)
        product_index += 1
    save_reviews_and_exit(driver, workbook, file_path, total_reviews_collected)

if __name__ == "__main__":
    main()