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
    try:
        workbook.save(filename=file_path)
        print(f"\nTotal number of collected reviews: {total_reviews_collected}")
        print(f"The Excel file has been saved at: {os.path.abspath(file_path)}")
    except Exception as e:
        print(f"\nError occurred while saving the Excel file: {e}")
    finally:
        driver.quit()
        exit(0)

def main():
    def signal_handler(sig, frame):
        print('\nYou pressed Ctrl+C!')
        save_reviews_and_exit(driver, workbook, file_path, total_reviews_collected)

    signal.signal(signal.SIGINT, signal_handler)

    user_input_item = input("Enter the item you want to scrape: ")
    max_pages = input("Enter the number of pages to collect reviews up to (leave blank to scrape all pages): ")
    max_pages = int(max_pages) if max_pages else float('inf')

    os.environ["PATH"] += os.pathsep + 'C:\\chromedriver.exe'
    chrome_options = Options()
    chrome_options.add_argument("--incognito")
    driver = webdriver.Chrome(options=chrome_options)

    base_url = f"https://www.daraz.com.bd/catalog/?q={user_input_item}&_keyori=ss&from=input&spm=a2a0e.home.search.go.285012f7rLpOTH"
    driver.get(base_url)

    current_dir = os.path.dirname(os.path.abspath(__file__))
    file_name = "daraz_reviews.xlsx"
    file_path = os.path.join(current_dir, file_name)

    try:
        workbook = load_workbook(filename=file_path)
        worksheet = workbook.active
    except FileNotFoundError:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = 'Reviews'
        worksheet.append(['Item', 'Review'])

    next_row = worksheet.max_row + 1
    product_index = 1
    total_reviews_collected = 0

    try:
        while True:
            driver.get(base_url)
            wait = WebDriverWait(driver, 10)
            products = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".product-card--vHfY9")))

            if product_index > len(products):
                break

            try:
                products[product_index - 1].click()
                time.sleep(3)

                current_height = 0
                while True:
                    driver.execute_script(f"window.scrollTo(0, {current_height});")
                    time.sleep(2)
                    reviews = driver.find_elements(By.CLASS_NAME, "review-content-sl")
                    if reviews:
                        break
                    current_height += 500
                    if current_height >= driver.execute_script("return document.body.scrollHeight;"):
                        print(f"\rSkipping item {product_index} due to missing reviews.", end="", flush=True)
                        product_index += 1
                        break

                if not reviews:
                    continue

                current_page = 1
                while current_page <= max_pages:
                    try:
                        reviews = WebDriverWait(driver, 10).until(
                            EC.presence_of_all_elements_located((By.CLASS_NAME, "review-content-sl"))
                        )
                        for review in reviews:
                            review_text = review.text
                            worksheet.append([user_input_item, review_text])
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
                    except Exception as e:
                        print(f"\rSkipping item {product_index} due to missing reviews or page number.", end="", flush=True)
                        break

                product_index += 1
            except Exception as e:
                print(f"\rSkipping item {product_index} due to an error: {e}", end="", flush=True)
                product_index += 1
                continue
    except Exception as e:
        print(f"\nAn error occurred: {e}")
    finally:
        save_reviews_and_exit(driver, workbook, file_path, total_reviews_collected)

if __name__ == "__main__":
    main()