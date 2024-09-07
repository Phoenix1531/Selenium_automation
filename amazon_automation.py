import os
import logging
from datetime import datetime
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import time

logging.basicConfig(filename='automation_log.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

driver = webdriver.Chrome()
driver.get("https://www.amazon.in")
driver.maximize_window()
logging.info("Amazon website opened and window maximized.")

products = ["Desktop","CPU","Keyboard and mouse combo"]

inputBox = driver.find_element(By.ID, "twotabsearchtextbox")
submitButton = driver.find_element(By.ID, "nav-search-submit-button")
logging.info("Located the search input box and submit button.")

screenshot_folder = os.path.join(os.getcwd(), 'Data')
if not os.path.exists(screenshot_folder):
    os.makedirs(screenshot_folder)

product_data = []

for product in products:
    try:
        inputBox.clear()
        inputBox.send_keys(product)
        submitButton.click()
        logging.info(f"Searching for {product}")

        WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, "s-result-sort-select"))
        )

        time.sleep(4)

        dropdown = driver.find_element(By.ID, "s-result-sort-select")
        select = Select(dropdown)
        select.select_by_value("review-rank")
        logging.info(f"Sorted results for {product} by review rank")

        time.sleep(5)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'h2.a-size-mini a.a-link-normal'))
        )

        link = driver.find_element(By.CSS_SELECTOR, 'h2.a-size-mini a.a-link-normal')
        driver.execute_script("arguments[0].removeAttribute('target');", link)
        link.click()
        logging.info(f"Clicked on the first result for {product}")
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, 'submit.add-to-cart'))
        )

        product_name = driver.find_element(By.ID, "productTitle").text
        product_price = driver.find_element(By.CSS_SELECTOR, ".a-price-whole").text
        product_rating = driver.find_element(By.CSS_SELECTOR, ".a-icon-alt").get_attribute('innerHTML')

        logging.info(f"Extracted details for {product}: {product_name}, {product_price}, {product_rating}")

        product_data.append({
            'Product Name': product_name,
            'Price': product_price,
            'Rating': product_rating
        })

        time.sleep(3)

        addToCart = driver.find_element(By.NAME, "submit.add-to-cart")
        addToCart.click()
        logging.info(f"Added {product} to cart")

        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

        screenshot_path = os.path.join(screenshot_folder,f"{product}_{timestamp}.png")
        driver.save_screenshot(screenshot_path)
        logging.info(f"Screenshot for {product} saved at {screenshot_path}")

        driver.back()
        inputBox = driver.find_element(By.ID, "twotabsearchtextbox")
        submitButton = driver.find_element(By.ID, "nav-search-submit-button")

    except Exception as e:
        logging.error(f"An error occurred for product '{product}': {e}")
        driver.back()

df = pd.DataFrame(product_data)
excel_path = os.path.join(os.getcwd(), 'product_data.xlsx')
df.to_excel(excel_path, index=False)
print(product_data)
logging.info(f"Product data saved to {excel_path}")

time.sleep(5)
driver.quit()
logging.info("Automation script completed and browser closed.")
