from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

from . import constants as const


class Torob(webdriver.Chrome):
    def __init__(self, teardown=False):
        self.teardown = teardown
        options = webdriver.ChromeOptions()
        options.add_experimental_option("excludeSwitches", ["enable_logging"])
        super(Torob, self).__init__(options=options)
        self.wait = WebDriverWait(self, 10)
        self.maximize_window()

    def __exit__(self, exc_type, exc, traceback):
        if self.teardown:
            self.quit()

    def land_first_page(self):
        self.get(const.BASE_URL)

    def search_box(self, query: str):
        self.search_text = query
        search_element = self.find_element(by=By.ID, value="search-query-input")
        search_element.clear()
        search_element.send_keys(self.search_text, Keys.RETURN)

    def sort_items(self, text: str):
        self.wait.until(
            EC.visibility_of_element_located(
                (
                    By.CSS_SELECTOR,
                    "div.jsx-2762017274.DropDownFilter_dropContainer__mAZBV",
                )
            )
        )
        self.find_element(
            By.CSS_SELECTOR,
            "div.jsx-2762017274.DropDownFilter_dropContainer__mAZBV",
        ).click()
        self.wait.until(
            EC.visibility_of_element_located(
                (
                    By.CSS_SELECTOR,
                    "div.dropdown__content ul.jsx-2762017274.dropdown-menuu",
                )
            )
        )
        self.find_element(By.PARTIAL_LINK_TEXT, text).click()
        self.wait.until(
            EC.invisibility_of_element_located(
                (By.CSS_SELECTOR, "div.dropdown__content")
            )
        )

    def _scroll_to_bottom(self):
        last_height = self.execute_script("return document.body.scrollHeight")
        while True:
            # Scroll down to the bottom
            self.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            # Wait for new content to load using WebDriverWait
            try:
                self.wait.until(
                    lambda driver: self.execute_script(
                        "return document.body.scrollHeight"
                    )
                    > last_height
                )
            except TimeoutException:
                print("No new content loaded after scrolling.")
                break

            # Calculate new scroll height and compare with last scroll height
            new_height = self.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break  # Exit if no new content is loaded
            last_height = new_height

    def find_commodities(self):
        self._scroll_to_bottom()
        names = self.find_elements(
            By.CSS_SELECTOR, "h2.ProductCard_desktop_product-name__JwqeK"
        )
        prices = self.find_elements(
            By.CSS_SELECTOR, "div.ProductCard_desktop_product-price-text__y20OV"
        )
        shops = self.find_elements(
            By.CSS_SELECTOR, "div.ProductCard_desktop_shops__mbtsF"
        )
        self.data_ = []
        for n, p, s in zip(names, prices, shops):
            name = n.get_attribute("innerHTML")
            price = p.get_attribute("innerHTML")
            shop = s.get_attribute("innerHTML")
            if self.search_text in name:
                self.data_.append({"name": name, "price": price, "shop": shop})

    def write_to_file(self, file_path="output.xlsx"):
        # Convert the list of dictionaries to a DataFrame
        df = pd.DataFrame(self.data_, columns=["name", "price", "shop"])
        return df


if __name__ == "__main__":
    # List of search queries
    search_queries = ["laptop", "phone", "monitor"]

    # Create an Excel writer object
    file_path = "all_products.xlsx"
    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        with Torob(teardown=True) as bot:
            for query in search_queries:
                bot.land_first_page()
                bot.search_box(query)
                bot.sort_items(text="محبوب")
                bot.find_commodities()
                df = bot.write_to_file()
                # Write the DataFrame to a sheet named after the query
                df.to_excel(writer, sheet_name=query, index=False)
                print(f"Data for '{query}' written to sheet '{query}' in {file_path}")

    print(f"All data written to {file_path}")
