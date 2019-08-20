from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec

from bs4 import BeautifulSoup

'''Web Scarping  for www.drivenow.com.au'''
'''
TASKS:
    1. We need to enter pick up date & time and drop off date & time for search window.
    2. Once we choose this and run macro, it should bring up www.drivenow.com.au website
        and enter these dates and time in search window. Pick up location is always Adelaide Airport.
        Then macro automatically click on search.
    3. It will bring up rates for 5 major car rentals, only those rental companies we need to know.

    4. At this stage, we need to scrap data from this webpage back into excel.

'''

# chrome_path = "D:/Softwares/webdriver/chromedriver"
# driver = webdriver.Firefox(
#     executable_path="D:/Softwares/webdriver/geckodriver")

url = "https://www.drivenow.com.au"


def fetch_data(driver, url):
    driver.get(url)

    # set pickup location to Adelaide Airport
    pickup_location = driver.find_element(By.ID,
                                          "s0-10-1-1-2-4-4-1-14[0]-3-0-pickupLocation")
    pickup_location.clear()
    pickup_location.send_keys("Adelaide Airport (ADL)")

    # select pickup location from suggestion
    if pickup_location.is_selected:
        print("location pickup selected")

        # wait for suggestions to appear
        wait = WebDriverWait(driver, 10)

        elem = wait.until(ec.visibility_of_element_located(
            (By.XPATH, "//div[@class='autocomplete-suggestions LocationAutocompleteSuggestions LocationAutocompleteSuggestions-s0-10-1-1-2-4-4-1-14[0]-3-0-1-6-pickupLocation X-LocationAutocompleteSuggestions X-LocationAutocompleteSuggestions-Pickup']")))
        # suggestion_elem = driver.find_element_by_xpath("//div[@class='autocomplete-suggestions LocationAutocompleteSuggestions LocationAutocompleteSuggestions-s0-10-1-1-2-4-4-1-14[0]-3-0-1-6-pickupLocation X-LocationAutocompleteSuggestions X-LocationAutocompleteSuggestions-Pickup']").click()

        # move cursor down and hit enter
        pickup_location.send_keys(Keys.DOWN)
        pickup_location.send_keys(Keys.RETURN)
        # pickup_location.click()

    # submit the form ->
    submit_btn = driver.find_element_by_xpath(
        "//button[@class='btn btn-success btn-lg btn-block X-SearchButton'][@type = 'button']")
    # submit_btn = driver.find_element_by_class_name("btn btn-success btn-lg btn-block X-SearchButton")
    submit_btn.send_keys(Keys.RETURN)

    # driver.getPageSource()

    WebDriverWait(driver, 10).until(
        ec.presence_of_element_located((By.TAG_NAME, 'body')))
    # driver.implicitly_wait(10)
    # show details in listview

    opt = WebDriverWait(driver, 10).until(ec.visibility_of_element_located(
        (By.XPATH, "/html/body/div[1]/div[1]/div/div/div[1]/div/div[2]/div/div[1]/div/div/div/div[3]")))
    opt.click()

    # listOption = driver.find_element(
    #     By.XPATH, "/html/body/div[1]/div[1]/div/div/div[1]/div/div[2]/div/div[1]/div/div/div/div[3]")
    # listOption.click()

    # WebDriverWait(driver, 10).until(ec.visibility_of_element_located(
    #     (By.XPATH, "//div[@class='X-DnSelectionMenu-Popup dn-selection-menu--popup']")))

    # select_supplier(driver)


def select_supplier(driver):

    supplier_list = WebDriverWait(driver, 10).until(ec.visibility_of_element_located(
        (By.XPATH, "//div[@class='X-DnSelectionMenu dn-selection-menu supplier X-supplier']")))
    supplier_list.click()

    avis = driver.find_element(
        By.XPATH, "/html/body/div[1]/div[1]/div/div/div[1]/div/div[2]/div/div[2]/div/div/div/div[2]/span[4]/div/div[1]/table/tr[3]/td[1]/i[2]")
    avis.click()

    # budget = driver.find_element(
    #     By.XPATH, "//table[@class='dn-selection-table']/tr[@class='X-SelectionMenu-Button dn-selection-option selected X-Option X-Option-3']")
    # budget.click()


if __name__ == '__main__':
    # chrome_path = "D:/Softwares/webdriver/chromedriver"
    driver = webdriver.Firefox(
        executable_path="D:/Softwares/webdriver/geckodriver")
    try:
        fetch_data(driver, url)
    finally:
        # driver.quit()
        pass
