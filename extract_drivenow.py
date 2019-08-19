from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec

'''Web Scarping  for www.drivenow.com.au'''
'''
TASKS:
    1. We need to enter pick up date & time and drop off date & time for search window.
    2. Once we chose this and run macro, it should bring up www.drivenow.com.au website 
        and enter these dates and time in search window. Pick up location is always Adelaide Airport. 
        Then macro automatically click on search.
    3. It will bring up rates for 5 major car rentals, only those rental companies we need to know.

    4. At this stage, we need to scrap data from this webpage back into excel.

'''

chrome_path = "D:/Softwares/webdriver/chromedriver"
driver = webdriver.Firefox(
    executable_path="D:/Softwares/webdriver/geckodriver")

url = "https://www.drivenow.com.au"
driver.get(url)

# set pickup location to Adelaide Airport
pickup_location = driver.find_element_by_id(
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


# elem = driver.find_element_by_name("q")
# elem.clear()
# elem.send_keys("nepal")
# elem.send_keys(Keys.RETURN)


# driver.close()
