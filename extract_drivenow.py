from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.common.exceptions import TimeoutException, NoSuchElementException

from bs4 import BeautifulSoup
import pandas as pd

from Vehicle import Ride

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
            (By.XPATH,
             "//div[@class='autocomplete-suggestions LocationAutocompleteSuggestions LocationAutocompleteSuggestions-s0-10-1-1-2-4-4-1-14[0]-3-0-1-6-pickupLocation X-LocationAutocompleteSuggestions X-LocationAutocompleteSuggestions-Pickup']")))
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

    # list_option = WebDriverWait(driver, 10).until(ec.visibility_of_element_located(
    #     (By.XPATH, "/html/body/div[1]/div[1]/div/div/div[1]/div/div[2]/div/div[1]/div/div/div/div[3]")))
    # list_option.click()

    list_option = WebDriverWait(driver, 10).until(ec.visibility_of_element_located(
        (By.XPATH, "//*[@id='X-Page-Nitro-Content']/div/div/div[1]/div/div[2]/div/div[1]/div/div/div/div[3]")))
    list_option.click()

    # close button banner to check items correctly
    close_banner(driver)

    # filter supplier for efficiency
    select_supplier(driver)

    # find Rides
    find_ride(driver)


def select_supplier(driver):
    table_field = "//*[@id='X-Page-Nitro-Content']/div/div/div[1]/div/div[2]/div/div[2]/div/div/div/div[2]/span[4]/div/div[1]/table/"
    avis_field = "img[@src='/webdata/images/supplier/logo/lrg/avis.gif']"
    budegt_field = "img[@src='/webdata/images/supplier/logo/budget-140-40.jpg']"
    europcar_field = "img[@src='/webdata/images/supplier/logo/lrg/europcar.gif']"
    hertz_field = "img[@src='/webdata/images/supplier/logo/hertz.gif']"
    thrifty_field = "img[@src='/webdata/images/supplier/logo/lrg/thrifty.gif']"

    supplier_list = WebDriverWait(driver, 10).until(ec.visibility_of_element_located(
        (By.XPATH, "//div[@class='X-DnSelectionMenu dn-selection-menu supplier X-supplier']")))
    supplier_list.click()

    # check for budget
    # check_supplier(
    #     driver, "//span/img[@src='/webdata/images/supplier/logo/budget-140-40.jpg']")

    try:
        # check avis if present
        check_supplier(
            driver, "//span/" + avis_field)
    except TimeoutException:
        print("Timed Out for avis: ")

    try:
        # check budget if present
        check_supplier(
            driver, "//span/" + budegt_field)
    except TimeoutException:
        print("Timed Out for Budget: ")

    try:
        # check europcar if present
        check_supplier(
            driver, "//td[2]/span/" + europcar_field)
    except TimeoutException:
        print("Timed Out for Europcar: ")

    try:
        # check and mark hertz if present
        check_supplier(
            driver, "//td[2]/span/" + hertz_field)
    except TimeoutException:
        print("Timed Out for Hertz: ")

    try:
        # check if thrifty is present and mark if present
        check_supplier(driver, "//td[2]/span/" + thrifty_field)
    except TimeoutException:
        print("Timed Out for Thrifty: ")

    # too slow while running in loops
    # for i in range(1, 10):

    #     try:
    #         # check budget if present
    #         check_supplier(
    #             driver, table + "tr[" + str(i) + "]/td[2]/span/img[@src='/webdata/images/supplier/logo/budget-140-40.jpg']")
    #     except TimeoutException:
    #         print("Timed Out for Budget: " + str(i))

    #     try:
    #         # check europcar if present
    #         check_supplier(
    #             driver, table + "tr[" + str(i) + "]/td[2]/span/img[@src='/webdata/images/supplier/logo/lrg/europcar.gif']")

    #     except TimeoutException:
    #         print("Timed Out for Europcar: " + str(i))

    #     try:
    #         # check and mark hertz if present
    #         check_supplier(
    #             driver, table + "tr[" + str(i) + "]/td[2]/span/img[@src='/webdata/images/supplier/logo/hertz.gif']")
    #     except TimeoutException:
    #         print("Timed Out for Hertz: " + str(i))

    #     try:
    #         # check if thrifty is present and mark if present
    #         check_supplier(driver, table + "tr[" + str(
    #             i) + "]/td[2]/span/img[@src='/webdata/images/supplier/logo/lrg/thrifty.gif']")
    #     except TimeoutException:
    #         print("Timed Out for Thrifty: " + str(i))

    # thrifty = WebDriverWait(driver, 10).until(ec.visibility_of_element_located(
    #     (By.XPATH, table + "tr[9]/td[2]/span/img[@src='/webdata/images/supplier/logo/lrg/thrifty.gif']")))
    # if thrifty:
    #     driver.find_element(By.XPATH, table+ "tr[9]").click()


# to filter supplier


def check_supplier(driver, xpath):
    # check item if present
    item = WebDriverWait(driver, 2).until(ec.visibility_of_element_located(
        (By.XPATH, xpath)))

    if item:
        item.click()


def close_banner(driver):
    WebDriverWait(driver, 10).until(ec.visibility_of_element_located(
        (By.XPATH, "/html/body/div[1]/div[2]/div/div[1]"))).click()


def find_ride(driver):
    # rides = WebDriverWait(driver, 10).until(ec.visibility_of_element_located(
    #     (By.XPATH, "/html/body/div[1]/div[1]/div/div/div[1]/div/div[4]/div")))

    rides_field = WebDriverWait(driver, 10).until(ec.visibility_of_element_located(
        (By.XPATH, "//*[@id='X-Page-Nitro-Content']/div/div/div[1]/div/div[4]/div")))

    rides_list = rides_field.find_elements(
        By.XPATH, "//div/div[@class='car-result']")
    # print(rides_field.text)

    # soup = BeautifulSoup(driver.page_source, 'lxml')

    # all_rides = driver.find_elements(By.XPATH, "//*[@id='X-Page-Nitro-Content']/div/div/div[1]/div/div[4]/div")
    print("Total Rides: ", len(rides_list))

    data_path = "//*[@id='X-Page-Nitro-Content']/div/div/div[1]/div/div[4]/div/div/div/div["

    # create pandas file
    df = pd.DataFrame(columns=['Name', 'Supplier', 'Type', 'Price'])

    for i in range(1, len(rides_list) + 1):

        try:
            price = driver.find_element(
            By.XPATH, data_path + str(i) + "]/div/div/table[1]/tr/td[3]/div[1]")
            print("Price: " + str(i) + " : " + price.text)

        except NoSuchElementException:
            print("Element not found Price:" + str(i))

        try:
            ride_name = rides_field.find_element(
                By.XPATH, data_path + str(i) + "]/div/div/table[1]/tr/td[2]/div[1]/span")
        except NoSuchElementException:
            print("Element not found Name:" + str(i))

        try:
            supplier_img = rides_field.find_element(
                By.XPATH, data_path + str(i) + "]/div/div/table[2]/tr/td[1]/div/img")
            src = supplier_img.get_attribute('src')
            srcs = src.split("/")
        except NoSuchElementException:
            print("Element not found Supplier Img:" + str(i))

        try:
            vehicle_type = rides_field.find_element(By.XPATH, data_path + str(i) + "]/div/div/table[1]/tr/td[2]/div[3]")
        except NoSuchElementException:
            print("Element not found Vehicle Type:" + str(i))

        else:
            print(ride_name.text + " supplier: " +
            srcs[-1] + " vehicle Type: " + vehicle_type.text)

            ride = Ride(ride_name.text,
                srcs[-1], vehicle_type=vehicle_type.text, price=price.text)
            
            data = {'Name': ride_name.text, 'Supplier': srcs[-1], 'Type':vehicle_type.text, 'Price': price.text}
            # new_data = pd.Series(data)
            df = df.append(pd.Series(data), ignore_index=True)
            
        print(ride)
    df.to_csv('data.csv')

    print(df)

if __name__ == '__main__':
    # chrome_path = "D:/Softwares/webdriver/chromedriver"
    driver = webdriver.Firefox(
        executable_path="D:/Softwares/webdriver/geckodriver")
    try:
        fetch_data(driver, url)
    finally:
        # driver.quit()
        pass
