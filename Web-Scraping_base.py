from bs4 import BeautifulSoup
import re
import numpy as np
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
from datetime import datetime

def get_data(url):
    """
    Function to retrieve the html data nested in the website given an url
    :param url: An internet url pointing to a website from which we wish to get data
    :return: A BeautifulSoup object containing the parsed html of the url
    """
    # We open an instance of chrome to search the url
    driver = webdriver.Chrome()
    driver.get(url)

    # To avoid a timeout we pause the code for 4 seconds (we repeat this often)
    time.sleep(3)

    # In order to access the data within the website we need to either accept or reject the cookies
    button = driver.find_element(By.XPATH,'//*[@id="onetrust-reject-all-handler"]')
    button.click()
    time.sleep(3)
    del button
    # During the testing phase it occurred few times that a popup appeared, preventing us from obtaining the data
    # Since it did not happen every time it was better to wrap it with an error handler
    try:
        # We wait until the popup appears and the close button becomes clickable
        close_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CLASS_NAME, 'popupCloseIcon'))
        )
        close_button.click()
    except:
        print("Popup not found or already closed.")
    # We switch the time frame for the economic calendar from 'Today' to 'This Week'
    button = driver.find_element(By.XPATH, '//*[@id="timeFrame_thisWeek"]')
    button.click()
    time.sleep(3)

    # To retrieve all the data from the economic calendar we have to scroll to the very bottom of the page
    scroll_to_bottom(driver)
    time.sleep(3)

    # We get the html file from the webpage using BeautifulSoup
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    # We close the chrome instance
    driver.quit()
    return soup

def scroll_to_bottom(driver):
    """
    A function designed to scroll down to the bottom of the page (it ensures that all the data is loaded)
    :param driver: It contains the current instance of the driver with the web page opened
    """

    bottom_height = driver.execute_script("return document.body.scrollHeight")
    stop = False
    while not stop:
        # Scroll down to the bottom
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        # Stop the process to let the data load
        time.sleep(2)
        # Calculate new bottom height and compare with previous one
        new_bottom_height = driver.execute_script("return document.body.scrollHeight")
        if new_bottom_height == bottom_height:
            stop = True
        bottom_height = new_bottom_height



def get_eco_calendar(url, soup):
    """
    The function returns all the economic calendar (including the area, impact, date, name and explanation of the
    indicator as well as the estimate, the previous value and the release) for the current week
    :param url: The url will be helpful to build the url for the web page of eah indicator
    :param soup: The BeautifulSoup object containing the html from the economic calendar web page
    :return: returns a dataframe containing the events and relevant related information
    """
    # Finding all the tables from the web page/soup object (the eco calendar is contained in a table)
    tables_list = soup.findAll('table')
    # We locate the economic calendar among all the data table founds
    for i in range(len(tables_list)):
        try:
            if tables_list[i]['id'] == 'economicCalendarData':
                eco_table = tables_list[i]
                break
        except:
            pass
    trs = eco_table.find("tbody").findAll("tr")
    caracs = []
    event_types = ['actual', 'forecast', 'previous']
    for tr in trs:
        # Handling errors is mandatory since the data does not only consists of indicators/events
        try:
            #print(tr['id'])
            # We store the id of the event which will help to retrieve data within the event
            event_id = tr['id'].split("_")[1]
            # Date of release of the economic data
            time_data = tr['data-event-datetime']
            # Country or geographic area of targeted by the data
            area = tr.find("td", {'class': 'left flagCur noWrap'}).find('span')['title']
            # Name of the indicator
            elem_name = tr.find("td", {'class': 'left event'}).find('a').text[1:]
            # Volatility/impact related to the release of the indicator
            elem_vol = tr.find("td", {'class': 'left textNum sentiment noWrap'}).findAll('i')
            vol = sum([1 if elem_vol[i]["class"][0] == elem_vol[0]["class"][0] else 0 for i in range(len(elem_vol))])
            events = []
            for i in range(len(event_types)):
                x = tr.find("td", {'class': re.compile('.*event-' + event_id + '-' + event_types[i] + '.*')}).text
                x = np.nan if x == '\xa0' else x
                events.append(x)
            # The link to the website explaining the indicator
            link = url + tr.find("td", {'class': 'left event'}).find('a')["href"]
            caracs.append([time_data, area, vol, elem_name, events[0], events[1], events[2], link])
        except:
            pass
    return pd.DataFrame(caracs,
                        columns=["Date", "Area", "Impact", "Indicator", "Actual", "Forecast", "Previous",
                                 "Link"])


if __name__ == "__main__":
    url = "https://uk.investing.com/economic-calendar/"
    data = get_data(url)
    df = get_eco_calendar(url.rsplit('/', 2)[0], data)
    df.to_excel("Economic_Calendar_"+ str(datetime.today().strftime('%d_%m_%Y'))+".xlsx", sheet_name='ECOD',
                index=False)
