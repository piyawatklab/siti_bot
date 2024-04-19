from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

from seleniumbase import Driver
import time
from logging import ERROR

import os
import glob
import shutil
import sys

import pandas as pd

from libs.settings import USERNAME_LOGIN , PASSWORD_LOGIN , SHEET_LINK

PWD = os.getcwd()

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
prefs = {"profile.default_content_settings.popups": 0,
            "download.default_directory": f'{PWD}/downloaded_files/',
            "directory_upgrade": True}
options.add_experimental_option("prefs",prefs)

data = []

def run(name):

    delete_order()

    order_list = search_order()
    
    if len(order_list) > 0 :
    
        # driver = Driver(uc=True)
        driver = Driver(uc=True,headless=True)
        driver.set_window_size(1920, 1080)

        url = 'https://sitirb.lyreco.com/sitiweb/FH1/dispatch.do?language=TH&country=TH'
        driver.get(url)
        time.sleep(3)

        driver.find_element(By.XPATH, "//input[@name='login']").send_keys(USERNAME_LOGIN)
        driver.find_element(By.XPATH, "//input[@name='password']").send_keys(PASSWORD_LOGIN)
        driver.find_element(By.XPATH, "//button[@id='submit']").click()
        time.sleep(3)
        driver.find_element(By.XPATH, "//button[@id='submitForm']").click()
        time.sleep(3)
        
        original_url = driver.current_url
        # print(original_url)

        start_index = original_url.find("/overview")

        for line in order_list:
            try:

                new_url = original_url[:start_index] + "/delivery/getDeliveryBySapId/" + line

                # print(new_url)
                driver.get(new_url)
                time.sleep(3)
                search = driver.find_element(By.XPATH, ".//tbody")
                row_search = search.find_elements(By.XPATH, ".//tr")
                # len(row_search)
                if len(row_search) > 0:
                    for row_line in row_search:
                        td_line = row_line.find_element(By.XPATH, ".//td[@class='input-sm']")
                        td_line.find_element(By.XPATH, ".//a").click()
                        time.sleep(3)
                        a_element = driver.find_elements(By.XPATH, ".//a[@class='blocking-download']")
                        if len(a_element) > 0 :
                            for download_pdf in a_element:
                                download_pdf.click()
                                time.sleep(3)
                        # time.sleep(3)
                        driver.find_element(By.XPATH, ".//button[@class='close']").click()
                        time.sleep(3)
                print(line,'Done')

            except Exception as e:
                time.sleep(1)
                print(line,'Error')

        driver.quit()
        move_file()

def search_order():

    df = pd.read_excel('downloaded_files/order.xlsx')

    data_array = df.values.tolist()
    data_array = [str(item[0]) for item in data_array]

    print(data_array)

    return data_array

def delete_order():

    files = glob.glob(f'downloaded_files/*pdf')
    for f in files:
        os.remove(f)

def move_file():
    try : 
        files = glob.glob(f'downloaded_files/*pdf')

        for source_file in files:
            file_name = os.path.basename(source_file)
            print(file_name)
            new_path = 'G:\\My Drive\\test_bot'
            # os.rename(source_file, f'{new_path}\\{file_name}')
            # shutil.move(source_file, f'{new_path}\\{file_name}')
            shutil.copy(source_file, f'{new_path}\\{file_name}')
            
    except Exception as e:
        time.sleep(1)
        print('move_file Error')

if __name__ == "__main__":

    name = 'default'
    if len(sys.argv) > 1:
        for arg in sys.argv[1:]:
            key, value = arg.split("=")
            if key == "name":
                name = value
    
    print("Name :", name)
    run(name)
    
    