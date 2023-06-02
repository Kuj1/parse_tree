import os
import json
import time
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver import Keys

import undetected_chromedriver as uc
from bs4 import BeautifulSoup

import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

UA = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) '\
                 'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.0.0 Safari/537.36'

data = os.path.join(os.getcwd(), 'kuvalda')
data_folder = os.path.join(os.getcwd(), 'result_data_kuvalda')
# user_data = os.path.join(os.getcwd(), 'user_data')
# drivers_dict = dict()

if not os.path.exists(data):
    os.mkdir(data)


def get_data():
    options = uc.ChromeOptions()
    # options.add_argument('--headless')

    driver = uc.Chrome(options=options)
    result_dict = {
        '1st': '',
        '2nd': '',
        '3d': '',
        '4th': '',
    }
    fourth_layer = list()

    with driver:
        driver.get('https://www.kuvalda.ru/')
        driver.find_element(By.CLASS_NAME, 'main-header__catalog').click()
        # menu_1_items = driver.find_elements(By.CLASS_NAME, 'menu__1')
        # for item_1 in menu_1_items:
            # count += 1
        # item_1.click()
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        menu_1_items = soup.find_all('div', class_='menu__1')
        for item in menu_1_items:
            menu_1_title = item.find('div', class_='menu__1-title').text
            # print(menu_1_title.strip())
            result_dict['1st'] = menu_1_title.strip()
            
            menu_2_items = item.find_all('div', class_='menu__2')
            for menu_item in menu_2_items:
                title = menu_item.find('a', class_='menu__2-title').text
                result_dict['2nd'] = title.strip()
                menu_3_items = menu_item.find_all('a', class_='menu__3 link')
                for menu_3_item in menu_3_items:
                    # print(menu_3_item.get('href'))
                    result_dict['3d'] = menu_3_item.text.strip()
                    url = f'https://www.kuvalda.ru{menu_3_item.get("href")}'
                    print(url)
                    driver.get(url)
                    new_soup = BeautifulSoup(driver.page_source, 'html.parser')
                    # promo_group = new_soup.find('div', class_='promo-groups__list')
                    if new_soup.find('div', class_='promo-groups__list'):
                        promo_group = new_soup.find('div', class_='promo-groups__list')
                        promo_items = promo_group.find_all('a')
                        result_dict['1st'] = menu_1_title.strip()
                        result_dict['2nd'] = title.strip()
                        result_dict['3d'] = menu_3_item.text.strip()
                        for promo_item in promo_items:
                            # fourth_layer.append(promo_item.get('title').strip())
                            result_dict['4th'] = promo_item.get('title').strip()
                            print(result_dict)
                            to_excel(result_dict)
                        fourth_layer.clear()
                    elif new_soup.find('div', class_='catalog-list__container'):
                        catalog_list = new_soup.find('div', class_='catalog-list__container').find_all('a')
                        result_dict['1st'] = menu_1_title.strip()
                        result_dict['2nd'] = title.strip()
                        result_dict['3d'] = menu_3_item.text.strip()
                        for catalog_item in catalog_list:
                            each_item_catalog = catalog_item.find('div', 'catalog-snippet__title').text
                            # fourth_layer.append(each_item_catalog.strip())
                            result_dict['4th'] = each_item_catalog.strip()
                            print(result_dict)
                            to_excel(result_dict)
                        # fourth_layer.clear()
                    else:
                        result_dict['4th'] = ''
                        print(result_dict)
                        to_excel(result_dict)
                        fourth_layer.clear()
                        continue
            # print(result_dict)

    

def to_excel(profile):
    table_name = "tree"
    result_table = os.path.join(data_folder, f'{table_name}.xlsx')
    name_of_sheet = "kuvalda"

    df = pd.DataFrame.from_dict(profile, orient='index')
    df = df.transpose()

    if os.path.isfile(result_table):
        workbook = openpyxl.load_workbook(result_table)
        sheet = workbook[f'{name_of_sheet}']

        for row in dataframe_to_rows(df, header=False, index=False):
            sheet.append(row)
        workbook.save(result_table)
        workbook.close()
    else:
        with pd.ExcelWriter(path=result_table, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name=f'{name_of_sheet}')


def main():
    get_data()


if __name__ == '__main__':
    main()
