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

data = os.path.join(os.getcwd(), 'vse.instr')
data_folder = os.path.join(os.getcwd(), 'result_data')
user_data = os.path.join(os.getcwd(), 'user_data')
drivers_dict = dict()

if not os.path.exists(data):
    os.mkdir(data)

### a = {
    # 'Instrum': {
    #    'Akk': {}
    # },
    # 
# }

def get_data():
    options = uc.ChromeOptions()
    options.add_argument('--headless')

    driver = uc.Chrome(options=options)
    result_dict = {
        '1st': '',
        '2nd': '',
        '3d': '',
        '4th': '',
    }
    first_cat = set()
    second_cat = set()
    third_cat = set()
    fourth_cat = set()

    result_tree = [
        'https://www.vseinstrumenti.ru/category/stroitelnyj-instrument-6474/',
        'https://www.vseinstrumenti.ru/category/elektrika-i-svet-6480/',
        'https://www.vseinstrumenti.ru/category/santehnicheskoe-oborudovanie-6750/',
        'https://www.vseinstrumenti.ru/category/ruchnoj-instrument-6481/',
        'https://www.vseinstrumenti.ru/category/oborudovanie-dlya-avtoservisa-i-garazha-6479/',
        'https://www.vseinstrumenti.ru/category/avtohimiya-4827/',
        'https://www.vseinstrumenti.ru/category/avtoaksessuary-5493/',
        'https://www.vseinstrumenti.ru/category/sadovaya-tehnika-i-instrument-6473/',
        'https://www.vseinstrumenti.ru/category/krepezh-2994/',
        'https://www.vseinstrumenti.ru/category/skobyanye-izdeliya-6661/',
        'https://www.vseinstrumenti.ru/category/metizy-170301/',
        'https://www.vseinstrumenti.ru/category/stroitelnye-materialy-12778/',
        'https://www.vseinstrumenti.ru/category/otdelochnye-materialy-12831/',
        'https://www.vseinstrumenti.ru/category/tovary-dlya-ofisa-i-doma-169730/',
        'https://www.vseinstrumenti.ru/category/tovary-dlya-doma-3788/',
        'https://www.vseinstrumenti.ru/category/snabzhenie-i-osnaschenie-ofisa-169810/',
        'https://www.vseinstrumenti.ru/category/sport-i-turizm-3351/',
        'https://www.vseinstrumenti.ru/category/stanki-6476/',
        'https://www.vseinstrumenti.ru/category/prisposobleniya-i-osnastka-dlya-stankov-4035/',
        'https://www.vseinstrumenti.ru/category/promyshlennye-komponenty-170819/',
        'https://www.vseinstrumenti.ru/category/klimaticheskoe-oborudovanie-6472/',
        'https://www.vseinstrumenti.ru/category/otopitelnoe-oborudovanie-12522/',
        'https://www.vseinstrumenti.ru/category/skladskoe-oborudovanie-i-tehnika-dlya-sklada-4100/',
        'https://www.vseinstrumenti.ru/category/oborudovanie-dlya-klininga-i-uborki-2976/',
        'https://www.vseinstrumenti.ru/category/stroitelnoe-oborudovanie-i-tehnika-6477/',
        'https://www.vseinstrumenti.ru/category/rashodnye-materialy-i-osnastka-6478/',
        'https://www.vseinstrumenti.ru/category/spetsodezhda-i-siz-4661/'
    ]
    with driver:
        for x in range(1, 28):
            try:
                with open(os.path.join(data, f'vse_instr_{x}.html'), 'r') as file:
                    soup = BeautifulSoup(file, 'html.parser')
                    tree_wrapper = soup.find('div', class_='tree-wrapper')
                    node = tree_wrapper.find_all('div', class_='node-wrapper')
                    count = len(node)
                    count_to = 0
                    for name_node in node:
                        count_to += 1
                        print(f'\n{count_to} / {count}\n')
                        name = name_node.find('div', class_='current-node -large').find('a').get('href')
                        url = f'https://www.vseinstrumenti.ru{name}'
                        print(url)
                        driver.get(url)
                        time.sleep(2)
                        try: 
                            driver.find_element(By.XPATH, '//div[@class="USFwIs"]')
                            new_new_soup = BeautifulSoup(driver.page_source, 'html.parser')
                            an_cats = new_new_soup.find('div', class_='USFwIs').find_all('div', class_='_5uKBsp xpfZ5m CQQqRL')
                            if an_cats:
                                for an_cat in an_cats:
                                    url_an_cat = an_cat.find('a').get('href')
                                    an_url = f'https://www.vseinstrumenti.ru{url_an_cat}'
                                    driver.get(an_url)
                                    new_soup = BeautifulSoup(driver.page_source, 'html.parser')
                                    bread_crumbs = new_soup.find('nav', attrs={'id': 'breadcrumbs-anchor'}).text.strip().split('/')
                                    print(f'NEW {an_url} //// {bread_crumbs}')
                                    if len(bread_crumbs) <= 3:
                                        continue
                                    elif len(bread_crumbs) <= 4:
                                        result_dict['1st'] = bread_crumbs[1].strip()
                                        result_dict['2nd'] = bread_crumbs[2].strip()
                                        result_dict['3d'] = bread_crumbs[3].strip()
                                    else:
                                        result_dict['1st'] = bread_crumbs[1].strip()
                                        result_dict['2nd'] = bread_crumbs[2].strip()
                                        result_dict['3d'] = bread_crumbs[3].strip()
                                        result_dict['4th'] = bread_crumbs[4].strip()
                            else:
                                an_cats = new_new_soup.find('div', class_='USFwIs').find_all('div', class_='z4xhEa xpfZ5m sW8Ck5 CQQqRL raAPf+')
                                for an_cat in an_cats:
                                    url_an_cat = an_cat.find('a').get('href')
                                    an_url = f'https://www.vseinstrumenti.ru{url_an_cat}'
                                    driver.get(an_url)
                                    new_soup = BeautifulSoup(driver.page_source, 'html.parser')
                                    bread_crumbs = new_soup.find('nav', attrs={'id': 'breadcrumbs-anchor'}).text.strip().split('/')
                                    print(f'NEW {an_url} //// {bread_crumbs}')
                                    if len(bread_crumbs) <= 3:
                                        continue
                                    elif len(bread_crumbs) <= 4:
                                        result_dict['1st'] = bread_crumbs[1].strip()
                                        result_dict['2nd'] = bread_crumbs[2].strip()
                                        result_dict['3d'] = bread_crumbs[3].strip()
                                        result_dict['4th'] = ''
                                    else:
                                        result_dict['1st'] = bread_crumbs[1].strip()
                                        result_dict['2nd'] = bread_crumbs[2].strip()
                                        result_dict['3d'] = bread_crumbs[3].strip()
                                        result_dict['4th'] = bread_crumbs[4].strip()
                                    to_excel(result_dict)
                        except Exception as ex:
                            print(f'\t[-] {ex}')
                            # continue
                        new_soup = BeautifulSoup(driver.page_source, 'html.parser')
                        bread_crumbs = new_soup.find('nav', attrs={'id': 'breadcrumbs-anchor'}).text.strip().split('/')
                        print(f'OLD {url} //// {bread_crumbs}')
                        if len(bread_crumbs) <= 3:
                            continue
                        elif len(bread_crumbs) <= 4:
                            result_dict['1st'] = bread_crumbs[1].strip()
                            result_dict['2nd'] = bread_crumbs[2].strip()
                            result_dict['3d'] = bread_crumbs[3].strip()
                            result_dict['4th'] = ''
                        else:
                            result_dict['1st'] = bread_crumbs[1].strip()
                            result_dict['2nd'] = bread_crumbs[2].strip()
                            result_dict['3d'] = bread_crumbs[3].strip()
                            result_dict['4th'] = bread_crumbs[4].strip()
                        count -= 1
                    
                        to_excel(result_dict)
                
            except Exception as ex:
                print(f'\t[-] {ex}')

def to_excel(profile):
    table_name = "tree"
    result_table = os.path.join(data_folder, f'{table_name}.xlsx')
    name_of_sheet = "vseinstr"

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
