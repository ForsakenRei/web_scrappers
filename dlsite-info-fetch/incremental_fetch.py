import asyncio
import nest_asyncio
import pyppeteer

import pandas as pd
from bs4 import BeautifulSoup

import requests
import os
import re
from datetime import datetime

file_path = 'H:/ASMR/'

current_list = os.listdir(file_path)
current_list = pd.DataFrame(current_list)
current_list.columns = ['Path_Name']
current_list['RJ_Code'] = current_list['Path_Name'].str.extract('(RJ\d{6})', expand=False)

old_file = './archive_list_2022-02-20.xlsx'
old_list = pd.read_excel(old_file)

archive_list = current_list[~current_list['RJ_Code'].isin(old_list['RJ_Code'])].reset_index().drop(columns=['index'])

work_list = []

def dlsite_info_extract(archive_list):      
    URL = 'https://www.dlsite.com/maniax/work/=/product_id/{}.html'.format(search_code)
    page = requests.get(URL)
    page_source = BeautifulSoup(page.content, 'html.parser')

    try:
        work_name = page_source.find('h1', id = 'work_name').text
        work_genre = page_source.find('span', {'class' : re.compile('icon_(GEN|ADL|R15)')}).text
        brand_name = page_source.find('span', {'class' : 'maker_name'}).text.replace('\n', '')
        work_cv = page_source.find('th', text = '声優').find_next_sibling('td').text
        work_cv = re.sub(r'\s', '', work_cv)   
        release_date = page_source.find('th', text = '販売日').find_next_sibling('td').text.replace('年', '/').replace('月', '/').replace('日', '')
        work_tag = page_source.find('th', text = 'ジャンル').find_next_sibling('td').text.replace('\n', ' ').strip().replace(' ', ',')
    except:
        work_tag = 'None'

    work_info = {}

    work_info['RJ_Code'] = search_code
    work_info['Brand_Name'] = brand_name
    work_info['Work_Name'] = work_name
    work_info['Work_CV'] = work_cv
    work_info['Work_Genre'] = work_genre
    work_info['Work_Tag'] = work_tag
    work_info['Release_Date'] = release_date

    work_list.append(work_info)

error_list = []

for i in range(len(archive_list)):
    search_code = archive_list['RJ_Code'][i]
    try:
        dlsite_info_extract(archive_list)
    except:
        error_rj = archive_list['RJ_Code'][i]
        error_list.append(error_rj)

sale_list = []

async def main():
        browser = await pyppeteer.launch()
        page = await browser.newPage()
        page_path = 'https://www.dlsite.com/maniax/work/=/product_id/{}.html'.format(search_code)
        await page.goto(page_path)

        page_content = await page.content()
        page_source = BeautifulSoup(page_content, 'html.parser')
        try:
                sale_count = page_source.find('dt', text = '販売数：').find_next_sibling('dd').text.replace(',' , '')
        except:
                sale_count = page_source.find('dt', text = '総販売数：').find_next_sibling('dd').text.replace(',' , '')
        sale_info = {}

        sale_info['RJ_Code'] = search_code
        sale_info['Sale_Count'] = sale_count
        sale_list.append(sale_info)
        await browser.close()

nest_asyncio.apply()

for i in (range(len(archive_list))):    
    try:
        search_code = archive_list['RJ_Code'][i]
        asyncio.get_event_loop().run_until_complete(main())       
    except:
        pass

sale_list = pd.DataFrame(sale_list)
work_list = pd.DataFrame(work_list)
error_list = pd.DataFrame(error_list)

combine_list = pd.merge(archive_list, work_list, how = 'left', on = 'RJ_Code')
combine_list = pd.merge(combine_list, sale_list, how = 'left',on = 'RJ_Code')
combine_list['Sale_Count'] = [re.findall(r'[0-9]{1,10}', x) for x in combine_list['Sale_Count']]
combine_list['Sale_Count'] = [max(x) for x in combine_list['Sale_Count']]
combine_list['Sale_Count'] = combine_list['Sale_Count'].astype(int)

output_list = pd.concat([old_list, combine_list]).reset_index().drop(columns=['index'])

date_time = datetime.now().strftime('%Y-%m-%d')
writer = pd.ExcelWriter('archive_list_{}.xlsx'.format(date_time))
output_list.to_excel(writer, sheet_name = 'Sheet1')
error_list.to_excel(writer, sheet_name = 'Sheet2')
writer.save()