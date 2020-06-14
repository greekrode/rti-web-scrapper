import re
import xlwt
import time
from multiprocessing import Process
from xlwt import Workbook
from selenium import webdriver

def write_to_sheet(sheet, row, stock_name, eps_value, bvps_value, per_value, pbv_value, current_price):
    sheet.write(row, 0, stock_name)
    sheet.write(row, 1, eps_value)
    sheet.write(row, 2, bvps_value)
    sheet.write(row, 3, per_value)
    sheet.write(row, 4, pbv_value)
    sheet.write(row, 5, current_price)

def init_driver():
    driver = webdriver.Chrome()
    driver.get('https://analytics.rti.co.id/?m_id=1&sub_m=s3&sub_sub_m=3')
    return driver

def back_to_home(driver):
    driver.back()
    driver.back()
def get_per_value(per_dirty_value):
    per_split = re.search('(\d+.\d+)', per_dirty_value)

    if per_split:
        per_value = per_split.group(0)
    return per_value

def get_bvps_value(bvps_dirty_value):
    bvps_split = re.search(': (-?\d*\.{0,1}\d+$)', bvps_dirty_value)
    if bvps_split:
        bvps_value = bvps_split.group(1)
    else:
        bvps_split = re.search(': \S+ (-?\d*\.{0,1}\d+$)', bvps_dirty_value)
        bvps_value = float(bvps_split.group(1)) * 14500
    return bvps_value

def get_pbv_value(pbv_dirty_value):
    pbv_split = re.search('(\d+.\d+)', pbv_dirty_value)

    if pbv_split:
        pbv_value = pbv_split.group(0)
    return pbv_value

def runInParallel(*fns):
      proc = []
      for fn in fns:
        p = Process(target=fn)
        p.start()
        proc.append(p)
      for p in proc:
            p.join()

def iterate(start, end, filename):
    driver = init_driver()

    wb = Workbook()
    sheet = wb.add_sheet('Sheet 1')

    for x in range(start, end):
        x_in_str = str(x+1)
        stock = driver.find_element_by_xpath('//*[@id="chgPctTable"]/tbody/tr['+x_in_str+']/td[1]')
        stock_name = stock.text
        current_price = driver.find_element_by_xpath('//*[@id="chgPctTable"]/tbody/tr['+x_in_str+']/td[3]').text
        stock.click()

        stock_profile = driver.find_element_by_xpath('//*[@id="ssm4"]')
        stock_profile.click()

        # Check that profiles data exist
        try:
            driver.find_element_by_xpath('//*[@id="dtable"]/tbody/tr[3]/td/form/table[3]')
        except:
            back_to_home(driver)
            continue

        eps_value = driver.find_element_by_xpath('//*[@id="dtable"]/tbody/tr[3]/td/form/table[2]/tbody/tr[3]/td/table/tbody/tr[2]/td[14]/b').text

        per_dirty_value = driver.find_element_by_xpath('//*[@id="dtable"]/tbody/tr[3]/td/form/table[2]/tbody/tr[3]/td/table/tbody/tr[2]/td[15]/b').text
        per_value = get_per_value(per_dirty_value)

        bvps_dirty_value = driver.find_element_by_xpath('//*[@id="dtable"]/tbody/tr[3]/td/form/table[3]/tbody/tr/td[2]/table/tbody/tr[4]/td/table/tbody/tr[2]/td[1]/table/tbody/tr[4]/td[2]/b').text
        bvps_value = get_bvps_value(bvps_dirty_value)

        pbv_dirty_value = driver.find_element_by_xpath('//*[@id="dtable"]/tbody/tr[3]/td/form/table[3]/tbody/tr/td[2]/table/tbody/tr[4]/td/table/tbody/tr[2]/td[2]/table/tbody/tr[4]/td[2]/b').text
        pbv_value = get_pbv_value(pbv_dirty_value)

        write_to_sheet(sheet, x, stock_name, eps_value, bvps_value, per_value, pbv_value, current_price)
        back_to_home(driver)

    wb.save(filename+'.xls')


if __name__ == "__main__":
    #  runInParallel(iterate2, iterate3, iterate4)
    iterate(0, 30, 'scrappedmsci')
