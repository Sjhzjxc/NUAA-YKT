#!/usr/bin/python
# -*- coding: utf-8 -*-
from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.support.ui import Select
import xlwt

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--disable-gpu')
driver = webdriver.Chrome(chrome_options=chrome_options)

username = ''  # your student Id
password = ''  # your password

xf = []
cz = []


def login():
    default_url = 'http://ykt.nuaa.edu.cn'
    driver.get(default_url)
    name = driver.find_element_by_id('TextBox1')
    name.send_keys(username)
    pwd = driver.find_element_by_id('TextBox2')
    pwd.send_keys(password)
    login_btn = driver.find_element_by_id('login')
    login_btn.click()


def save():
    soup = BeautifulSoup(driver.page_source, 'lxml')
    consume = soup.find('table', {'id': 'dlconsume'})
    if consume is not None:
        trs = consume.find_all('tr')
        for tr in trs:
            if (tr.has_attr('style') and 'color:#333333' not in tr['style'] and 'color:#284775' not in tr['style']) or \
                    not tr.has_attr('style'):
                continue
            item = []
            plus = False
            for td in tr.find_all('td', {'align': 'center'}):
                if '充值' in td.string:
                    plus = True
                item.append(td.string.replace(' ', '').replace('\r\n', '').replace('\t', '').replace('\n', ''))
            if item.__len__() == 0:
                continue
            if plus:
                cz.append(tuple(item))
            else:
                xf.append(tuple(item))


def get_consume():
    driver.find_element_by_xpath('//*[@id="Table4"]/tbody/tr[3]/td/a').click()
    Select(driver.find_element_by_id('DropDownList1')).select_by_index(3)
    driver.find_element_by_id('Button1').click()
    save()
    count = int(driver.find_element_by_id('Label6').text)
    print(count)
    i = 2
    page_num = 1
    while True:
        count = int(driver.find_element_by_id('Label6').text)
        line = 12
        total_page_num = count // 10 + 1
        z_page_num = (total_page_num // 10 + 1) * 10
        if page_num >= total_page_num:
            line = count % 10 + 2
        if page_num == total_page_num // 10 * 10 + 1:
            i = z_page_num - total_page_num + 4
        if page_num == total_page_num:
            break
        href = driver.find_element_by_xpath(
            '//*[@id="dlconsume"]/tbody/tr[' + str(line) + ']/td/table/tbody/tr/td[' + str(i) + ']/a'
        )
        if href.text == '...':
            i = 3
        href.click()
        save()
        i += 1
        page_num += 1


def save_to_xls(path):
    workbook = xlwt.Workbook(encoding='utf-8')
    book_sheet1 = workbook.add_sheet('消费', cell_overwrite_ok=True)
    book_sheet2 = workbook.add_sheet('充值', cell_overwrite_ok=True)
    for i, row in enumerate(tuple(xf)):
        for j, col in enumerate(row):
            book_sheet1.write(i, j, col)
    for i, row in enumerate(tuple(cz)):
        for j, col in enumerate(row):
            book_sheet2.write(i, j, col)
    workbook.save(path)


if __name__ == '__main__':
    login()
    get_consume()
    save_to_xls('grade.xls')
    driver.close()
