#!/usr/bin/env python
#                    _ooOoo_
#                   o8888888o
#                   88" . "88
#                   (| -_- |)
#                   O\  =  /O
#                ____/`---'\____
#              .'  \\|     |  `.
#             /  \\|||  :  |||  \
#            /  _||||| -:- |||||-  \
#            |   | \\\  -  / |   |
#            | \_|  ''\---/''  |   |
#            \  .-\__  `-`  ___/-. /
#          ___`. .'  /--.--\  `. . __
#       ."" '<  `.___\_<|>_/___.'  >'"".
#      | | :  `- \`.;`\ _ /`;.`/ - ` : | |
#      \  \ `-.   \_ __\ /__ _/   .-` /  /
# ======`-.____`-.___\_____/___.-`____.-'======
#                    `=---='

# ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
#          God Bless          No bugs
#           Author:             kbl

import os
import sys
from os.path import abspath

import requests
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

from lxml import html

import win32com.client
import openpyxl

from src import send_mail

###################################################
#################### CONSTANTS ####################
###################################################
###################################################

# EMAIL_ADDR_FROM = ""
# EMAIL_ADDR_PASSWORD = ""
# EMAIL_ADDR_TO = "test@ya.ru"

URL_PATH = "https://yandex.ru/"
EURO_LINK_XPATH = '//*[@id="wd-_topnews"]/div/div[3]/div/div/div[2]/a'
DOLLAR_LINK_XPATH = '//*[@id="wd-_topnews"]/div/div[3]/div/div/div[1]/a'
TABLE_XPATH = '//*[@id="neo-page"]/div/div[2]/div[1]/div[1]/div[1]/div[2]/div/div[2]/div/div[2]'

BASE_PATH = os.path.split(abspath(__file__))[0]
RESULT_PATH = os.path.join(BASE_PATH,"result.xlsx")

###################################################
###################### METHODS ####################
###################################################
###################################################

def get_link(xpath_link):
    page = requests.get(URL_PATH, verify=False)
    _element = None
    _link = None
    if page.status_code == 200:
        tree = html.fromstring(page.content)
        tmp_element = tree.xpath(xpath_link)
        if tmp_element:
            _element = tmp_element[0]

        if _element is not None:
            _element_values = _element.values()
            if len(_element_values) > 1:
                _link = _element_values[1]
    return _link

def get_data(link):
    page = requests.get(link, verify=False)
    _data_element = None
    _data_elements_children = None
    result_data = list()
    if page.status_code == 200:
        tree = html.fromstring(page.content)
        tmp_element = tree.xpath(TABLE_XPATH)
        if tmp_element:
            _data_element = tmp_element[0]
        if(_data_element is not None):
            _data_elements_children = _data_element.getchildren()
        if _data_elements_children is not None:
            for elements in _data_element.getchildren():
                tmp_elem_data = elements.getchildren()
                if len(tmp_elem_data) != 3:
                    print("Error parse :( Кажется, я Вам не подхожу, не сработало (")
                    return []
                #result_data.append([a.text.replace("'","") for a in tmp_elem_data])
                try:
                    result_data.append([tmp_elem_data[0].text, float(tmp_elem_data[1].text.replace("'","").replace(",",".")), float(tmp_elem_data[2].text.replace("'","").replace(",","."))])
                except ValueError:
                    result_data.append([a.text for a in tmp_elem_data])
    return result_data

def create_xlsx(dollar_data, euro_data, xlsx_path):
    xlsx_data = list()
    for row_index in range(len(dollar_data)):
        xlsx_data.append([*dollar_data[row_index], *euro_data[row_index]])
    result_flag = True
    try:
        Excel = win32com.client.Dispatch("Excel.Application")
        wb = Excel.Workbooks.Add()
        sheet = wb.ActiveSheet

        #заполняем лист
        for data in range(len(xlsx_data)):
            for row_data in range(len(xlsx_data[data])):
                sheet.Cells(data+1, row_data+1).Value = xlsx_data[data][row_data]

        # добавляем формулу
        sheet.Cells(1,"G").Value = "Соотношение EUR/USD"
        for i in range(2,12):
            sheet.Cells(i,"G").Value =  '=(E{0}/B{0})'.format(i)

        #автоширина
        range1 = sheet.Range("A1:G11")
        range1.Columns("A:G").EntireColumn.AutoFit()

        #изменим на рубли
        sheet.Columns("B").EntireColumn.NumberFormatLocal = '#,##₽'
        sheet.Columns("E").EntireColumn.NumberFormatLocal = '#,##₽'

        #check autosum
        try:
            sheet.Cells(1,"H").Value = '=СУММ(B:B)'
            value = int(sheet.Cells(1,"H").Value)
            sheet.Cells(1,"H").Value = ""
        except Exception as error_check_autosum:
            print("Autosum is bad. Defines not as numbers")
            print(error_check_autosum)

    except Exception as error_create_xlsx:
        print("Error in create_xlsx method")
        print(error_create_xlsx)
        print("Save and exiting")
        result_flag = False
    finally:
        wb.SaveAs(xlsx_path)
        wb.Close(SaveChanges=True)
        Excel.Application.Quit()
    return result_flag

def create_message_text(line_count):
    end_word = ['.','а.','и.','и.','и.','.','.','.','.','.']
    base_template = "Сгенерирован excel файл, размером {0} строк{1}"
    result_string = ""
    #заканчивается на 1
    if line_count % 10 == 1:
        #если не заканчивается на 11
        if(str(line_count)[-2:] != '11'):
            result_string = base_template.format(line_count, end_word[1])
        else:
            result_string = base_template.format(str(line_count), end_word[0])
    #заканчивается на 2,3,4
    elif line_count % 10 == 2 or line_count % 10 == 3 or line_count % 10 == 4:
        #не заканчивается на 12,13,14
        part_end = str(line_count)[-2:]
        if part_end != '12' or part_end != '13' or part_end != '14':
            result_string = base_template.format(line_count, end_word[line_count % 10])
        else:
            result_string = base_template.format(line_count, end_word[0])
    else:
        result_string = base_template.format(line_count, end_word[0])
    return result_string

###################################################
################### START SCRIPT ##################
###################################################
###################################################

if __name__ == "__main__":

    if(len(sys.argv) != 4):
        print("Usage {0} email_from email_password email_to")
        exit(0)
    else:
        EMAIL_ADDR_FROM = sys.argv[1]
        EMAIL_ADDR_PASSWORD = sys.argv[2]
        EMAIL_ADDR_TO = sys.argv[3]


    try:
        os.unlink(RESULT_PATH)
    except FileNotFoundError:
        pass

    dollar_link = get_link(DOLLAR_LINK_XPATH)
    dollar_data = []
    bad_flag = False
    try:
        if dollar_link is not None:
            dollar_data = get_data(dollar_link)
            if not dollar_data:
                #take two ;)
                dollar_data = get_data(dollar_link)
                if not dollar_data:
                    raise OSError
            for el_index in range(len(dollar_data[0])):
                dollar_data[0][el_index]+= " (USD)"
        else:
            raise OSError
    except OSError:
        bad_flag = True


    euro_link = get_link(EURO_LINK_XPATH)
    euro_data = []
    try:
        if euro_link is not None:
            euro_data = get_data(euro_link)
            if not euro_data:
                #take two ;)
                euro_data = get_data(euro_link)
                if not euro_data:
                    raise OSError
            for el_index in range(len(euro_data[0])):
                euro_data[0][el_index]+= " (EUR)"
        else:
            raise OSError
    except OSError:
        bad_flag = True

    if bad_flag or len(dollar_data) != len(euro_data):
        print("Прошу прощения за затраченное время")
        print("Что-то не работает, но правда работало :( Можете попробовать запустить ещё раз, дело может быть в сети")
        exit(0)

    if(create_xlsx(dollar_data, euro_data, RESULT_PATH)):
        message_text = create_message_text(len(dollar_data))
        send_mail.send_email(EMAIL_ADDR_FROM, EMAIL_ADDR_PASSWORD, EMAIL_ADDR_TO,"Тестовое задание Гринатом", message_text ,[RESULT_PATH])
    else:
        print("Нечего отправлять на почту")
