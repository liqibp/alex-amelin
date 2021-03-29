# -*- coding: utf-8 -*-
import requests
from requests.exceptions import ConnectionError
from time import sleep
import json
import datetime
from xlrd import open_workbook
import xlsxwriter
import csv
import openpyxl as op
import xlrd
import os

# Метод для корректной обработки строк в кодировке UTF-8 как в Python 3, так и в Python 2
import sys

if sys.version_info < (3,):
    def u(x):
        try:
            return x.encode("utf8")
        except UnicodeDecodeError:
            return x
else:
    def u(x):
        if type(x) == type(b''):
            return x.decode('utf8')
        else:
            return x

# Получаем вчерашнюю дату
def getYesterday(): 
    today=datetime.date.today() 
    oneday=datetime.timedelta(days=1) 
    yesterday=today-oneday  
    return yesterday
 
date1 = str(getYesterday())

# --- Входные данные ---
# Адрес сервиса Reports для отправки JSON-запросов (регистрозависимый)
ReportsURL = 'https://api.direct.yandex.com/json/v5/reports'

# OAuth-токен пользователя, от имени которого будут выполняться запросы
token = 'AQAAAAAye9OuAAWoSe5kDEDFiE9IkbFVe-zMvf0'

# Логин клиента рекламного агентства
# Обязательный параметр, если запросы выполняются от имени рекламного агентства
clientLogin = 'marketdo4a-hbr-riverstart'

# --- Подготовка запроса ---
# Создание HTTP-заголовков запроса
headers = {
           # OAuth-токен. Использование слова Bearer обязательно
           "Authorization": "Bearer " + token,
           # Логин клиента рекламного агентства
           "Client-Login": clientLogin,
           # Язык ответных сообщений
           "Accept-Language": "ru",
           # Режим формирования отчета
           "processingMode": "auto",
           # Формат денежных значений в отчете
           "returnMoneyInMicros": "false",
           # Не выводить в отчете строку с названием отчета и диапазоном дат
           "skipReportHeader": "true"
           # Не выводить в отчете строку с названиями полей
           # "skipColumnHeader": "true"
           # Не выводить в отчете строку с количеством строк статистики
           # "skipReportSummary": "true"
           }

# Создание тела запроса
body = {
    "params": {
        "SelectionCriteria": {
            "DateFrom": "2019-01-01",
            "DateTo": "2019-06-06"
        },
        "Goals": [
            "5125289"
        ],
        "FieldNames": [
            "CampaignName",
            "Impressions",
            "Clicks",
            "Ctr",
            "Cost",
            "Conversions",
            "Revenue"
        ],
        "ReportName": u(clientLogin),
        "ReportType": "CAMPAIGN_PERFORMANCE_REPORT",
        "DateRangeType": "CUSTOM_DATE",
        "Format": "TSV",
        "IncludeVAT": "NO",
        "IncludeDiscount": "NO"
    }
}

# Кодирование тела запроса в JSON
body = json.dumps(body, indent=4)

# --- Запуск цикла для выполнения запросов ---
# Если получен HTTP-код 200, то выводится содержание отчета
# Если получен HTTP-код 201 или 202, выполняются повторные запросы
while True:
    try:
        req = requests.post(ReportsURL, body, headers=headers)
        req.encoding = 'utf-8'  # Принудительная обработка ответа в кодировке UTF-8
        if req.status_code == 400:
            print("Параметры запроса указаны неверно или достигнут лимит отчетов в очереди")
            print("RequestId: {}".format(req.headers.get("RequestId", False)))
            print("JSON-код запроса: {}".format(u(body)))
            print("JSON-код ответа сервера: \n{}".format(u(req.json())))
            break
        elif req.status_code == 200:
            format(u(req.text))
            break
        elif req.status_code == 201:
            print("Отчет успешно поставлен в очередь в режиме офлайн")
            retryIn = int(req.headers.get("retryIn", 60))
            print("Повторная отправка запроса через {} секунд".format(retryIn))
            print("RequestId: {}".format(req.headers.get("RequestId", False)))
            sleep(retryIn)
        elif req.status_code == 202:
            print("Отчет формируется в режиме офлайн")
            retryIn = int(req.headers.get("retryIn", 60))
            print("Повторная отправка запроса через {} секунд".format(retryIn))
            print("RequestId:  {}".format(req.headers.get("RequestId", False)))
            sleep(retryIn)
        elif req.status_code == 500:
            print("При формировании отчета произошла ошибка. Пожалуйста, попробуйте повторить запрос позднее")
            print("RequestId: {}".format(req.headers.get("RequestId", False)))
            print("JSON-код ответа сервера: \n{}".format(u(req.json())))
            break
        elif req.status_code == 502:
            print("Время формирования отчета превысило серверное ограничение.")
            print("Пожалуйста, попробуйте изменить параметры запроса - уменьшить период и количество запрашиваемых данных.")
            print("JSON-код запроса: {}".format(body))
            print("RequestId: {}".format(req.headers.get("RequestId", False)))
            print("JSON-код ответа сервера: \n{}".format(u(req.json())))
            break
        else:
            print("Произошла непредвиденная ошибка")
            print("RequestId:  {}".format(req.headers.get("RequestId", False)))
            print("JSON-код запроса: {}".format(body))
            print("JSON-код ответа сервера: \n{}".format(u(req.json())))
            break

    # Обработка ошибки, если не удалось соединиться с сервером API Директа
    except ConnectionError:
        # В данном случае мы рекомендуем повторить запрос позднее
        print("Произошла ошибка соединения с сервером API")
        # Принудительный выход из цикла
        break

    # Если возникла какая-либо другая ошибка
    except:
        # В данном случае мы рекомендуем проанализировать действия приложения
        print("Произошла непредвиденная ошибка")
        # Принудительный выход из цикла
        break

#Задаем параметры книги xls
# file = open("marketdo4a-chlb-riverstart.csv", "w")
# string = req.text
# string = string.replace(",", ".")
# string = string.replace("\t", ",")
# print(string)
# file.write(string)
# file.close()
file = open("marketdo4a-hbr-riverstart.tsv", "w")
string = req.text
string = string.replace(".", ",")
file.write(string)
file.close()

#Запись tsv данных в xlsx
from xlsxwriter.workbook import Workbook
# Add some command-line logic to read the file names.
tsv_file = 'marketdo4a-hbr-riverstart.tsv'
xlsx_file = 'marketdo4a-hbr-riverstart.xlsx'

# Create an XlsxWriter workbook object and add a worksheet.
workbook = Workbook(xlsx_file)
worksheet = workbook.add_worksheet()

# Create a TSV file reader.
tsv_reader = csv.reader(open(tsv_file, 'rt'), delimiter='\t')

# Read the row data from the TSV file and write it to the XLSX file.
for row, data in enumerate(tsv_reader):
    worksheet.write_row(row, 0, data)

# Close the XLSX file.
workbook.close()

#Обрабатываем полученный xlsx файл
excel_data_file = xlrd.open_workbook('marketdo4a-hbr-riverstart.xlsx')
sheet = excel_data_file.sheet_by_index(0)
row_numbers = sheet.nrows
print(row_numbers)

wb = op.load_workbook('marketdo4a-hbr-riverstart.xlsx')
ws = wb['Sheet1']

#переводим заголовки
ws.cell(row = 1, column = 1).value = 'Кампания'
ws.cell(row = 1, column = 2).value = 'Показы'
ws.cell(row = 1, column = 3).value = 'Переходы/Клики'
ws.cell(row = 1, column = 4).value = 'CTR'
ws.cell(row = 1, column = 5).value = 'Бюджет'
ws.cell(row = 1, column = 6).value = 'Продажи с сайта'
ws.cell(row = 1, column = 7).value = 'Доход с сайта'
ws.cell(row = 1, column = 8).value = 'CPC'
ws.cell(row = 1, column = 9).value = 'CPO'
ws.cell(row = 1, column = 10).value = 'CV'

true_row_number = row_numbers - 2

#информация о логике циклов дана в цикле расчет сро
#Цикл расчета cpc (считает python)
campaign_name = ws['A2']
impressions = ws['B2']
clicks = ws['C2']
ctr_val = ws['D2']
budget_col = ws['E2']
ecomm_sales = ws['F2']
money = ws['G2']
cpc = ws['H2']
cpo = ws['I2']
cv = ws['J2']

a = 1
for i in range(true_row_number):
    a += 1
    clicks = ws['C' + str(a)]
    budget_col = ws['E' + str(a)]
    val_if_zero = 0
    if clicks.value == 0:
        ws.cell(row = a, column = 8).value = val_if_zero
    else:
        result_cpc = budget_col.value / clicks.value
        ws.cell(row = a, column = 8).value = result_cpc
    pass

#Цикл расчета СPO (считает python)
campaign_name = ws['A2'] #объявляем название показателей как переменных (A2 потому что не берем строку с заголовками столбцов)
impressions = ws['B2']
clicks = ws['C2']
ctr_val = ws['D2']
budget_col = ws['E2']
ecomm_sales = ws['F2']
money = ws['G2']
cpc = ws['H2']
cpo = ws['I2']
cv = ws['J2']

a = 1 #переменная 'a' - динамическая, используется для отсчета номера ряда в экселе. 
for i in range(true_row_number):
    a += 1 #на каждой итерации цикла прибавляем единицу к динамической переменной чтобы записывать данные в нужный ряд
    budget_col = ws['E' + str(a)] #в данном случае считаем cpo, поэтому берем два показателя, и указываем для них координаты. Причем строка - это динамическая переменная а.
    ecomm_sales = ws['F' + str(a)]
    val_if_zero = 0
    if ecomm_sales.value == '--' or ecomm_sales.value == 0: #если нет данных, или стоит 0, будем в результат тоже записывать 0
        ws.cell(row = a, column = 6).value = val_if_zero
        ws.cell(row = a, column = 9).value = val_if_zero
    else:
        result_cpo = budget_col.value / ecomm_sales.value #в остальных случаях расчет возможен, поэтому пишем формулу расчета cpo(при чем делимое и делитель постоянно меняются за счет динамической переменной a)
        ws.cell(row = a, column = 9).value = result_cpo #результат равен частному из предыдущей строки, записываем его в ячейку таблицы.
    pass

#Цикл расчета CV (считает python)
campaign_name = ws['A2']
impressions = ws['B2']
clicks = ws['C2']
ctr_val = ws['D2']
budget_col = ws['E2']
ecomm_sales = ws['F2']
money = ws['G2']
cpc = ws['H2']
cpo = ws['I2']
cv = ws['J2']

a = 1
for i in range(true_row_number):
    a += 1
    clicks = ws['C' + str(a)]
    ecomm_sales = ws['F' + str(a)]
    val_if_zero = 0
    if clicks.value == '--' or clicks.value == 0:
        ws.cell(row = a, column = 10).value = val_if_zero
    else:
        result_cv = ecomm_sales.value / clicks.value
        ws.cell(row = a, column = 10).value = result_cv
    pass

#приводим к нулю показатели дохода если дохода нет
campaign_name = ws['A2']
impressions = ws['B2']
clicks = ws['C2']
ctr_val = ws['D2']
budget_col = ws['E2']
ecomm_sales = ws['F2']
money = ws['G2']
cpc = ws['H2']
cpo = ws['I2']
cv = ws['J2']

a = 1
for i in range(true_row_number):
    a += 1
    money = ws['G' + str(a)]
    val_if_zero = 0
    if money.value == '--':
        ws.cell(row = a, column = 7).value = val_if_zero
    else:
        pass
    pass
    
wb.save('marketdo4a-hbr-riverstart.xlsx')
wb.close()
os.remove('marketdo4a-hbr-riverstart.tsv')

