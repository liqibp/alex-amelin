# -*- coding: utf-8 -*-
import requests
from requests.exceptions import ConnectionError
from time import sleep
import json

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

# --- Входные данные ---
# Адрес сервиса Reports для отправки JSON-запросов (регистрозависимый)
ReportsURL = 'https://api-sandbox.direct.yandex.com/json/v5/reports'

# OAuth-токен пользователя, от имени которого будут выполняться запросы
token = 'AQAAAAAgPg5nAAWGVbgmuyXKTUE3sqKkzwHx7Uo'

# Логин клиента рекламного агентства
# Обязательный параметр, если запросы выполняются от имени рекламного агентства
clientLogin = ''

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
           "returnMoneyInMicros": "false"
           # Не выводить в отчете строку с названием отчета и диапазоном дат
           # "skipReportHeader": "true",
           # Не выводить в отчете строку с названиями полей
           # "skipColumnHeader": "true",
           # Не выводить в отчете строку с количеством строк статистики
           # "skipReportSummary": "true"
           }

# Мой код, принимает данные от пользователя (дату) и передает их в тело запроса
print('Введите начало и конец периода. Даты задаются в формате DD-MM-YYYY')
dd1 = int(input("Enter the start date(day): "))
if dd1 > 31:
    while dd1 > 31:
        print('Неверный день. только 1 - 31')
        dd1 = int(input("Enter the start date(day): "))
    else:
        pass
else:
    pass

day1 = dd1
day1 = str(day1)
if dd1 < 10:
    day1 = "0" + day1
else:
    pass

mm1 = int(input("Enter the start date(month): "))
if dd1 == 29 or dd1 == 30 or dd1 == 31:
    while mm1 == 2:
        print('В феврале нет столько дней. введите другое значение')
        mm1 = int(input("Enter the start date(month): "))
    else:
        pass
elif dd1 == 31:
    while mm1 == 4 or mm1 == 6 or mm1 == 9 or mm1 == 11:
        print('В месяце нет столько дней. введите другое значение')
        mm1 = int(input("Enter the start date(month): "))
    else:
        pass
elif mm1 > 12:
    while mm1 > 12:
        print('Неверный месяц. введите заново')
        mm1 = int(input("Enter the start date(month): "))
    else:
        pass
else:
    pass

month1 = mm1
month1 = str(month1)
if mm1 < 10:
    month1 = "0" + month1
else:
    pass

yyyy1 = int(input("Enter the start date(year): "))

print("okay")

dd2 = int(input("Enter the end date(day): "))
if dd2 > 31:
    while dd1 > 31:
        print('Неверный день. только 1 - 31')
        dd2 = int(input("Enter the end date(day): "))
    else:
        pass
else:
    pass

day2 = dd2
day2 = str(day2)
if dd2 < 10:
    day2 = "0" + day2
else:
    pass

mm2 = int(input("Enter the end date(month): "))
if dd2 == 29 or dd2 == 30 or dd2 == 31:
    while mm2 == 2:
        print('В феврале нет столько дней. введите другое значение')
        mm2 = int(input("Enter the end date(month): "))
    else:
        pass
elif dd2 == 31:
    while mm2 == 4 or mm2 == 6 or mm2 == 9 or mm2 == 11:
        print('В месяце нет столько дней. введите другое значение')
        mm2 = int(input("Enter the end date(month): "))
    else:
        pass
elif mm2 > 12:
    while mm2 > 12:
        print('Неверный месяц. введите заново')
        mm2 = int(input("Enter the end date(month): "))
    else:
        pass
else:
    pass

month2 = mm2
month2 = str(month2)
if mm2 < 10:
    month2 = "0" + month2
else:
    pass

yyyy2 = int(input("Enter the end date(year): "))
if dd2 < dd1 and mm2 <= mm1 and yyyy2 <= yyyy1:
    while dd2 < dd1 and mm2 <= mm1 and yyyy2 <= yyyy1:
        print('Год не может быть меньше при такой дате. Введите другой год')
        yyyy2 = int(input("Enter the end date(year): "))
    else:
        pass
else:
    pass

print("okay, send the request...")

str(yyyy1)
str(yyyy2)

date1 = (str(yyyy1) + "-" + str(month1) + "-" + str(day1))
date2 = (str(yyyy2) + "-" + str(month2) + "-" + str(day2))

# Создание тела запроса
body = {
    "params": {
        "SelectionCriteria": {
            "DateFrom": date1,
            "DateTo": date2
        },
        "FieldNames": [
            "Date",
            "CampaignName",
            "Cost"
        ],
        "ReportName": u("Отчет по расходам"),
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
            print("Отчет создан успешно")
            print("RequestId: {}".format(req.headers.get("RequestId", False)))
            print("Содержание отчета: \n{}".format(u(req.text)))
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

input("\n\nНажмите Enter чтобы выйти .")
