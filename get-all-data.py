# -*- coding: utf-8 -*-
import requests
from requests.exceptions import ConnectionError
from time import sleep
import json
import datetime
import openpyxl
from xlrd import open_workbook
import xlwt
import xlsxwriter

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

clients = ['marketdo4a-blr-riverstart','marketdo4a-brn-riverstart','marketdo4a-brzn-riverstart','marketdo4a-chita-riverstart','marketdo4a-chlb-riverstart','marketdo4a-ekb-riverstart','marketdo4a-hbr-riverstart','marketdo4a-irk-riverstart','marketdo4a-klg-riverstart','marketdo4a-kmr-riverstart','marketdo4a-kna-riverstart','marketdo4a-krd-riverstart','marketdo4a-krr-riverstart','marketdo4a-kstrm-riverstart','marketdo4a-kzn-riverstart','marketdo4a-mgd-riverstart','marketdo4a-mkp-riverstart','marketdo4a-msk-riverstart','marketdo4a-nn-riverstart','marketdo4a-novur-riverstart','marketdo4a-nvch-riverstart','marketdo4a-nvkz-riverstart','marketdo4a-nvr-riverstart','marketdo4a-nvsb-riverstart','marketdo4a-nzhnv-riverstart','marketdo4a-oren-riverstart','marketdo4a-perm-riverstart','marketdo4a-riverstart','marketdo4a-rnd-riverstart','marketdo4a-simf-riverstart','marketdo4a-smr-riverstart','marketdo4a-sochi-riverstart','marketdo4a-spb-riverstart','marketdo4a-srg-riverstart','marketdo4a-tomsk-riverstart','marketdo4a-tumen-riverstart','marketdo4a-ufa-riverstart','marketdo4a-ulud-riverstart','marketdo4a-ulyan-riverstart','marketdo4a-ussur-riverstart','marketdo4a-uszsah-riverstart','marketdo4a-vldm-riverstart','marketdo4a-vldv-riverstart','marketdo4a-vlg-riverstart','marketdo4a-vn-riverstart','marketdo4a-vrn-riverstart','marketdo4a-yakut-riverstart']
# Получаем вчерашнюю дату
def getYesterday(): 
    today=datetime.date.today() 
    oneday=datetime.timedelta(days=1) 
    yesterday=today-oneday  
    return yesterday
 
date1 = str(getYesterday())

for i in clients:
    # --- Входные данные ---
    # Адрес сервиса Reports для отправки JSON-запросов (регистрозависимый)
    ReportsURL = 'https://api.direct.yandex.com/json/v5/reports'

    # OAuth-токен пользователя, от имени которого будут выполняться запросы
    token = 'AQAAAAAye9OuAAWoSe5kDEDFiE9IkbFVe-zMvf0'

    # Логин клиента рекламного агентства
    # Обязательный параметр, если запросы выполняются от имени рекламного агентства
    clientLogin = i

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
    files = i + '.xls'
    file = open(files, "w")
    file.write(req.text)
    file.close()
    #workbook = xlsxwriter.Workbook('marketdo4a-chlb-riverstart.xls')
    #worksheet_analysis = workbook.add_worksheet('marketdo4a-chlb-riverstart')
    #worksheet_analysis.write("A1","Название кампании")
    #worksheet_analysis.write("B1","Показы")
    #worksheet_analysis.write("C1","Клики")
    #worksheet_analysis.write("D1","Расход")
    #worksheet_analysis.write("E1","CTR%")
    #worksheet_analysis.write("F1","CPC") #средняя цена клика, расход / клики
    #worksheet_analysis.write("G1","CPO") # бюджет / конверсии
    #worksheet_analysis.write("H1","CV") # конверсии / клики
    #worksheet_analysis.write("I1","Конверсии") # берем из метрики ym:s:ecommercePurchases
    #workbook.close()

    # Далее скрипт должен читать файл xls. 
    # Первым делом он читает ячейку на содержание в ней значений. Ячейка - последняя в строке. если в ней есть значение - в следующую ячейку справа записываем значение переменной clientLogin
    # После, в следующую ячейку записываем значение переменной cpc - которое должен посчитать скрипт. Расчет происходит следующим образом - считываем ячейку в строке, в которой забито значение Cost и ячейки со значением clicks. Передаем их в новые переменные, после считаем значение cpc = cost / clicks. И так с CPO, CV.

