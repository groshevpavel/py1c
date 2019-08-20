import time
import datetime
import os
import json
import csv


import o1c
import _query # запросы к 1с в одном файле

import logging
_ABSPATH = os.path.dirname(os.path.abspath(__file__))
_LOG = os.path.join(_ABSPATH, __file__.split('.')[0] + '.log')


# from openpyxl import Workbook
# import xlwt


def save_xlsx():
    # =============================== EXCEL openpyxl
    wb = Workbook()
    ws = wb.active
    ws.title = 'TDSheet'

    # добавляем заголовки столбцов
    ws.append(o.columns)

    for d in data:
        ws.append(d)

    wb.save('e:\\1.xlsx')

def save_xls():
    # =============================== EXCEL xlwt
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('TDSheet')

    # добавляем заголовки столбцов
    for n, c in enumerate(o.columns):
        ws.write(0,n,c)

    for r, d in enumerate(data, 1):
        for c in xrange(len(o.columns)):
            value = d[c]
            ws.write(r,c, value)

    wb.save('e:\\1.xls')

def save_csv(csv_data = None, csvfile="e:/1.csv"):
    if csv_data:
        with open(csvfile, 'wb') as f:
            # json.dump(o.all_(), f)
            writer = csv.writer(f, dialect='excel', delimiter=";")
            # writer.writerow([s.encode("windows-1251") for s in o.columns])
            writer.writerow([s.encode("utf8") for s in o.columns])

            try: # если попадаются данные с русскими буквами конвертируем и записываем еще раз
                writer.writerows( o.converted_csv_data(csv_data, "utf8", convert_floats=True) )
            except UnicodeEncodeError:
                writer.writerows(csv_data)


o = o1c.O1c(_query._CONN_STRING_PROD)

# o.make_query(_query.contragents_suppliers_sql)

# headers = ['CAG_GUID','CAG_UASCODE','CAG_NameShort','CAG_INN','CAG_KPP','CAG_OKPO','CAG_NameFull','CAG_tel','CAG_Email',
# 'CAG_FactAddress','CAG_JuryAddress','CAG_KIS_code', 'CAG_Manager']
# o.savecsv(filename = 'test__contragents.csv', index=1, headers=headers)

# headers = ['SUP_1C_GUID','SUP_1C_PartnerGUID','SUP_Name','SUP_Code']
# o.savecsv(filename = 'test__suppliers.csv', index=2, headers=headers)



# o.make_query(_query.nomen_podrazdeleniya_groups)
# headers = None
# o.savecsv(filename = 'test__nomen.csv', index=0, headers=headers)
# o.savecsv(filename = 'test__podrazd.csv', index=1, headers=headers)
# o.savecsv(filename = 'test__groups.csv', index=2, headers=headers)



# o.make_query(_query.postuplenia_sql)

# days_from_yesterday = 1
# startdate, enddate = o.ndays_from_yesterday(days_from_yesterday) # за N дней начиная со вчера

# o.setp(ur"НачалоПериода", startdate)
# o.setp(ur"КонецПериода", enddate)

# o.savecsv(filename = 'test__postuplenia.csv')

# o.make_query(_query.bizreg_test_sql)
# o.make_query((u'R:\\Формат Аптеки\\Технологии\\1C\\1С_Запросы',u'НомеклатураС_ЗШК_НДС.txt',))
# o.setp(ur"Родитель", 0x1ab8b824-acfd-11e7-813d-00155dbdb007)

o.make_query(_query.sklad_list_sql)

o.savecsv(filename = './test/test__sklad000.csv')

# data = o.all_() # массив кортежей

# =============================== ПОЛУЧАЕМ ДАННЫЕ
# print u"%s\tВыборка данных завершена.."% o.get_now_str()

# edo_path = ur"R:/Формат Аптеки/Технологии/ЭДО/xls"
# edo_path = os.path.dirname(os.path.abspath(__file__))
# csv_filename = "edo_%s.csv" % o.get_now_str("%d-%m-%Y_%H%M%S")
# csv_fullfilename = os.path.join(edo_path, csv_filename)

# print u"%s\tСохраняю '%s'"% (o.get_now_str(), csv_fullfilename)
# save_csv(data, csvfile = csv_fullfilename)
# # save_xlsx()
