import os
import sys
import xlrd
import datetime
from datetime import date
import calendar
import mail
import loggs
def dayweek(par):
    #Определение дня недели
    dayweek = calendar.day_abbr[datetime.date.weekday(datetime.date.today())]

    if par == '-s':
        days = ['пн', 'вт', 'ср', 'чт', 'пт', 'сб', 'вс']
        return days[date.weekday(date.today())]

    if par == '-k':
        days = ['пнд', 'вт', 'ср', 'чт', 'птн', 'сб', 'вс']
        return days[date.weekday(date.today())]

class ParsingXLS(object):
    def parsing(parampars = None):
        os.chdir(sys.path[0])
        if parampars =='-s':
            try:
                saleid = {'отъем': 400, 'Свиньи товарные до 95 кг ': 609, 'Свиньи товарные 95 - 110 кг ': 609,
                          'Свиньи товарные 110-120 кг ': 609,
                          'Свиньи товарные 120-130 кг ': 609, 'Свиньи товарные от 130 кг ': 609,
                          'Выбраковка с откорма 50-70': 606,
                          'Выбраковка с откорма 70 и выше': 605, 'Брак с откорма от 50 кг': 604,
                          'Свиноматки основные до 200 кг': 114,
                          'Свиноматки основные от  200 кг': 114, 'Свинки до 200 кг': 302, 'Свинки от 200 кг': 312,
                          'Свиноматки брак': 112,
                          'Крипторхи': 615, 'Брак хряки до 150 кг': 610, 'Брак хряки от 150 кг': 611,
                          'тех. Брак хряки до 150 кг': 201,
                          'тех. Брак хряки от 150 кг': 203, 'Свинки ремонтные': 304, 'хряки ремонтные': 306}
                wb = xlrd.open_workbook('tmp/svin_temp.xlsm', formatting_info=False, encoding_override="utf-8", on_demand=True)
                sheet = wb.sheet_by_index(1)
                result = open('tmp/svin_res.r', 'a', encoding='utf-8')
                for rownum in range(sheet.nrows):
                    row = sheet.row_values(rownum)
                    if row[2] == dayweek('-s'):
                        if row[3] >= 1.0:
                            rown = int(row[0])
                            weeknum = int(row[1])
                            xldate = row[3]
                            date = xlrd.xldate_as_datetime(xldate, wb.datemode).strftime("%d/%m/%y")
                            xltime = row[6]
                            if xltime == '':
                                continue
                            pointin = row[4]
                            if pointin == '':
                                continue
                            sales = str(row[5])
                            if sales == '':
                                continue
                            pointout = row[10]
                            if pointout == '':
                                continue
                            driver = row[16]
                            if driver == '':
                                continue
                            carnumber = row[17]
                            if carnumber == '':
                                continue
                            trailernumber = str(row[19])
                            if trailernumber == '':
                                continue
                            not_cr = row[13]
                            if not_cr != '':
                                continue
                            slid = saleid.get(sales)
                            year, month, day, hour, minute, second = xlrd.xldate_as_tuple(xltime, wb.datemode)
                            ti = datetime.time(hour,minute).hour
                            if slid == None:
                                slid = ''
                            time = xlrd.xldate_as_datetime(xltime, wb.datemode).strftime(f'{ti}:%M')
                            tr = trailernumber.split()
                            tnum = ''.join(tr)
                            result.write(("%s;%s;%s;%s;%s;%s;%s;%s;%s;%s\n" %
                                            (date, time, pointin, sales.strip(), pointout, driver, carnumber, tnum,row[20], slid)))
                result.close()
            except: loggs.log_error(f'Данные из файла svin_temp.xlsm не получены, {sys.stderr}')
            else:loggs.log_info('Данные из файла korm_temp.xlsm получены')
        if parampars =='-k':
            try:
                wb = xlrd.open_workbook('tmp/korm_temp.xls', formatting_info=False, encoding_override='utf-8',
                                        on_demand=True)
                data = wb.sheet_by_index(0)
                result = open('tmp/korm_res.r', 'a', encoding='utf-8')
                result.flush()
                curr_date = datetime.date.today()
                for rownum in range(data.nrows):
                    if rownum < 4:
                        continue
                    if rownum > 8830:
                        break
                    row = data.row_values(rownum)
                    if not row[5]:
                        continue
                    year, month, day, hour, minute, second = xlrd.xldate_as_tuple(row[1], wb.datemode)
                    dat = datetime.date(year, month, day)
                    if dat != curr_date:
                        continue
                    if row[5] >= 1.0:
                        row[5] = row[5] - 1.0
                    xldate = row[1]
                    date = xlrd.xldate_as_datetime(xldate, wb.datemode).strftime("%d/%m/%y")
                    xltime = row[5]
                    xl_numrow = int(row[2])
                    if xltime == '': continue
                    markofkorm = row[6]
                    if markofkorm == '': continue
                    tonofmark = row[7]
                    if tonofmark == '': continue
                    typekorm = row[8]
                    if typekorm == '': continue
                    out = row[9]
                    if out == '': continue
                    insale = row[10]
                    if insale == '': continue
                    driver = row[17]
                    if driver == '': continue
                    carnumber = row[18]
                    if carnumber == '': continue
                    trailernumber = row[19]
                    if trailernumber == '': continue
                    try:
                        wtf = int(row[16])
                    except:
                        wtf = row[16]
                    if wtf == '': continue
                    xldoubletime = row[80]
                    hr = time = xlrd.xldate_as_datetime(xltime, wb.datemode).hour
                    time = xlrd.xldate_as_datetime(xltime, wb.datemode).strftime(f'{hr}:%M')
                    try:
                        doubletime = xlrd.xldate_as_datetime(xldoubletime, wb.datemode).strftime("%Y-%m-%d")
                    except:
                        doubletime = ''
                    result.write(("%s;%s;%s;%s;%s;%s;%s;%s;%s;%s;%s;%s;%s\n" %
                                  (date, time, markofkorm, tonofmark, typekorm, out, insale, driver, carnumber,
                                   trailernumber, wtf, xl_numrow, doubletime)))
                result.close()
            except: loggs.log_error(f'Данные из файла korm_temp.xls не получены, {sys.stderr}')
            else:loggs.log_info('Данные из файла korm_temp.xls получены')
    def bugtracking(parampars = None):
        os.chdir(sys.path[0])
        if parampars == '-s':
            try:
                saleid = {'отъем': 400, 'Свиньи товарные до 95 кг ': 609, 'Свиньи товарные 95 - 110 кг ': 609,
                          'Свиньи товарные 110-120 кг ': 609,
                          'Свиньи товарные 120-130 кг ': 609, 'Свиньи товарные от 130 кг ': 609,
                          'Выбраковка с откорма 50-70 ': 606,
                          'Выбраковка с откорма 70 и выше': 605, 'Брак с откорма от 50 кг': 604,
                          'Свиноматки основные до 200 кг': 114,
                          'Свиноматки основные от  200 кг': 114, 'Свинки до 200 кг': 302, 'Свинки от 200 кг': 312,
                          'Свиноматки брак': 112,
                          'Крипторхи': 615, 'Брак хряки до 150 кг': 610, 'Брак хряки от 150 кг': 611,
                          'тех. Брак хряки до 150 кг': 201,
                          'тех. Брак хряки от 150 кг': 203, 'Свинки ремонтные': 304, 'хряки ремонтные': 306}
                wb = xlrd.open_workbook('tmp/svin_temp.xlsm', formatting_info=False, encoding_override="utf-8",
                                        on_demand=True)
                sheet = wb.sheet_by_index(1)
                errfile = open('tmp/errfile.src', 'a', encoding='utf-8')
                loggs.log_info('Начат поиск ошибок в файле svin_temp.xlsm')
                for rownum in range(sheet.nrows):
                    row = sheet.row_values(rownum)
                    if row[2] == dayweek('-s'):
                        if row[3] >= 1.0:
                            xltime = row[6]
                            if xltime == '':
                                errfile.write(f'В строке {str(rownum)} не указано время<br>\n')
                            pointin = row[4]
                            if pointin == '':
                                errfile.write(f'Ошибка в строке {str(rownum)} не указана конечная точка<br>\n')
                            sales = row[5]
                            if sales == '':
                                errfile.write(f'Ошибка в строке {str(rownum)} не указано наименование товара<br>\n')
                            pointout = row[10]
                            if pointout == '':
                                errfile.write(f'Ошибка в строке {str(rownum)} не указана точка отправления<br>\n')
                            driver = row[16]
                            if driver == '':
                                errfile.write(f'Ошибка в строке {str(rownum)} не указан водитель<br>\n')
                            carnumber = row[17]
                            if carnumber == '':
                                errfile.write(f'Ошибка в строке {str(rownum)} не указан номер машины<br>\n')
                            trailernumber = row[19]
                            if trailernumber == '':
                                errfile.write(
                                    f'Ошибка в строке {str(rownum)} отсутствует информация по номеру прицепа<br>\n')
                errfile.close()
            except: loggs.log_error(f'Поиск ошибок в файле svin_temp.xlsm не выполнен, {sys.stderr}')
            else:loggs.log_info('Поиск ошибок в файле svin_temp.xlsm выполнен')
            err_ms = open('tmp/errfile.src', 'r', encoding='utf-8')
            try:
                mail.err_mail_msg(err_ms.read(), param='-s')
            except:
                loggs.log_info(f'Ошбка отправки сообщения с ощибками, {sys.stderr}')
            err_ms.close()
            clr_err_ms = open('tmp/errfile.src', 'w', encoding='utf-8')
            clr_err_ms.close()
        if parampars=='-k':
            try:
                wb = xlrd.open_workbook('tmp/korm_temp.xls', formatting_info=False, encoding_override="utf-8",
                                        on_demand=True)
                sheet = wb.sheet_by_index(0)
                errfile = open('tmp/errfile.src', 'a', encoding='utf-8')
                loggs.log_info('Начат поиск ошибок в файле korm_temp.xls')
                for rownum in range(sheet.nrows):
                    row = sheet.row_values(rownum)
                    if row[0] == dayweek('-k'):
                        if row[4] != '' and row[5] != '' and row[6] != '' and row[7] != '' and row[8] != '' and row[
                            9] != '' and row[10] != '':
                            xltime = row[5]
                            if xltime == '':
                                errfile.write(f'В строке {str(rownum)} не указано время<br>\n')
                            markofkorm = row[6]
                            if markofkorm == '':
                                errfile.write(f'В строке {str(rownum)} не указана марка корма<br>\n')
                            tonofmark = row[7]
                            if tonofmark == '':
                                errfile.write(f'В строке {str(rownum)} не указан вес по маркам<br>\n')
                            typekorm = row[8]
                            if typekorm == '':
                                errfile.write(f'В строке {str(rownum)} не указан тип корма<br>\n')
                            out = row[9]
                            if out == '':
                                errfile.write(f'В строке {str(rownum)} не указано откуда корм<br>\n')
                            insale = row[10]
                            if insale == '':
                                errfile.write(f'В строке {str(rownum)} точка прибытия по маркам<br>\n')
                            driver = row[17]
                            if driver == '':
                                errfile.write(f'В строке {str(rownum)} отсутствуют данные по водителю<br>\n')
                            carnumber = row[18]
                            if carnumber == '':
                                errfile.write(f'В строке {str(rownum)} не указан номер машины<br>\n')
                            trailernumber = row[19]
                            if trailernumber == '':
                                errfile.write(f'В строке {str(rownum)} отсутствует информация по прицепу<br>\n')
                            xldoubletime = row[80]
                            if xldoubletime == '':
                                errfile.write(f'В строке {str(rownum)} не продублироваанна дата<br>\n')
                            cpmarkofkorm = str(markofkorm).count('+')
                            cptonofmark = str(tonofmark).count('+')
                            cptypekorm = str(typekorm).count('+')
                            if cpmarkofkorm != cptonofmark != cptypekorm:
                                errfile.write(
                                    f'В строке {str(rownum)} допущена ошибка {markofkorm} {tonofmark} {typekorm} не соответствуют друг другу<br> \n')
                errfile.close()
            except: loggs.log_error(f'Поиск ошибок в файле korm_temp.xls не выполнен, {sys.stderr}')
            else:loggs.log_info('Поиск ошибок в файле korm_temp.xls выполнен')
            err_ms = open('tmp/errfile.src', 'r', encoding='utf-8')
            try:
                mail.err_mail_msg(err_ms.read(), param='-k')
            except:loggs.log_info(f'Ошбка отправки сообщения с ощибками, {sys.stderr}')
            err_ms.close()
            clr_err_ms = open('tmp/errfile.src', 'w', encoding='utf-8')
            clr_err_ms.close()