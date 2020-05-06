import os,sys,xlrd,datetime,calendar,smtplib, shutil
from datetime import date
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
def dayweek(par):
    #Определение дня недели
    dayweek = calendar.day_abbr[datetime.date.weekday(datetime.date.today())]

    if par == '-s':
        days = ['пн', 'вт', 'ср', 'чт', 'пт', 'сб', 'вс']
        return days[date.weekday(date.today())]

    if par == '-k':
        days = ['пнд', 'вт', 'ср', 'чт', 'птн', 'сб', 'вс']
        return days[date.weekday(date.today())]
class ParseXls(object):
    def __init__(self, parampars):
        self.parampars = parampars
        self.dayweek = calendar.day_abbr[datetime.date.weekday(datetime.date.today())]
        self.tempfolder = f'{sys.path[0]}/tmp'
        self.saleid = {'отъем': 400, 'Свиньи товарные до 95 кг ': 609, 'Свиньи товарные 95 - 110 кг ': 609,
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
        self.errfile = open('tmp/errfile.src', 'a', encoding='utf-8')
        self.svinresult = open(f'{self.tempfolder}/svin_res.r', 'a', encoding='utf-8')
        self.kormresult = open(f'{self.tempfolder}/korm_res.r', 'a', encoding='utf-8')
        self.info = log_info()
        self.error = log_error()
        self.curr_date = datetime.date.today()
        if self.parampars == '-s':
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
            try:
                wb = xlrd.open_workbook(f'{self.tempfolder}/svin_temp.xlsm', formatting_info=False, encoding_override="utf-8",
                                            on_demand=True)
                sheet = wb.sheet_by_index(1)
                for rownum in range(sheet.nrows):
                    row = sheet.row_values(rownum)
                    if row[2] == dayweek('-s'):
                        if row[3] >= 1.0:
                            xldate = row[3]
                            date = xlrd.xldate_as_datetime(xldate, wb.datemode).strftime("%d/%m/%y")
                            xltime = row[6]
                            if xltime == '':
                                    self.errfile.write(f'В строке {str(rownum)} не указано время<br>\n')
                                    continue
                            pointin = row[4]
                            if pointin == '':
                                    self.errfile.write(f'Ошибка в строке {str(rownum)} не указана конечная точка<br>\n')
                                    continue
                            sales = str(row[5])
                            if sales == '':
                                    self.errfile.write(f'Ошибка в строке {str(rownum)} не указано наименование товара<br>\n')
                                    continue
                            pointout = row[10]
                            if pointout == '':
                                    self.errfile.write(f'Ошибка в строке {str(rownum)} не указана точка отправления<br>\n')
                                    continue
                            driver = row[16]
                            if driver == '':
                                    self.errfile.write(f'Ошибка в строке {str(rownum)} не указан водитель<br>\n')
                                    continue
                            carnumber = row[17]
                            if carnumber == '':
                                    self.errfile.write(f'Ошибка в строке {str(rownum)} не указан номер машины<br>\n')
                                    continue
                            trailernumber = str(row[19])
                            if trailernumber == '':
                                    self.errfile.write(f'Ошибка в строке {str(rownum)} отсутствует информация по номеру прицепа<br>\n')
                                    continue
                            not_cr = row[13]
                            if not_cr != '':
                                continue
                            slid = saleid.get(sales)
                            year, month, day, hour, minute, second = xlrd.xldate_as_tuple(xltime, wb.datemode)
                            ti = datetime.time(hour, minute).hour
                            if slid == None:
                                slid = ''
                            time = xlrd.xldate_as_datetime(xltime, wb.datemode).strftime(f'{ti}:%M')
                            tr = trailernumber.split()
                            tnum = ''.join(tr)
                            self.svinresult.write(("%s;%s;%s;%s;%s;%s;%s;%s;%s;%s\n" %
                                            (date, time, pointin, sales.strip(), pointout, driver, carnumber, tnum,
                                               row[20], slid)))
                self.svinresult.close()
                self.errfile.close()
            except:
                log_error(f'Данные из файла svin_temp.xlsm не получены, {sys.stderr}')
            else:
                log_info('Данные из файла korm_temp.xlsm получены')
            err_ms = open('tmp/errfile.src', 'r', encoding='utf-8')
            try:
                ErrSend('-s', err_ms.read(), ['am.fesenko@agrohold.ru', 'a.v.borodin@agrohold.ru'], '192.168.100.238',
                        25)
            except:
                log_error(f'Ошбка отправки сообщения с ощибками, {sys.stderr}')
            err_ms.close()
        if self.parampars == '-k':
            try:
                wb = xlrd.open_workbook(f'tmp/korm_temp.xls', formatting_info=False, encoding_override='utf-8',
                                            on_demand=True)
                data = wb.sheet_by_index(0)
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
                    if dat != self.curr_date:
                        continue
                    if row[5] >= 1.0:
                        row[5] = row[5] - 1.0
                    xldate = row[1]
                    date = xlrd.xldate_as_datetime(xldate, wb.datemode).strftime("%d/%m/%y")
                    xltime = row[5]
                    xl_numrow = int(row[2])
                    if xltime == '': continue
                    markofkorm = row[6]
                    if markofkorm == '':
                        self.errfile.write(f'В строке {str(rownum)} не указана марка корма<br>\n')
                        continue
                    tonofmark = row[7]
                    if tonofmark == '':
                        errfile.write(f'В строке {str(rownum)} не указан вес по маркам<br>\n')
                        continue
                    typekorm = row[8]
                    if typekorm == '':
                        self.errfile.write(f'В строке {str(rownum)} не указан тип корма<br>\n')
                        continue
                    out = row[9]
                    if out == '':
                        self.errfile.write(f'В строке {str(rownum)} не указано откуда корм<br>\n')
                        continue
                    insale = row[10]
                    if insale == '':
                        errfile.write(f'В строке {str(rownum)} точка прибытия по маркам<br>\n')
                        continue
                    driver = row[17]
                    if driver == '':
                        errfile.write(f'В строке {str(rownum)} отсутствуют данные по водителю<br>\n')
                        continue
                    carnumber = row[18]
                    if carnumber == '':
                        errfile.write(f'В строке {str(rownum)} не указан номер машины<br>\n')
                        continue
                    trailernumber = row[19]
                    if trailernumber == '':
                        errfile.write(f'В строке {str(rownum)} отсутствует информация по прицепу<br>\n')
                        continue
                    try:
                        wtf = int(row[16])
                    except:
                        wtf = row[16]
                    if wtf == '':
                        continue
                    xldoubletime = row[80]
                    if xldoubletime == '':
                        errfile.write(f'В строке {str(rownum)} не продублироваанна дата<br>\n')
                    hr = time = xlrd.xldate_as_datetime(xltime, wb.datemode).hour
                    time = xlrd.xldate_as_datetime(xltime, wb.datemode).strftime(f'{hr}:%M')
                    try:
                        doubletime = xlrd.xldate_as_datetime(xldoubletime, wb.datemode).strftime("%Y-%m-%d")
                    except:
                        doubletime = ''
                    self.kormresult.write(("%s;%s;%s;%s;%s;%s;%s;%s;%s;%s;%s;%s;%s\n" %
                                    (date, time, markofkorm, tonofmark, typekorm, out, insale, driver, carnumber,
                                    trailernumber, wtf, xl_numrow, doubletime)))
                self.kormresult.close()
                self.errfile.close()
            except:
                log_error(f'Данные из файла korm_temp.xls не получены, {sys.stderr}')
            else:
                log_info('Данные из файла korm_temp.xls получены')
            err_ms = open('tmp/errfile.src', 'r', encoding='utf-8')
        try:
            ErrSend('-k',err_ms.read(),['am.fesenko@agrohold.ru', 'a.v.borodin@agrohold.ru'],'192.168.100.238',25)
        except:
            log_error(f'Ошбка отправки сообщения с ощибками, {sys.stderr}')
        err_ms.close()
class ErrSend(object):
    def __init__(self, type, msg, mail_list, host, port):
        self.type = str(type)
        self.msg = msg
        self.mail_list = mail_list
        self.host = str(host)
        self.port = int(port)
        self.settings = {'smtp':self.host, 'port':self.port, 'frm':'xls.script@agrohold.ru'}
        message = MIMEMultipart("alternative")
        if self.type == '-s':
            message[
                "Subject"] = f'Отчет об ошибках функции get_svin на момент {datetime.datetime.today().strftime(" %Y/%m/%d; %H:%M;")}'
        if self.type == '-k':
            message[
                "Subject"] = f'Отчет об ошибках функции get_korm на момент {datetime.datetime.today().strftime(" %Y/%m/%d; %H:%M;")}'
        message["From"] = self.settings.get('frm')
        message["To"] = ', '.join(self.mail_list)
        html = f"<html><body><p><strong>{self.msg}</strong></p></body></html>"
        # Сделать их текстовыми\html объектами MIMEText
        message.attach(MIMEText(msg, "plain"))
        message.attach(MIMEText(html, "html"))
        mailObj = smtplib.SMTP(self.settings.get('smtp'), self.settings.get('port'))
        mailObj.sendmail(self.settings.get('frm'), self.mail_list, message.as_string())
        mailObj.quit()
class Mount(object):
    def __init__(self, param, path, mount_point):
        self.param = param
        self.path = path
        self.mount_point = mount_point
        self.info = log_info()
        self.error = log_error()
        today = datetime.datetime.today()
        year = today.strftime("%Y")
        week = today.isocalendar()[1]
        try:
            os.system(f"mount.cifs -o username=s_korm,password=www_8888,domain=agrohold.ru {self.path} {self.mount_point}")
        except:
            log_error('Ошибка при монтировании раздела')
        else:
            log_info('Раздел примонтирован')
        if self.param == '-k':
            try:
                shutil.copy(f'/mnt/winshare/КОРМА/{year}/неделя {week}/ОБЩИЙ график.xls', 'tmp/korm_temp.xls')
            except:
                log_error('Ошибка при копировании файла ОБЩИЙ график.xls')
        if self.param == '-s':
            try:
                shutil.copy(f'/mnt/winshare/График отгрузок/{year}/неделя-{week}/График свиновозов.xlsm','tmp/svin_temp.xlsm')
            except:
                log_error('Ошибка при копировании файла График свиновозов.xlsm')
        try:
            os.system(f"umount {self.mount_point}")
        except:
            log_error('Раздел не размонтирован')
        else:
            log_info('Раздел размонтирован')
def verify_ip(user_ip):
    #Проверка ip адресса на доступ к функциям
    granted_ip = ['192.168.100.101', '192.168.100.4','192.168.100.12',
                      '192.168.100.6','192.168.100.104','192.168.100.102',
                      '192.168.102.209','192.168.100.100','192.168.100.103',
                      '192.168.100.193','192.168.14.41','192.168.96.3',
                      '192.168.96.6','192.168.96.11','192.168.96.12',
                      '192.168.96.13','192.168.96.16','192.168.96.120',
                      '127.0.0.1','192.168.1.2','192.168.1.100',
                      '192.168.12.137','192.168.100.2','192.168.100.202',
                      '192.168.100.208','192.168.100.252']
    verify = granted_ip.count(user_ip)
    return verify
def cleaner():
    tempfile = {'kormview': 'tmp/korm_res.r', 'svinview': 'tmp/svin_res.r', 'errfile': 'tmp/errfile.src',
                'kormtemp': 'tmp/korm_temp.xls', 'svintemp': 'tmp/svin_temp.xlsm'}
    #try:
    for temp in tempfile.values():
        if os.path.isfile(temp) == True:
            os.remove(temp)
    #except:
        #loggs.log_error(
            #f'Ошибка при удалении {tempfile.get("svinview")}, {tempfile.get("kormview")}, {tempfile.get("svintemp")}, {tempfile.get("kormtemp")}, {tempfile.get("errfile")}')
def run(param):
    if param == '-s':
        cleaner()
        Mount('-s', '//192.168.100.10/Documents/Общая', '/mnt/winshare')
        ParseXls('-s').parsing()
    if param == '-k':
        cleaner()
        Mount('-k', '//192.168.100.10/Documents/Общая', '/mnt/winshare')
        ParseXls('-k').parsing()
def log_info(msg):
    logfile = open(file='tmp/info.log', mode='a', encoding='utf-8')
    logfile.write(f'{datetime.datetime.today().strftime("[%Y-%m-%d|%H:%M]")} {msg}\n')
def log_error(msg):
    logfile = open(file='tmp/error.log', mode='a', encoding='utf-8')
    logfile.write(f'{datetime.datetime.today().strftime("[%Y-%m-%d|%H:%M]")} {msg}\n')