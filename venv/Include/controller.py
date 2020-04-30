import os
import shutil
import datetime
import time
import sys
import loggs
#from mount import Mount
#from mount import TempCreate
#from mount import CheckSize
def mountdir():
    try:
        os.system("mount.cifs -o username=s_korm,password=www_8888,domain=agrohold.ru //192.168.100.10/Documents/Общая /mnt/winshare")
    except:loggs.log_error('Ошибка при монтировании раздела')
    else:loggs.log_info('Раздел примонтирован')

def umountdir():
    try:
        os.system("umount /mnt/winshare")
    except:loggs.log_error('Раздел не размонтирован')
    else:loggs.log_info('Раздел размонтирован')

def copytmp(paramtmp):
    today = datetime.datetime.today()
    year = today.strftime("%Y")
    week = today.isocalendar()[1]

    if paramtmp == '-k':
        t = f'/mnt/winshare/КОРМА/{year}/неделя {week}/ОБЩИЙ график.xls'
        try:
            shutil.copy(t, 'tmp/korm_temp.xls')
        except: loggs.log_error('Ошибка при копировании файла ОБЩИЙ график.xls')

    if paramtmp == '-s':
        try:
            shutil.copy(f'/mnt/winshare/График отгрузок/{year}/неделя-{week}/График свиновозов.xlsm', 'tmp/svin_temp.xlsm')
        except: loggs.log_error('Ошибка при копировании файла График свиновозов.xlsm')


def sizeorig(paramdir):
    today = datetime.datetime.today()
    year = today.strftime("%Y")
    week = today.isocalendar()[1]

    if paramdir == '-k':
        file = f'/mnt/winshare/КОРМА/{year}/неделя {week}/ОБЩИЙ график.xls'
    if paramdir == '-s':
        file = f'/mnt/winshare/График отгрузок/{year}/неделя-{week}/График свиновозов.xlsm'
    try:
        return os.path.getsize(file)
    except:loggs.log_error('Ошибка при получении размера оригинального файла')
    else:loggs.log_info('Размеры оригинального файла получены')

def mounting(param):
    if param == '-s':
        mountdir()
        if not os.path.isfile('tmp/svin_temp.xlsm') or sizeorig('-s') == os.path.getsize('tmp/svin_temp.xlsm'):
            try:
                copytmp(param)
            except:loggs.log_error('Ошибка при копировании файла svin_temp.xlsm')
            else:loggs.log_info('Файл svin_temp.xlsm скопирован')
        umountdir()

    if param == '-k':
        mountdir()
        if not os.path.isfile('tmp/korm_temp.xls') or sizeorig('-k') == os.path.getsize('tmp/korm_temp.xls'):
            try:
                copytmp(param)
            except:loggs.log_error('Ошибка при копировании файла korm_temp.xls')
            else:loggs.log_info('Файл korm_temp.xls скопирован')
        umountdir()

