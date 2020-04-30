import os
import loggs
def cleanTemp(keyType = None):
    tempfile = ['tmp/korm_res.r','tmp/svin_res.r', 'tmp/errfile.src', 'tmp/korm_temp.xls', 'tmp/svin_temp.xlsm']
    try:
        for temp in tempfile:
            if os.path.isfile(temp) == True:
                os.remove(temp)
    except:
        loggs.log_error(f'Ошибка при удалении {svinview}, {kormview}, {svintemp}, {kormtemp}, {errfile}')
