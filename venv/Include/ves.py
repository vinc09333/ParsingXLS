# -*- coding: cp1251 -*-
import controller
from parsingxls import ParsingXLS
#Класс Run, класс который объединяет в себе вссе модули необходимые для работы.
# Функция run_operating принимает в себя param = '-s' или '-k'.
#В соответствии с указанным параметром запускаются все необходимые функции в которые передается указанный параметр
class Run(object):
    def run_operating(param = None):
        if param == '-s':
            #controller.mounting('-s')
            ParsingXLS.parsing('-s')
            ParsingXLS.bugtracking('-s')
        if param == '-k':
            #controller.mounting('-k')
            ParsingXLS.parsing('-k')
            ParsingXLS.bugtracking('-k')