# https://python.su/forum/topic/303/?page=2
# https://bevice.ru/posts/1178.html
# https://infostart.ru/public/82509/
# https://habr.com/post/139272/
# — Если вы выбираете поле типа Дата, то оно буде возвращено как объект PyTime. 
# Это специальный тип данных для передачи даты-времени в COM-соединении. 
# С ним не так удобно работать, как с привычным datetime. 
# Если передать этот объект в int(), то вернется timestamp, 
# из которого потом можно получить datetime методом fromtimestamp().

import time
import datetime
import os
import inspect
import json
import csv

import pythoncom
# import pywintypes
from win32com.client import Dispatch, DispatchEx, gencache

import logging
import logging.handlers
_log_level = logging.DEBUG
_LOGGER_NAME = __name__

_ABSPATH = os.path.dirname(os.path.abspath(__file__))


get1 = lambda obj, attr: getattr(obj, attr.encode('cp1251'))


class O1c( object ):
    def __init__(self, conn_str, auto=True):
        if not conn_str:
            raise ValueError('Ошибка инициализации: Необходимо установить строку подключения к ИБ 1С')
        else:
            self.conn_str = conn_str

        if not _no_db:
            self.db = mssqlDb(**_DATABASE)
            self.conn = self.db.connect()
            self.cur = self.conn.cursor()

        self.V83 = None
        self.insert_data = [] # массив кортежей для db.executemany
        self.bad_data = {} # словарь плохих данных, ключ - номер плохой записи, значение - содерж записи
        self.columns_count = 0 # количество столбцов в одной записи из результата запроса
        self.total_rows = 0 # количество записей которое вернул запрос
        self.query_text = "" # текст 1С SQL запроса

        self.timings = [] # массив временных меток выполнения
        self.t() # задаем первоначальную метку

        self.logger_name = _LOGGER_NAME
        self.log = self.init_logger(_log_level)

        # по умолчанию посли инициализации класса сразу создаем КОМ-подключение к 1С
        if auto:
            self.get_V83()

    def t(self, comment = ''):
        _t = time.time()
        
        if not comment:
            callerframerecord = inspect.stack()[1]
            frame = callerframerecord[0]
            info = inspect.getframeinfo(frame)
            
            comment = "Вызов: %s" % info.function

        if len(self.timings) < 1:
            tdict = (_t, 0, comment, ) # временая метка, дельта тайм, комментарий
        else:
            _dt = _t - self.timings[-1][0] # из последней записи берем первый элемент кортежа - время
            tdict = (_t, _dt , comment, ) # временая метка, дельта тайм, комментарий
            
            
        self.timings.append(tdict)


    def init_logger(self, level=logging.INFO):
        # def handleError(self, record):
        #     """
        #     Handle errors which occur during an emit() call.
        #
        #     This method should be called from handlers when an exception is
        #     encountered during an emit() call. If raiseExceptions is false,
        #     exceptions get silently ignored. This is what is mostly wanted
        #     for a logging system - most users will not care about errors in
        #     the logging system, they are more interested in application errors.
        #     You could, however, replace this with a custom handler if you wish.
        #     The record which was being processed is passed in to this method.
        #     """
        #     raiseExceptions = True
        #
        #     if raiseExceptions and sys.stderr:  # see issue 13807
        #         ei = sys.exc_info()
        #         try:
        #             traceback.print_exception(ei[0], ei[1], ei[2], None, sys.stderr)
        #             sys.stderr.write('Logged from file %s, line %s\n' % (record.filename, record.lineno))
        #         except IOError:
        #             print ("LOGGER emit error! Try to restart")
        #             logging.shutdown()
        #             reload(logging)
        #             self.log = self.init_logger(_log_level)
        #             # pass    # see issue 5971
        #         finally:
        #             del ei

        log = logging.getLogger(self.logger_name)
        log.setLevel(level)
        # formatter = logging.Formatter('%(levelname)-8s [%(asctime)s] %(message)s')
        formatter = logging.Formatter('%(asctime)s %(levelname)-9s %(message)s')
        
        # logger_filename = os.path.join(_SAVEPATH, '%s.log' % self.logger_name) # отключено логирование на сетевой диск
        logger_filename = os.path.join(_ABSPATH, '%s.log' % self.logger_name)

        try:
            # handler = logging.FileHandler(logger_filename)
            handler = logging.handlers.RotatingFileHandler(logger_filename, maxBytes=1024*1024, backupCount=5)
        except:
            # если сетевой путь недоступен для записи логов генерим лог в папке со скриптом
            # handler = logging.FileHandler('%s.log' % self.logger_name)
            handler = logging.handlers.RotatingFileHandler('%s.log' % self.logger_name, maxBytes=1024*1024, backupCount=5)

        handler.setLevel(level)
        handler.setFormatter(formatter)
        log.addHandler(handler)
        handler = logging.StreamHandler()
        handler.setLevel(level)
        handler.setFormatter(formatter)
        log.addHandler(handler)

        # logging.Handler.handleError = handleError
        return log


    def get_V83(self):
        if self.conn_str:
            try:
                self.log.debug('Начало инициализации 1С СОМ-соединения')
                self.V83 = gencache.EnsureDispatch("V83.COMConnector").Connect(self.conn_str) # http://timgolden.me.uk/python/win32_how_do_i.html
                self.log.debug('Установлено 1С COM-соединение')
                # self.t('Установлено COM-подключение к 1С')
            except Exception as e:
                self.log.critical("Не удалось инициализировать подлючение к 1С, работа невозможна!")
                raise SystemError('Ошибка инициализации: Не удалось инициализировать подлючение к 1С, работа невозможна! %s'%e)
                
        else:
            self.log.crtitical("Не задана строка подключения к 1С, работа невозможна!")
            raise SystemError('Ошибка инициализации: Необходимо установить текст 1С запроса')

    def check_V83(self):
        return False if self.V83 == None else True

    def wipe_query(self):
        if hasattr(self, 'query'): # повторное создание нового объекта Запрос - удаляем прошлый объект Запрос
            del self.query

    def wipe_result(self):
        """
            Удаляем данные от предыдущих запросов
                вызывается в момент инициализации нового объекта Запрос = значит пользователь готовиться получать новые данные
                вызывается при задании новых параметров запроса = значит пользователь готовиться получать новые данные по уже созданному объекту запроса
        """
        if hasattr(self, 'result'): # повторное создание нового объекта Запрос - удаляем результат прошлого запроса
            del self.result

    def show_exception(self, e):
        err_code = e[0]
        err_text = e[1]
        err_tuple = e[2]
        err_last_flag = e[3]
        
        err_1c_code = err_tuple[0]
        err_1c_version = err_tuple[1]
        err_1c_text = err_tuple[2]
        err_1c_flag0_bool = err_tuple[3]
        err_1c_flag1_int = err_tuple[4]
        err_1c_flag2_int = err_tuple[5]
        
        res = "%s (%i, %i)" %(err_1c_text, err_code, err_1c_code)
        return res.encode("windows-1251", errors="ignore")

    def load_query_text_from_file(self, params):
        assert isinstance(params, (list, tuple,)), "Вызов функции с неверным типом параметров - ожидается список или кортеж"

        loadpath = os.path.join(*params)
        
        assert os.path.exists(loadpath), "Файл не найден - %s" % loadpath
        
        with open(loadpath, 'r') as file:
            return file.read()

    def make_query(self, query_text):
        if not query_text:
            self.log.crtitical("Ошибка создания объекта Запрос: Необходимо установить строку запроса к ИБ 1С")
            raise ValueError('Ошибка создания объекта Запрос: Необходимо установить строку запроса к ИБ 1С')

        if isinstance(query_text, (tuple, list,)):
            query_text = self.load_query_text_from_file(query_text)

        if query_text == self.query_text:
            if self.query:
                return

        self.query_text = query_text

        self.wipe_query()
        self.wipe_result()

        try:
            if not self.check_V83():
                self.get_V83()
            else:
                self.log.debug('Создаем 1С объект "Запрос"')
                self.query = self.V83.NewObject("Query", self.query_text)
                self.log.debug('1С объект "Запрос" создан успешно')
        except Exception as e:
            err = self.show_exception(e)
            raise SystemError('Ошибка создания объекта Запрос, работа невозможна! %s'% err)
                        
        self.t('Создан объект Запрос 1С')

    def check_query(self):
        if not hasattr(self, 'query'):
            self.log.critical("Вызов метода Выполнить без созданного объекта Запрос, работа невозможна")
            raise SystemError('Вызов метода Выполнить без созданного объекта Запрос, работа невозможна')
        return True

    def setp(self, pname, pvalue):
        """
            Установка параметров запроса
        """
        self.check_query()
        
        if hasattr(self, 'result'): # повторное создание нового объекта Запрос - удаляем результат прошлого запроса
            del self.result

        try:
            self.query.SetParameter(pname, pvalue)
            self.log.debug('Установлен параметр запроса "%s" = %s' %(pname, pvalue))
        except Exception as e:
            self.log.critical("Ошибка установки параметра 1С запроса: %s" % e)
            raise SystemError("Ошибка установки параметра 1С запроса")


    def executebatch(self):
        self.check_query()
        try:
            self.log.debug('Начало выполнения пакетного запроса')
            queryresult = self.query.ExecuteBatch()
            self.log.debug('Пакетный запрос выполнен')
        except Exception as e:
            err = "Ошибка выполнения пакетного 1С запроса: %s" % self.show_exception(e)
            self.log.critical(err)
            raise SystemError(err)

        columns = []
        for r in range(len(queryresult)):
            columns_count = queryresult[r].Columns.Count()
            cols = []
            for i in range(columns_count):
                name = queryresult[r].Columns.Get(i).Name
                cols.append(name)
            columns.append(tuple(cols))

        self.result = queryresult
        self.columns = columns
        self.index = len(columns)

        # self.log.info('Пакетный запрос вернул набор из %i сущностей, столбцы: %s' %(len(columns), '|'.join([", ".join(p) for p in columns]) )
        # columns_names = " >>>".join([", ".join([str(b).encode('utf-8', errors='replace') for b in aa]) for aa in columns]) # '1, 2| 3, 4| 5, 6'
        columns_names = " >>>".join([", ".join([b for b in aa]) for aa in columns]) # '1, 2| 3, 4| 5, 6'
        self.log.info('Пакетный запрос вернул набор из %i сущностей' % len(columns) )
        self.log.info('столбцы: >>>%s' % columns_names )

    def yieldbatch_dict(self, index=0):
        if not hasattr(self, 'result'):
            self.executebatch()
            # err = 'yieldbatch_dict: нет результатов запроса, сначала выполните пакетный запрос'
            # self.log.error(err)
            # raise SystemError(err)

        # если индекс отрицательный то берем сущность "с краю" аля питоник..
        if index < 0:
            index = self.index + (index + 1)

        try:
            res = self.result[index]
        except IndexError:
            err = 'yieldbatch_dict: В результатах пакетного запроса нет сущностей с индексом %i' % index
            self.log.error(err)
            raise IndexError(err)

        try:
            columns = self.columns[index]
        except IndexError:
            err = 'yieldbatch_dict: В результатах пакетного запроса нет данных по именам столбцов из сущности с индексом %i' % index
            self.log.error(err)
            raise IndexError(err)


        selection = res.Select()

        rec_num = 0 # количество извлеченных записей
        while selection.Next():
            _dict = {}
            # print 50*'='
            for c in columns:
                value = get1(selection, c)
                oldvalue = value
                value = self.V83.XMLString(value) # преобразование через внутреннюю 1С функцию XMLString
                # print type(value), value, type(oldvalue), oldvalue
                _dict[c] = value
            rec_num += 1
            yield _dict


        if index > 0:
            self.log.info('Извлечено %i записей из сущности с индексом %i' % (rec_num, index))
        else:
            self.log.info('Извлечено %i записей' % rec_num)

        self.total = rec_num


    def yieldbatch_tuple(self, index=0):
        for d in self.yieldbatch_dict(index):
            yield tuple([d[column] for column in self.columns[index]])


    def getbatch_headers(self, index=0):
        if not hasattr(self, 'columns'):
            self.executebatch()

        return self.columns[index]


    def savecsv(self, filename=None, data=None, index = 0, headers = None, enc = "windows-1251", delimiter=";", oper_func = None):
        """
            Сохранение данных полученных из запроса в csv
                filename    = имя файла для сохранения данных
                data        = массив кортежей с данными для сохранения
                index       = индекс массива из результатов пакетного запроса, умолч = 0
                enc         = кодировка для выгрузки, умолч = "windows-1251"
                delimiter   = разделитель для csv файла, умолч = ';'
                oper_func   = функция преобразования (если задана) данных полученных из запроса
        """
        if filename is None:
            filename = 'query_%s_index_%i' %(self.get_now_str("%Y%m%d-%H%M%S"), index)
            err = 'Не указано имя файла csv для сохранения данных, "%s" назначено принудительно' % filename
            self.log.warning(err)
            # raise SystemError(err)

        if data is None:
            if not hasattr(self, 'result'):
                self.executebatch()

            # data = [d for d in self.yieldbatch_tuple(index)]
            data = self.yieldbatch_tuple(index)

        with open(filename, 'w', newline = "") as f:
            # json.dump(o.all_(), f)
            writer = csv.writer(f, dialect='excel', delimiter=delimiter)
            # writer.writerow([s.encode("windows-1251") for s in o.columns])
            if headers is None:
                headers = self.getbatch_headers(index)
            # elif len(headers) != len(data):
                # err = 'Длина переданного массива заголовков и набора данных в data не соответствуют'
                # self.log.error(err)
                # raise IndexError(err)


            writer.writerow([h.encode(enc) for h in headers])

            try: # если попадаются данные с русскими буквами конвертируем и записываем еще раз
                writer.writerows( self.converted_csv_data(source = data, enc=enc, convert_floats=True, oper_func = oper_func) )
            except UnicodeEncodeError:
                writer.writerows(data)


    def execute(self):
        self.check_query()
        try:
            self.query_result = self.query.Execute()
            self.result = self.query_result.Unload()
        except Exception as e:
            self.log.critical("Ошибка выполнения 1С запроса: %s" % e)
            raise SystemError("Ошибка выполнения 1С запроса: %s" % e)

        self.total = len(self.result)

        columns_count = self.result.Columns.Count()
        self.columns = []
        for i in range(columns_count):
            name = self.result.Columns.Get(i).Name
            self.columns.append(name)

        columns_names = ", ".join(['"%s"'%f for f in self.columns])
        self.log.info("1C запрос вернул: %i записей, %i колонок (%s)" % (self.total, self.columns_count, columns_names))

        self.t('Запрос 1С выполнен')
            
        
    def yield_dict(self, is_rownum = False, oper_func = None):

        rownum = 0
        for t in self.yield_tuple(is_rownum = is_rownum, oper_func = oper_func):
            _dict = {}

            for column_position, column_name in enumerate(self.columns):
                _dict[column_name] = t[column_position]

            rownum += 1
            yield _dict

        self.t('Объект Запрос: извлечено %i записей (словарей)' % rownum)

    def yield_tuple(self, is_rownum = False, oper_func = None):
        """
            Возвращаем по одной записи из результата запроса
            
            Записи возвращаются в виде кортежа, если выставлен флаг is_rownum - первая запись в кортеже номер записи
        """
        if self.total < 1:
            return
        
        if not hasattr(self, 'result') or not self.result:
            self.execute()
            # self.log.critical("Запрос не вернул данных, извлчение данных невозможно")
            # raise ValueError('Запрос не вернул данных, извлчение данных невозможно')

        

        bad_data = {}
        for rownum, row in enumerate(self.result):

            _tuple = []
            if is_rownum:
                _tuple.append(rownum)

            for c in self.columns:
                attr = get1(row, c)
                
                if attr is None:
                    _tuple.append(None)
                
                if isinstance(attr, (str, unicode, )):
                    if len(attr) > 0:
                        _tuple.append( attr.strip() )
                    else:
                        _tuple.append( 'ПУСТО' )
                
                if isinstance(attr, (int, float, bool, )):
                    _tuple.append(attr)
                    
                if 'time' in unicode(type(attr)):
                    try:
                        time_string = time.strftime("%Y-%m-%d %H:%M:%S. 0", time.localtime(int(attr)))
                    except Exception: # если тайм объект возвращает нечто несуразное
                        time_string = None

                    _tuple.append(time_string)

            columns_len = len(self.columns) + 1 if is_rownum else len(self.columns)
                
            if len(_tuple) != columns_len: # вернулось неверное количество данных, какаято из ячеек пуста
                # print rownum, len(_tuple), _tuple
                _dict = {c : unicode(get1(row, c)) if not None else "empty" for c in self.columns}
                bad_data[rownum] = _dict # словарь словарей: ключ - номер записи, значение - словарь в котором ключи: назв. колонок, значения - значения записи
                continue
            else:
                # insert_data.append(tuple(_tuple)) # конвертируем в кортеж и суём в список
                _tuple = tuple(_tuple)
                if oper_func and hasattr(oper_func, '__call__'): # если задана oper_func, то
                    _tuple = oper_func(_tuple) # выполняем oper_func и результат возвращаем в yield
                    # print _tuple, type(_tuple)
                yield _tuple # конвертируем в кортеж и суём в список
        
        self.t('Объект Запрос: извлечено %i записей (кортежей)' % rownum)

        if bad_data:
            jsonfilename = '_bad.json'
            with open(jsonfilename, 'w') as outfile:
                json.dump(bad_data, outfile)
            
            self.log.error("%i записей не внесено, не внесенные данные сохранены в %s" %(len(bad_data), jsonfilename))
            self.bad_data = True

            self.t('Объект Запрос: из них плохих %i записей' % len(bad_data))


    def all_(self, is_rownum = False, oper_func = None):
        """
            Сразу возвращаем массив кортежей
        """
        _list = []
        for f in self.yield_tuple(is_rownum = is_rownum, oper_func = oper_func):
            _list.append(f)
        
        return _list



    def yield_date(self, startdate = datetime.datetime(2018, 3, 1, 0, 0, 0, 0),  days_to_split = 7):
        """
        Параметры:
            1й - datetime.datetime  - начало выборки
            2й - int                - размер деления в днях
        """
        finaldate = ( datetime.datetime.now() - datetime.timedelta(1) ).replace(hour = 23, minute = 59, second = 59, microsecond = 999999)
        
        while startdate < finaldate:
            enddate = (startdate + datetime.timedelta(days_to_split - 1)).replace(hour=23, minute=59, second=59, microsecond = 999999)
            
            if finaldate - startdate < datetime.timedelta(days_to_split):
                enddate = finaldate

            yield (startdate, enddate)

            startdate = (startdate + datetime.timedelta(days_to_split)).replace(hour = 0, minute = 0, second = 0, microsecond = 0)


    def ndays_from_yesterday(self, n = 7):
        yesterday = ( datetime.datetime.now() - datetime.timedelta(1) ).replace(hour = 23, minute = 59, second = 59, microsecond = 999999)
        nday = (yesterday - datetime.timedelta(n-1)).replace(hour = 0, minute = 0, second = 0, microsecond = 0)
        return (nday, yesterday)

    def get_now_str(self, format_str = "%d-%m-%Y %H:%M:%S" ):
        return time.strftime(format_str, time.localtime())

    def converted_csv_data(self, source = None, enc = "windows-1251", convert_floats = False, oper_func = None):
        """
            convert_csv_data(self, source = o.all_(), enc = "windows-1251")
            source - источник (массив кортежей)
            enc - кодировка в которую конвертируем
            convert_floats - если нужны числа с плавающей точкой с запятой в качестве делителя десятичной части
        """
        def localize_floats(row):
            return [
                str(el).replace('.', ',') if isinstance(el, float) else el 
                for el in row
            ]
            
        if not source:
            source = self.all_()

        csv_data = []
        for t in source:
            _t = []

            if convert_floats:
                t = localize_floats(t)

            for i in t:
                if isinstance(i, bool):
                    _t.append(unicode(i))
                    continue
                if isinstance(i, (str, unicode,)):
                    try:
                        _t.append(i.encode(enc, errors='replace'))
                    except Exception as e:
                        self.log.critical("convert_csv_data(): Ошибка кодировки, исключение: %s "%e)
                        return
                    continue
                _t.append(i)

            if oper_func:
                _t = oper_func(_t)

            csv_data.append(_t)
        return csv_data





if __name__ == '__main__':
    pass
