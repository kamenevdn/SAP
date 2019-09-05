import sys, win32com.client
import pandas as pd


class sap():
    def __init__(self):
        self.connection = win32com.client.GetObject("SAPGUI").GetScriptingEngine.Connections[0]
    
    def session(self):
        return self.connection.Sessions[0]
    
    def error_print(self):
        """Вывод системных ошибок"""
        print('Ошибка')
        print(sys.exc_info())
        return
    
    def run(self, transaction):
        """Функция запуска транзакции
        transaction - str, код транзакции.
        """
        self.session().StartTransaction(Transaction=transaction)
        return
    
    def check_system(self, system_id):
        """Проверяем, в какой системе выполняется скрипт.
        Возвращает ошибку, если заявленная система не равна системе исполнения.
        system_id - str, идентификатор системы, в которой предполагается выполнение скрипта.
        """
        try:
            session = self.connection.Sessions[0]
            system = session.findById("wnd[0]").text.split('(')[0]
        except:
            self.error_print()
        finally:
            session = None
            if system != system_id:
                raise ValueError('Не та система!', system)
    
    def alvscroll(self, grid, max_rows=30):
        rows = grid.rowcount
        
        #Если строк в ALV больше, чем max_rows, то устраиваем прокрутку
        if rows > max_rows:
            visible_row = max_rows + 1
            
        #Скроллим по max_rows строк за раз    
            while visible_row < rows:
                grid.firstVisibleRow = visible_row
                visible_row = visible_row + max_rows

    def read_alv(self, grid, cols=None, tech_names = False, max_rows=30):
        """Функция для чтения экранной ALV-таблицы.
        Возвращает объект типа pandas.DataFrame со всеми колонками ALV (или только из массива cols).
        Параметры:
        grid: идентификатор ALV-grid таблицы на экране запущенной сессии.
              Например: grid = session.findById("wnd[0]/usr/cntlALV0101/shellcont/shell")
              
        cols: список технических названий полей таблицы, которые необходимо прочитать.
              Например: ['MATNR','MATKL']
              Если не задан, то используются все выведенные поля ALV-таблицы.
              
        tech_names: boolean, если True, то в качестве названий столбцов используются технические имена полей.
                    Если False, то берется краткое название.
        
        max_rows: количество строк для скроллинга таблицы за один раз. Необходимо для подгрузки длинных таблиц. По умолчанию 30.
        """
        rows = grid.rowcount
        
        self.alvscroll(grid, max_rows)#Если строк в ALV больше, чем max_rows, то устраиваем прокрутку
        
        data = []
        
        #Если список колонок не задан, то читаем все колонки с экрана
        if cols == None:
            cols = list(grid.ColumnOrder)
            
        cols_name = []
        if tech_names:
            cols_name = cols
        else:
            for col in cols:
                cols_name.append(grid.GetColumnTitles(col)[1])
            
        #Читаем каждую строку ALV-таблицы по указанным именам столбцов
        for row in range(0,rows):
            row_data = {}
            
            for col_name in cols:         
                col_value = grid.getcellvalue(row, col_name)
                if grid.GetColumnDataType(col_name) == 'int':
                    row_data[col_name] = self.replace_minus_int(col_value)
                #Под decimal может быть таймштамп, поэтому если в строке ':', то пропускаем
                elif grid.GetColumnDataType(col_name) == 'decimal' and ':' not in col_value:
                    row_data[col_name] = self.replace_minus_float(col_value)
                else:
                    row_data[col_name] = col_value
                    
                
            data.append(row_data)
        
        #Записываем таблицу в датафрейм
        dataframe = pd.DataFrame(data, columns = cols)
        dataframe.columns = cols_name
        
        return dataframe
    
    def replace_minus_int(self, string):
        """Функция замены минуса после символьного значения на минус перед int числом"""
        string = string.replace('.','')
        string = string.replace(' ','')
        string = string.replace(',','.')
        if len(string)>0:
            if string[-1] == '-':
                int_string = -int(string.replace('-',''))
            else:
                int_string = int(string)
            return int_string
        else:
            return 0
    
    def replace_minus_float(self, string):
        """Функция замены минуса после символьного значения на минус перед int числом"""
        string = string.replace('.','')
        string = string.replace(' ','')
        string = string.replace(',','.')
        if len(string)>0:
            if string[-1] == '-':
                float_string = -float(string.replace('-',''))
            else:
                float_string = float(string)
            return float_string
        else:
            return 0

def alvscroll(grid, max_rows=30):
    rows = grid.rowcount
    
    #Если строк в ALV больше, чем max_rows, то устраиваем прокрутку
    if rows > max_rows:
        visible_row = max_rows + 1
        
    #Скроллим по max_rows строк за раз    
        while visible_row < rows:
            grid.firstVisibleRow = visible_row
            visible_row = visible_row + max_rows

def read_alv(grid, cols, cols_name, max_rows=30):
    """Функция для чтения экранной ALV-таблицы.
    Возвращает объект типа pandas.DataFrame с колонками из массива cols.
    Параметры:
    grid: идентификатор ALV-grid таблицы на экране запущенной сессии. Например: grid = session.findById("wnd[0]/usr/cntlALV0101/shellcont/shell")
    cols: список технических названий полей таблицы, которые необходимо прочитать. Например: ['MATNR','MATKL']
    max_rows: количество строк для скроллинга таблицы за один раз. Необходимо для подгрузки длинных таблиц. По умолчанию 30.
    """
    rows = grid.rowcount
    
    alvscroll(grid, max_rows)#Если строк в ALV больше, чем max_rows, то устраиваем прокрутку
    
    data = []
    
    #Читаем каждую строку ALV-таблицы по указанным именам столбцов
    for row in range(0,rows):
        row_data = {}
        
        for col_name in cols:         
            col_value = grid.getcellvalue(row, col_name)
            row_data[col_name] = col_value
            
        data.append(row_data)
    
    #Записываем таблицу в датафрейм
    dataframe = pd.DataFrame(data, columns = cols)
    dataframe.columns = cols_name
    
    return dataframe




def replace_minus_int(string):
    """Функция замены минуса после символьного значения на минус перед int числом"""
    string = string.replace('.','')
    string = string.replace(' ','')
    if string[-1] == '-':
        string = -int(string.replace('-',''))
    else:
        string = int(string)
    return(string)
        
class script:
    def __init__(self):
#        self.sess = None
#        self.connection = None
#        self.application = None
#        self.SapGuiAuto = None
        """Функция, открывающая и возвращающая текущую SAP-сессию"""
        self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
        
        if not type(self.SapGuiAuto) == win32com.client.CDispatch:
            return
            
        self.application = self.SapGuiAuto.GetScriptingEngine
        
        if not type(self.application) == win32com.client.CDispatch:
            self.SapGuiAuto = None
            return
            
        self.connection = self.application.Children(0)
        
        if not type(self.connection) == win32com.client.CDispatch:
            self.application = None
            self.SapGuiAuto = None
            return
            
        self.sess = self.connection.Children(0)
        
        if not type(self.sess) == win32com.client.CDispatch:
            self.connection = None
            self.application = None
            self.SapGuiAuto = None
            return
        
#        return self.sess

    def session(self):
        """Функция, открывающая и возвращающая текущую SAP-сессию"""
      
        return self.sess
    
    def run(self, transaction):
        """Функция запуска транзакции
        Параметры:
        transaction: код транзакции, string
        """
        self.sess.StartTransaction(Transaction=transaction)
        return
    
    def session_close(self):
        """Закрытие текущей сессии скрипта"""
        self.sess = None
        self.connection = None
        self.application = None
        self.SapGuiAuto = None 
        return self.sess, self.connection, self.application, self.SapGuiAuto
    
    def error_print(self):
        """Вывод системных ошибок"""
        print('Ошибка')
        print(sys.exc_info())
        return

