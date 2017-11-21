# -*- coding: UTF-8 -*-
import os
import os.path
import logging
import logging.config
#import io
import sys
import configparser
#import time
import shutil
import openpyxl                       # Для .xlsx
#import xlrd                          # для .xls
from   price_tools import getCellXlsx, getCell, quoted, dump_cell, currencyType, subInParentheses



def convert_sheet( dealerName, sheet):
    confName = ('cfg_'+dealerName+'.cfg').lower()
    csvFName = ('csv_'+dealerName+'.csv').lower()
    if not os.path.exists( confName ) :
        log.error( 'Нет файла конфигурации '+confName)
        return
    # Прочитать конфигурацию из файла
    in_columns_j, out_columns = config_read( confName )
    ssss = []
#   for i in range(1, sheet.nrows) :                                  # xls
    for i in range(1, sheet.max_row +1) :                             # xlsx
        i_last = i
        try:
            impValues = getXlsxString(sheet, i, in_columns_j)
            if impValues['цена'] == '0' :                           # Пустая строка
                pass
                print( 'Пустая строка. i=', i )
    
            else :                                                  # Информационная строка
                sss = []                                            # формируемая строка для вывода в файл
                for outColName in out_columns.keys() :
                    shablon = out_columns[outColName]
                    for key in impValues.keys():
                        if shablon.find(key) >= 0 :
                            shablon = shablon.replace(key, impValues[key])
                    if (outColName == 'закупка') and ('*' in shablon) :
                        #print(shablon)
                        kkkk = float( shablon[  shablon.find('*')+1 : ] )
                        #print(kkkk)
                        vvvv = float( shablon[ :shablon.find('*')     ] )
                        #print(vvvv)
                        shablon = str( kkkk * float( vvvv ) )
                    sss.append( quoted( shablon))
                ssss.append(','.join(sss))
                    
        except Exception as e:
            log.debug('Exception: <' + str(e) + '> при обработке строки ' + str(i) +'<' + '>' )
            raise e

    log.info('Обработано ' +str(i_last)+ ' строк.')
    f2 = open( csvFName, 'w', encoding='cp1251')
    strHeader = ','.join( out_columns.keys() ) + ','
    f2.write( strHeader + '\n' )
    data = ',\n'.join(ssss) +','
    bbbb = data.encode(encoding='cp1251', errors='replace')
    data = bbbb.decode(encoding='cp1251')
    f2.write(data)
    f2.close()
    if os.path.exists('c://AV_PROM/prices/'+dealerName) : shutil.copy2( csvFName, 'c://AV_PROM/prices/'+dealerName+'/'+csvFName)




def appendSensor( shablon, impValues):
    ss = impValues['тип_сенсора']
    if ss != 'нет' :  shablon = shablon + '\nтип сенсора: ' + ss
    ss = impValues['количество_точек_касания']
    if ss != 'нет' :  shablon = shablon + '\nколичество точек касания: ' + ss
    return shablon



def currencyType( row, col, sheet ):
    '''
    Функция анализирует "формат ячейки" таблицы excel, является ли он "денежным"
    и какая валюта указана в этом формате.
    Распознаются не все валюты и способы их описания.
    '''
    c = sheet.cell( row=row, column=col )
    '''                                                  # -- для XLS
    xf = sheet.book.xf_list[c.xf_index]
    fmt_obj = sheet.book.format_map[xf.format_key]
    fmt_str = fmt_obj.format_str
    '''                                                  # -- для XLSX
    fmt_str = c.number_format

    if 'р' in fmt_str:
        val = 'RUB'
    elif '\xa3' in fmt_str:
        val = 'GBP'
    elif chr(8364) in fmt_str:
        val = 'EUR'
    elif (fmt_str.find('USD')>=0) or (fmt_str.find('[$$')>=0) :
        val = 'USD'
    else:
        val = ''
    return val



def getXlsxString(sh, i, in_columns_j):
    impValues = {}
    for item in in_columns_j.keys() :
        j = in_columns_j[item]
        if item in ('закупка','продажа','цена') :
            if getCellXlsx(row=i, col=j, isDigit='N', sheet=sh).find('Call') >=0 :
                impValues[item] = '0.1'
            else :
                impValues[item] = getCellXlsx(row=i, col=j, isDigit='Y', sheet=sh)
            #print(sh, i, sh.cell( row=i, column=j).value, sh.cell(row=i, column=j).number_format, currencyType(sh, i, j))
        elif item == 'валюта_по_формату':
            impValues[item] = currencyType(row=i, col=j, sheet=sh)
        else:
            impValues[item] = getCellXlsx(row=i, col=j, isDigit='N', sheet=sh)
    return impValues



def getXlsString(sh, i, in_columns_j):
    impValues = {}
    for item in in_columns_j.keys() :
        j = in_columns_j[item] -1
        if item in ('закупка','продажа','цена') :
            if getCell(row=i, col=j, isDigit='N', sheet=sh).find('Звоните') >=0 :
                impValues[item] = '0.1'
            else :
                impValues[item] = getCell(row=i, col=j, isDigit='Y', sheet=sh)
            #print(sh, i, sh.cell( row=i, column=j).value, sh.cell(row=i, column=j).number_format, currencyType(sh, i, j))
        elif item == 'валюта_по_формату':
            impValues[item] = currencyType(row=i, col=j, sheet=sh)
        else:
            impValues[item] = getCell(row=i, col=j, isDigit='N', sheet=sh)
    return impValues



def config_read( cfgFName ):
    log.debug('Reading config ' + cfgFName )
    
    config = configparser.ConfigParser()
    if os.path.exists(cfgFName):     config.read( cfgFName, encoding='utf-8')
    else : log.debug('Не найден файл конфигурации.')

    # в разделе [cols_in] находится список интересующих нас колонок и номера столбцов исходного файла
    in_columns_names = config.options('cols_in')
    in_columns_j = {}
    for vName in in_columns_names :
        if ('' != config.get('cols_in', vName)) :
            in_columns_j[vName] = config.getint('cols_in', vName) 
    
    # По разделу [cols_out] формируем перечень выводимых колонок и строку заголовка результирующего CSV файла
    out_columns_names = config.options('cols_out')
    out_columns = {}
    for vName in out_columns_names :
        if ('' != config.get('cols_out', vName)) :
            out_columns[vName] = config.get('cols_out', vName) 

    return in_columns_j, out_columns



def convert2csv( dealerName ):
    fileNameIn = 'new_'+dealerName+'.xlsx'
    book = openpyxl.load_workbook(filename = fileNameIn, read_only=False, keep_vba=False, data_only=False)  # xlsx
    sheet = book.worksheets[0]                                                                              # xlsx                               
    log.info('-------------------  '+sheet.title +'  ----------')                                           # xlsx
#   sheetNames = book.get_sheet_names()                                                                     # xlsx

#   book = xlrd.open_workbook( fileNameIn.encode('cp1251'), formatting_info=True)                       # xls
#   sheet = book.sheets()[0]                                                                            # xls
#   log.info('-------------------  '+sheet.name +'  ----------')                                        # xls
    convert_sheet( dealerName, sheet)



def make_loger():
    global log
    logging.config.fileConfig('logging.cfg')
    log = logging.getLogger('logFile')



def main( dealerName):
    make_loger()
    log.info('         '+dealerName )
    convert2csv( dealerName )
    if os.path.exists( 'python.log') : shutil.copy2( 'python.log', 'c://AV_PROM/prices/' + dealerName +'/python.log')
    if os.path.exists( 'python.1'  ) : shutil.copy2( 'python.log', 'c://AV_PROM/prices/' + dealerName +'/python.1'  )



if __name__ == '__main__':
    myName = os.path.basename(os.path.splitext(sys.argv[0])[0])
    mydir    = os.path.dirname (sys.argv[0])
    print(mydir, myName)
    main( 'crestron')
