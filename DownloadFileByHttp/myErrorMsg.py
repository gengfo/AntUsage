'''
@author: Yinkan LI
@version: 2.0
@see: http://www.python-excel.org/
@note: install python 2.7.2, xlrd, xlwt and xlutils. for windows, add python to PATH
'''

IN_PATH = 'D:\GengFo\MyProgs\GitHub\AntUsage\DownloadFileByHttp\my.xls'
OUT_PATH = 'D:/project-git/arp-iris4/ARP_APP_Server/ARP_WAR/WebContent/js/com.oocl.ir4.arp.web/common/ErrorMessage.js'
OUT_PATH_JAVA = 'D:/project-git/arp-iris4/ARP_APP_Server/ARP_Domain/src/com/oocl/ir4/arp/constant/common/ErrorCodeConstant.java'

from xlrd import open_workbook, cellname

book = open_workbook(IN_PATH)
msgList = {}
for sheet in book.sheets():
    if sheet.name == 'Invoice' or sheet.name == 'Payment' or sheet.name == 'Common' or sheet.name == 'Financial Rule' or sheet.name == 'Invoice Matching' or sheet.name == 'Report' or sheet.name == 'Tax Invoice':
        #print sheet.nrows
        for row_index in range(1, sheet.nrows):
            #print sheet.cell(row_index, 1).value
            if sheet.cell(row_index, 1).value is not None and sheet.cell(row_index, 1).value != '':
                #c = str(int(sheet.cell(row_index, 1).value))
                c = str(sheet.cell(row_index, 1).value)
                m = str(sheet.cell(row_index, 8).value).replace('\n', '<br/>').replace('\"', '\'')
                if m == '':
                	m = str(sheet.cell(row_index, 4).value).replace('\n', '<br/>').replace('\"', '\'')
                print c + '    ->    ' + m
                msgList[c] = m

file = open(OUT_PATH, 'w')
header = '''Ext.ns("com.oocl.ir4.arp.web.common");
Ext.ns("arp");
arp.Msg = {
/**
 * @class arp.Msg
 * @author Yinkan Li
 */
 '''
file.writelines(header)

msgs = []
for msg in msgList:
    row = msg + ':' + '"' + msgList[msg] + '"'
    msgs.append(row)
    msgs.append(',')
msgs.pop()
file.writelines(msgs)

footer = '''
};
arp.Msg.get = function (code, encode) {
    var msg = '', args = [], i, l, needEncode = true;
    code = [].concat(code);
    if (code.length < 1) {
        return '';
    }
    if (typeof(encode) === 'boolean') {
        needEncode = encode;
    }
    for (i = typeof(encode) === 'boolean' ? 2 : 1, l = arguments.length; i < l; i++) {
        args = args.concat(arguments[i]);
    }
    for (i = 0, l = code.length; i < l; i++) {
        var message = arp.Msg[code[i]], tail = needEncode ? '<br/>' : '';
        if (Ext.isDefined(message)) {
            msg = msg + message.replace(/\{(\d+)\}/g, function (m, t) {
                return args[t];
            }) + tail;
        } else {
            msg = msg + 'Message ' + code + ' is undefined.' + tail;
        }
    }
    return msg;
};
    '''
file.writelines(footer)
file.close()
print 'Done JS'

file_java = open(OUT_PATH_JAVA, 'w')
header = '''package com.oocl.ir4.arp.constant.common;
public final class ErrorCodeConstant {'''
file_java.writelines(header)

msgs = []
for msg in msgList:
    row = 'public static final String ' + msg + ' = ' + '"' + msg + '";'
    msgs.append(row)
file_java.writelines(msgs)
footer = '}'
file_java.writelines(footer)
file_java.close()
print 'Done JAVA'


'''
#import
from xlrd import open_workbook, cellname

#workbook
read a xls file                :  book = open_workbook('Book1.xls')
get all sheets                 :  book.sheets
get a sheet                    :  book.sheet_by_index(index)
get sheet count                :  book.nsheets

#sheet
get row count                :  sheet.nrows
get column count            :  sheet.ncols
get sheet name                :  sheet.name

#cell
get cell                    :  sheet.cell(row_index, col_index)
get cell actual value        :  sheet.cell(row_index, col_index).value
get cell name(eg:A1)        :  cellname(row_index, col_index)
get cell value type            :  sheet.cell_type(row_index, col_index)
get cell value                :  sheet.cell_value(row_index, col_index)
get a list of cell in row        :  sheet.row_slice(start_index, end_index)
get a list of cell in col        :  sheet.col_slice(start_index, end_index)
get a list of cell type in row    :  sheet.row_types(start_index, end_index)
get a list of cell type in col    :  sheet.col_types(start_index, end_index)
get a list of cell value in row    :  sheet.row_values(start_index, end_index)
get a list of cell value in col    :  sheet.col_values(start_index, end_index)

#read and write txt file
#read:
file_object = open('thefile.txt')
try:
    all_the_text = file_object.read()
finally:
    file_object.close()
     
#write:
file_object = open('thefile.txt', 'w')
file_object.write(all_the_text)
file_object.writelines(list_of_text_strings)   #better performance
file_object.close()
'''
