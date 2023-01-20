import xlrd

#基本数据
half = "QWERTYUIOPASDFGHJKLZXCVBNMqwertyuiopasdfghjklzxcvbnm1234567890,.?:;()=-+*/（）「」『』<>&$@#%^*[]"
shorter = "iIl1()![]"
colorDict = {
    'BS':'white',
    'HS':'#A9A7A1',
    'B':'#9FFBFC',
    'H':'#ED9595',
    'S':'#86CAFF',
    'L':'#EEA9EF',
    'F':'#80F4D7',
    'Y':'#FFF29F',
    'C':'#84CC35',
    'CN':'#E6C378'
}
elements = ['BS','HS','B','H','S','L','F','Y','C','CN']

#读取表格数据
excel = xlrd.open_workbook("设置.xls")
table = excel.sheet_by_index(0)
LineNum = table.nrows
ColNum = table.ncols

settings = table.col_values(1)
settings[0] = int(settings[0])
if(settings[1] != '*'): settings[1] = int(settings[1])
settings[2] = int(settings[2])
if(settings[3] != '*'): settings[3] = int(settings[3])
print(settings)
