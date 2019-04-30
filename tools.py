
import string
import itertools
import os, xlrd
import datetime
from pony.orm import *

db = Database()

# 自动生成数据库表的列数和表名
models = [{'colnum':12,'classname':'Nodata'},{'colnum':11,'classname':'Nameerr'},
            {'colnum':12,'classname':'Nodatasrc'},{'colnum':11,'classname':'Nameerrsrc'}]


def gen_model(colnum, classname):
    class_str = ["class %s(db.Entity):" % classname, ]
    for i,colname in zip(range(colnum),itertools.combinations('s'+string.ascii_lowercase,2)):
        class_str.append('    %s = Optional(str, nullable=True)' % ''.join(colname))
    class_str = '\n'.join(class_str)
    return class_str

def gen_all_models(models):
    class_strs = []
    for model in models:
        class_strs.append(gen_model(**model))
    return '\n'.join(class_strs)


class_strs = gen_all_models(models)
exec(class_strs)

# set_sql_debug(True)
db.bind(provider='sqlite', filename='my.db', create_db=True)
db.generate_mapping(create_tables=True)

def get_files(directory):
    files = []
    files = os.listdir(directory)
    files = [f for f in files if f.endswith('.xls') or f.endswith('.xlsx')]
    files = [os.path.join(directory,f) for f in files]
    return files

def save_datas_xlsx(filename,datas):
    #将一张表的信息写入电子表格中XLSX文件格式
    import xlsxwriter
    w = xlsxwriter.Workbook(filename)
    w_sheet = w.add_worksheet('sheet1')
    for rowi,row in enumerate(datas):
        for coli,celld in enumerate(row):
            w_sheet.write(rowi,coli,celld)
    w.close()

@db_session
def gather_file_datas(tab_obj,filename,row_deal_function=None,grid_end=0,start_row=1):
    """start_row＝1 有一行标题行；gred_end=1 末尾行不导入"""
    """row_del_function 为每行的数据类型处理函数，不传则对数据类型不作处理 """
    wb = xlrd.open_workbook(filename)
    ws = wb.sheets()[0]
    nrows = ws.nrows
    for i in range(start_row,nrows-grid_end):
        datas = ws.row_values(i)
        if row_deal_function:
            datas = row_deal_function(datas)
        datas = {''.join(k):v for k,v in zip(itertools.combinations('s'+string.ascii_lowercase,2),datas) if v}
        tab_obj(**datas)

def gather_dir_datas(dir, tab_obj, row_deal_function=None, grid_end=0, start_row=1):
    if not os.path.exists(mydir):
    print('Directory is not exist.')
    return
    filenames = get_files(mydir)
    for filename in filenames:
        gather_file_datas(tab_obj,filename,row_deal_function=None,grid_end=0,start_row=1)
