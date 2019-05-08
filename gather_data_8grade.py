import os
import string
import csv
import xlrd
from db_mod import *
import xlwt

# 导入初二在校生（excel格式）到数据库中
# 表存放子目录：gradey8之中

GRADE_KS = ('sch','grade','sclass','gsrid','ssrid',
    'dsrid','name','idcode','regtype','sex','nation')

def get_files(directory):
    files = []
    files = os.listdir(directory)
    files = [f for f in files if f.endswith('.xls') or f.endswith('.xlsx')]
    files = [os.path.join(directory,f) for f in files]
    return files

@db_session
def gath_data(tab_obj,ks,chg_dir,grid_end=1,start_row=1):
    """start_row＝1 有一行标题行；gred_end=1 末尾行不导入"""
    files = get_files(chg_dir)
    for file in files:
        wb = xlrd.open_workbook(file)
        ws = wb.sheets()[0]
        nrows = ws.nrows
        for i in range(start_row,nrows-grid_end):
            datas = ws.row_values(i)
            datas = {k:v for k,v in zip(ks,datas) if v}
            tab_obj(**datas)

def check_idcode(stud):
    '''身份证号验证程序'''
    all_chars = string.digits + 'X'
    idcode = stud.idcode
    if not all(c in string.digits for c in idcode[:-1]) or idcode[-1] not in all_chars:
        if 'x' in idcode:
            return 'x应为大写！'
        else:
            return '包含不正确的字符！'

    checkcodes = ['1', '0', 'X','9', '8', '7', '6', '5', '4', '3', '2']
    wi = [7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2]
    s = sum((i*int(j) for i,j in zip(wi,idcode[:-1])))
    checkcode = checkcodes[s % 11]
    info = ''
    if 'x' in idcode:
        print(stud.name, stud.idcode, 'x应为大写！')
    ck = 'X' if idcode[-1] == 'x' else idcode[-1]
    if checkcode != ck:
        return '校验出错！'
    if int(idcode[-2]) % 2 == 0 and stud.sex == '男':
        return '性别码出错！'
    if int(idcode[-2]) % 2 == 1 and stud.sex == '女':
        return '性别码出错！'

@db_session
def check_stud_idcode():
    print('身份证号码校验错误信息：')
    for s in select(s for s in GradeY8):
        if s.idcode and s.idcode[:-1].isdigit():
            if 'x' in s.idcode:
                print(s.idcode, s.name, s.sch)
                s.idcode = s.idcode.upper()
            ret = check_idcode(s)
            if ret:
                print(ret,':',s.sch,s.name,s.idcode,s.sex)

def save_datas_xlsx(filename,datas):
    #将一张表的信息写入电子表格中XLSX文件格式
    import xlsxwriter
    w = xlsxwriter.Workbook(filename)
    w_sheet = w.add_worksheet('sheet1')
    for rowi,row in enumerate(datas):
        for coli,celld in enumerate(row):
            w_sheet.write(rowi,coli,celld)
    w.close()


def save_datas(filename,datas):
    #将一张表的信息写入电子表格中
    #文件内容不能太大，否则输出文件会出错。
    w = xlwt.Workbook(encoding='utf-8')
    ws = w.add_sheet('sheet1')
    for rowi,row in enumerate(datas):
        rr = ws.row(rowi)
        for coli,celld in enumerate(row):
            if isinstance(celld,int):
                rr.set_cell_number(coli,celld)
            elif isinstance(celld,str):
                rr.set_cell_text(coli,celld)
    #rr = ws.row(4)
    w.save(filename)

@db_session
def dump_bm(sch=None):
    if sch:
        datas = [['全国学籍号', '学籍号', '姓名', '身份证号', '毕业年份', '届别', '性别'],]
        for s in select(s for s in GradeY8 if s.sch == sch):
            datas.append([s.gsrid, s.ssrid, s.name, s.idcode, '2020', '应届', s.sex])
        # print(sch)
        # save_datas_xlsx(sch+'.xlsx', datas)
        save_datas(sch+'.xls', datas)
    else:
        for sch in select(s.sch for s in GradeY8):
            datas = [['全国学籍号', '学籍号', '姓名', '身份证号', '毕业年份', '届别', '性别'],]
            for s in select(s for s in GradeY8 if s.sch == sch):
                datas.append([s.gsrid, s.ssrid, s.name, s.idcode, '2020', '应届', s.sex])
            # print(sch)
            # save_datas_xlsx(sch+'.xlsx', datas)
            save_datas(sch+'.xls', datas)


if __name__ == '__main__':
    db.bind(**DB_PARAMS)
    db.generate_mapping(create_tables=True)

    # gath_data(GradeY8,GRADE_KS,'data\\gradey8',0) # 末尾行无多余数据
    # check_stud_idcode()
    dump_bm(sch='泗县泗州学校')