import os
import csv
import string
import xlrd
from db_mod import *


# 导入初三中考报名所有学生
# 增加届别字段
# 增加班级字段
SIGN_KS = ('signid','name','sex','idcode',
    'schcode','graduation_year','classcode','classnum','nation',
    'gsrid','ssrid','birth','regtype','hometel','addr','periodtel','sch','zhtype')

sch_datas = {
    '5000' : '宿州泗县招生办',
    '5001' : '宿州泗县中学',
    '5002' : '宿州泗县二中',
    '5003' : '宿州泗县三中',
    '5004' : '宿州泗县思源学校',
    '5005' : '宿州泗县草庙中学',
    '5006' : '宿州泗县黑塔中学',
    '5007' : '宿州泗县刘圩中学',
    '5008' : '宿州泗县山头中学',
    '5009' : '宿州泗县大庄中学',
    '5010' : '宿州泗县瓦坊中学',
    '5011' : '宿州泗县黄圩中学',
    '5012' : '宿州泗县杨集中学',
    '5013' : '宿州泗县屏山中学',
    '5014' : '宿州泗县长沟中学',
    '5015' : '宿州泗县草沟中学',
    '5016' : '宿州泗县丁湖中学',
    '5017' : '宿州泗县城南中学',
    '5018' : '宿州泗县墩集中学',
    '5019' : '宿州泗县双语中学',
    '5020' : '宿州泗县泗州学校',
    '5021' : '宿州泗县育才学校',
    '5022' : '宿州泗县灵童学校',
    '5023' : '宿州泗州实验学校',
    '5024' : '宿州泗县风华学校',
    '5025' : '宿州泗县光明学校',

    }

# 转学代码zhtype说明：
# 1 县外转入，2 县外转入，县内转
# 3 县内转学，4 无转学记录 0 历届


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
            if 'graduation_year' in datas:
                datas['graduation_year'] = str(int(datas['graduation_year']))
            datas['sch'] = sch_datas[datas['schcode']]
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
    ck = 'X' if idcode[-1] == 'x' else idcode[-1]
    if checkcode != ck:
        return '校验出错！'
    if int(idcode[-2]) % 2 == 0 and stud.sex == '1':
        return '性别码出错！'
    if int(idcode[-2]) % 2 == 1 and stud.sex == '2':
        return '性别码出错！'

@db_session
def check_stud_idcode():
    print('身份证号码校验错误信息：')
    for s in select(s for s in SignAll):
        if s.idcode and s.idcode[:-1].isdigit():
            ret = check_idcode(s)
            if ret:
                print(ret,':',s.sch,s.name,s.idcode,s.sex)

if __name__ == '__main__':
    db.bind(**DB_PARAMS)
    db.generate_mapping(create_tables=True)

    gath_data(SignAll,SIGN_KS,'data//sign',0) # 末尾行无多余数据
    check_stud_idcode()
