import xlrd
from db_mod import *


# 处理仅有外地转入记录学生的定向
# 户籍地信息

def save_datas_xlsx(filename,datas):
    #将一张表的信息写入电子表格中XLSX文件格式
    import xlsxwriter
    w = xlsxwriter.Workbook(filename)
    w_sheet = w.add_worksheet('sheet1')
    for rowi,row in enumerate(datas):
        for coli,celld in enumerate(row):
            w_sheet.write(rowi,coli,celld)
    w.close()

def get_data(filename,headline=True):
     #既可以打开xls类型的文件，也可以打开xlsx类型的文件
    start = 1 if headline else 0
    datas = []
    w = xlrd.open_workbook(filename)
    ws = w.sheets()[0]
    nrows = ws.nrows
    for i in range(start,nrows):
        data = ws.row_values(i)
        datas.append(data)
    return datas

@db_session
def get_studs():
    data_title = ['中考报名号','姓名','性别','身份证号','学校','班级代码','是否本地户口']
    datas = [data_title, ]
    for stud in select(s for s in SignAll if s.zhtype==1 and s.idcode[:6] != '341324'):
        datas.append([stud.signid,stud.name,stud.sex,stud.idcode,stud.sch,stud.classcode])
    save_datas_xlsx('县外转入需确定户籍地.xlsx', datas)

#本地户口填1，其余不填

@db_session
def import_regaddr(filename='县外转入需确定户籍地.xlsx'):
    datas = get_data(filename)
    for data in datas:
        stud = select(s for s in SignAll if s.signid==data[0]).first()
        if stud:
            if data[-1] and int(data[-1]) == 1:
                stud.regaddr = True
            else:
                stud.regaddr = False
        else:
            print(data, '无此考生！')



if __name__ == '__main__':
    db.bind(**DB_PARAMS)
    db.generate_mapping(create_tables=True)

    # get_studs()
    import_regaddr()