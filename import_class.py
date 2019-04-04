import xlrd
from db_mod import *

## 导入临时数据（班级进入数据库表）
@db_session
def import_classcode(start_row=1,grid_end=0):
    wb = xlrd.open_workbook('bj.xls')
    ws = wb.sheets()[0]
    nrows = ws.nrows
    for i in range(start_row,nrows-grid_end):
        signid,classcode = ws.row_values(i)
        stud = select(s for s in SignAll if s.signid==signid).first()
        if stud:
            stud.classcode = classcode

## 导入临时数据（届别进入数据库表）
@db_session
def import_graduation(start_row=1,grid_end=0):
    wb = xlrd.open_workbook('jb.xls')
    ws = wb.sheets()[0]
    nrows = ws.nrows
    for i in range(start_row,nrows-grid_end):
        signid,graduation_year = ws.row_values(i)
        stud = select(s for s in SignAll if s.signid==signid).first()
        if stud:
            stud.graduation_year = str(int(graduation_year))

if __name__ == '__main__':
    db.bind(**DB_PARAMS)
    db.generate_mapping(create_tables=True)
    # import_classcode()
    import_graduation()