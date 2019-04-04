from db_mod import *

# 提供给招办的预报名（从网上报名）名单的处理程序


# # \copy gradey18 (sch,grade,sclass,gsrid,ssrid,dsrid,name,idcode,sex) from g:\grade32018\grade183.csv with csv

# #  \copy studzhall (gsrid,dsrid,idcode,name,sex,birth,sch,zhtype,optdate,zhsrc,zhdes) from 1516zh.csv with csv

# #  \copy studzhall (gsrid,dsrid,idcode,name,sex,birth,sch,zhtype,optdate,zhsrc,zhdes) from 1617zh.csv with csv

# #  \copy studzhall (gsrid,dsrid,idcode,name,sex,birth,sch,zhtype,optdate,zhsrc,zhdes) from 1718zh.csv with csv

# #  \copy keyinfochg (ssrid,oname,name,osex,sex,obirth,birth,oidcode,idcode,sch,grade,sclass) from d:\work2018\grade32018\keyinfochg.csv with csv

# # delete from studzhall where zhtype='跨省就学转学(入)';

# # select distinct(zhtype) from studzhall;

def save_datas_xlsx(filename,datas):
    #将一张表的信息写入电子表格中XLSX文件格式
    import xlsxwriter
    w = xlsxwriter.Workbook(filename)
    w_sheet = w.add_worksheet('sheet1')
    for rowi,row in enumerate(datas):
        for coli,celld in enumerate(row):
            w_sheet.write(rowi,coli,celld)
    w.close()


def is_out(zh_data):
    """判别是否有县外转学记录"""
    out_flags = ('市内转','跨市转','跨省转','跨境转')
    for zh in zh_data:
        for out_flag in out_flags:
            if zh.zhtype and zh.zhtype.startswith(out_flag):
                return True

def has_local(zh_data):
    """判别是否有县内转学记录"""
    for zh in zh_data:
        if zh.zhtype and zh.zhtype.startswith('县区内转'):
            return True

# 清除小学和高中学籍变动
@db_session
def clear_studzhall():
    for zh_stud in select(s for s in StudZhAll):
        if not ((zh_stud.zhsrc and '初中' in zh_stud.zhsrc) 
            or (zh_stud.zhdes and '初中' in zh_stud.zhdes)):
            zh_stud.delete()

# 清除小学阶段的关键信息变更
@db_session
def clear_keyinfo():
    for keyinfo in select(s for s in KeyInfoChg):
        if '小学' in keyinfo.grade:
            keyinfo.delete()

# 插入曾用身份证号
@db_session
def insert_oidcode():
    for keyinfo in select(s for s in KeyInfoChg):
        if keyinfo.oidcode and keyinfo.oidcode != keyinfo.idcode:
            if count(select(s for s in GradeY18 \
                if s.idcode == keyinfo.idcode)) == 1:
                stud = select(s for s in GradeY18 \
                    if s.idcode == keyinfo.idcode).first()
                stud.oidcode = keyinfo.oidcode

# 获取指定学生的转学记录

@db_session
def get_zh_recos(stud):
    zh_recos = []
    if stud.idcode:
        # 依据现身份证
        zhs = select(zh for zh in StudZhAll 
            if zh.idcode.upper() == stud.idcode.upper())[:]
        if zhs:
            zh_recos.extend(zhs)
        # 依据曾用身份证
        if stud.oidcode:
            zhs = select(zh for zh in StudZhAll
                if zh.idcode.upper() == stud.oidcode.upper())[:]
            if zhs:
                zh_recos.extend(zhs)
    else:
        print(stud.name,stud.sch,'No id number!')
    # 依据国家学籍号
    zhs = select(zh for zh in StudZhAll 
        if zh.gsrid == stud.gsrid)[:]
    if zhs:
        zh_recos.extend(zhs)
    return set(zh_recos)

@db_session
def set_zh_from_out():
    for stud in select(s for s in GradeY18):
        zh_recos = get_zh_recos(stud)
        if zh_recos and is_out(zh_recos):
            stud.outzh = 1

@db_session
def set_outzh_local():
    for stud in select(s for s in GradeY18 if s.outzh==1):
        zh_recos = get_zh_recos(stud)
        if zh_recos and has_local(zh_recos):
            stud.outzh = 2

@db_session
def set_localzh():
    for stud in select(s for s in GradeY18 if s.outzh==None):
        zh_recos = get_zh_recos(stud)
        if zh_recos and has_local(zh_recos):
            stud.localzh = 1


# 1县外转入              outzh=1 
# 2县外转入且县内转入    outzh=2
# 3县内转学              localzh=1
# 4无转学记录            ouzh=None and localzh=None


# 将数据导出为excel
# @db_session
# def get_sch_data_xls():
#     schs = select(s.sch for s in GradeY18)

#     # 导出各校无从县外转入的应届生（从各校报名）,每学校一个文件
#     for sch in schs:
#         datas = [['学校','年级','班级','全国学籍号','学籍号','地区学号','姓名','身份证号','性别'],]
#         query = select([s.sch,s.grade,s.sclass,s.gsrid,s.ssrid,s.dsrid,s.name,s.idcode,s.sex] 
#             for s in GradeY18 if s.sch == sch and s.outzh == None)[:]
#         datas.extend(query)
#         save_datas_xlsx('.'.join((sch,'xlsx')),datas)
#         print(count(select(s for s in GradeY18 if s.sch == sch and s.outzh == None)),sch)

#     # 导出所有学校有县外转入记录的应届生（从招办报名）
#     datas = [['学校','年级','班级','全国学籍号','学籍号','地区学号','姓名','身份证号','性别'],]
#     query = select([s.sch,s.grade,s.sclass,s.gsrid,s.ssrid,s.dsrid,s.name,s.idcode,s.sex] 
#         for s in GradeY18 if s.outzh != None)[:]
#     datas.extend(query)
#     save_datas_xlsx('.'.join(('招办报名名单','xlsx')),datas)
#     print('招办报名人数：',count(select(s for s in GradeY18 if s.outzh != None)))

#     # 导出从各校报名的总名单
#     datas = [['学校','年级','班级','全国学籍号','学籍号','地区学号','姓名','身份证号','性别'],]
#     query = select(s for s in GradeY18 if s.outzh == None).order_by(GradeY18.sch)
#     for s in query:
#         datas.append([s.sch,s.grade,s.sclass,s.gsrid,s.ssrid,s.dsrid,s.name,s.idcode,s.sex])
#     save_datas_xlsx('.'.join(('各校报名总名单','xlsx')),datas)
#     print('各校报名总人数：',count(select(s for s in GradeY18 if s.outzh == None)))


if __name__ == '__main__':
    db.bind(**DB_PARAMS)
    db.generate_mapping(create_tables=True)

    clear_studzhall()
    clear_keyinfo()
    insert_oidcode()
    set_zh_from_out()
    set_outzh_local()
    set_localzh()
    # get_sch_data_xls()