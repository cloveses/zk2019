import datetime
from db_mod import *
from sch_id import code_sches

# 对参加中考报名的全部中考报名考生进行资格审查


# zhtype=1  县外转入              outzh=1 
# zhtype=2  县外转入且县内转学    outzh=2
# zhtype=3  县内有转学            localzh=1
# zhtype=4  无转学记录            ouzh=None and localzh=None
# zhtype=0  无应届学籍学生审查时视为历届生（包括外省学生回乡报考和历届生）
#           外省回乡报考的应届只能由招办依照考生提供的材料审查

# 注意：
# 运行完后对转学记录中身份证号为空的学生进行查询和审查

# 审查时注意：
# 转学代码不是1或2，但从招办报名的；
# 历届以应届身份报名(转学代码为0，届别标志为应届的)
# 招办报名且转学代码为3(通过从招办报名不正当获取定向指标)


#查询报名学校与学籍学校不一致学生 和 应届学生报名填写成历届
@db_session
def get_sch_diff():
    for signstud in select(s for s in SignAll):
        gradeystud = select(s for s in GradeY18 if s.idcode.upper()==signstud.idcode.upper()).first()
        if gradeystud:
            signsch = code_sches[signstud.schcode]
            if gradeystud.sch != signsch:
                print(signsch, gradeystud.sch)
                print(signstud.name, '中考报名所在学校', signstud.sch, '学籍所在学校', gradeystud.sch)
            # else:
            #     print(signstud.id, end=',')
            if signstud.graduation_year != '2019':
                print('应届学生报名为历届：', signstud.name, signstud.idcode)

@db_session
def get_exchange_idcode():
    for stud in select(s for s in GradeY18):
        if stud.oidcode:
            repeat_stud = select(s for s in GradeY18 if s.idcode.upper()==stud.oidcode.upper()).first()
            if repeat_stud:
                print(stud.name, stud.oidcode, repeat_stud.sch, '旧身份证号，有人使用：', repeat_stud.name, repeat_stud.idcode, repeat_stud.sch)


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
def set_signall_zhtype():
    year = datetime.datetime.now().year
    for stud in SignAll.select():
        gradey18 = select(s for s in GradeY18 if s.idcode.upper()==stud.idcode.upper()).first()
        if gradey18:
            if int(stud.graduation_year) != year:
                print('审查 应届,报名为历届，需核查：',stud.idcode,stud.name,stud.sch)
            if gradey18.outzh == 1:
                stud.zhtype = 1
            elif gradey18.outzh == 2:
                stud.zhtype = 2
            elif gradey18.localzh == 1:
                stud.zhtype = 3
            elif gradey18.outzh == None and gradey18.localzh == None:
                stud.zhtype = 4
        else:
            stud.zhtype = 0
            if int(stud.graduation_year) == year:
                print('审查 历届,报名为应届，需核查：',stud.idcode,stud.name,stud.sch)
            # 查不到身份证号，即为历届生
            # print(stud.idcode,stud.name,'no find.')

@db_session
def get_datas():
    data_title = ['中考报名号','姓名','性别','身份证号','学校','班级代码','定向与否','备注']

    all_datas = [data_title,]

    datas = [data_title,]

    schs = select(s.sch for s in SignAll)[:]
    # print(schs)
    schs.remove('泗县招生办')
    for sch in schs:
        datas = datas[:1]

        # 获取享受定向的应届生名单
        query = select([s.signid,s.name,s.sex,s.idcode,s.sch,s.classcode] 
            for s in SignAll if s.zhtype==4 and s.sch==sch)[:]
        datas.extend(query)

        # 获取应届生因有县转学记录而不享受定向名单
        zh_datas = select([s.signid,s.name,s.sex,s.idcode,s.sch,s.classcode] 
            for s in SignAll if s.zhtype==3 and s.sch==sch)[:]
        zh_datas = [list(zh_data) for zh_data in zh_datas]
        for zh_data  in zh_datas:
            zh_data.extend(('不享受定向','应届有县内转学记录'))
        datas.extend(zh_datas)

        #获取历届生而不享受定向名单
        pr_datas = select([s.signid,s.name,s.sex,s.idcode,s.sch,s.classcode] 
            for s in SignAll if s.zhtype==0 and s.sch==sch)[:]
        pr_datas = [list(pr_data) for pr_data in pr_datas]
        for pr_data in pr_datas:
            pr_data.extend(('不享受定向','历届生'))

        datas.extend(pr_datas)

        save_datas_xlsx('.'.join((sch+'定向审查结果','xlsx')),datas)

        all_datas.extend(datas[1:])

    sch = '泗县招生办'
    datas = datas[:1]
    query = select([s.signid,s.name,s.sex,s.idcode,s.sch,s.classcode] 
        for s in SignAll if s.zhtype !=2 and s.zhtype !=0 and s.sch==sch)[:]
    datas.extend(query)

    zh_datas = select([s.signid,s.name,s.sex,s.idcode,s.sch,s.classcode] 
        for s in SignAll if s.zhtype==2)[:]
    zh_datas = [list(zh_data) for zh_data in zh_datas]
    for zh_data in zh_datas:
        zh_data.extend(('不享受定向','应届同时有县外、县内转学记录'))
    datas.extend(zh_datas)

    pr_datas = select([s.signid,s.name,s.sex,s.idcode,s.sch,s.classcode] 
        for s in SignAll if s.zhtype==0 and s.sch==sch)[:]
    pr_datas = [list(pr_data) for pr_data in pr_datas]
    for pr_data in pr_datas:
        pr_data.extend(('不享受定向','历届生'))
    datas.extend(pr_datas)

    save_datas_xlsx('.'.join((sch+'定向审查结果','xlsx')),datas)

    all_datas.extend(datas[1:])
    save_datas_xlsx(('全县定向审查结果.xlsx'),all_datas)

# @db_session
# def get_all_data():
#     data_title = ['中考报名号','姓名','性别','身份证号','学校','学校代码']

#     datas = [data_title,]
#     query = select([s.signid,s.name,s.sex,s.idcode,s.sch] 
#         for s in SignAll if s.zhtype==0)[:]
#     datas.extend(query)
#     save_datas_xlsx('.'.join(('全县历届学生名单','xlsx')),datas)

#     # 对于招办报名学生，只审查有无县内转学记录，其余情况由招办处理
#     datas = datas[:1]
#     query = select([s.signid,s.name,s.sex,s.idcode,s.sch] 
#         for s in SignAll if s.zhtype !=2 and s.sch=='泗县招生办')[:]
#     datas.extend(query)
#     save_datas_xlsx('.'.join(('招办报名无县内转学记录名单','xlsx')),datas)

#     datas = datas[:1]
#     query = select([s.signid,s.name,s.sex,s.idcode,s.sch] 
#         for s in SignAll if s.zhtype==2)[:]
#     datas.extend(query)
#     save_datas_xlsx('.'.join(('招办报名有县内转学记录名单','xlsx')),datas)

#     schs = select(s.sch for s in SignAll)[:]
#     # print(schs)
#     schs.remove('泗县招生办')
#     for sch in schs:
#         datas = datas[:1]
#         query = select([s.signid,s.name,s.sex,s.idcode,s.sch] 
#             for s in SignAll if s.zhtype==4 and s.sch==sch)[:]
#         datas.extend(query)
#         save_datas_xlsx('.'.join((sch+'应届享受定向名单','xlsx')),datas)

#         datas = datas[:1]
#         query = select([s.signid,s.name,s.sex,s.idcode,s.sch] 
#             for s in SignAll if s.zhtype==3 and s.sch==sch)[:]
#         datas.extend(query)
#         save_datas_xlsx('.'.join((sch+'应届不享受定向名单','xlsx')),datas)

if __name__ == '__main__':
    db.bind(**DB_PARAMS)
    db.generate_mapping(create_tables=True)

    # get_sch_diff()
    # get_exchange_idcode()
    set_signall_zhtype()
    # get_datas()
