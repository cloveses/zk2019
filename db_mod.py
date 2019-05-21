from pony.orm import *

__author__ = "cloveses"

DB_PARAMS = {
    'provider':'postgres',
    'user':'postgres',
    'password':'123456',
    'host':'192.168.7.233',
    'database':'sc2019'
}
db = Database()

# 所有转学记录
class StudZhAll(db.Entity):
    gsrid = Optional(str,nullable = True)
    dsrid = Optional(str,nullable = True)
    idcode = Optional(str,nullable = True)
    name = Required(str)
    sex = Required(str)
    birth = Optional(str,nullable = True)
    sch = Required(str)
    zhtype = Required(str)
    optdate = Optional(str,nullable = True)
    zhsrc = Optional(str,nullable = True)
    zhdes = Optional(str,nullable = True)

# 三年级学生记录
class GradeY18(db.Entity):
    sch = Required(str)
    grade = Required(str)
    sclass = Required(str)
    gsrid = Optional(str,nullable = True)
    ssrid = Optional(str,nullable = True)
    dsrid = Optional(str,nullable = True)
    name = Required(str)
    idcode = Optional(str,nullable = True)
    oidcode = Optional(str,nullable = True)
    regtype = Optional(str,nullable=True)
    sex = Required(str)
    nation =  Optional(str,nullable=True)
    enter = Optional(str,nullable=True)
    outzh = Optional(int,nullable = True) #null 无县外转入记录 1 有县外转学记录 2 有县外转学记录且有县内转学记录
    localzh = Optional(int,nullable = True) # outzh=None localzh=1 有县内转学记录 outzh=None localzh=None 

# 关键信息变更表
class KeyInfoChg(db.Entity):
    ssrid = Optional(str,nullable = True)
    oname = Required(str)
    name = Required(str)
    osex = Required(str)
    sex = Required(str)
    obirth = Required(str)
    birth = Required(str)
    oidcode = Optional(str,nullable = True)
    idcode = Required(str)
    sch = Required(str)
    grade = Required(str)
    sclass = Required(str)

# 所有中考报名学生信息
class SignAll(db.Entity):
    signid = Required(str)
    name = Required(str)
    sex = Required(str)
    idcode = Required(str)
    sch = Required(str)
    schcode = Required(str)
    graduation_year = Optional(str)
    classcode = Optional(str,nullable = True)
    classnum = Required(str)
    nation = Required(str)
    gsrid = Required(str)
    ssrid = Required(str)
    birth = Required(str)
    regtype = Required(str)
    hometel = Required(str)
    addr = Required(str)
    periodtel = Required(str)
    zhtype = Optional(int,nullable = True)
    regaddr = Optional(bool, nullable=True)
    #regaddr 户籍地，本地为True,用于审查外地转入本地学生定向享受情况

    # 1 县外转入，2 县外转入，县内转
    # 3 县内转学，4 无转学记录
    # 0 非应届生

# 八年级学生记录
class GradeY8(db.Entity):
    sch = Required(str)
    grade = Required(str)
    sclass = Required(str)
    gsrid = Optional(str,nullable = True)
    ssrid = Optional(str,nullable = True)
    dsrid = Optional(str,nullable = True)
    name = Required(str)
    idcode = Optional(str,nullable = True)
    regtype = Optional(str,nullable=True) #户口类型
    sex = Required(str)
    nation =  Optional(str,nullable=True)

