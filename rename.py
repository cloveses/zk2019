import shutil, os
from ph_models import *

#把身份证号命名的照片改为用中考报名号用的照片


@db_session
def main():
    for stud in select(s for s in StudPh):
        opath = os.path.join('ksPic', stud.schcode, 'Z'+stud.idcode+'.jpg')
        npath = os.path.join('photo', stud.signid+'.jpg')
        shutil.copyfile(opath, npath)
if __name__ == '__main__':
    main()