
# 各考点考试日期及考试安排,即考试编排

arrange_datas =(
    ('泗县一中', '4月20日上午', (('宿州泗县中学', '女'), ('宿州泗县泗州学校', '女'))), #女
    ('泗县一中', '4月20日下午', (('宿州泗县中学', '男'), ('宿州泗县泗州学校', '男'))), #男
    ('泗县一中', '4月21日上午', ('宿州泗县灵童学校', '宿州泗县黑塔中学')),
    ('泗县一中', '4月21日下午', ('宿州泗县山头中学', '宿州泗县刘圩中学', '宿州泗县光明学校')),

    ('泗县二中', '4月20日上午', (('宿州泗县二中','女'), '宿州泗县丁湖中学')),
    ('泗县二中', '4月20日下午', (('宿州泗县二中','男'), '宿州泗县草沟中学')),
    ('泗县二中', '4月21日上午', ('宿州泗县墩集中学', '宿州泗县育才学校')),
    ('泗县二中', '4月21日下午', ('宿州泗县草庙中学', '宿州泗县城南中学', '宿州泗县思源学校')),

    ('泗县三中', '4月20日上午', (('宿州泗县三中', '女'), '宿州泗县长沟中学')), #女
    ('泗县三中', '4月20日下午', (('宿州泗县三中', '男'), '宿州泗县双语中学')), #男
    ('泗县三中', '4月21日上午', ('宿州泗县屏山中学','宿州泗县大庄中学')),
    ('泗县三中', '4月21日下午', ('宿州泗县黄圩中学', '宿州泗县杨集中学', '宿州泗县瓦坊中学',
                                '宿州泗县招生办', '宿州泗州实验学校')),
    )

date_seq = ('4月20日上午','4月20日下午','4月21日上午','4月21日下午')
sch_seq = ('泗县一中','泗县二中','泗县三中')
# 准考证号前缀
PREFIX = 19250000