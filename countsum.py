import os
import xlrd
import xlsxwriter

## 统计定向学生数和学生总数

def get_data(filename,headline=1):
     #既可以打开xls类型的文件，也可以打开xlsx类型的文件
    #w = xlrd.open_workbook('text.xls')
    #w = xlrd.open_workbook('acs.xlsx')
    datas = []
    w = xlrd.open_workbook(filename)
    ws = w.sheets()[0]
    nrows = ws.nrows
    for i in range(headline,nrows):
        data = ws.row_values(i)
        datas.append(data)
    #    print(datas)
    return datas

def save_datas_xlsx(filename,datas):
    #将一张表的信息写入电子表格中XLSX文件格式
    w = xlsxwriter.Workbook(filename)
    w_sheet = w.add_worksheet('sheet1')
    for rowi,row in enumerate(datas):
        for coli,celld in enumerate(row):
            w_sheet.write(rowi,coli,celld)
    w.close()


def countit(filename='全县定向审查结果.xlsx'):
    datas = get_data(filename)
    rets = dict()
    for data in datas:
        if data[4] in rets:
            if data[6] == '不享受定向':
                rets[data[4]][0] += 1
                rets[data[4]][1] += 1
            else:
                rets[data[4]][0] += 1
        else:
            if data[6] == '不享受定向':
                rets[data[4]] = [1,1]
            else:
                rets[data[4]] = [1,0]
    results = [[k,v[0],v[1]] for k,v in rets.items()]
    save_datas_xlsx('sumres.xlsx',results)

if __name__ == '__main__':
    countit()
