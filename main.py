#-*- coding:gbk -*-

import xlrd
import xlwt
import os


class Student:
    name = ''
    clas = ['' for i in range(0,9)]
    mark = ['' for i in range(0,9)]
    grade = ['' for i in range(0,9)]

    def __init__(self):
        self.name=''
        self.clas=['' for i in range(0,9)]
        self.mark=['' for i in range(0,9)]
        self.grade=['' for i in range(0,9)]



filename = ('����.xlsx','��ѧ.xlsx','Ӣ��.xlsx','����.xlsx','��ʷ.xlsx','����.xlsx','����.xlsx','��ѧ.xlsx','����.xlsx')
sub = ('����','��ѧ','Ӣ��','����','��ʷ','����','����','��ѧ','����')
status=0
isReg = [False for i in range(0,1400)]
Batch = 2015

def Run():
    stu = ['' for i in range(0,1400)]
    for i in range(0,9):
        if os.path.exists(filename[i]):
            data = xlrd.open_workbook(filename[i])
            table = data.sheets()[0]
            for k in range(1, table.nrows):
                num = int(table.cell(k, 0).value)
                num %= 10000
                if not isReg[num]:
                    tmp = Student()
                    isReg[num] = True
                else:
                    tmp=stu[num]
                tmp.name = table.cell(k, 1).value
                tmp.mark[i] = table.cell(k, 2).value
                tmp.grade[i] = table.cell(k, 3).value
                tmp.clas[i] = table.cell(k, 4).value

                stu[num] = tmp
        else:
            print((filename[i]+'������').decode('gbk'))
    return stu


def WriteFile(stu):
    cnt=1
    wb = xlwt.Workbook()
    ws = wb.add_sheet('result')
    ws.write(0, 0, 'ѧ��'.decode('gbk'))
    ws.write(0, 1, '����'.decode('gbk'))
    for i in range(1,10):
        ws.write(0, i * 3 - 1, sub[i - 1].decode('gbk'))
        ws.write(0, i * 3, '�ȵ�'.decode('gbk'))
        ws.write(0, i * 3 + 1, '�ֲ��'.decode('gbk'))
    ws.write(0, 29, '�ܷ�'.decode('gbk'))
    ws.write(0, 30, '�ȵ�'.decode('gbk'))
    for i in range(0,1400):
        if isReg[i]:
            ws.write(cnt, 1, stu[i].name)
            ws.write(cnt, 0, Batch*10000+i)
            for j in range(1,10):
                ws.write(cnt, j*3-1, stu[i].mark[j-1])
                ws.write(cnt, j*3, stu[i].grade[j-1])
                ws.write(cnt, j*3+1, stu[i].clas[j-1])
            cnt = cnt + 1
    wb.save('example.xls')




if __name__ == '__main__':
    student = Run()
    WriteFile(student)
    if status == 9:
        print('���β����ɹ������ڵ�ǰĿ¼�²����ļ���'.decode('gbk'))
    else:
        print('���β���ʧ�ܣ������ļ����ļ���ʽ�����ԣ�'.decode('gbk'))