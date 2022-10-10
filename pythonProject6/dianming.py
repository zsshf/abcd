import random
import xlrd
import xlwt
import pandas as pd

p = [0.8, 0.2]
x = [0, 1]


def sort():
    for i in range(0, len(l1)):
        b = l1[i]  # b为子列表
        sum = 0
        count = 0
        route = 0
        for j in range(0, len(b) - 1):
            if type(b[j]) == int:
                sum += b[j]  # 到勤次数
                count += 1  # 总次数
        route = (count - sum) / count  # 缺勤率
        b.append(float('{:.02}'.format(route)))  # 把缺勤率加入列表
    # 排序
    m = len(b) - 1  # m代表缺勤率下标
    sortlist = sorted(l1, key=lambda x: x[m], reverse=True)  # 根据缺勤率进行排序
    return sortlist


# 抽点名单
def nameList():
    student = []  # 存放抽点名单
    count = 0  # 缺勤率在80%及以上的学生人数
    # 得到缺勤率在百分之80以上的学生
    for i in range(0, len(sortlist)):
        b = l1[i]  # b为子列表
        if b[len(b) - 1] >= 0.8:
            count += 1
            m = count
            student.append(b[0])
    '''
    # 还有0-3位同学由于各种原因缺席,所以在缺勤率在百分之80以下的学生里随机抽点三位同学
    resultlist = random.sample(range(count, 90), 3)
    for j in resultlist:
        b=l1[j]
        student.append(b[0]
    '''
    # 缺勤率在百分之80往下的前三个学生
    for j in range(m, m + 3):
        b = sortlist[j]
        student.append(b[0])
    # 返回抽点名单
    return student


# 计算有效点名率E=有效点名次数/总点名次数
def accumulate():
    count = 0
    for i in range(0, len(namelist)):
        b = sortlist[i]
        # 和第20次记录比对，因为是在排好序的列表中进行比对，子列表多了缺勤率
        # #所以第20次记录在子列表的倒数第2个位置
        if b[len(b) - 2] == 0:
            count += 1  # 有效点名次数

    e = count / len(namelist)  # e=有效点名次数/总点名次数
    return e


def read_xls(path: str):
    data_excel = xlrd.open_workbook(path)
    table = data_excel.sheets()[0]
    n_rows = table.nrows
    l1 = []
    for i in range(n_rows - 1):
        data = table.row_values(i + 1, start_colx=0, end_colx=None)
        l1.append(data)
    return l1


def to_int(l: list):
    for i in range(len(l)):
        for j in range(1,len(l[i])):
            if l[i][j] % 1 == 0:
                l[i][j]=int(l[i][j])
    return l


for i in range(65, 70):
    print(chr(i) + '班:')
    path = 'D:\pyCharm\pythonProject\\' + chr(i) + '班表格.xls'
    l1 = []
    l1 = read_xls(path)
    l1 = to_int(l1)
    print("生成学生到勤记录：", l1)
    sortlist = sort()
    print("根据缺勤率排好序的列表：", sortlist)
    namelist = nameList()
    print("抽点名单：", namelist)
    E = accumulate()
    print('E={:.02%}'.format(E))
