import random
import xlwt
import xlrd

import pandas as pd
from matplotlib import pyplot as plt

p = [0.8, 0.2]
x = [0, 1]

def p_random(x, p):  # 按概率生成0,1
    start = 0
    random_num = random.random()
    for idx, score in enumerate(p):
        start += score
        if random_num <= start:
            break
    return x[idx]


##n = p_random(x, p)
##print(n)

def makelist(name: str):
    l = [[] for i in range(90)]
    for i in range(90):  # 生成学号
        if i < 9:
            l[i].append(str(name)+ '000' + str(i + 1))
        else:
            l[i].append(str(name) + '00' + str(i + 1))
    n1 = random.randint(5, 8)  # 5-8人缺勤80的课
    for i in range(n1):
        for j in range(21):
            n2 = p_random(x, p)
            l[i].append(n2)
    for i in range(90 - n1):
        for j in range(21):
            l[i + n1].append(1)
    for i in range(21):
        n4 = random.randint(0, 3)
        for k in range(n4):
            n3 = random.randint(n1, 89)
            l[n3][i + 1] = 0
    #创建一个Worlbook对象，相当于创建一个Excel文件
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    #创建好的excel文件一般有sheet1,sheet2,sheet3
    #此处表示将sheet1命名为"name+班数据"
    sheet = book.add_sheet(str(name) + '班数据', cell_overwrite_ok=True)
    col = []
    col.append('学号')
    for i in range(1, 22):
        col.append('第' + str(i) + '次记录')
    for i in range(0, 22):
        sheet.write(0, i, col[i])
    for i in range(90):
        data = l[i]
        for j in range(22):
            sheet.write(i + 1, j, data[j])
    savepath = 'C:/users/yanyan/PycharmProjects/abcd/pythonProject6/test_data//' + str(name) + '班表格.xls'
    book.save(savepath)
    # print(n1)  # 打印n1
    # print(l)  # 打印表格


for i in range(0, 100):#班级：A到E班
    makelist(i)

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

m=0
# 抽点名单
def nameList():
    global m
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
    # 缺勤率在百分之80下面的前三个学生
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
        # 和第21次记录比对，因为是在排好序的列表中进行比对，子列表多了缺勤率
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
X=[]
e=[]
for i in range(0, 100):
    print(str(i) + '班:')
    path = 'C:/users/yanyan/PycharmProjects/abcd/pythonProject6/test_data//' + str(i) + '班表格.xls'
    l1 = []
    l1 = read_xls(path)
    l1 = to_int(l1)
    #print("生成学生到勤记录：", l1)
    sortlist = sort()
    #print("根据缺勤率排好序的列表：", sortlist)
    namelist = nameList()
    print("抽点名单：", namelist)
    e.append(accumulate())

    #print('E={:.02%}'.format(e[]))

    X.append(i)
c=0
for i in range(0,100):
    c += e[i]
e_avg=c/100
print("e_avg=",e_avg)

pd.DataFrame({'E': e},
             index=X).plot.line()  # 图形横坐标默认为数据索引index。
#
plt.savefig(r'p1.png', dpi=200)
plt.show()  # 显示当前正在编译的图像

plt.bar(  # 设置x和y
    X, e,

    # 设置柱子宽度
    width=0.5,

    # 设置柱子颜色
    color="blue",

    # 设置legend的名称
    label="y")
plt.xlabel(  # x标签的名称
    "次数",

    # x标签的字体大小
    size=12,

    # x标签的字体颜色
    color="black")

plt.ylabel("E",
           size=12,
           color="black")
plt.savefig(r'p2.png', dpi=200)
plt.show()
