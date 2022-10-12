import random
import xlwt

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
            l[i].append(name + '000' + str(i + 1))
        else:
            l[i].append(name + '00' + str(i + 1))
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
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = book.add_sheet(name + '班数据', cell_overwrite_ok=True)
    col = []
    col.append('学号')
    for i in range(1, 22):
        col.append('第' + str(i) + '记录')
    for i in range(0, 22):
        sheet.write(0, i, col[i])
    for i in range(90):
        data = l[i]
        for j in range(22):
            sheet.write(i + 1, j, data[j])
    savepath = 'D:\pyCharm\pythonProject\\' + name + '班表格.xls'
    book.save(savepath)
    # print(n1)  # 打印n1
    # print(l)  # 打印表格


for i in range(65, 70):
    makelist(chr(i))
