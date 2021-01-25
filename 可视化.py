import numpy as np
from matplotlib import pyplot as plt
import matplotlib
import xlrd


def pie_paint(size, name):
    matplotlib.rcParams['font.sans-serif'] = ['KaiTi']
    plt.figure(figsize=(6, 9))
    # 定义标签
    labels = [u'乐观', u'感激', u'悲伤', u'不满', u'忧虑', u'害怕']
    # 计算占比，使用tf-idf作为数据依据; 配色
    # size = [1.3051, 0.3281, 0.9292, 0.8658, 1.1228, 4.1034]
    colors = ['orange', 'yellowgreen', 'lightblue', 'red', 'purple', 'gray']
    # 间隙
    explode = [0.05, 0.05, 0.05, 0.05, 0.05, 0.05]
    # 饼状图返回值，饼状图外文本，饼状图内部文本
    a, b, c = plt.pie(size, explode, labels, colors,
                      labeldistance=1, autopct='%3.1f%%', shadow=False,
                      startangle=90, pctdistance=0.5)
    for i in b:
        i.set_size = 30
    for j in c:
        j.set_size = 20
    plt.axis('equal')
    plt.legend()
    plt.savefig('./'+name+'.jpg')


def line_paint(path):
    matplotlib.rcParams['font.sans-serif'] = ['KaiTi']
    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_index(0)
    x = np.array([1, 2, 3, 4])
    Y = np.zeros((6, 4), dtype=np.float)
    fig = plt.figure(figsize=(10, 8))
    for i in range(0, 4):
        Y[0][i] = (str(sheet.cell_value(1+i*7, 2)))
        Y[1][i] = (str(sheet.cell_value(2+i*7, 2)))
        Y[2][i] = (str(sheet.cell_value(3+i*7, 2)))
        Y[3][i] = (str(sheet.cell_value(4+i*7, 2)))
        Y[4][i] = (str(sheet.cell_value(5+i*7, 2)))
        Y[5][i] = (str(sheet.cell_value(6+i*7, 2)))
    plt.plot(x, Y[0], 'y', label='乐观', linewidth=2)
    plt.plot(x, Y[1], 'g', label='感激', linewidth=2)
    plt.plot(x, Y[2], 'b', label='悲伤', linewidth=2)
    plt.plot(x, Y[3], 'r', label='不满', linewidth=2)
    plt.plot(x, Y[4], 'c', label='忧虑', linewidth=2)
    plt.plot(x, Y[5], 'k', label='害怕', linewidth=2)
    plt.title("tf-idf/time")
    plt.legend()
    plt.savefig('./折线图.jpg')


if __name__ == '__main__':
    book = xlrd.open_workbook('./result.xls')
    sheet = book.sheet_by_index(0)
    # 批量画饼状图
    for i in range(0, 4):
        data_list = []
        for j in range(1+i*7, 7+i*7):
            data_list.append(str(sheet.cell_value(j, 4)))
        pie_paint(data_list, sheet.cell_value(1+i*7, 0))

    # 画折线图
    line_paint('./result.xls')
