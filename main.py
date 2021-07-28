# coding=utf-8

import pandas as pd
import re
import os
from bokeh.plotting import figure, output_file, show
from bokeh.io import export_png


# Press the green button in the gutter to run the script.

def mkdir(path):
    folder = os.path.exists(path)

    if not folder:  # 判断是否存在文件夹如果不存在则创建为文件夹
        os.makedirs(path)  # makedirs 创建文件时如果路径不存在会创建这个路径
        print("---  new folder...  ---")

        print("---  OK  ---")
    else:
        print("---  There is this folder!  ---")


def generate_html(filename, dir_name):
    path = dir_name + '/' + filename
    print(path)
    data = pd.read_excel(path, names=["key", "value"], sheet_name=0, skiprows=3)
    data_x = data["key"].tolist()
    for i in range(len(data_x)):
        data_x[i] = re.findall('\d+', data_x[i])[0]
    data_y = data["value"].tolist()
    data_x.reverse()
    data_y.reverse()
    # 输出html文件
    mkdir("生成的图表/" + dir_name)
    output_file("生成的图表/" + dir_name + '/' + filename.split('.xls')[0] + ".html")

    # 创建一个画布，并设置标题和坐标标签
    p = figure(title=filename, x_axis_label='x', y_axis_label='y')

    # 添加一条线，饼干设置图例和线粗细
    p.line(data_x, data_y, legend_label="Temp.", line_width=4)
    # 生成html文件，并在浏览器中打开

    show(p)


def generate_png(filename, dir_name):
    path = dir_name + '/' + filename
    print(path)
    data = pd.read_excel(path, names=["key", "value"], sheet_name=0, skiprows=3)
    data_x = data["key"].tolist()
    for i in range(len(data_x)):
        data_x[i] = re.findall('\d+', data_x[i])[0]
    data_y = data["value"].tolist()
    data_x.reverse()
    data_y.reverse()
    # 输出html文件
    mkdir("生成的图表/" + dir_name)

    # 创建一个画布，并设置标题和坐标标签
    p = figure(title=filename, x_axis_label='x', y_axis_label='y')

    # 添加一条线，饼干设置图例和线粗细
    p.line(data_x, data_y, legend_label="Temp.", line_width=4)
    # 生成html文件，并在浏览器中打开

    export_png(
        p,
        filename="a.png"
    )


def work(dir_name):
    for _, _, item in os.walk("./" + dir_name):
        for filename in item:
            generate_html(filename, dir_name)


if __name__ == '__main__':
    work("需要可视化的文件")
