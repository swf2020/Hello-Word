import os
import pandas as pd
import xlwt
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog
import sys
import  os

# 加载窗口
class MyWindow(QtWidgets.QWidget):
    def __init__(self):
        super(MyWindow, self).__init__()

    def msg(self):
        fileName1, filetype = QFileDialog.getOpenFileName(self, "选取文件", "./",
                                                          "All Files (*);;Excel Files (*.xls)")  # 设置文件扩展名过滤,注意用双分号间隔
        return fileName1

# 将文件读取出来放一个列表里面
def merge(filename):
    pwd = os.path.dirname(filename)  # 获取文件目录
    file_list = [] # 新建列表，存放文件名
    dfs = [] # 新建列表存放每个文件数据(依次读取多个相同结构的Excel文件并创建DataFrame)
    for root,dirs,files in os.walk(pwd): # 第一个为起始路径，第二个为起始路径下的文件夹，第三个是起始路径下的文件。
      for file in files:
        file_path = os.path.join(root, file)
        file_list.append(file_path) # 使用os.path.join(dirpath, name)得到全路径
        df = pd.read_excel(file_path) # 将excel转换成DataFrame
        dfs.append(df)

    df = pd.concat(dfs) # 将多个DataFrame合并为一个
    df.to_excel(os.path.dirname(filename) + '/result.xls', index=False) # 写入excel文件，不包含索引数据

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    myshow = MyWindow()
    filename = myshow.msg()  # 加载指定的文件
    merge(filename)
