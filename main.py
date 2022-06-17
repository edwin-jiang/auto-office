from openpyxl import Workbook,load_workbook
import pandas as pd


def createExcelFile():
    # 创建一个工作簿对象
    wb = Workbook()
    # 在索引为0的位置创建一个名为mytest的sheet页
    ws = wb.create_sheet('mytest',0)
    # 对sheet页设置一个颜色（16位的RGB颜色）
    ws.sheet_properties.tabColor = 'ff72BA'
    # 将创建的工作簿保存为Mytest.xlsx
    wb.save('Mytest.xlsx')
    # 最后关闭文件
    wb.close()

def openExcelFile():
    # 加载工作簿
    wb2 = load_workbook('Mytest.xlsx')
    # 获取sheet页
    ws2 = wb2['mytest']
    ws3 = wb2.get_sheet_by_name('mytest')
    # 打印sheet页的颜色属性值
    print('color:',ws2.sheet_properties.tabColor)
    wb2.close()

openExcelFile()