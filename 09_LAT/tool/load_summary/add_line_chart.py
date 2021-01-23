# -*- coding: utf-8 -*-
# @Time    : 2019/10/17 15:54
# @Author  : CE
# @File    : add_line_chart.py

import openpyxl
from openpyxl.chart import (
    Series,
    LineChart,
    Reference)

class excel_operation:

    def __int__(self,
                worksheet,
                title,
                row_min,
                row_max,
                col_min,
                col_max):

        self.worksheet = worksheet
        self.row_min   = row_min
        self.row_max   = row_max
        self.col_min   = col_min
        self.col_max   = col_max
        self.title     = title


    def add_line_chart(sheet, min_col, min_row, max_col, max_row, position):
        c1 = LineChart()
        c1.title = 'Dynamic Power Curve'  # 图的标题
        c1.style = 12  # 线条的style
        c1.y_axis.title = 'Power(kw)'  # y坐标的标题
        c1.x_axis.title = "Wind Speed(m/s)"  # x坐标的标题
        data = Reference(sheet, min_col=min_col, min_row=min_row, max_col=max_col,
                         max_row=max_row)  # 图像的数据 起始行、起始列、终止行、终止列
        c1.add_data(data, titles_from_data=True, from_rows=True)
        dates = Reference(sheet, min_col=2, min_row=1, max_col=max_col)
        c1.set_categories(dates)
        sheet.add_chart(c1, position)  # 将图表添加到 sheet中

