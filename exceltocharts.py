"""
Created by Chavaz on 2021/6/4.
实现指定excel文件指定数据格式实现饼图、折线图、趋势图图表展示功能
pip install pyecharts
pip install xlrd
"""

from pyecharts import options as opts
from pyecharts.charts import Pie, Line
from pyecharts.commons.utils import JsCode
import xlrd

def readExcelData():
    print("1.饼图\n2.折线图\n3.趋势图")
    cc = int(input("请输入想要展示的图表:"))

    #key值
    keyData=['tag_name','annotation','reg_type','address','byte_size','data_type','operate_flag','data_seq']
    # 读取excel表的数据
    workbook = xlrd.open_workbook(r'1.xls')
    # 选取需要读取数据的那一页
    sheet = workbook.sheet_by_index(0)
    # 获得行数和列数
    rows = sheet.nrows
    cols = sheet.ncols
    # 创建一个数组用来存储excel中的数据
    p = []
    labels = []
    x_data = []
    y_data = []
    xaxis_data = []
    for i in range(0, rows):
        d={}
        for j in range(0,cols):
            q=keyData[j] #自己设置的key
            d[q] = sheet.cell(i, j).value #具体的数据

        if cc == 1:
            d.pop('reg_type')

        ap = []
        label = []
        # face = []

        for k, v in d.items():
            label.append(v)
            # if isinstance(v, float):  # excel中的值默认是float通过'"%s":%d'，'"%s":"%s"'格式化数组
            #     ap.append('"%s":"%s"' % (k, str(int(v)))) #转为字符串
            # else:
            #     ap.append('"%s":"%s"' % (k, v))
        s = '{%s}' % (','.join(ap))  #把list转成用，分隔的字符串 在把字符串套了一个花括号
        x_data.append(label[0])
        y_data.append(label[1])
        if cc == 3:
            xaxis_data.append(label[2])

        labels.append(label)
        p.append(s)

    t = '[%s]' % (','.join(p))  # 格式化

    if cc == 1:
        c = (
            Pie()
                .add("", labels)
                .set_colors(["blue", "green", "yellow", "red", "pink", "orange", "purple"])
                .set_global_opts(title_opts=opts.TitleOpts(title="饼图"))
                .set_series_opts(label_opts=opts.LabelOpts(formatter="{b}: {c}"))
                .render("pie_set_color.html")
        )
    elif cc == 2:
        (
            Line()
                .set_global_opts(
                tooltip_opts=opts.TooltipOpts(is_show=False),
                xaxis_opts=opts.AxisOpts(type_="category"),
                yaxis_opts=opts.AxisOpts(
                    type_="value",
                    axistick_opts=opts.AxisTickOpts(is_show=True),
                    splitline_opts=opts.SplitLineOpts(is_show=True),
                ),
            )
                .add_xaxis(xaxis_data=x_data)
                .add_yaxis(
                series_name="",
                y_axis=y_data,
                symbol="emptyCircle",
                is_symbol_show=True,
                label_opts=opts.LabelOpts(is_show=False),
                is_smooth=True,
            )
                .render("basic_line_chart.html")
        )
    elif cc == 3:
        js_formatter = """function (params) {
                console.log(params);
                return '降水量  ' + params.value + (params.seriesData.length ? '：' + params.seriesData[0].data : '');
            }"""

        (
            Line(init_opts=opts.InitOpts(width="1600px", height="800px"))
                .add_xaxis(
                xaxis_data
            )
                .add_yaxis(
                series_name="趋势图",
                is_smooth=True,
                symbol="emptyCircle",
                is_symbol_show=False,
                color="#6e9ef1",
                y_axis=y_data,
                label_opts=opts.LabelOpts(is_show=False),
                linestyle_opts=opts.LineStyleOpts(width=2),
            )
                .set_global_opts(
                legend_opts=opts.LegendOpts(),
                tooltip_opts=opts.TooltipOpts(trigger="none", axis_pointer_type="cross"),
                xaxis_opts=opts.AxisOpts(
                    type_="category",
                    axistick_opts=opts.AxisTickOpts(is_align_with_label=True),
                    axisline_opts=opts.AxisLineOpts(
                        is_on_zero=False, linestyle_opts=opts.LineStyleOpts(color="#d14a61")
                    ),
                    axispointer_opts=opts.AxisPointerOpts(
                        is_show=True, label=opts.LabelOpts(formatter=JsCode(js_formatter))
                    ),
                ),
                yaxis_opts=opts.AxisOpts(
                    type_="value",
                    splitline_opts=opts.SplitLineOpts(
                        is_show=True, linestyle_opts=opts.LineStyleOpts(opacity=1)
                    ),
                ),
            )
                .render("multiple_x_axes.html")
        )

readExcelData()