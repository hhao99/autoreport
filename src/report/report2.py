from pptx import Presentation
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.util import Inches, Pt
from pptx import Presentation
from pptx.chart.data import ChartData, CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_TICK_MARK, XL_TICK_LABEL_POSITION, XL_MARKER_STYLE
from pptx.util import Cm  # Inches
from pandas import np
from pptx.enum.chart import XL_LEGEND_POSITION

from pptx.dml.color import RGBColor
import pandas as pd

if __name__ == '__main__':

    df_vw = pd.read_excel(r'c:/auto-report/Database_small_demo.xlsx', 0)
    df_mkt = pd.read_excel(r'c:/auto-report/Database_small_demo.xlsx', 1)

    print("df_vw=" + str(df_vw.shape[0]))
    view_rules = 'OEM'  # or OEM
    all_mkt = False  # true分母是否为全市场，false分母为fuel_type细分市场
    oem = df_vw.groupby('OEM').size().reset_index()["OEM"]
    brand = df_vw.groupby('Brand').size().reset_index()["Brand"]
    fuel_type = df_vw.groupby('Fuel_Type').size().reset_index()["Fuel_Type"]
    fuel_type_group = df_vw.groupby('Fuel_Type_Group').size().reset_index()["Fuel_Type_Group"]

    print(oem)
    print(brand)
    print(fuel_type)
    print(fuel_type_group)
    # 此处应该对 oem brand fuel_type三个数组进行过滤，把用户不需要的删除掉，默认是全选，需要通过读取web端的配置文件
    # oem = ['FAW-VW', 'JAC-VW', 'JV TBD', 'SAIC-VW']
    # brand = ['Audi', 'Cupra', 'Jetta', 'Sihao', 'Skoda', 'VW']
    # fuel_type = ['ICE', 'BEV', 'PHEV']
    oem_filter = oem
    brand_filter = brand
    fuel_type_filter = fuel_type  # 根据私有fitler决定哪个fuel_type需要保留

    PR_Status = 'PR67.OP'
    start_year = 2018
    year_span = 9
    end_year = start_year + year_span
    # oem = ['FAW-VW', 'JAC-VW', 'JV TBD', 'SAIC-VW']
    # brand = ['Audi', 'Cupra', 'Jetta', 'Sihao', 'Skoda', 'VW']
    # 根据所有filter过滤大众数据
    df_vw_filter = \
        df_vw[(df_vw['PR_Status'] == 'PR67.OP') & (df_vw['YEAR'] >= start_year) & (df_vw['YEAR'] <= end_year) \
              & (df_vw['OEM'].isin(oem)) & (df_vw['Brand'].isin(brand)) & (df_vw['Fuel_Type'].isin(fuel_type))]
    print("df_vw_filter=" + str(df_vw_filter.shape[0]))

    # 全销售市场数据分组聚合(ice bev phev)-------------------------
    df_mkt_sum = df_mkt.groupby('Fuel_type').agg({'Volume': np.sum}).reset_index()

    # 获得需要显示的年份数组，在图标中最为重要y轴坐标
    data_years = df_vw_filter.groupby(['YEAR']).size().reset_index().sort_values(['YEAR'], ascending=[True])['YEAR']
    print(data_years)

    # 大众市场销量数据分组统计
    df_vw_group_brand_year = []
    df_vw_group_oem_year = []
    if view_rules == 'Brand':  # 如果按brand显示
        df_vw_group_brand_year = df_vw_filter.groupby(['Brand', 'YEAR']).agg({'Volume': np.sum}).reset_index()
        print(df_vw_group_brand_year)
    else:  # 如果按OEM显示
        df_vw_group_oem_year = df_vw_filter.groupby(['OEM', 'YEAR']).agg({'Volume': np.sum}).reset_index()
        print(df_vw_group_oem_year)

    prs = Presentation('c:/auto-report/cover.pptx')
    title_only_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_only_slide_layout)
    shapes = slide.shapes

    # 开始创建点线图-----------------------------------
    chart_data_line = ChartData()
    chart_data_line.categories = data_years
    if all_mkt:
        mkt_volume = df_mkt_sum[(df_mkt_sum['Fuel_type'].isin(fuel_type))].agg({'Volume': np.sum}) \
            .reset_index().iloc[0, 1]
    else:
        mkt_volume = df_mkt_sum[(df_mkt_sum['Fuel_type'].isin(fuel_type_filter))].agg({'Volume': np.sum}) \
            .reset_index().iloc[0, 1]

    # 添加线图
    if view_rules == 'Brand':
        for brand in brand_filter:
            total_volumes = df_vw_group_brand_year[(df_vw_group_brand_year['Brand'] == brand)] \
                .groupby(['YEAR']).agg({'Volume': np.sum}).reset_index() \
                .sort_values(['YEAR'], ascending=[True])['Volume']
            series_mkt_rate = [vol / mkt_volume for vol in total_volumes]
            chart_data_line.add_series(brand, series_mkt_rate)
    else:
        for oem in oem_filter:
            total_volumes = df_vw_group_oem_year[(df_vw_group_oem_year['OEM'] == oem)] \
                .groupby(['YEAR']).agg({'Volume': np.sum}).reset_index() \
                .sort_values(['YEAR'], ascending=[True])['Volume']
            series_mkt_rate = [vol / mkt_volume for vol in total_volumes]
            chart_data_line.add_series(oem, series_mkt_rate)

    x, y, cx, cy = Cm(1), Cm(5.5), Cm(24), Cm(4)
    chart_line = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE, x, y, cx, cy, chart_data_line
    ).chart

    chart_line.has_legend = True
    chart_line.legend.include_in_layout = False
    chart_line.legend.position = XL_LEGEND_POSITION.LEFT
    chart_line.legend.font.size = Pt(10)

    for line_serie in chart_line.series:
        line_serie.smooth = True
        line_serie.marker.style = XL_MARKER_STYLE.CIRCLE
        line_serie.data_labels.show_value = True
        line_serie.data_labels.number_format = '0.0%'
        line_serie.data_labels.font.size = Pt(10)

    value_axis_line = chart_line.value_axis
    value_axis_line.has_major_gridlines = False
    value_axis_line.major_tick_mark = XL_TICK_MARK.NONE
    value_axis_line.tick_label_position = XL_TICK_LABEL_POSITION.NONE
    value_axis_line.format.line.dash_style = MSO_LINE_DASH_STYLE.ROUND_DOT
    value_axis_line.visible = False

    category_axis_line = chart_line.category_axis
    category_axis_line.has_major_gridlines = False
    category_axis_line.major_tick_mark = XL_TICK_MARK.NONE
    category_axis_line.tick_label_position = XL_TICK_LABEL_POSITION.NONE
    category_axis_line.format.line.dash_style = MSO_LINE_DASH_STYLE.ROUND_DOT
    category_axis_line.visible = False

    # 开始创建表格
    if view_rules == 'Brand':
        rows = len(brand_filter) * 2 + 1
    else:
        rows = len(oem_filter) * 2 + 1

    cols = len(data_years) + 1
    table_width = 24
    table_height = 3
    top = Cm(9.5)
    left = Cm(1.5)  # Inches(2.0)
    width = Cm(table_width)  # Inches(6.0)
    height = Cm(table_height)  # Inches(0.8)

    # 添加表格到幻灯片 --------------------
    table = shapes.add_table(rows, cols, left, top, width, height).table

    # 给data_year加E
    data_years_E = [str(year) + 'E' if year > start_year else str(year) for year in data_years]

    # 设置单元格宽度
    columns_width = table_width / cols - 0.1
    for i in range(cols):
        if i == 0:
            table.columns[i].width = Cm(columns_width + 0.6)
        else:
            table.columns[i].width = Cm(columns_width)  # Inches(2.0)

    row_height = table_height / rows
    for i in range(rows):
        table.rows[i].height = Cm(row_height)  # Inches(2.0)

    # 设置标题行
    for i in range(cols):
        if i == 0:
            if view_rules == 'Brand':
                table.cell(0, i).text = 'Vol.by Brand'
            else:
                table.cell(0, i).text = 'Vol.by OEM'
        else:
            table.cell(0, i).text = str(data_years_E[i - 1])

    # 填充表格数据
    if view_rules == 'Brand':
        print('maoyadong' + str(len(brand_filter)))
        for brand_idx, brand in enumerate(brand_filter):
            row_idx = brand_idx * 2
            print(row_idx)
            total_volumes = df_vw_group_brand_year[(df_vw_group_brand_year['Brand'] == brand)] \
                .groupby(['YEAR']).agg({'Volume': np.sum}).reset_index() \
                .sort_values(['YEAR'], ascending=[True])['Volume']
            series_mkt_rate = [vol / mkt_volume for vol in total_volumes]
            table.cell(row_idx + 1, 0).text = brand
            table.cell(row_idx + 2, 0).text = 'MKT%'
            for col_idx, vol in enumerate(total_volumes):
                table.cell(row_idx + 1, col_idx + 1).text = format(vol / 1000, '.0f')
                table.cell(row_idx + 2, col_idx + 1).text = format(series_mkt_rate[col_idx] * 100, '.1%')
    else:
        for oem_idx, oem in enumerate(oem_filter):
            row_idx = oem_idx * 2
            print(row_idx)
            total_volumes = df_vw_group_oem_year[(df_vw_group_oem_year['OEM'] == oem)] \
                .groupby(['YEAR']).agg({'Volume': np.sum}).reset_index() \
                .sort_values(['YEAR'], ascending=[True])['Volume']
            series_mkt_rate = [vol / mkt_volume for vol in total_volumes]
            table.cell(row_idx + 1, 0).text = oem
            table.cell(row_idx + 2, 0).text = 'MKT%'
            for col_idx, vol in enumerate(total_volumes):
                table.cell(row_idx + 1, col_idx + 1).text = format(vol / 1000, '.0f')
                table.cell(row_idx + 2, col_idx + 1).text = format(series_mkt_rate[col_idx] * 100, '.1%')


    # 调整table每个cell的字体
    def iter_cells(table):
        for row in table.rows:
            for cell in row.cells:
                yield cell


    for cell in iter_cells(table):
        if cell.text.strip() == '':
            cell.text = r'/'
        for paragraph in cell.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(10)

    # 开始添加注释文本框
    left = Cm(1)  # left，top为相对位置
    top = Cm(4)
    width = Cm(2)  # width，height为文本框的大小
    height = Cm(1)

    # 在指定位置添加文本框
    textbox = shapes.add_textbox(left, top, width, height)
    tf = textbox.text_frame

    # 在文本框中写入文字
    para = tf.add_paragraph()  # 新增段落
    para.text = "Volume\n'000units"  # 向段落写入文字
    para.line_spacing = 1.5  # 1.5 倍的行距
    para.font.size = Pt(10)

prs.save('c:/auto-report/template_tmp2.pptx')
print("maoyadong")

