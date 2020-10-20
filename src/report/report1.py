from pptx import Presentation
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.util import Inches, Pt
from pptx import Presentation
from pptx.chart.data import ChartData, CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_TICK_MARK, XL_TICK_LABEL_POSITION, XL_MARKER_STYLE, XL_LABEL_POSITION
from pptx.util import Cm  # Inches
from pandas import np
from pptx.enum.chart import XL_LEGEND_POSITION
from collections import Counter, OrderedDict
from pptx.dml.color import RGBColor
import pandas as pd
from report_common import ReportBase
import report_common
import report_logger


# if __name__ == '__main__':
class Report1(ReportBase):

    def Run(self):
        df_vw = self.df_vw
        df_mkt = self.df_mkt

        PR_Status_actual = self.PR_Status_actual
        PR_Status_local = self.PR_Status_local
        PR_Status = self.PR_Status
        year_filter = self.year_filter
        oem_filter = self.oem_filter
        brand_filter = self.brand_filter
        build_type_filter = self.build_type_filter
        shapes = self.shapes
        year_categories = self.year_categories

        # 根据所有filter过滤大众数据和mkt数据
        df_vw_filter = self.df_vw_filter
        df_mkt_filter = self.df_mkt_filter


        fuel_type = self.fuel_type_group_all
        fuel_type_group = self.fuel_type_group_all

        # 全销售市场数据分组聚合-------------------------
        df_mkt_sum = df_mkt_filter.groupby(['Fuel_type', 'Year']).agg({'Volume': np.sum}).reset_index()
        df_vw_sum = df_vw_filter.groupby(['Fuel_Type_Group', 'YEAR']).agg({'Volume': np.sum}).reset_index()
        fuel_year_vol_dict = OrderedDict()
        fuel_year_rate_dict = OrderedDict()
        for fuel in fuel_type_group:
            year_vol_dict = OrderedDict()
            year_rate_dict = OrderedDict()
            fuel_type_calculate = report_common.get_fuel_from_fuelTypeGroup(fuel)
            for year in year_filter:
                vw_volume = 0
                mkt_volume = 0
                mkt_rate = 0
                if not df_vw_sum[(df_vw_sum['Fuel_Type_Group'] == fuel) & (df_vw_sum['YEAR'] == year)].empty:
                    vw_volume = \
                        df_vw_sum[(df_vw_sum['Fuel_Type_Group'] == fuel) & (df_vw_sum['YEAR'] == year)] \
                            .reset_index().at[0, 'Volume']
                else:
                    report_logger.record_log('report1', 'vw', PR_Status_local, year, '', fuel)

                if not df_mkt_sum[(df_mkt_sum['Fuel_type'].isin(fuel_type_calculate)) & (df_mkt_sum['Year'] == year)].empty:
                    mkt_volume = \
                        df_mkt_sum[(df_mkt_sum['Fuel_type'].isin(fuel_type_calculate)) & (df_mkt_sum['Year'] == year)] \
                            .agg({'Volume': np.sum}).reset_index().iat[0, 1]
                else:
                    report_logger.record_log('report1', 'mkt', PR_Status_local, year, fuel_type_calculate)

                if mkt_volume > 0:
                    mkt_rate = vw_volume / mkt_volume

                year_vol_dict[year] = vw_volume
                year_rate_dict[year] = mkt_rate
            fuel_year_vol_dict[fuel] = year_vol_dict
            fuel_year_rate_dict[fuel] = year_rate_dict
        print('fuel_year_vol_dict=' + str(fuel_year_vol_dict))
        print('fuel_year_rate_dict=' + str(fuel_year_rate_dict))

        total_year_vol_dict = OrderedDict()
        total_year_rate_dict = OrderedDict()
        for year in year_filter:
            all_vw_volume = 0
            all_mkt_volume = 0
            all_mkt_rate = 0

            df_all_vw_volume = \
                df_vw_sum[(df_vw_sum['Fuel_Type_Group'].isin(fuel_type_group)) & (df_vw_sum['YEAR'] == year)]
            if not df_all_vw_volume.empty:
                all_vw_volume = df_all_vw_volume.agg({'Volume': np.sum}).reset_index().iat[0, 1]
            else:
                report_logger.record_log('report1', 'vw', PR_Status_local, year, '', fuel_type_group)

            df_all_mkt_volume = \
                df_mkt_sum[(df_mkt_sum['Fuel_type'].isin(fuel_type)) & (df_mkt_sum['Year'] == year)]
            if not df_all_mkt_volume.empty:
                all_mkt_volume = df_all_mkt_volume.agg({'Volume': np.sum}).reset_index().iat[0, 1]
            else:
                report_logger.record_log('report1', 'mkt', PR_Status_local, year, fuel_type)

            if all_mkt_volume >= 0:
                all_mkt_rate = all_vw_volume / all_mkt_volume

            total_year_rate_dict[year] = all_mkt_rate
            total_year_vol_dict[year] = all_vw_volume
        print('total_year_rate_dict=' + str(total_year_rate_dict))


        top_base = 3
        # 开始创建点线图-----------------------------------
        chart_data_line = ChartData()
        chart_data_line.categories = year_filter
        series_line1 = total_year_rate_dict.values()
        chart_data_line.add_series('MS%', series_line1)

        x, y, cx, cy = Cm(1), Cm(top_base), Cm(24), Cm(4)
        chart_line = shapes.add_chart(
            XL_CHART_TYPE.LINE, x, y, cx, cy, chart_data_line
        ).chart

        chart_line.has_legend = True
        chart_line.legend.include_in_layout = False
        chart_line.legend.position = XL_LEGEND_POSITION.LEFT
        chart_line.legend.font.size = Pt(10)
        chart_line.chart_title.text_frame.clear()

        for line_serie in chart_line.series:
            line_serie.smooth = True
            line_serie.marker.style = XL_MARKER_STYLE.CIRCLE
            line_serie.data_labels.show_value = True
            line_serie.data_labels.number_format = '0.0%'
            line_serie.data_labels.font.size = Pt(10)
            line_serie.data_labels.position = XL_LABEL_POSITION.ABOVE

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

        # 开始创建柱状堆积图---------------------------------
        chart_data_stack = CategoryChartData()
        chart_data_stack.categories = year_categories
        # chart_data_stack.categories = [2016,2017,2018,2019,2020,2021,2022,2023,2024,2025,2026,2027]

        for fuel in fuel_type_group:
            series_Volumes = fuel_year_vol_dict[fuel].values()
            series_Volumes = [vol / 1000 for vol in series_Volumes]
            chart_data_stack.add_series(fuel, series_Volumes)
        chart_data_stack.add_series('', [vol * 0.1 / 1000 for vol in total_year_vol_dict.values()])
        # chart_data_stack.add_series('', [300,400,500,600,700,800,900,1000,1100,1200,1300,1400])

        x, y, cx, cy = Cm(1), Cm(top_base + 3), Cm(24), Cm(7)
        graphic_frame_stack = shapes.add_chart(
            XL_CHART_TYPE.COLUMN_STACKED, x, y, cx, cy, chart_data_stack
        )

        chart_stack = graphic_frame_stack.chart
        # chart_stack.has_title = True
        # chart_stack.chart_title.has_text_frame = True
        # chart_stack.chart_title.text_frame.text = "maoyadong"
        # chart_stack.chart_title.text_frame.paragraphs[0].font.size = Pt(10)

        chart_stack.has_legend = True
        chart_stack.legend.position = XL_LEGEND_POSITION.LEFT  # XL_LEGEND_POSITION.CORNER
        chart_stack.legend.include_in_layout = False
        chart_stack.legend.font.size = Pt(10)

        for series_stack in chart_stack.series:
            series_stack.data_labels.show_value = True
            series_stack.data_labels.number_format = '0'
            series_stack.data_labels.font.size = Pt(10)

        series_stack2 = chart_stack.series[2]
        series_stack2.data_labels.font.size = Pt(10)
        series_stack2.data_labels.show_value = True
        series_stack2.format.fill.patterned()  # 设置为可更改填充颜色的
        series_stack2.format.fill.solid()  # 设置为纯色填充
        series_stack2.format.fill.fore_color.rgb = RGBColor(255, 255, 255)  # 白色

        for idx, year in enumerate(year_filter):
            series_point = series_stack2.points[idx]
            series_point.data_label.has_text_frame = True
            paragraphs1 = series_point.data_label.text_frame.paragraphs[0]
            run1 = paragraphs1.add_run()
            run1.text = format(total_year_vol_dict[year] / 1000, '.0f')
            # run1.font.size = font_size

        value_axis_stack = chart_stack.value_axis
        value_axis_stack.has_major_gridlines = False
        value_axis_stack.major_tick_mark = XL_TICK_MARK.NONE
        value_axis_stack.tick_label_position = XL_TICK_LABEL_POSITION.NONE
        value_axis_stack.format.line.dash_style = MSO_LINE_DASH_STYLE.ROUND_DOT
        value_axis_stack.visible = False

        category_axis_stack = chart_stack.category_axis
        category_axis_stack.has_major_gridlines = False
        category_axis_stack.major_tick_mark = XL_TICK_MARK.NONE
        category_axis_stack.tick_label_position = XL_TICK_LABEL_POSITION.LOW
        category_axis_stack.tick_labels.font.size = Pt(10)
        category_axis_stack.format.line.dash_style = MSO_LINE_DASH_STYLE.SOLID
        category_axis_stack.visible = True

        # value_axis_stack.has_title = True
        # value_axis_stack.axis_title.has_text_frame = True
        # value_axis_stack.axis_title.text_frame.text = "False positive"
        # value_axis_stack.axis_title.text_frame.paragraphs[0].font.size = Pt(10)

        # 开始创建表格
        rows = 3
        cols = len(year_filter) + 1
        table_width = 24
        table_height = 3
        top = Cm(top_base + 10)
        left = Cm(1)  # Inches(2.0)
        width = Cm(table_width)  # Inches(6.0)
        height = Cm(table_height)  # Inches(0.8)

        # 添加表格到幻灯片 --------------------
        table = shapes.add_table(rows, cols, left, top, width, height).table

        # 设置单元格宽度
        columns_width = table_width / cols
        for i in range(cols):
            table.columns[i].width = Cm(columns_width)  # Inches(2.0)

        row_height = table_height / rows
        for i in range(rows):
            table.rows[i].height = Cm(row_height)  # Inches(2.0)

        # 设置标题行
        for i in range(cols):
            if i == 0:
                table.cell(0, i).text = 'MS%'
            else:
                table.cell(0, i).text = str(year_categories[i - 1])

        # 填充表格数据
        for row_idx, fuel_group in enumerate(fuel_type_group):
            table.cell(row_idx + 1, 0).text = fuel_group + ' MKT%'
            for col_idx, year in enumerate(year_filter):
                rate = fuel_year_rate_dict[fuel_group][year]
                table.cell(row_idx + 1, col_idx + 1).text = format(rate, '.1%')

        # 调整table每个cell的字体
        report_common.set_table_format(table, Pt(10), r'/')

        # 开始添加注释文本框
        left = Cm(1)  # left，top为相对位置
        top = Cm(top_base - 1)
        width = Cm(2)  # width，height为文本框的大小
        height = Cm(1)

        # 在指定位置添加文本框
        textbox = shapes.add_textbox(left, top, width, height)
        report_common.set_textbox_format(textbox, Pt(8), "Volume\n'000units")


