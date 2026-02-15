"""
شیت ۳: تحلیل ماهانه
"""
import pandas as pd
from openpyxl.chart import BarChart, LineChart, AreaChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import ColorScaleRule

from config.constants import MONTH_NAMES
from config.colors import CF_COLORS, CHART_SPECIFIC
from styles.excel_styles import write_table, auto_width
from charts.chart_helpers import ChartPlacer, apply_single_color


def build(wb, kpi):
    """ساخت شیت تحلیل ماهانه"""
    ws = wb.create_sheet('تحلیل ماهانه')
    ws.sheet_view.rightToLeft = True

    headers = [
        'ماه', 'تعداد', 'میانگین (ساعت)', 'میانه', 'انحراف معیار',
        'حداقل', 'حداکثر', 'درصد به‌موقع', 'رشد میانگین (%)', 'تغییر به‌موقع',
    ]
    rows = []
    for m in kpi.available_months:
        if m in kpi.kpi_month.index:
            rd = kpi.kpi_month.loc[m]
            rows.append([
                MONTH_NAMES[m], int(rd['تعداد_مرسوله']),
                rd['میانگین_ساعت'], rd['میانه_ساعت'],
                rd['انحراف_معیار'], rd['حداقل_ساعت'], rd['حداکثر_ساعت'],
                rd['درصد_بموقع'],
                rd['رشد_میانگین'] if pd.notna(rd['رشد_میانگین']) else '-',
                rd['تغییر_بموقع'] if pd.notna(rd['تغییر_بموقع']) else '-',
            ])

    lr = write_table(ws, headers, rows, 1, '📅 تحلیل عملکرد ماهانه')
    auto_width(ws, len(headers))

    # Conditional Formatting
    ws.conditional_formatting.add(
        f'C3:C{lr - 1}',
        ColorScaleRule(
            start_type='min', start_color=CF_COLORS['good_start'],
            mid_type='percentile', mid_value=50, mid_color=CF_COLORS['mid'],
            end_type='max', end_color=CF_COLORS['bad_end'],
        ),
    )
    ws.conditional_formatting.add(
        f'H3:H{lr - 1}',
        ColorScaleRule(
            start_type='min', start_color=CF_COLORS['bad_end'],
            mid_type='percentile', mid_value=50, mid_color=CF_COLORS['mid'],
            end_type='max', end_color=CF_COLORS['good_start'],
        ),
    )

    placer = ChartPlacer(start_row=lr + 1)

    # نمودار ۱: روند خطی میانگین
    lc1 = LineChart()
    lc1.title = "روند میانگین زمان تحویل ماهانه"
    lc1.y_axis.title = "ساعت"
    lc1.style = 10
    lc1.width = 30
    lc1.height = 15
    lc1.add_data(
        Reference(ws, min_col=3, min_row=2, max_row=lr - 1),
        titles_from_data=True,
    )
    lc1.set_categories(Reference(ws, min_col=1, min_row=3, max_row=lr - 1))
    if lc1.series:
        lc1.series[0].graphicalProperties.line.width = 28000
    lc1.dataLabels = DataLabelList()
    lc1.dataLabels.showVal = True
    ws.add_chart(lc1, placer.next_anchor())

    # نمودار ۲: ترکیبی تعداد + درصد به‌موقع
    combo = BarChart()
    combo.title = "تعداد مرسوله و درصد تحویل به‌موقع (ماهانه)"
    combo.style = 10
    combo.width = 30
    combo.height = 15
    combo.add_data(
        Reference(ws, min_col=2, min_row=2, max_row=lr - 1),
        titles_from_data=True,
    )
    combo.set_categories(Reference(ws, min_col=1, min_row=3, max_row=lr - 1))
    combo.y_axis.title = "تعداد"

    lo = LineChart()
    lo.add_data(
        Reference(ws, min_col=8, min_row=2, max_row=lr - 1),
        titles_from_data=True,
    )
    lo.y_axis.axId = 200
    lo.y_axis.title = "درصد"
    if lo.series:
        lo.series[0].graphicalProperties.line.width = 28000
        lo.series[0].graphicalProperties.solidFill = CHART_SPECIFIC['combo_line']
    combo += lo
    ws.add_chart(combo, placer.next_anchor())

    # نمودار ۳: Area Chart
    area = AreaChart()
    area.title = "روند تجمعی تعداد مرسولات ماهانه"
    area.style = 10
    area.width = 30
    area.height = 15
    area.add_data(
        Reference(ws, min_col=2, min_row=2, max_row=lr - 1),
        titles_from_data=True,
    )
    area.set_categories(Reference(ws, min_col=1, min_row=3, max_row=lr - 1))
    ws.add_chart(area, placer.next_anchor())

    # نمودار ۴: میانگین و میانه ماهانه
    lc_mm = LineChart()
    lc_mm.title = "مقایسه میانگین و میانه ماهانه"
    lc_mm.y_axis.title = "ساعت"
    lc_mm.style = 10
    lc_mm.width = 30
    lc_mm.height = 15
    lc_mm.add_data(
        Reference(ws, min_col=3, min_row=2, max_row=lr - 1),
        titles_from_data=True,
    )
    lc_mm_med = LineChart()
    lc_mm_med.add_data(
        Reference(ws, min_col=4, min_row=2, max_row=lr - 1),
        titles_from_data=True,
    )
    if lc_mm_med.series:
        lc_mm_med.series[0].graphicalProperties.line.dashStyle = "dash"
    lc_mm.set_categories(Reference(ws, min_col=1, min_row=3, max_row=lr - 1))
    lc_mm += lc_mm_med
    ws.add_chart(lc_mm, placer.next_anchor())

    return ws
