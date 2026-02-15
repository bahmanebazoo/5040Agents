"""
شیت ۵: توزیع ساعتی
"""
from openpyxl.chart import BarChart, LineChart, AreaChart, Reference
from openpyxl.formatting.rule import DataBarRule

from config.colors import CF_COLORS, CHART_SPECIFIC
from styles.excel_styles import write_table, auto_width
from charts.chart_helpers import ChartPlacer, colorize_category_bars


def build(wb, kpi):
    """ساخت شیت توزیع ساعتی"""
    ws = wb.create_sheet('توزیع ساعتی')
    ws.sheet_view.rightToLeft = True

    headers = ['ساعت', 'تعداد ورود', 'تعداد تحویل', 'درصد ورود', 'درصد تحویل']
    rows = []
    total_e = kpi.hourly_entry.sum()
    total_d = kpi.hourly_delivery.sum()
    for h in range(24):
        e = int(kpi.hourly_entry.get(h, 0))
        d = int(kpi.hourly_delivery.get(h, 0))
        rows.append([
            f'{h:02d}:00', e, d,
            round(e / total_e * 100, 1) if total_e > 0 else 0,
            round(d / total_d * 100, 1) if total_d > 0 else 0,
        ])

    lr = write_table(ws, headers, rows, 1, '⏰ توزیع ساعتی ورود و تحویل مرسولات')
    auto_width(ws, len(headers))

    ws.conditional_formatting.add(
        f'B3:B{lr - 1}',
        DataBarRule(start_type='min', end_type='max', color=CF_COLORS['databar_blue']),
    )
    ws.conditional_formatting.add(
        f'C3:C{lr - 1}',
        DataBarRule(start_type='min', end_type='max', color='ED7D31'),
    )

    placer = ChartPlacer(start_row=lr + 1)

    # نمودار ۱: ستونی - ساعتی → چندرنگ (دسته‌بندی مفهومی = ساعت)
    # اما اینجا دو سری داریم (ورود + تحویل) → رنگ ثابت هر سری
    ch1 = BarChart()
    ch1.type = "col"
    ch1.title = "توزیع ساعتی ورود و تحویل"
    ch1.y_axis.title = "تعداد"
    ch1.style = 10
    ch1.width = 35
    ch1.height = 15
    ch1.add_data(
        Reference(ws, min_col=2, min_row=2, max_row=lr - 1),
        titles_from_data=True,
    )
    ch1.add_data(
        Reference(ws, min_col=3, min_row=2, max_row=lr - 1),
        titles_from_data=True,
    )
    ch1.set_categories(Reference(ws, min_col=1, min_row=3, max_row=lr - 1))
    if len(ch1.series) >= 2:
        ch1.series[0].graphicalProperties.solidFill = CHART_SPECIFIC['entry_bar']
        ch1.series[1].graphicalProperties.solidFill = CHART_SPECIFIC['delivery_bar']
    ws.add_chart(ch1, placer.next_anchor())

    # نمودار ۲: خطی مقایسه
    lc = LineChart()
    lc.title = "مقایسه خطی توزیع ساعتی ورود و تحویل"
    lc.y_axis.title = "تعداد"
    lc.style = 10
    lc.width = 35
    lc.height = 15
    lc.add_data(
        Reference(ws, min_col=2, min_row=2, max_row=lr - 1),
        titles_from_data=True,
    )
    lc.add_data(
        Reference(ws, min_col=3, min_row=2, max_row=lr - 1),
        titles_from_data=True,
    )
    lc.set_categories(Reference(ws, min_col=1, min_row=3, max_row=lr - 1))
    ws.add_chart(lc, placer.next_anchor())

    # نمودار ۳: Area
    area = AreaChart()
    area.title = "توزیع تجمعی ساعتی"
    area.style = 10
    area.width = 35
    area.height = 15
    area.add_data(
        Reference(ws, min_col=2, min_row=2, max_row=lr - 1),
        titles_from_data=True,
    )
    area.add_data(
        Reference(ws, min_col=3, min_row=2, max_row=lr - 1),
        titles_from_data=True,
    )
    area.set_categories(Reference(ws, min_col=1, min_row=3, max_row=lr - 1))
    ws.add_chart(area, placer.next_anchor())

    return ws
