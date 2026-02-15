"""
شیت ۸: رتبه‌بندی نمایندگی‌ها
"""
from openpyxl.chart import BarChart, RadarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import ColorScaleRule

from config.colors import CHART_SPECIFIC, CF_COLORS
from styles.excel_styles import write_table, auto_width, style_header
from charts.chart_helpers import ChartPlacer, apply_single_color, add_data_labels


def build(wb, kpi, scoring):
    """ساخت شیت رتبه‌بندی"""
    ws = wb.create_sheet('رتبه‌بندی نمایندگی‌ها')
    ws.sheet_view.rightToLeft = True

    headers = [
        'رتبه', 'نمایندگی', 'تعداد', 'میانگین', 'درصد به‌موقع',
        'انحراف معیار', 'امتیاز نهایی', 'سطح عملکرد',
    ]
    rows = []
    for rank, (agent, rd) in enumerate(scoring.iterrows(), 1):
        rows.append([
            rank, agent, rd['تعداد_مرسوله'], rd['میانگین_ساعت'],
            rd['درصد_بموقع'], rd['انحراف_معیار'],
            rd['امتیاز_نهایی'], rd['سطح_عملکرد'],
        ])

    lr = write_table(ws, headers, rows, 1, '🏆 رتبه‌بندی جامع نمایندگی‌ها (امتیاز ترکیبی)')
    auto_width(ws, len(headers))

    ws.conditional_formatting.add(
        f'G3:G{lr - 1}',
        ColorScaleRule(
            start_type='min', start_color=CF_COLORS['bad_end'],
            mid_type='percentile', mid_value=50, mid_color=CF_COLORS['mid'],
            end_type='max', end_color=CF_COLORS['good_start'],
        ),
    )

    placer = ChartPlacer(start_row=lr + 1)

    # نمودار امتیاز (تک‌رنگ - محور = نمایندگی)
    ch = BarChart()
    ch.type = "bar"
    ch.title = "رتبه‌بندی نمایندگی‌ها (امتیاز ترکیبی)"
    ch.x_axis.title = "امتیاز"
    ch.style = 10
    ch.width = 35
    ch.height = 22
    ch.add_data(
        Reference(ws, min_col=7, min_row=2, max_row=lr - 1),
        titles_from_data=True,
    )
    ch.set_categories(Reference(ws, min_col=2, min_row=3, max_row=lr - 1))
    apply_single_color(ch, CHART_SPECIFIC['ranking_bar'])
    add_data_labels(ch, show_val=True)
    ws.add_chart(ch, placer.next_anchor(custom_height=24))

    # ─── Radar Top 5 ───
    top5 = scoring.head(5).index.tolist()

    sr_radar = placer.current_row + 2
    ws.cell(row=sr_radar, column=1, value='معیار')
    for ci, agent in enumerate(top5, 2):
        ws.cell(row=sr_radar, column=ci, value=agent)

    radar_metrics = ['امتیاز میانگین', 'درصد به‌موقع', 'پایداری (انحراف معیار)']
    for mi, metric in enumerate(radar_metrics):
        r = sr_radar + 1 + mi
        ws.cell(row=r, column=1, value=metric)
        for ci, agent in enumerate(top5, 2):
            if metric == 'امتیاز میانگین':
                ws.cell(row=r, column=ci, value=round(scoring.loc[agent, 'score_avg'], 1))
            elif metric == 'درصد به‌موقع':
                ws.cell(row=r, column=ci, value=scoring.loc[agent, 'درصد_بموقع'])
            else:
                ws.cell(row=r, column=ci, value=round(scoring.loc[agent, 'score_std'], 1))

    radar = RadarChart()
    radar.type = "filled"
    radar.title = "مقایسه رادار ۵ نمایندگی برتر"
    radar.style = 10
    radar.width = 25
    radar.height = 18
    for ci in range(2, 2 + len(top5)):
        radar.add_data(
            Reference(ws, min_col=ci, min_row=sr_radar,
                      max_row=sr_radar + len(radar_metrics)),
            titles_from_data=True,
        )
    radar.set_categories(
        Reference(ws, min_col=1, min_row=sr_radar + 1,
                  max_row=sr_radar + len(radar_metrics))
    )
    ws.add_chart(radar, f"A{sr_radar + len(radar_metrics) + 2}")

    return ws
