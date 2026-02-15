"""
شیت ۹: مرسولات پرت (Outliers)
"""
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import ColorScaleRule

from config.colors import CHART_SPECIFIC, CF_COLORS
from styles.excel_styles import write_table, auto_width
from charts.chart_helpers import ChartPlacer, apply_single_color, add_data_labels


def build(wb, kpi):
    """ساخت شیت مرسولات پرت"""
    ws = wb.create_sheet('مرسولات پرت')
    ws.sheet_view.rightToLeft = True

    df_out = kpi.df_outliers

    headers = [
        'سریال فاکتور', 'نمایندگی', 'تاریخ ورود', 'تاریخ تحویل',
        'اختلاف (ساعت)', 'ماه',
    ]
    rows = []
    for _, rd in df_out.sort_values('اختلاف_ساعت', ascending=False).iterrows():
        rows.append([
            rd['سریال فاکتور'], rd['نمایندگی'],
            str(rd['تاریخ ورود به نمایندگی']), str(rd['تاریخ تحویل']),
            round(rd['اختلاف_ساعت'], 2), rd.get('نام_ماه', ''),
        ])

    lr = write_table(
        ws, headers, rows, 1,
        f'⚠️ مرسولات پرت (اختلاف بیش از {round(kpi.outlier_threshold, 1)} '
        f'ساعت) - تعداد: {len(df_out)}',
    )
    auto_width(ws, len(headers))

    ws.conditional_formatting.add(
        f'E3:E{lr - 1}',
        ColorScaleRule(
            start_type='min', start_color=CF_COLORS['mid'],
            end_type='max', end_color=CF_COLORS['bad_end'],
        ),
    )

    # تعداد پرت به تفکیک نمایندگی
    outlier_by_agent = df_out.groupby('نمایندگی').size().sort_values(ascending=False)
    sr = lr + 2
    ws.cell(row=sr, column=1, value='نمایندگی')
    ws.cell(row=sr, column=2, value='تعداد پرت')
    for i, (agent, cnt) in enumerate(outlier_by_agent.items()):
        ws.cell(row=sr + 1 + i, column=1, value=agent)
        ws.cell(row=sr + 1 + i, column=2, value=int(cnt))

    placer = ChartPlacer(start_row=sr + len(outlier_by_agent) + 2)

    ch = BarChart()
    ch.type = "bar"
    ch.title = "تعداد مرسولات پرت به تفکیک نمایندگی"
    ch.style = 10
    ch.width = 30
    ch.height = 16
    ch.add_data(
        Reference(ws, min_col=2, min_row=sr, max_row=sr + len(outlier_by_agent)),
        titles_from_data=True,
    )
    ch.set_categories(
        Reference(ws, min_col=1, min_row=sr + 1, max_row=sr + len(outlier_by_agent))
    )
    apply_single_color(ch, CHART_SPECIFIC['outlier_bar'])
    add_data_labels(ch, show_val=True)
    ws.add_chart(ch, placer.next_anchor())

    return ws
