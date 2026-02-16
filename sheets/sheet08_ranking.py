"""
شیت ۸: رتبه‌بندی نمایندگی‌ها
"""
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import (
    Paragraph, ParagraphProperties, CharacterProperties, RichTextProperties,
)
from openpyxl.formatting.rule import ColorScaleRule

from config.colors import CF_COLORS
from styles.excel_styles import write_table, auto_width
from charts.chart_helpers import ChartPlacer


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

    # ───────────────────────────────────────────────
    # نمودار رتبه‌بندی (ستونی عمودی با رنگ‌بندی هوشمند)
    # ───────────────────────────────────────────────
    chart = BarChart()
    chart.type = "col"  # ✅ ستونی عمودی
    chart.title = "رتبه‌بندی نمایندگی‌ها (امتیاز ترکیبی)"
    chart.y_axis.title = "امتیاز"
    chart.x_axis.title = "نمایندگی"
    chart.width = 36
    chart.height = 20
    chart.style = 10
    chart.legend = None  # حذف Legend

    chart.add_data(
        Reference(ws, min_col=7, min_row=2, max_row=lr - 1),
        titles_from_data=True,
    )
    chart.set_categories(
        Reference(ws, min_col=2, min_row=3, max_row=lr - 1)
    )

    # ── چرخش مورب اسم نمایندگی‌ها (دقیقاً مثل sheet02) ──
    chart.x_axis.txPr = RichText(
        p=[Paragraph(
            pPr=ParagraphProperties(
                defRPr=CharacterProperties(sz=900, b=False)
            ),
            endParaRPr=CharacterProperties(sz=900),
        )],
        bodyPr=RichTextProperties(
            rot=-45 * 60000,  # -45 درجه
            spcFirstLastPara=True,
            vertOverflow='ellipsis',
            vert='horz',
            wrap='square',
            anchor='ctr',
            anchorCtr=True,
        ),
    )
    chart.x_axis.tickLblPos = 'low'
    chart.x_axis.delete = False

    # ── رنگ‌بندی میله‌ها: سبز (۵ برتر)، قرمز (۵ ضعیف‌تر)، آبی (بقیه) ──
    total_agents = len(scoring)
    series = chart.series[0]

    for i in range(total_agents):
        pt = DataPoint(idx=i)
        if i < 5:
            # ۵ نفر برتر: سبز
            pt.graphicalProperties.solidFill = "27AE60"
        elif i >= total_agents - 5:
            # ۵ نفر ضعیف‌تر: قرمز
            pt.graphicalProperties.solidFill = "E74C3C"
        else:
            # بقیه: آبی
            pt.graphicalProperties.solidFill = "4472C4"
        series.data_points.append(pt)

    # ── برچسب‌های داده با چرخش عمودی (90 درجه) ──
    dl = DataLabelList()
    dl.showVal = True
    dl.showPercent = False
    dl.showCatName = False
    dl.showSerName = False

    # تنظیم چرخش عمودی برچسب‌ها
    dl.txPr = RichText(
        p=[Paragraph(
            pPr=ParagraphProperties(
                defRPr=CharacterProperties(sz=900, b=False)
            ),
        )],
        bodyPr=RichTextProperties(
            rot=-90 * 60000,  # -90 درجه (عمودی)
            vert='horz',
        ),
    )

    series.dLbls = dl

    # جایگذاری نمودار در شیت
    ws.add_chart(chart, placer.next_anchor(custom_height=22))

    return ws
