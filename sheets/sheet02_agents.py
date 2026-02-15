"""
شیت ۲: عملکرد نمایندگی‌ها
+ نمودار پراکندگی (Scatter/Bubble)

فاصله‌گذاری:
  - ۲۴px از سمت راست (در RTL = از ستون A)
  - ۹۶px بین هر دو نمودار
"""
from openpyxl.chart import BarChart, LineChart, ScatterChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import (
    Paragraph, ParagraphProperties, CharacterProperties, RichTextProperties,
)
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule

from config.colors import (
    SINGLE_BAR_COLOR, CHART_SPECIFIC, CF_COLORS, SCATTER_COLORS,
)
from styles.excel_styles import write_table, auto_width
from charts.chart_helpers import ChartPlacer, apply_single_color


# ──────────────────────────────────────────────
# توابع کمکی داخلی
# ──────────────────────────────────────────────

def _remove_legend_and_rotate_x(chart, rotation=-45):
    """
    ۱. حذف Legend از نمودار میله‌ای
    ۲. چرخش مورب لیبل‌های محور X
    """
    chart.legend = None

    chart.x_axis.txPr = RichText(
        p=[Paragraph(
            pPr=ParagraphProperties(
                defRPr=CharacterProperties(sz=900, b=False)
            ),
            endParaRPr=CharacterProperties(sz=900),
        )],
        bodyPr=RichTextProperties(
            rot=rotation * 60000,
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


def _rotate_x_labels(chart, rotation=-45):
    """فقط چرخش لیبل‌های محور X (بدون حذف Legend) — برای نمودار ترکیبی"""
    chart.x_axis.txPr = RichText(
        p=[Paragraph(
            pPr=ParagraphProperties(
                defRPr=CharacterProperties(sz=900)
            ),
            endParaRPr=CharacterProperties(sz=900),
        )],
        bodyPr=RichTextProperties(
            rot=rotation * 60000,
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


# ──────────────────────────────────────────────
# تابع اصلی ساخت شیت
# ──────────────────────────────────────────────

def build(wb, kpi):
    """ساخت شیت عملکرد نمایندگی‌ها"""
    ws = wb.create_sheet('عملکرد نمایندگی‌ها')
    ws.sheet_view.rightToLeft = True

    kpi_agent = kpi.kpi_agent

    # ─── جدول ───
    headers = [
        'نمایندگی', 'تعداد', 'میانگین (ساعت)', 'میانه (ساعت)',
        'انحراف معیار', 'حداقل', 'حداکثر', 'Q1', 'Q3', 'IQR',
        'درصد به‌موقع',
    ]
    rows = []
    for agent, rd in kpi_agent.iterrows():
        rows.append([
            agent, rd['تعداد_مرسوله'], rd['میانگین_ساعت'], rd['میانه_ساعت'],
            rd['انحراف_معیار'], rd['حداقل_ساعت'], rd['حداکثر_ساعت'],
            rd['Q1'], rd['Q3'], rd['IQR'], rd['درصد_بموقع'],
        ])

    lr = write_table(ws, headers, rows, 1, '📋 عملکرد کامل نمایندگی‌ها')
    auto_width(ws, len(headers))

    # ─── Conditional Formatting ───
    ws.conditional_formatting.add(
        f'C3:C{lr - 1}',
        ColorScaleRule(
            start_type='min', start_color=CF_COLORS['good_start'],
            mid_type='percentile', mid_value=50, mid_color=CF_COLORS['mid'],
            end_type='max', end_color=CF_COLORS['bad_end'],
        ),
    )
    ws.conditional_formatting.add(
        f'K3:K{lr - 1}',
        ColorScaleRule(
            start_type='min', start_color=CF_COLORS['bad_end'],
            mid_type='percentile', mid_value=50, mid_color=CF_COLORS['mid'],
            end_type='max', end_color=CF_COLORS['good_start'],
        ),
    )
    ws.conditional_formatting.add(
        f'B3:B{lr - 1}',
        DataBarRule(
            start_type='min', end_type='max',
            color=CF_COLORS['databar_blue'],
        ),
    )

    # ══════════════════════════════════════════════════════════════
    # نمودارها — فاصله‌گذاری دقیق:
    #   ۲۴px از سمت راست
    #   ۹۶px فاصله عمودی بین هر نمودار
    # ══════════════════════════════════════════════════════════════
    placer = ChartPlacer(
        start_row=lr + 1,
        right_margin_px=24,
        gap_px=96,
        default_chart_height_rows=18,
        anchor_col=1,       # ستون A (1-based)
    )

    # ═══════════════════════════════════════════
    # نمودار ۱: میانگین زمان تحویل
    # ═══════════════════════════════════════════
    ch1 = BarChart()
    ch1.type = "col"
    ch1.title = "میانگین زمان تحویل به تفکیک نمایندگی (ساعت)"
    ch1.y_axis.title = "ساعت"
    ch1.style = 10
    ch1.width = 38
    ch1.height = 16
    ch1.add_data(
        Reference(ws, min_col=3, min_row=2, max_row=lr - 1),
        titles_from_data=True,
    )
    ch1.set_categories(Reference(ws, min_col=1, min_row=3, max_row=lr - 1))
    ch1.shape = 4
    apply_single_color(ch1, SINGLE_BAR_COLOR)
    _remove_legend_and_rotate_x(ch1, rotation=-45)
    placer.place_chart(ws, ch1)

    # ═══════════════════════════════════════════
    # نمودار ۲: درصد به‌موقع
    # ═══════════════════════════════════════════
    ch2 = BarChart()
    ch2.type = "col"
    ch2.title = "درصد تحویل به‌موقع (≤۲۴ ساعت) به تفکیک نمایندگی"
    ch2.y_axis.title = "درصد"
    ch2.style = 10
    ch2.width = 38
    ch2.height = 16
    ch2.add_data(
        Reference(ws, min_col=11, min_row=2, max_row=lr - 1),
        titles_from_data=True,
    )
    ch2.set_categories(Reference(ws, min_col=1, min_row=3, max_row=lr - 1))
    apply_single_color(ch2, CHART_SPECIFIC['ontime_bar'])
    _remove_legend_and_rotate_x(ch2, rotation=-45)
    placer.place_chart(ws, ch2)

    # ═══════════════════════════════════════════
    # نمودار ۳: میانگین + میانه (Combo)
    # ═══════════════════════════════════════════
    ch3 = BarChart()
    ch3.title = "مقایسه میانگین و میانه زمان تحویل"
    ch3.y_axis.title = "ساعت"
    ch3.style = 10
    ch3.width = 38
    ch3.height = 16
    ch3.add_data(
        Reference(ws, min_col=3, min_row=2, max_row=lr - 1),
        titles_from_data=True,
    )
    ch3.set_categories(Reference(ws, min_col=1, min_row=3, max_row=lr - 1))
    apply_single_color(ch3, SINGLE_BAR_COLOR)

    line_median = LineChart()
    line_median.add_data(
        Reference(ws, min_col=4, min_row=2, max_row=lr - 1),
        titles_from_data=True,
    )
    if line_median.series:
        line_median.series[0].graphicalProperties.line.width = 25000
        line_median.series[0].graphicalProperties.solidFill = CHART_SPECIFIC['median_line']
    ch3 += line_median
    _rotate_x_labels(ch3, rotation=-45)   # Legend باقی می‌ماند (Combo)
    placer.place_chart(ws, ch3)

    # ═══════════════════════════════════════════
    # نمودار ۴: انحراف معیار
    # ═══════════════════════════════════════════
    ch4 = BarChart()
    ch4.type = "col"
    ch4.title = "انحراف معیار زمان تحویل (پایداری عملکرد)"
    ch4.y_axis.title = "ساعت"
    ch4.style = 10
    ch4.width = 38
    ch4.height = 16
    ch4.add_data(
        Reference(ws, min_col=5, min_row=2, max_row=lr - 1),
        titles_from_data=True,
    )
    ch4.set_categories(Reference(ws, min_col=1, min_row=3, max_row=lr - 1))
    apply_single_color(ch4, CHART_SPECIFIC['std_bar'])
    _remove_legend_and_rotate_x(ch4, rotation=-45)
    placer.place_chart(ws, ch4)

    # ═══════════════════════════════════════════
    # نمودار ۵: Scatter — از جدول موجود بالا
    # ═══════════════════════════════════════════
    _build_scatter_chart(ws, kpi_agent, lr, placer)

    return ws


def _build_scatter_chart(ws, kpi_agent, table_last_row, placer: ChartPlacer):
    """
    نمودار پراکندگی:
    - محور X = تعداد سفارشات  → ستون B (col=2) جدول بالا
    - محور Y = میانگین ساعت   → ستون C (col=3) جدول بالا
    - اندازه مارکر = حجم سفارشات (نرمال‌شده)

    ✅ بدون نوشتن داده تکراری — مستقیم از جدول اصلی
    """
    agents = list(kpi_agent.index)
    max_count = kpi_agent['تعداد_مرسوله'].max()
    min_marker = 8
    max_marker = 35

    data_start_row = 3

    scatter = ScatterChart()
    scatter.title = "پراکندگی: تعداد سفارشات vs میانگین زمان تحویل"
    scatter.x_axis.title = "تعداد سفارشات"
    scatter.y_axis.title = "میانگین ساعت تحویل"
    scatter.style = 10
    scatter.width = 38
    scatter.height = 20

    for i, agent in enumerate(agents):
        r = data_start_row + i

        x_ref = Reference(ws, min_col=2, min_row=r, max_row=r)
        y_ref = Reference(ws, min_col=3, min_row=r, max_row=r)

        scatter.add_data(y_ref, titles_from_data=False)

        current_series = scatter.series[-1]
        current_series.xvalues = x_ref
        current_series.tx = SeriesLabel(v=agent)

        # اندازه مارکر متناسب با حجم
        count = kpi_agent.loc[agent, 'تعداد_مرسوله']
        if max_count > 0:
            ratio = count / max_count
            marker_size = int(min_marker + ratio * (max_marker - min_marker))
        else:
            marker_size = min_marker

        current_series.graphicalProperties.line.noFill = True

        color = SCATTER_COLORS[i % len(SCATTER_COLORS)]
        current_series.graphicalProperties.solidFill = color

        from openpyxl.chart.marker import Marker
        current_series.marker = Marker(symbol='circle', size=marker_size)
        current_series.marker.graphicalProperties.solidFill = color

    # جایگذاری با دقت پیکسلی
    placer.place_chart(ws, scatter, custom_height_rows=22)
