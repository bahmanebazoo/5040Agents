"""
charts/chart_helpers.py
توابع کمکی نمودارها + ChartPlacer با فاصله‌گذاری پیکسلی دقیق

واحدها:
    1 EMU = 1/914400 اینچ
    1 px  = 9525 EMU  (در 96 DPI)
"""
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList
from openpyxl.utils import get_column_letter


# ──────────────────────────────────────────────
# ثوابت تبدیل
# ──────────────────────────────────────────────
EMU_PER_PIXEL = 9525                    # 1 px = 9525 EMU (96 DPI)
DEFAULT_ROW_HEIGHT_PX = 15              # ارتفاع پیش‌فرض هر ردیف اکسل


def px_to_emu(px: int) -> int:
    """تبدیل پیکسل به EMU"""
    return int(px * EMU_PER_PIXEL)


# ──────────────────────────────────────────────
# ChartPlacer — جایگذاری دقیق نمودارها
# ──────────────────────────────────────────────
class ChartPlacer:
    """
    جایگذاری نمودارها با فاصله پیکسلی دقیق.

    استراتژی:
        ۱. محاسبه ردیف و آفست عمودی بر اساس gap_px
        ۲. استفاده از ws.add_chart با cell string
        ۳. بلافاصله بعد، تنظیم colOff و rowOff روی anchor نمودار

    پارامترها:
        start_row        : اولین ردیف مجاز (1-based)
        right_margin_px  : فاصله از سمت راست (px) → در RTL = فاصله از ستون A
        gap_px           : فاصله عمودی بین نمودارها (px)
        default_chart_height_rows : ارتفاع پیش‌فرض نمودار (تعداد ردیف)
        anchor_col       : ستون شروع نمودار (1-based). 1 = A
    """

    def __init__(
        self,
        start_row: int = 1,
        right_margin_px: int = 24,
        gap_px: int = 96,
        default_chart_height_rows: int = 18,
        anchor_col: int = 1,
    ):
        self._start_row = start_row
        self._right_margin_px = right_margin_px
        self._gap_px = gap_px
        self._default_chart_height_rows = default_chart_height_rows
        self._anchor_col = anchor_col        # 1-based

        # وضعیت داخلی
        self._current_row = start_row        # 1-based: ردیف بعدی مجاز
        self._is_first = True

        # مقادیر EMU از قبل محاسبه شده
        self._right_margin_emu = px_to_emu(right_margin_px)
        self._gap_emu = px_to_emu(gap_px)

    # ── Properties ──

    @property
    def current_row(self) -> int:
        """ردیف فعلی (1-based)"""
        return self._current_row

    @current_row.setter
    def current_row(self, value: int):
        self._current_row = value

    # ── محاسبه موقعیت ──

    def _compute_target_row_and_offset(self) -> tuple:
        """
        محاسبه ردیف هدف (1-based) و offset عمودی (EMU).

        Returns:
            (target_row_1based, row_offset_emu)
        """
        if self._is_first:
            return self._current_row, 0

        # gap_px را به ردیف‌های کامل + باقیمانده تقسیم می‌کنیم
        gap_full_rows = self._gap_px // DEFAULT_ROW_HEIGHT_PX
        gap_remainder_px = self._gap_px % DEFAULT_ROW_HEIGHT_PX

        target_row = self._current_row + gap_full_rows
        row_offset_emu = px_to_emu(gap_remainder_px)

        return target_row, row_offset_emu

    # ── روش اصلی: place_chart ──

    def place_chart(self, ws, chart, custom_height_rows: int = None):
        """
        جایگذاری نمودار با دقت پیکسلی.

        مراحل:
          1. محاسبه ردیف و آفست
          2. ws.add_chart با cell string
          3. تنظیم colOff و rowOff روی anchor

        Args:
            ws: شیت اکسل
            chart: نمودار openpyxl
            custom_height_rows: ارتفاع سفارشی (ردیف)
        """
        target_row, row_off_emu = self._compute_target_row_and_offset()
        col_letter = get_column_letter(self._anchor_col)

        # ── مرحله ۱: اضافه کردن نمودار با cell reference ──
        cell_ref = f"{col_letter}{target_row}"
        ws.add_chart(chart, cell_ref)

        # ── مرحله ۲: تنظیم آفست‌های پیکسلی روی anchor ──
        # وقتی ws.add_chart صدا زده می‌شود، openpyxl یک TwoCellAnchor
        # می‌سازد و آن را در chart.anchor ذخیره می‌کند.
        # anchor._from یک AnchorMarker است با col, colOff, row, rowOff
        anchor = chart.anchor

        if hasattr(anchor, '_from'):
            # TwoCellAnchor — حالت عادی
            anchor._from.colOff = self._right_margin_emu
            anchor._from.rowOff = row_off_emu
        elif hasattr(anchor, 'col'):
            # اگر مستقیم AnchorMarker باشد
            anchor.colOff = self._right_margin_emu
            anchor.rowOff = row_off_emu

        # ── مرحله ۳: به‌روزرسانی ردیف فعلی ──
        chart_h = custom_height_rows or self._default_chart_height_rows
        self._current_row = target_row + chart_h
        self._is_first = False

    # ── روش ساده (بدون آفست پیکسلی) — برای سازگاری عقب‌گرد ──

    def next_anchor(self, custom_height: int = None) -> str:
        """
        فقط cell string برمی‌گرداند (بدون آفست پیکسلی).
        برای مواقعی که place_chart قابل استفاده نیست.

        ⚠️ توصیه: از place_chart استفاده کنید.
        """
        target_row, _ = self._compute_target_row_and_offset()
        col_letter = get_column_letter(self._anchor_col)

        chart_h = custom_height or self._default_chart_height_rows
        self._current_row = target_row + chart_h
        self._is_first = False

        return f"{col_letter}{target_row}"


# ──────────────────────────────────────────────
# توابع کمکی نمودار
# ──────────────────────────────────────────────

def apply_single_color(chart, color_hex: str):
    """اعمال رنگ واحد به تمام سری‌های نمودار"""
    for s in chart.series:
        s.graphicalProperties.solidFill = color_hex


def colorize_pie(chart, colors: list, count: int):
    """رنگ‌آمیزی برش‌های نمودار دایره‌ای"""
    if not chart.series:
        return
    for i in range(min(count, len(colors))):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = colors[i]
        chart.series[0].data_points.append(pt)


def colorize_category_bars(chart, num_categories: int):
    """
    رنگ‌آمیزی هر ستون نمودار میله‌ای با رنگ متفاوت.

    برای نمودارهایی که یک سری دارند ولی هر دسته باید رنگ جدا داشته باشد.
    رنگ‌ها از پالت پیش‌فرض زیر استفاده می‌شوند:
    """
    CATEGORY_PALETTE = [
        "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5",
        "70AD47", "264478", "9B57A0", "636363", "EB7E30",
        "2F5496", "BF9000", "43682B", "C55A11", "7030A0",
        "D34817", "00B0F0", "00B050", "FF0000", "FFC000",
        "8064A2", "F79646", "4BACC6", "C0504D", "9BBB59",
        "8DB4E2", "F4B183", "C5E0B4", "FFD966", "D6B4E0",
        "A9D08E", "BDD7EE", "F8CBAD", "E2EFDA", "FFE699",
    ]

    if not chart.series:
        return

    series = chart.series[0]
    for i in range(num_categories):
        pt = DataPoint(idx=i)
        color = CATEGORY_PALETTE[i % len(CATEGORY_PALETTE)]
        pt.graphicalProperties.solidFill = color
        series.data_points.append(pt)


def add_data_labels(
    chart,
    show_val: bool = False,
    show_percent: bool = False,
    show_cat: bool = False,
    separator: str = "\n",
    font_size: int = 900,
):
    """
    افزودن برچسب‌های داده به نمودار.
    separator = "\\n" → تعداد و درصد در دو خط جدا
    """
    if not chart.series:
        return
    dl = DataLabelList()
    dl.showVal = show_val
    dl.showPercent = show_percent
    dl.showCatName = show_cat
    dl.showSerName = False
    dl.separator = separator
    chart.series[0].dLbls = dl
