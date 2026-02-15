"""
پالت رنگ‌ها و تنظیمات رنگی پروژه
تغییر رنگ‌ها فقط از این فایل انجام می‌شود
"""

# ─── رنگ‌های اصلی اکسل (هدر، عنوان، حاشیه) ───
EXCEL_COLORS = {
    'header_bg':    '2F5496',
    'header_font':  'FFFFFF',
    'title_bg':     '1F3864',
    'light_blue':   'D6E4F0',
    'light_green':  'E2EFDA',
    'light_red':    'FCE4EC',
    'border':       'B4C6E7',
    'good':         '92D050',
    'warning':      'FFEB3B',
    'bad':          'FF5252',
    'alt_row':      'F2F7FB',
}

# ─── رنگ ثابت نمودارهای ستونی (تک‌رنگ - برای نمایندگی‌ها) ───
SINGLE_BAR_COLOR = '4472C4'

# ─── رنگ‌های نمودارهای ستونی با دسته‌بندی مفهومی ───
CATEGORY_PALETTE = [
    '4472C4', 'ED7D31', 'A5A5A5', 'FFC000',
    '5B9BD5', '70AD47', '264478', '9B57A0',
    '636363', 'FF6B6B', '48C9B0', 'F39C12',
]

# ─── رنگ‌های نمودار دایره‌ای: وضعیت تحویل ───
PIE_STATUS_COLORS = ['92D050', 'FF5252', 'FFEB3B']

# ─── رنگ‌های نمودار دایره‌ای: بازه‌های زمانی ───
PIE_TIMERANGE_COLORS = [
    '27AE60', '2ECC71', 'F1C40F', 'E67E22',
    'E74C3C', 'C0392B', '8E44AD',
]

# ─── رنگ‌های نمودار دایره‌ای: سطح عملکرد ───
PIE_PERFORMANCE_COLORS = [
    '27AE60', '2ECC71', 'F1C40F', 'E74C3C', '8B0000',
]

# ─── رنگ‌های خاص نمودارها ───
CHART_SPECIFIC = {
    'ontime_bar':       '70AD47',
    'std_bar':          'E74C3C',
    'median_line':      'ED7D31',
    'entry_bar':        '4472C4',
    'delivery_bar':     'ED7D31',
    'weekday_count':    '3498DB',
    'weekday_avg':      'E67E22',
    'ranking_bar':      '2ECC71',
    'outlier_bar':      'E74C3C',
    'daily_bar':        '3498DB',
    'combo_line':       'E74C3C',
    'scatter_marker':   '2980B9',
}

# ─── رنگ‌های Conditional Formatting ───
CF_COLORS = {
    'good_start':   '92D050',
    'mid':          'FFEB3B',
    'bad_end':      'FF5252',
    'databar_blue': '5B9BD5',
    'databar_red':  'E74C3C',
}

# ─── رنگ‌های Scatter/Bubble Chart ───
SCATTER_COLORS = [
    '2980B9', 'E74C3C', '27AE60', 'F39C12', '8E44AD',
    '1ABC9C', 'D35400', '2C3E50', 'C0392B', '16A085',
    '2ECC71', 'E67E22', '9B59B6', '34495E', 'F1C40F',
    '3498DB', 'E91E63', '00BCD4', 'FF5722', '795548',
    '607D8B', 'CDDC39', 'FF9800', '673AB7', '009688',
    '4CAF50', 'FFC107', 'FF4081', '00E5FF', 'B388FF',
    'EA80FC', '80CBC4', 'FFAB91', 'A1887F', 'CE93D8',
    '81D4FA', 'C5E1A5', 'FFF176', 'FFCC80', 'EF9A9A',
]
