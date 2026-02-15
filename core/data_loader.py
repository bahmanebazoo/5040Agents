"""
خواندن، پاکسازی و پردازش داده‌ها
Single Responsibility: فقط مسئول آماده‌سازی داده
"""
import pandas as pd
import numpy as np
import jdatetime
from datetime import datetime
import warnings

from config.constants import (
    MONTH_NAMES, WEEKDAY_NAMES, TIME_RANGE_ORDER, ONTIME_THRESHOLD,
)

warnings.filterwarnings('ignore')


def _jalali_to_gregorian_datetime(jalali_str):
    """تبدیل تاریخ شمسی به میلادی"""
    try:
        if pd.isna(jalali_str):
            return None
        jalali_str = str(jalali_str).strip()
        parts = jalali_str.split(' ')
        date_part = parts[0]
        time_part = parts[1] if len(parts) > 1 else '00:00:00'
        dc = date_part.split('/')
        jy, jm, jd = int(dc[0]), int(dc[1]), int(dc[2])
        tc = time_part.split(':')
        hour, minute = int(tc[0]), int(tc[1])
        second = int(tc[2]) if len(tc) > 2 else 0
        g_date = jdatetime.date(jy, jm, jd).togregorian()
        return datetime(g_date.year, g_date.month, g_date.day, hour, minute, second)
    except Exception:
        return None


def _extract_jalali_parts(jalali_str):
    """استخراج سال، ماه، روز شمسی"""
    try:
        if pd.isna(jalali_str):
            return None, None, None
        parts = str(jalali_str).strip().split(' ')[0].split('/')
        return int(parts[0]), int(parts[1]), int(parts[2])
    except Exception:
        return None, None, None


def _get_jalali_weekday(jalali_str):
    """روز هفته شمسی (0=شنبه ... 6=جمعه)"""
    try:
        if pd.isna(jalali_str):
            return None
        parts = str(jalali_str).strip().split(' ')[0].split('/')
        jy, jm, jd = int(parts[0]), int(parts[1]), int(parts[2])
        return jdatetime.date(jy, jm, jd).weekday()
    except Exception:
        return None


def _time_category(hours):
    """تعیین بازه زمانی تحویل"""
    if pd.isna(hours):
        return 'نامشخص'
    if hours <= 6:
        return '۰-۶ ساعت'
    elif hours <= 12:
        return '۶-۱۲ ساعت'
    elif hours <= 18:
        return '۱۲-۱۸ ساعت'
    elif hours <= 24:
        return '۱۸-۲۴ ساعت'
    elif hours <= 48:
        return '۲۴-۴۸ ساعت'
    elif hours <= 72:
        return '۴۸-۷۲ ساعت'
    else:
        return 'بیش از ۷۲ ساعت'


def _shift_category(hour):
    """تعیین شیفت ورود"""
    if pd.isna(hour):
        return 'نامشخص'
    if 6 <= hour < 14:
        return 'صبح (۶-۱۴)'
    elif 14 <= hour < 22:
        return 'عصر (۱۴-۲۲)'
    else:
        return 'شب (۲۲-۶)'


def load_and_process(filepath: str) -> pd.DataFrame:
    """
    خواندن فایل اکسل و اعمال تمام تبدیلات

    Returns:
        pd.DataFrame: دیتافریم پردازش‌شده
    """
    print("📂 در حال خواندن فایل...")
    df = pd.read_excel(filepath)
    print(f"✅ {len(df)} ردیف خوانده شد | ستون‌ها: {list(df.columns)}")

    print("⏳ در حال پردازش تاریخ‌ها...")

    # تبدیل تاریخ
    df['ورود_میلادی'] = df['تاریخ ورود به نمایندگی'].apply(_jalali_to_gregorian_datetime)
    df['تحویل_میلادی'] = df['تاریخ تحویل'].apply(_jalali_to_gregorian_datetime)

    # اختلاف ساعت
    df['اختلاف_ساعت'] = (
            (df['تحویل_میلادی'] - df['ورود_میلادی']).dt.total_seconds() / 3600
    )

    # حذف مقادیر منفی
    df = df[df['اختلاف_ساعت'] >= 0].copy()

    # استخراج اجزای تاریخ
    df[['سال_شمسی', 'ماه_شمسی', 'روز_شمسی']] = df['تاریخ ورود به نمایندگی'].apply(
        lambda x: pd.Series(_extract_jalali_parts(x))
    )
    df['نام_ماه'] = df['ماه_شمسی'].map(MONTH_NAMES)
    df['روز_هفته'] = df['تاریخ ورود به نمایندگی'].apply(_get_jalali_weekday)
    df['نام_روز_هفته'] = df['روز_هفته'].map(WEEKDAY_NAMES)

    # ساعت ورود و تحویل
    df['ساعت_ورود'] = df['ورود_میلادی'].dt.hour
    df['ساعت_تحویل'] = df['تحویل_میلادی'].dt.hour

    # بازه و وضعیت و شیفت
    df['بازه_تحویل'] = df['اختلاف_ساعت'].apply(_time_category)
    df['وضعیت_تحویل'] = df['اختلاف_ساعت'].apply(
        lambda x: 'به‌موقع' if x <= ONTIME_THRESHOLD else 'تأخیردار'
    )
    df['شیفت_ورود'] = df['ساعت_ورود'].apply(_shift_category)

    print(f"✅ پردازش تمام شد! داده‌های معتبر: {len(df)}")
    return df
