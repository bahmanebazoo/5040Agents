"""
محاسبه تمام KPIها
Single Responsibility: فقط مسئول محاسبات آماری
"""
import pandas as pd
import numpy as np
from dataclasses import dataclass, field
from typing import Dict, List, Any

from config.constants import (
    MONTH_NAMES, WEEKDAY_NAMES, TIME_RANGE_ORDER, ONTIME_THRESHOLD,
)


@dataclass
class KPIResult:
    """نتیجه محاسبات KPI - تمام داده‌های مورد نیاز شیت‌ها"""
    df: pd.DataFrame

    # نمایندگی
    kpi_agent: pd.DataFrame = None

    # ماهانه
    kpi_month: pd.DataFrame = None
    available_months: List[int] = field(default_factory=list)
    month_cols_ordered: List[str] = field(default_factory=list)

    # ماتریس‌ها
    pivot_avg: pd.DataFrame = None
    pivot_count: pd.DataFrame = None
    pivot_ontime: pd.DataFrame = None
    pivot_median: pd.DataFrame = None

    # ساعتی
    hourly_entry: pd.Series = None
    hourly_delivery: pd.Series = None

    # روز هفته
    weekday_stats: pd.DataFrame = None

    # بازه تحویل
    time_range_dist: pd.Series = None
    pivot_timerange: pd.DataFrame = None

    # شیفت
    shift_stats: pd.DataFrame = None

    # پرت‌ها
    Q1_global: float = 0
    Q3_global: float = 0
    IQR_global: float = 0
    outlier_threshold: float = 0
    df_outliers: pd.DataFrame = None

    # روند روزانه
    daily_trend: pd.DataFrame = None


def calculate_all(df: pd.DataFrame) -> KPIResult:
    """محاسبه تمام KPIها و برگرداندن به‌صورت یک شیء واحد"""
    print("📊 در حال محاسبه KPIها...")

    result = KPIResult(df=df)

    # ─── ماه‌های موجود ───
    result.available_months = sorted(
        df['ماه_شمسی'].dropna().unique().astype(int)
    )
    result.month_cols_ordered = [
        MONTH_NAMES[m] for m in result.available_months
    ]

    # ─── KPI نمایندگی ───
    kpi_agent = df.groupby('نمایندگی').agg(
        تعداد_مرسوله=('سریال فاکتور', 'count'),
        میانگین_ساعت=('اختلاف_ساعت', 'mean'),
        میانه_ساعت=('اختلاف_ساعت', 'median'),
        انحراف_معیار=('اختلاف_ساعت', 'std'),
        حداقل_ساعت=('اختلاف_ساعت', 'min'),
        حداکثر_ساعت=('اختلاف_ساعت', 'max'),
        Q1=('اختلاف_ساعت', lambda x: x.quantile(0.25)),
        Q3=('اختلاف_ساعت', lambda x: x.quantile(0.75)),
    ).round(2)

    ontime_by_agent = df.groupby('نمایندگی').apply(
        lambda x: round(
            (x['اختلاف_ساعت'] <= ONTIME_THRESHOLD).sum() / len(x) * 100, 1
        )
    ).rename('درصد_بموقع')

    kpi_agent = kpi_agent.join(ontime_by_agent)
    kpi_agent['IQR'] = (kpi_agent['Q3'] - kpi_agent['Q1']).round(2)
    kpi_agent = kpi_agent.sort_values('میانگین_ساعت')
    result.kpi_agent = kpi_agent

    # ─── KPI ماهانه ───
    kpi_month = df.groupby('ماه_شمسی').agg(
        تعداد_مرسوله=('سریال فاکتور', 'count'),
        میانگین_ساعت=('اختلاف_ساعت', 'mean'),
        میانه_ساعت=('اختلاف_ساعت', 'median'),
        انحراف_معیار=('اختلاف_ساعت', 'std'),
        حداقل_ساعت=('اختلاف_ساعت', 'min'),
        حداکثر_ساعت=('اختلاف_ساعت', 'max'),
    ).round(2).sort_index()

    ontime_by_month = df.groupby('ماه_شمسی').apply(
        lambda x: round(
            (x['اختلاف_ساعت'] <= ONTIME_THRESHOLD).sum() / len(x) * 100, 1
        )
    ).rename('درصد_بموقع')

    kpi_month = kpi_month.join(ontime_by_month)
    kpi_month['رشد_میانگین'] = kpi_month['میانگین_ساعت'].pct_change().mul(100).round(1)
    kpi_month['تغییر_بموقع'] = kpi_month['درصد_بموقع'].diff().round(1)
    result.kpi_month = kpi_month

    # ─── ماتریس‌ها ───
    mco = result.month_cols_ordered

    pivot_avg = df.pivot_table(
        values='اختلاف_ساعت', index='نمایندگی',
        columns='نام_ماه', aggfunc='mean'
    ).round(2)
    result.pivot_avg = pivot_avg[[c for c in mco if c in pivot_avg.columns]]

    pivot_count = df.pivot_table(
        values='سریال فاکتور', index='نمایندگی',
        columns='نام_ماه', aggfunc='count'
    ).fillna(0).astype(int)
    result.pivot_count = pivot_count[[c for c in mco if c in pivot_count.columns]]

    pivot_ontime = df.pivot_table(
        values='اختلاف_ساعت', index='نمایندگی', columns='نام_ماه',
        aggfunc=lambda x: round(
            (x <= ONTIME_THRESHOLD).sum() / len(x) * 100, 1
        )
    )
    result.pivot_ontime = pivot_ontime[[c for c in mco if c in pivot_ontime.columns]]

    pivot_median = df.pivot_table(
        values='اختلاف_ساعت', index='نمایندگی',
        columns='نام_ماه', aggfunc='median'
    ).round(2)
    result.pivot_median = pivot_median[[c for c in mco if c in pivot_median.columns]]

    # ─── ساعتی ───
    result.hourly_entry = df.groupby('ساعت_ورود').size().reindex(range(24), fill_value=0)
    result.hourly_delivery = df.groupby('ساعت_تحویل').size().reindex(range(24), fill_value=0)

    # ─── روز هفته ───
    weekday_stats = df.groupby('روز_هفته').agg(
        تعداد=('سریال فاکتور', 'count'),
        میانگین_ساعت=('اختلاف_ساعت', 'mean'),
    ).round(2).sort_index()
    weekday_stats['نام_روز'] = weekday_stats.index.map(WEEKDAY_NAMES)
    weekday_ontime = df.groupby('روز_هفته').apply(
        lambda x: round(
            (x['اختلاف_ساعت'] <= ONTIME_THRESHOLD).sum() / len(x) * 100, 1
        )
    ).rename('درصد_بموقع')
    result.weekday_stats = weekday_stats.join(weekday_ontime)

    # ─── بازه تحویل ───
    time_range_dist = df['بازه_تحویل'].value_counts()
    result.time_range_dist = time_range_dist.reindex(
        [t for t in TIME_RANGE_ORDER if t in time_range_dist.index]
    )
    pivot_tr = pd.crosstab(df['نمایندگی'], df['بازه_تحویل'])
    result.pivot_timerange = pivot_tr[
        [c for c in TIME_RANGE_ORDER if c in pivot_tr.columns]
    ]

    # ─── شیفت ───
    shift_stats = df.groupby('شیفت_ورود').agg(
        تعداد=('سریال فاکتور', 'count'),
        میانگین_ساعت=('اختلاف_ساعت', 'mean'),
    ).round(2)
    shift_ontime = df.groupby('شیفت_ورود').apply(
        lambda x: round(
            (x['اختلاف_ساعت'] <= ONTIME_THRESHOLD).sum() / len(x) * 100, 1
        )
    ).rename('درصد_بموقع')
    result.shift_stats = shift_stats.join(shift_ontime)

    # ─── پرت‌ها ───
    result.Q1_global = df['اختلاف_ساعت'].quantile(0.25)
    result.Q3_global = df['اختلاف_ساعت'].quantile(0.75)
    result.IQR_global = result.Q3_global - result.Q1_global
    result.outlier_threshold = result.Q3_global + 1.5 * result.IQR_global
    result.df_outliers = df[df['اختلاف_ساعت'] > result.outlier_threshold].copy()

    # ─── روند روزانه ───
    daily = df.groupby(['سال_شمسی', 'ماه_شمسی', 'روز_شمسی']).agg(
        تعداد=('سریال فاکتور', 'count'),
        میانگین_ساعت=('اختلاف_ساعت', 'mean'),
    ).round(2).reset_index()
    daily['تاریخ'] = daily.apply(
        lambda r: f"{int(r['سال_شمسی'])}/{int(r['ماه_شمسی']):02d}/{int(r['روز_شمسی']):02d}",
        axis=1,
    )
    result.daily_trend = daily

    print("✅ KPIها محاسبه شد!")
    return result
