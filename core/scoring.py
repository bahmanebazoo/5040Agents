"""
امتیازدهی و رتبه‌بندی نمایندگی‌ها
"""
import pandas as pd
from config.constants import SCORING_WEIGHTS


def performance_level(score: float) -> str:
    """تعیین سطح عملکرد بر اساس امتیاز"""
    if score >= 80:
        return '⭐ عالی'
    elif score >= 60:
        return '✅ خوب'
    elif score >= 40:
        return '⚠️ متوسط'
    elif score >= 20:
        return '❌ ضعیف'
    else:
        return '🔴 بحرانی'


def calculate_scores(kpi_agent: pd.DataFrame) -> pd.DataFrame:
    """
    محاسبه امتیاز ترکیبی و رتبه‌بندی

    Returns:
        pd.DataFrame: با ستون‌های score_avg, score_ontime, score_std,
                      امتیاز_نهایی, سطح_عملکرد
    """
    scoring = kpi_agent.copy()

    # نرمال‌سازی (0 تا 100)
    _range_avg = scoring['میانگین_ساعت'].max() - scoring['میانگین_ساعت'].min()
    if _range_avg > 0:
        scoring['score_avg'] = 100 - (
                (scoring['میانگین_ساعت'] - scoring['میانگین_ساعت'].min()) / _range_avg * 100
        )
    else:
        scoring['score_avg'] = 100

    scoring['score_ontime'] = scoring['درصد_بموقع']

    _range_std = scoring['انحراف_معیار'].max() - scoring['انحراف_معیار'].min()
    if _range_std > 0:
        scoring['score_std'] = 100 - (
                (scoring['انحراف_معیار'] - scoring['انحراف_معیار'].min()) / _range_std * 100
        )
    else:
        scoring['score_std'] = 100

    # امتیاز نهایی وزن‌دار
    w = SCORING_WEIGHTS
    scoring['امتیاز_نهایی'] = (
            scoring['score_avg'] * w['avg'] +
            scoring['score_ontime'] * w['ontime'] +
            scoring['score_std'] * w['std']
    ).round(1)

    scoring = scoring.sort_values('امتیاز_نهایی', ascending=False)
    scoring['سطح_عملکرد'] = scoring['امتیاز_نهایی'].apply(performance_level)

    return scoring
