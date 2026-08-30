from datetime import date, datetime
from zoneinfo import ZoneInfo


def local_today() -> date:
    """業務ブックの基準タイムゾーン（日本時間）で当日を返す。"""
    return datetime.now(ZoneInfo("Asia/Tokyo")).date()
